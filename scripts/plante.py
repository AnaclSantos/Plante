#!/usr/bin/env python3
"""Confere a planilha V2 do Plante e gera o Painel_Plante.html.

Uso:
  python plante.py conferir <planilha.xlsx>
  python plante.py painel   <planilha.xlsx> <modelo.html> <saida.html>
  python plante.py bruto    <planilha.xlsx> <saida.json>      (dados brutos, para testes)

Só usa a biblioteca padrão do Python (sem pacotes extras, sem rede).
"""
import sys, os, re, json, zipfile, posixpath, datetime
import xml.etree.ElementTree as ET

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'pr': 'http://schemas.openxmlformats.org/package/2006/relationships'}
RID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
ABAS = {  # aba -> (chave nos dados, tabela, colunas obrigatórias)
    'Pilares': ('pilares', 'tbPilares', ['id_Pilar', 'Texto_Pilar']),
    'Blocos': ('blocos', 'tbBlocos', ['id_Pilar', 'id_Bloco', 'Texto_Bloco', 'Bloco_Curto']),
    'Linhas': ('linhas', 'tbLinhas', ['id_Pilar', 'id_Bloco', 'id_Linha', 'Texto_Linha']),
    'Acoes': ('acoes', 'tbAcoes', ['id_Linha', 'id_Acao', 'Texto_Acao', 'Setor', 'Pagina_Rota']),
    'Iniciativas': ('iniciativas', 'tbIniciativas', ['Codigo', 'id_Acao', 'Iniciativa', 'Ano_Inicio', 'Ano_Fim', 'Inicio_Ate_2025', 'Tipo', 'Sugestao', 'Responsavel', 'Descricao']),
    'MensagensPNE': ('mensagens', 'tbMensagensPNE', ['id_Linha', 'Pagina_Rota', 'Acoes', 'Mensagem']),
    'InstrumentosFinanceiros': ('instrumentos', 'tbInstrumentos', ['id_Linha', 'Instrumento']),
    'Indicadores': ('indicadores', 'tbIndicadores', ['id_Linha', 'Indicador', 'Tipo', 'Fonte', 'Memoria_Calculo']),
    'Glossario': ('glossario', 'tbGlossario', ['Sigla', 'Nome']),
}


def col_idx(ref):
    n = 0
    for ch in re.match(r'[A-Z]+', ref).group(0):
        n = n * 26 + ord(ch) - 64
    return n - 1


def ler_xlsx(caminho):
    z = zipfile.ZipFile(caminho)
    nomes = set(z.namelist())
    shared = []
    if 'xl/sharedStrings.xml' in nomes:
        for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('m:si', NS):
            shared.append(''.join(t.text or '' for t in si.iter('{%s}t' % NS['m'])))
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = {r.get('Id'): r.get('Target') for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels')).findall('pr:Relationship', NS)}
    abas = {}
    for s in wb.find('m:sheets', NS).findall('m:sheet', NS):
        alvo = rels[s.get(RID)]
        path = alvo.lstrip('/') if alvo.startswith('/') else posixpath.normpath(posixpath.join('xl', alvo))
        root = ET.fromstring(z.read(path))
        linhas = {}
        for row in root.iter('{%s}row' % NS['m']):
            for c in row.findall('m:c', NS):
                t, v = c.get('t'), c.find('m:v', NS)
                if t == 's' and v is not None:
                    val = shared[int(v.text)]
                elif t == 'inlineStr':
                    val = ''.join(x.text or '' for x in c.iter('{%s}t' % NS['m']))
                else:
                    val = v.text if v is not None else None
                linhas.setdefault(int(re.search(r'\d+', c.get('r')).group(0)), {})[col_idx(c.get('r'))] = val
        tabelas = []
        relp = posixpath.join(posixpath.dirname(path), '_rels', posixpath.basename(path) + '.rels')
        if relp in nomes:
            for r in ET.fromstring(z.read(relp)).findall('pr:Relationship', NS):
                if r.get('Type', '').endswith('/table'):
                    tg = r.get('Target'); t = ET.fromstring(z.read(tg.lstrip('/') if tg.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(path), tg))))
                    tabelas.append(t.get('name') or t.get('displayName'))
        abas[s.get('name')] = {'linhas': linhas, 'tabelas': tabelas}
    return abas


def registros(aba):
    linhas = aba['linhas']
    if not linhas:
        return [], []
    ordem = sorted(linhas)
    cab = linhas[ordem[0]]
    ncol = max(cab) + 1 if cab else 0
    nomes = [(cab.get(i) or '').strip() for i in range(ncol)]
    regs = []
    for r in ordem[1:]:
        cel = linhas[r]
        if all(cel.get(i) in (None, '') for i in range(ncol)):
            continue
        regs.append({nomes[i]: ('' if cel.get(i) is None else str(cel.get(i))) for i in range(ncol) if nomes[i]})
    return nomes, regs


def bruto(caminho):
    abas = ler_xlsx(caminho)
    dados = {'atualizado': datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')}
    for aba, (chave, _, _) in ABAS.items():
        dados[chave] = registros(abas[aba])[1] if aba in abas else []
    return abas, dados


def txt(v):
    return ' '.join(str(v or '').replace('​', '').split())


def conferir(caminho):
    abas, d = bruto(caminho)
    problemas, avisos = [], []
    for aba, (chave, tab, cols) in ABAS.items():
        if aba not in abas:
            problemas.append(f'Falta a aba **{aba}**.'); continue
        nomes = registros(abas[aba])[0]
        faltam = [c for c in cols if c not in nomes]
        if faltam:
            problemas.append(f'Na aba **{aba}** faltam as colunas: {", ".join(faltam)}.')
    if 'Pontos Focais' in abas:
        avisos.append('A planilha tem a aba **Pontos Focais** (nomes e e-mails). Num repositório público ela fica visível: envie uma cópia sem essa aba.')
    linhas = {txt(r.get('id_Linha')) for r in d['linhas']}
    acoes = {txt(r.get('id_Acao')): txt(r.get('id_Linha')) for r in d['acoes']}
    for a, l in acoes.items():
        if l not in linhas: problemas.append(f'A ação {a} aponta para a linha {l}, que não existe na aba Linhas.')
    sem_ano, fora, invert, codigos = 0, 0, [], {}
    for r in d['iniciativas']:
        c = txt(r.get('Codigo')); a = txt(r.get('id_Acao'))
        if a not in acoes: problemas.append(f'A iniciativa {c or txt(r.get("Iniciativa"))[:40]} aponta para a ação {a}, que não existe na aba Acoes.')
        if c: codigos[c] = codigos.get(c, 0) + 1
        try:
            i = int(float(txt(r.get('Ano_Inicio')) or 0)); f = int(float(txt(r.get('Ano_Fim')) or 0))
        except ValueError:
            i = f = 0
        if not i and not f: sem_ano += 1
        elif (i and not 2025 <= i <= 2029) or (f and not 2025 <= f <= 2029): fora += 1
        elif i and f and f < i: invert.append(c)
    dup = [c for c, n in codigos.items() if n > 1]
    if dup: problemas.append('Códigos de iniciativa repetidos: ' + ', '.join(dup[:8]) + ('…' if len(dup) > 8 else '') + '.')
    if sem_ano: avisos.append(f'{sem_ano} iniciativa(s) sem Ano_Inicio e Ano_Fim: ficam fora do Gantt.')
    if fora: avisos.append(f'{fora} iniciativa(s) com ano fora de 2025–2029: o Gantt corta no limite do ciclo.')
    if invert: avisos.append('Iniciativas com Ano_Fim antes de Ano_Inicio: ' + ', '.join(invert[:8]) + '.')
    lin_ind = {txt(r.get('id_Linha')) for r in d['indicadores']}
    falta_ind = sorted(linhas - lin_ind)
    if falta_ind: avisos.append('Linhas sem indicador: ' + ', '.join(falta_ind) + '.')
    sug = sum(1 for r in d['iniciativas'] if txt(r.get('Sugestao')).lower() in ('sim', 's', 'x', '1', 'true'))
    out = ['## Conferência da planilha do Plante (V2)', '',
           f'**Resultado:** {"pronta para gerar o painel" if not problemas else "corrigir antes de gerar"} ({len(problemas)} problema(s), {len(avisos)} aviso(s)).', '',
           '| Item | Total |', '| --- | --- |',
           f'| Linhas de ação | {len(linhas)} |', f'| Ações | {len(acoes)} |', f'| Iniciativas | {len(d["iniciativas"])} |',
           f'| Iniciativas sugeridas | {sug} |', f'| Mensagens PNE | {len(d["mensagens"])} |',
           f'| Instrumentos financeiros | {len(d["instrumentos"])} |', f'| Indicadores | {len(d["indicadores"])} |',
           f'| Siglas no glossário | {len(d["glossario"])} |', '']
    if problemas: out += ['### Problemas', ''] + [f'- {p}' for p in problemas[:40]] + ['']
    if avisos: out += ['### Avisos', ''] + [f'- {a}' for a in avisos] + ['']
    print('\n'.join(out))
    return problemas


def painel(caminho, modelo, saida):
    _, dados = bruto(caminho)
    html = open(modelo, encoding='utf-8').read()
    marca = '/*DADOS_PLANTE*/null'
    if marca not in html:
        sys.exit('O modelo não tem o espaço reservado /*DADOS_PLANTE*/null. Use o Painel_Plante_modelo.html original.')
    open(saida, 'w', encoding='utf-8').write(html.replace(marca, json.dumps(dados, ensure_ascii=False).replace('</', '<\\/')))
    print(f'Painel gerado: {saida} ({len(dados["iniciativas"])} iniciativas, dados de {dados["atualizado"]} UTC).')


if __name__ == '__main__':
    if len(sys.argv) >= 3 and sys.argv[1] == 'conferir':
        sys.exit(1 if conferir(sys.argv[2]) else 0)
    elif len(sys.argv) >= 5 and sys.argv[1] == 'painel':
        painel(sys.argv[2], sys.argv[3], sys.argv[4])
    elif len(sys.argv) >= 4 and sys.argv[1] == 'bruto':
        json.dump(bruto(sys.argv[2])[1], open(sys.argv[3], 'w', encoding='utf-8'), ensure_ascii=False)
    else:
        print(__doc__)
