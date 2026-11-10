#!/usr/bin/env python3
"""Confere a planilha do Plante e, se pedido, gera o Painel_Plante.html.

Uso:
  python plante.py conferir <planilha.xlsx>
  python plante.py painel   <planilha.xlsx> <modelo.html> <saida.html>

Só usa a biblioteca padrão do Python (sem pacotes extras, sem rede).
"""
import sys, os, re, json, zipfile, posixpath, datetime
import xml.etree.ElementTree as ET

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'pr': 'http://schemas.openxmlformats.org/package/2006/relationships'}
RID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'

ABAS = {  # aba -> (tabela esperada, colunas obrigatórias)
    'Pilares': ('tbPilares', ['id_Pilar', 'Texto_Pilar']),
    'Blocos': ('tbBlocos', ['id_Pilar', 'id_Bloco', 'Texto_Bloco']),
    'Linhas de acao': ('tbLinhas', ['id_Pilar', 'id_Bloco', 'id_Linha', 'Texto_Linha', 'Mensagens PNE', 'Indicador']),
    'Iniciativas': ('tbIniciativas', ['Id Pilar', 'Id_Bloco', 'Id_Linha', 'Id_Ação', 'Texto_Ação', 'Iniciativa', 'Data_Inicio',
                                      'Data_Fim', 'Responsável', 'Status', 'Descrição', 'Setor', 'Semáforo']),
    'Glossario': ('tbGlossario', ['Sigla', 'Nome']),
}
CHAVE = {'Pilares': 'pilares', 'Blocos': 'blocos', 'Linhas de acao': 'linhas', 'Iniciativas': 'iniciativas', 'Glossario': 'glossario'}
SITUACOES = {'Finalizado', 'Em Andamento', 'Não Iniciado'}


def col_idx(ref):
    letters = re.match(r'[A-Z]+', ref).group(0)
    n = 0
    for ch in letters:
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
        path = posixpath.normpath(posixpath.join('xl', alvo)) if not alvo.startswith('/') else alvo.lstrip('/')
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
                r = int(re.search(r'\d+', c.get('r')).group(0))
                linhas.setdefault(r, {})[col_idx(c.get('r'))] = val
        # tabelas desta aba
        tabelas = []
        relp = posixpath.join(posixpath.dirname(path), '_rels', posixpath.basename(path) + '.rels')
        if relp in nomes:
            for r in ET.fromstring(z.read(relp)).findall('pr:Relationship', NS):
                if r.get('Type', '').endswith('/table'):
                    tp = posixpath.normpath(posixpath.join(posixpath.dirname(path), r.get('Target')))
                    t = ET.fromstring(z.read(tp))
                    tabelas.append({'nome': t.get('name') or t.get('displayName'), 'ref': t.get('ref')})
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
        if all((cel.get(i) in (None, '')) for i in range(ncol)):
            continue
        regs.append({nomes[i]: ('' if cel.get(i) is None else str(cel.get(i))) for i in range(ncol) if nomes[i]})
    return nomes, regs


def txt(v):
    if v is None:
        return ''
    return ' '.join(str(v).replace('​', '').split())


def norm_acao(a):
    a = txt(a)
    return a[:-2] + a[-1] if len(a) >= 2 and a[-2] == '.' else a


def data(v):
    t = txt(v)
    if not t:
        return None
    m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})', t)
    if m:
        return datetime.date(int(m[3]), int(m[2]), int(m[1]))
    if re.match(r'^\d+(\.\d+)?$', t):
        return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(float(t)))
    m = re.match(r'^(\d{4})-(\d{2})-(\d{2})', t)
    if m:
        return datetime.date(int(m[1]), int(m[2]), int(m[3]))
    return 'inválida'


def conferir(caminho):
    abas = ler_xlsx(caminho)
    problemas, avisos, dados = [], [], {}
    for aba, (tab, cols) in ABAS.items():
        if aba not in abas:
            problemas.append(f'Falta a aba **{aba}**.')
            continue
        nomes, regs = registros(abas[aba])
        dados[aba] = regs
        tabs = [t['nome'] for t in abas[aba]['tabelas']]
        if tab not in tabs and not os.environ.get('GITHUB_ACTIONS'):  # nomes de Tabela só importam para o Power Automate
            problemas.append(f'A aba **{aba}** não tem a Tabela **{tab}** (tabelas encontradas: {", ".join(tabs) or "nenhuma"}). Veja o Passo 1 do guia.')
        faltam = [c for c in cols if c not in nomes]
        if faltam:
            problemas.append(f'Na aba **{aba}** faltam as colunas: {", ".join(faltam)}.')
    if 'Pontos Focais' in abas:
        avisos.append('A planilha tem a aba **Pontos Focais** (nomes e e-mails). Num repositório público, ela fica visível para qualquer pessoa: envie uma cópia sem essa aba.')
    ini = dados.get('Iniciativas', [])
    pil = {txt(r.get('id_Pilar')) for r in dados.get('Pilares', [])}
    blo = {txt(r.get('id_Bloco')) for r in dados.get('Blocos', [])}
    lin = {txt(r.get('id_Linha')) for r in dados.get('Linhas de acao', [])}
    acoes = {norm_acao(r.get('Id_Ação')) for r in ini if txt(r.get('Id_Ação'))}
    sem_resp = [r for r in ini if not txt(r.get('Responsável'))]
    sem_sit = [r for r in ini if not txt(r.get('Status'))]
    sit_fora = sorted({txt(r.get('Status')) for r in ini if txt(r.get('Status')) and txt(r.get('Status')) not in SITUACOES})
    formato = sorted({txt(r.get('Id_Ação')) for r in ini if txt(r.get('Id_Ação')) != norm_acao(r.get('Id_Ação'))})
    for r in ini:
        a = norm_acao(r.get('Id_Ação'))
        if not a:
            problemas.append(f'Iniciativa sem Id_Ação: "{txt(r.get("Iniciativa"))[:60]}".')
            continue
        l = re.sub(r'[a-z]+$', '', a).rstrip('.')
        if lin and l not in lin:
            problemas.append(f'A ação {a} aponta para a linha {l}, que não existe na aba Linhas de acao.')
        if blo and txt(r.get('Id_Bloco')) not in blo:
            problemas.append(f'A ação {a} usa o bloco {txt(r.get("Id_Bloco"))}, que não existe na aba Blocos.')
        if pil and txt(r.get('Id Pilar')) not in pil:
            problemas.append(f'A ação {a} usa o pilar {txt(r.get("Id Pilar"))}, que não existe na aba Pilares.')
        di, df = data(r.get('Data_Inicio')), data(r.get('Data_Fim'))
        if 'inválida' in (di, df):
            problemas.append(f'A ação {a} tem uma data que não é uma data válida.')
        elif di and df and df < di:
            avisos.append(f'A ação {a} termina ({df:%d/%m/%Y}) antes de começar ({di:%d/%m/%Y}).')
    if sit_fora:
        avisos.append('Situações fora da lista (Finalizado, Em Andamento, Não Iniciado), que o painel trata como texto livre: ' + ', '.join(sit_fora) + '.')
    if formato:
        avisos.append(f'{len(formato)} códigos de ação no formato "1.2.1.a" (o painel corrige para "1.2.1a"): ' + ', '.join(formato[:6]) + ('…' if len(formato) > 6 else '') + '.')
    if len(ini) > 256 and not os.environ.get('GITHUB_ACTIONS'):
        avisos.append(f'Há {len(ini)} iniciativas: a ação "Listar Iniciativas" do fluxo precisa da Paginação ativada (limite 5000).')
    out = ['## Conferência da planilha do Plante', '',
           f'**Resultado:** {"pronta para atualizar o painel" if not problemas else "corrigir antes de atualizar"} '
           f'({len(problemas)} problema(s), {len(avisos)} aviso(s)).', '',
           '| Item | Total |', '| --- | --- |',
           f'| Pilares | {len(pil)} |', f'| Blocos | {len(blo)} |', f'| Linhas de ação | {len(lin)} |',
           f'| Ações | {len(acoes)} |', f'| Iniciativas | {len(ini)} |',
           f'| Iniciativas com situação | {len(ini) - len(sem_sit)} |', f'| Iniciativas sem responsável | {len(sem_resp)} |',
           f'| Siglas no glossário | {len(dados.get("Glossario", []))} |', '']
    if problemas:
        out += ['### Problemas (impedem ou distorcem a atualização)', ''] + [f'- {p}' for p in problemas[:40]] + ['']
    if avisos:
        out += ['### Avisos', ''] + [f'- {a}' for a in avisos] + ['']
    print('\n'.join(out))
    return abas, dados, problemas


def painel(caminho, modelo, saida):
    abas = ler_xlsx(caminho)
    bruto = {'atualizado': datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')}
    for aba, chave in CHAVE.items():
        bruto[chave] = registros(abas[aba])[1] if aba in abas else []
    html = open(modelo, encoding='utf-8').read()
    marca = '/*DADOS_PLANTE*/null'
    if marca not in html:
        sys.exit('O modelo não tem o espaço reservado /*DADOS_PLANTE*/null. Use o Painel_Plante_modelo.html original.')
    js = json.dumps(bruto, ensure_ascii=False).replace('</', '<\\/')
    open(saida, 'w', encoding='utf-8').write(html.replace(marca, js))
    print(f'Painel gerado: {saida} ({len(bruto["iniciativas"])} iniciativas, dados de {bruto["atualizado"]} UTC).')


if __name__ == '__main__':
    if len(sys.argv) >= 3 and sys.argv[1] == 'conferir':
        conferir(sys.argv[2])
    elif len(sys.argv) >= 5 and sys.argv[1] == 'painel':
        painel(sys.argv[2], sys.argv[3], sys.argv[4])
    else:
        print(__doc__)
