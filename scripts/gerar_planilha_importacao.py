#!/usr/bin/env python3
"""
Gera planilhas de importação de matrículas no mesmo padrão da
"PLANILHA UNICA CONSOLIDADA - SCC 2209 MATRICULAS.xlsx" (que importa sem erro).

Regras aplicadas (copiadas do que a planilha consolidada faz):
  * todas as células gravadas como TEXTO (inline string), sem células numéricas;
  * datas como texto dd/mm/aaaa (DATA ABERTURA, DATA do ato, DATA DE NASCIMENTO);
  * AREA TOTAL como texto com vírgula decimal (ex.: 75,90); área digitada como
    "26.04" que o Google transformou em data é reconstruída (26,04) só quando o
    mesmo número aparece na DESCRICAO;
  * célula vazia vira "NÃO CONSTA" (ex.: coluna NOTAS);
  * PROTOCOLO só com número (até 7 dígitos) ou número com ponto de milhar;
    qualquer outro conteúdo vira "NÃO CONSTA" (o original fica no relatório);
    o PROTOCOLO da parte (ATO_PARTES) acompanha o do ato a que ela pertence;
  * CPF/CNPJ/CEP gravados como número pelo Google perdem zero à esquerda:
    o zero é recolocado só quando o resultado é um CPF/CNPJ válido (CEP: 8 dígitos);
  * CNPJ/CGC com menos de 11 dígitos (impossível) vira "NÃO CONSTA";
    CPF DO CONJUGE igual ao CPF da própria pessoa vira "NÃO CONSTA";
  * ESTADO por extenso vira a sigla (Pernambuco -> PE), como na consolidada;
  * linhas exatamente repetidas são removidas;
  * sem as ~344 mil células vazias "invisíveis" e sem linhas vazias no fim das abas;
  * layout, estilos, larguras, filtro e painel congelado copiados da consolidada.

Matrículas que precisam de decisão humana NÃO entram na planilha de importação
e vão para PENDENTES_conferencia_manual_NAO_IMPORTAR.xlsx:
  * a mesma matrícula transcrita mais de uma vez com dados diferentes;
  * imóvel sem descrição (urbano/rural) ou com as duas;
  * numeração de atos repetida (ex.: dois R-3), pois renumerar mudaria o
    número oficial do livro;
  * atos que não começam no número 1 (provável página de continuação);
  * descrição do imóvel que é trecho de outro texto;
  * parte que não se liga a um ato;
  * DATA ABERTURA anterior a 1976 (antes da Lei 6.015/73);
  * ato dizendo "fica cancelada a presente matrícula" com ATIVA = ATIVA;
  * numeração de ato muito acima da quantidade de atos (ex.: R-10 com 3 atos);
  * casos achados na conferência manual (lista REVISAR_MANUAL abaixo).

Também gera uma planilha PILOTO pequena (subconjunto da principal) para testar
a importação antes de mandar tudo, já que o sistema não mostra log de erro.

Uso:
  python3 scripts/gerar_planilha_importacao.py \\
      --modelo "PLANILHA UNICA CONSOLIDADA - SCC 2209 MATRICULAS.xlsx" \\
      --situacao "Situacao_Matriculas_por_Livro (2).xlsx" \\
      --ja-importadas "PLANILHA UNICA CONSOLIDADA - SCC 2209 MATRICULAS.xlsx" "LIVRO 2-AE (2).xlsx" \\
      --fontes "Cópia de PLANILHA SANTA 2-AE a 2-AZ.xlsx" ... \\
      --saida importacao_corrigida
"""
import argparse, collections, copy, datetime, os, re, sys, unicodedata, zipfile
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.utils import get_column_letter

NC = 'NÃO CONSTA'
NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
RID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
SHEETS = ['MATRÍCULA', 'T. URBANO', 'T. RURAL', 'MATRICULA_ATO', 'ATO_PARTES', 'P. FÍSICA', 'P. JURÍDICA']
DATE_COLS = {('MATRÍCULA', 'DATA ABERTURA'), ('MATRICULA_ATO', 'DATA'), ('P. FÍSICA', 'DATA DE NASCIMENTO')}
AREA_COLS = {('T. URBANO', 'AREA TOTAL'), ('T. RURAL', 'AREA TOTAL')}
PAD = {  # (aba, coluna): tipo de documento -> só para células NUMÉRICAS na origem
    ('P. JURÍDICA', 'CNPJ/CGC'): 'cnpj', ('P. JURÍDICA', 'CPF REPRESENTANTE 1'): 'cpf',
    ('P. JURÍDICA', 'CPF REPRESENTANTE 2'): 'cpf', ('P. FÍSICA', 'CPF'): 'cpf',
    ('P. FÍSICA', 'CPF DO CONJUGE'): 'cpf', ('ATO_PARTES', 'DOCUMENTO DA PARTE'): 'doc',
    ('P. FÍSICA', 'CEP'): 'cep', ('P. JURÍDICA', 'CEP'): 'cep',
    ('T. URBANO', 'CEP'): 'cep', ('T. RURAL', 'CEP'): 'cep',
}
UF = {'acre': 'AC', 'alagoas': 'AL', 'amapa': 'AP', 'amazonas': 'AM', 'bahia': 'BA', 'ceara': 'CE',
      'distrito federal': 'DF', 'espirito santo': 'ES', 'goias': 'GO', 'maranhao': 'MA', 'mato grosso': 'MT',
      'mato grosso do sul': 'MS', 'minas gerais': 'MG', 'para': 'PA', 'paraiba': 'PB', 'parana': 'PR',
      'pernambuco': 'PE', 'piaui': 'PI', 'rio de janeiro': 'RJ', 'rio grande do norte': 'RN',
      'rio grande do sul': 'RS', 'rondonia': 'RO', 'roraima': 'RR', 'santa catarina': 'SC',
      'sao paulo': 'SP', 'sergipe': 'SE', 'tocantins': 'TO'}
ESTADO_COLS = {('P. FÍSICA', 'ESTADO'), ('P. JURÍDICA', 'ESTADO')}
# Achados da conferência manual (verificação independente das planilhas geradas).
REVISAR_MANUAL = {
    '2745': 'descrição do imóvel é trecho de outro texto ("Transportado do Livro nº 2 AE, fls. 04 verso. suctivas...")',
    '3118': 'AREA TOTAL igual ao número da matrícula (3118) e TIPO AREA "m"',
    '6343': 'AREA TOTAL 48800 m² não bate com as medidas da descrição (cerca de 488 m²)',
    '6077': 'MATRICULA ANTERIOR contém a qualificação de uma pessoa (texto na coluna errada)',
    '3860': 'AREA TOTAL 70,00 ha, mas a descrição diz 33,9 ha',
}
ILLEGAL = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')
EXCEL_EPOCH = datetime.date(1899, 12, 30)


# --------------------------------------------------------------------------- leitura
def ler_xlsx(path):
    """Lê o XML cru: {aba: [(linha, {coluna: (tipo, valor, formato)})]}."""
    z = zipfile.ZipFile(path)
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = {r.get('Id'): r.get('Target') for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
    ss = []
    if 'xl/sharedStrings.xml' in z.namelist():
        for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('m:si', NS):
            ss.append(''.join(t.text or '' for t in si.iter('{%s}t' % NS['m'])))
    st = ET.fromstring(z.read('xl/styles.xml'))
    nf = st.find('m:numFmts', NS)
    numfmts = {int(n.get('numFmtId')): n.get('formatCode') for n in (nf if nf is not None else [])}
    xfs = [int(x.get('numFmtId', 0)) for x in st.find('m:cellXfs', NS)]
    out = {}
    for sh in wb.find('m:sheets', NS):
        tgt = rels[sh.get(RID)].lstrip('/')
        tgt = tgt if tgt.startswith('xl/') else 'xl/' + tgt
        rows = []
        for _, el in ET.iterparse(z.open(tgt)):
            if el.tag != '{%s}row' % NS['m']:
                continue
            r = {}
            for c in el.findall('m:c', NS):
                col = re.match(r'[A-Z]+', c.get('r')).group(0)
                t = c.get('t', 'n')
                v = c.find('m:v', NS)
                if t == 's':
                    val = ss[int(v.text)] if v is not None else None
                elif t == 'inlineStr':
                    isel = c.find('m:is', NS)
                    val = ''.join(x.text or '' for x in isel.iter('{%s}t' % NS['m'])) if isel is not None else None
                else:
                    val = v.text if v is not None else None
                s = int(c.get('s', 0))
                fid = xfs[s] if s < len(xfs) else 0
                r[col] = (t, val, numfmts.get(fid, 'builtin%d' % fid))
            rows.append((int(el.get('r')), r))
            el.clear()
        out[sh.get('name')] = rows
    return out


def chave_matricula(cell):
    if cell is None or cell[1] is None:
        return None
    s = str(cell[1]).strip()
    if not s:
        return None
    try:
        d = Decimal(s)
        if d == d.to_integral_value():
            return str(int(d))
    except InvalidOperation:
        pass
    s2 = s.replace('.', '')
    return s2 if s2.isdigit() else s


def norm_txt(s):
    s = unicodedata.normalize('NFKD', s or '')
    s = ''.join(ch for ch in s if not unicodedata.combining(ch))
    return re.sub(r'\s+', ' ', s).strip().lower()


def cpf_ok(d):
    if len(d) != 11 or len(set(d)) == 1:
        return False
    for n in (9, 10):
        soma = sum(int(d[i]) * (n + 1 - i) for i in range(n))
        if (soma * 10 % 11) % 10 != int(d[n]):
            return False
    return True


def cnpj_ok(d):
    if len(d) != 14 or len(set(d)) == 1:
        return False
    for n in (12, 13):
        pesos = list(range(n - 7, 1, -1)) + list(range(9, 1, -1))
        soma = sum(int(d[i]) * pesos[i] for i in range(n))
        dv = 11 - soma % 11
        if (0 if dv >= 10 else dv) != int(d[n]):
            return False
    return True


def completar_zeros(out, tipo):
    """Recoloca zeros perdidos só quando o resultado é documento válido."""
    if tipo == 'cep':
        return out.zfill(8) if len(out) == 7 else out
    if tipo == 'cpf' and len(out) in (9, 10) and cpf_ok(out.zfill(11)):
        return out.zfill(11)
    if tipo == 'cnpj' and len(out) in (12, 13) and cnpj_ok(out.zfill(14)):
        return out.zfill(14)
    if tipo == 'doc':
        if len(out) in (9, 10) and cpf_ok(out.zfill(11)):
            return out.zfill(11)
        if len(out) in (12, 13) and cnpj_ok(out.zfill(14)):
            return out.zfill(14)
    return out


# --------------------------------------------------------------------------- conversão
class Log:
    def __init__(self):
        self.ajustes = []      # (matricula, aba, arquivo, linha, coluna, antes, depois, regra)
        self.contagem = collections.Counter()

    def add(self, mat, aba, origem, col, antes, depois, regra, detalhar=True):
        self.contagem[(mat, regra)] += 1
        if detalhar:
            self.ajustes.append((mat, aba, origem[0], origem[1], col, antes, depois, regra))


def serial_para_data(x):
    d = EXCEL_EPOCH + datetime.timedelta(days=int(x))
    return d.strftime('%d/%m/%Y')


def texto_data(s):
    m = re.fullmatch(r'\s*(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})\s*', s)
    if not m:
        return None
    d, mo, y = map(int, m.groups())
    try:
        datetime.date(y, mo, d)
    except ValueError:
        return None
    return '%02d/%02d/%04d' % (d, mo, y)


def converter(aba, col, cell, mat, origem, log, ctx):
    """Converte uma célula de origem para o texto final."""
    t, v, fmt = cell if cell is not None else (None, None, None)
    if v is None or (isinstance(v, str) and v.strip() == ''):
        log.add(mat, aba, origem, col, '' if v is None else v, NC, 'vazio -> NÃO CONSTA', detalhar=False)
        return NC
    numerico = t not in ('s', 'inlineStr', 'str', 'e', 'b')
    if numerico:
        try:
            x = Decimal(str(v))
        except InvalidOperation:
            numerico = False
    if numerico:
        if (aba, col) in DATE_COLS and 1 <= x <= 2958465:
            out = serial_para_data(x)
            log.add(mat, aba, origem, col, str(v), out, 'data numérica -> texto dd/mm/aaaa', detalhar=False)
            return out
        if (aba, col) in AREA_COLS:
            formato_data = 'd' in (fmt or '').lower() and 'm' in (fmt or '').lower()
            if formato_data and x > 20000:
                # o Google leu "26.04" como a data 26/04: reconstrói 26,04 e confere na DESCRICAO
                d = EXCEL_EPOCH + datetime.timedelta(days=int(x))
                out = '%02d,%02d' % (d.day, d.month)
                ctx.setdefault('area_reconstruida', []).append(out)
                log.add(mat, aba, origem, col, str(v), out, 'área que o Google converteu em data -> reconstruída')
                return out
            if x == x.quantize(Decimal('0.01')):
                out = str(x.quantize(Decimal('0.01'))).replace('.', ',')
            else:
                out = format(x.normalize(), 'f').replace('.', ',')
            if formato_data:
                log.add(mat, aba, origem, col, str(v), out, 'área com formato de data -> texto com vírgula')
            else:
                log.add(mat, aba, origem, col, str(v), out, 'área numérica -> texto com vírgula', detalhar=False)
            return out
        if x == x.to_integral_value():
            out = str(int(x))
        else:
            out = format(x.normalize(), 'f').replace('.', ',')
            log.add(mat, aba, origem, col, str(v), out, 'número decimal -> texto com vírgula')
        if (aba, col) in PAD and out.isdigit():
            novo = completar_zeros(out, PAD[(aba, col)])
            if novo != out:
                log.add(mat, aba, origem, col, out, novo, 'zero à esquerda restaurado (documento válido)')
                out = novo
        if (aba, col) == ('P. JURÍDICA', 'CNPJ/CGC') and out.isdigit() and len(out) < 11:
            log.add(mat, aba, origem, col, out, NC, 'CNPJ/CGC impossível (menos de 11 dígitos) -> NÃO CONSTA')
            return NC
        return out
    s = str(v)
    s2 = ILLEGAL.sub('', s)
    if s2 != s:
        log.add(mat, aba, origem, col, repr(s), s2, 'caractere de controle removido')
    s = s2.strip()
    if s == '':
        return NC
    if norm_txt(s) in ('nao consta', 'n/c'):
        if s != NC:
            log.add(mat, aba, origem, col, s, NC, 'variação de NÃO CONSTA padronizada')
        return NC
    if (aba, col) in DATE_COLS:
        d = texto_data(s)
        if d and d != s:
            log.add(mat, aba, origem, col, s, d, 'data em texto padronizada dd/mm/aaaa')
            return d
        if not d:
            log.add(mat, aba, origem, col, s, s, 'ATENÇÃO: data em texto fora do padrão (mantida)')
    if (aba, col) in AREA_COLS and re.fullmatch(r'\d+\.\d{1,2}', s):
        novo = s.replace('.', ',')
        log.add(mat, aba, origem, col, s, novo, 'área com ponto decimal -> vírgula')
        return novo
    if (aba, col) in ESTADO_COLS and norm_txt(s) in UF:
        novo = UF[norm_txt(s)]
        if novo != s:
            log.add(mat, aba, origem, col, s, novo, 'estado por extenso -> sigla', detalhar=False)
        return novo
    if (aba, col) == ('P. JURÍDICA', 'CNPJ/CGC') and s.isdigit() and len(s) < 11:
        log.add(mat, aba, origem, col, s, NC, 'CNPJ/CGC impossível (menos de 11 dígitos) -> NÃO CONSTA')
        return NC
    if s.startswith('='):
        log.add(mat, aba, origem, col, s, s, 'texto começando com = gravado como texto')
    return s


def limpar_protocolo(p):
    if p == NC:
        return NC
    if re.fullmatch(r'0+\d{1,7}', p) and len(p.lstrip('0')) <= 7:
        return p.lstrip('0')
    if re.fullmatch(r'\d{1,7}', p) or re.fullmatch(r'\d{1,3}(\.\d{3})+', p):
        return p
    return NC


def norm_id(s):
    m = re.fullmatch(r'\s*(R|AV)\s*[-.]?\s*0*(\d+)\s*', (s or '').upper())
    return '%s-%d' % (m.group(1), int(m.group(2))) if m else (s or '').strip().upper()


# --------------------------------------------------------------------------- montagem
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--modelo', required=True)
    ap.add_argument('--situacao', required=True)
    ap.add_argument('--ja-importadas', nargs='*', default=[])
    ap.add_argument('--fontes', nargs='+', required=True)
    ap.add_argument('--saida', required=True)
    a = ap.parse_args()
    os.makedirs(a.saida, exist_ok=True)

    # ---- quais matrículas faltam no sistema, por categoria do relatório
    ws = openpyxl.load_workbook(a.situacao, read_only=True, data_only=True)['Por Livro']
    hdr = None
    cat = {}
    livro_de = {}
    for row in ws.iter_rows(values_only=True):
        if hdr is None:
            hdr = list(row)
            continue
        if not row[0] or str(row[0]).startswith('TOTAL'):
            continue
        def nums(colname):
            v = row[hdr.index(colname)]
            return [int(x.replace('.', '')) for x in re.findall(r'\d[\d.]*', str(v))] if v else []
        for n in nums('DIGITADAS — quais são'):
            cat[str(n)] = 'DIGITADA'; livro_de[str(n)] = row[0]
        for n in nums('PLANILHA — quais são'):
            cat[str(n)] = 'PLANILHA'; livro_de[str(n)] = row[0]
        for n in nums('FALTANTES — quais são'):
            cat[str(n)] = 'FALTANTE'; livro_de[str(n)] = row[0]
    ja = set()
    for f in a.ja_importadas:
        for _, r in ler_xlsx(f)['MATRÍCULA'][1:]:
            k = chave_matricula(r.get('A'))
            if k:
                ja.add(k)
    faltam = {k: v for k, v in cat.items() if k not in ja}

    # ---- modelo (cabeçalhos e estilos)
    modelo_raw = ler_xlsx(a.modelo)
    headers = {}
    for aba in SHEETS:
        h = modelo_raw[aba][0][1]
        cols = sorted((c for c in h if h[c][1]), key=lambda c: (len(c), c))
        headers[aba] = [h[c][1] for c in cols]

    # ---- datas de abertura de todas as matrículas (para checar a sequência)
    datas_abertura = {}
    def guardar_datas(dados):
        rows = dados['MATRÍCULA']
        h = {cell[1].strip(): c for c, cell in rows[0][1].items() if cell[1]}
        for _, r in rows[1:]:
            k = chave_matricula(r.get('A'))
            cell = r.get(h.get('DATA ABERTURA'))
            if not k or not k.isdigit() or not cell or cell[1] is None:
                continue
            t, v, _f = cell
            try:
                if t not in ('s', 'inlineStr', 'str'):
                    d = EXCEL_EPOCH + datetime.timedelta(days=int(Decimal(v)))
                else:
                    tv = texto_data(str(v))
                    d = datetime.datetime.strptime(tv, '%d/%m/%Y').date() if tv else None
            except Exception:
                d = None
            if d and d.year >= 1976:
                datas_abertura[int(k)] = d
    for f in a.ja_importadas:
        guardar_datas(ler_xlsx(f))

    # ---- lê fontes e agrupa linhas por matrícula
    linhas = collections.defaultdict(lambda: collections.defaultdict(list))  # mat -> aba -> [(origem, {col: cell})]
    for f in a.fontes:
        dados = ler_xlsx(f)
        guardar_datas(dados)
        nome = os.path.basename(f)
        for aba in SHEETS:
            rows = dados[aba]
            h = {cell[1].strip(): c for c, cell in rows[0][1].items() if cell[1]}
            faltando = [n for n in headers[aba] if n not in h]
            if faltando:
                sys.exit(f'{nome} / {aba}: colunas ausentes {faltando}')
            for rn, r in rows[1:]:
                k = chave_matricula(r.get('A'))
                if k is None or k not in faltam:
                    continue
                linhas[k][aba].append(((nome, rn), {n: r.get(h[n]) for n in headers[aba]}))

    com_fonte = {k for k in linhas if linhas[k]['MATRÍCULA']}
    log = Log()
    saida = {cat_: {aba: [] for aba in SHEETS} for cat_ in ('importar', 'digitadas')}
    pendentes = []         # (mat, livro, categoria, motivo, detalhe)
    pend_linhas = {aba: [] for aba in SHEETS}
    atencao = []           # (mat, motivo, detalhe)

    for mat in sorted(linhas, key=lambda x: int(x) if x.isdigit() else 10**9):
        L = linhas[mat]
        conv = {}
        ctx = {}
        for aba in SHEETS:
            conv[aba] = [(o, [converter(aba, n, cells[n], mat, o, log, ctx) for n in headers[aba]]) for o, cells in L[aba]]

        def pendente(motivo, detalhe=''):
            pendentes.append((mat, livro_de.get(mat, ''), faltam[mat], motivo, detalhe))
            for aba in SHEETS:
                for o, vals in conv[aba]:
                    pend_linhas[aba].append([o[0], o[1]] + vals)

        if not conv['MATRÍCULA']:
            continue  # matrícula só aparece em abas filhas: ignorada (não há cabeçalho da matrícula)

        # -- matrícula repetida (na mesma fonte ou em fontes diferentes)
        n_m = len(conv['MATRÍCULA'])
        if n_m > 1:
            iguais = len({tuple(v) for _, v in conv['MATRÍCULA']}) == 1
            filhos_ok = iguais and all(
                all(c % n_m == 0 for c in collections.Counter(tuple(v) for _, v in conv[aba]).values())
                for aba in SHEETS[1:])
            if not filhos_ok:
                origens = '; '.join(f'{o[0]} linha {o[1]}' for o, _ in conv['MATRÍCULA'])
                pendente('matrícula aparece mais de uma vez com dados diferentes', origens)
                continue
            for aba in SHEETS:
                vistos = set(); novo = []
                for o, v in conv[aba]:
                    if tuple(v) in vistos:
                        log.add(mat, aba, o, '(linha)', 'linha repetida', 'removida', 'cópia idêntica da matrícula removida')
                        continue
                    vistos.add(tuple(v)); novo.append((o, v))
                conv[aba] = novo

        # -- linhas exatamente repetidas nas abas filhas
        for aba in SHEETS[1:]:
            vistos = set(); novo = []
            for o, v in conv[aba]:
                if tuple(v) in vistos:
                    log.add(mat, aba, o, '(linha)', 'linha repetida', 'removida', 'linha idêntica removida')
                    continue
                vistos.add(tuple(v)); novo.append((o, v))
            conv[aba] = novo

        # -- exatamente uma descrição de imóvel (urbano OU rural)
        nu, nr = len(conv['T. URBANO']), len(conv['T. RURAL'])
        if nu + nr != 1:
            pendente('imóvel sem descrição única', f'{nu} linha(s) em T. URBANO e {nr} em T. RURAL')
            continue

        # -- conferência manual explícita
        if mat in REVISAR_MANUAL:
            pendente('achado na conferência manual', REVISAR_MANUAL[mat])
            continue
        HM = headers['MATRÍCULA']
        mrow = conv['MATRÍCULA'][0][1]
        prop = (conv['T. URBANO'] or conv['T. RURAL'])[0][1]
        prop_h = headers['T. URBANO'] if conv['T. URBANO'] else headers['T. RURAL']
        descricao = prop[prop_h.index('DESCRICAO')]
        sem_confirmacao = [a for a in ctx.get('area_reconstruida', [])
                           if a not in descricao and a.replace(',', '.') not in descricao]
        if sem_confirmacao:
            pendente('área convertida em data pelo Google sem confirmação na descrição', ', '.join(sem_confirmacao))
            continue
        palavra = norm_txt(descricao).split(' ')[0] if descricao != NC else ''
        inicios_ok = ('um', 'uma', 'uns', 'umas', 'o', 'a', 'os', 'as', 'dois', 'duas', 'tres', 'parte',
                      'terreno', 'lote', 'casa', 'imovel', 'gleba', 'area', 'sitio', 'fazenda',
                      'predio', 'apartamento', 'loja', 'sala', 'galpao', 'unidade')
        if descricao.startswith('[') or (descricao[:1].islower() and not palavra.startswith(inicios_ok)):
            pendente('descrição do imóvel parece trecho de outro texto', descricao[:120])
            continue
        dab = mrow[HM.index('DATA ABERTURA')]
        if texto_data(dab) == dab and int(dab[-4:]) < 1976:
            pendente('DATA ABERTURA anterior a 1976 (antes das matrículas da Lei 6.015/73)', dab)
            continue

        H = headers['MATRICULA_ATO']; iP, iT, iO, iX, iD, iA = (H.index(x) for x in ('PROTOCOLO', 'TIPO ATO', 'ORDEM', 'TEXTO', 'DATA', 'ATO'))
        atos = conv['MATRICULA_ATO']
        ordens = [v[iO] for _, v in atos]
        if any(not o.isdigit() for o in ordens):
            pendente('ORDEM de ato não numérica', ', '.join(o for o in ordens if not o.isdigit()))
            continue
        nums = [int(o) for o in ordens]
        rotulos = ' '.join('%s-%s(%s)' % (v[iT], v[iO], v[iD]) for _, v in atos)
        if len(set(nums)) != len(nums):
            pendente('numeração de atos repetida (renumerar mudaria o número oficial do livro)', rotulos)
            continue
        if nums and min(nums) > 1:
            pendente('atos não começam no número 1 (provável página de continuação)', rotulos)
            continue
        faltando = sorted(set(range(1, max(nums) + 1)) - set(nums)) if nums else []
        seguidos = any(b == a + 1 for a, b in zip(faltando, faltando[1:]))
        if seguidos:
            pendente('faltam dois ou mais atos seguidos no meio da matrícula', rotulos)
            continue
        if nums and max(nums) >= 10 and len(nums) * 2 <= max(nums):
            pendente('numeração de ato muito acima da quantidade de atos', rotulos)
            continue
        if mrow[HM.index('ATIVA')] == 'ATIVA' and any('cancelada a presente matricula' in norm_txt(v[iX]) for _, v in atos):
            pendente('ato cancela a matrícula, mas ATIVA = ATIVA', rotulos)
            continue
        for o, v in atos:
            p = limpar_protocolo(v[iP])
            if p != v[iP]:
                log.add(mat, 'MATRICULA_ATO', o, 'PROTOCOLO', v[iP], p, 'protocolo não numérico -> NÃO CONSTA')
                v[iP] = p
        ids = ['%s-%d' % (v[iT].strip().upper(), int(v[iO])) for _, v in atos]

        HP = headers['ATO_PARTES']; jI, jN, jP = (HP.index(x) for x in ('ID', 'NOME DA PARTE', 'PROTOCOLO'))
        problema = None
        mapa_partes = []
        for o, v in conv['ATO_PARTES']:
            pid = norm_id(v[jI])
            if pid not in ids:
                problema = f'parte "{v[jN]}" com ID {v[jI]}: nenhum ato com esse número'
                break
            mapa_partes.append(ids.index(pid))
        if problema:
            pendente('parte não se liga a um ato', problema)
            continue
        for (o, v), i in zip(conv['ATO_PARTES'], mapa_partes):
            if v[jI] != ids[i]:
                log.add(mat, 'ATO_PARTES', o, 'ID', v[jI], ids[i], 'ID da parte padronizado')
                v[jI] = ids[i]
            ap_ = atos[i][1][iP]
            if v[jP] != ap_:
                log.add(mat, 'ATO_PARTES', o, 'PROTOCOLO', v[jP], ap_, 'protocolo da parte = protocolo do ato')
                v[jP] = ap_
        vistos = set(); novo = []
        for o, v in conv['ATO_PARTES']:
            if tuple(v) in vistos:
                log.add(mat, 'ATO_PARTES', o, '(linha)', 'linha repetida', 'removida', 'linha idêntica removida')
                continue
            vistos.add(tuple(v)); novo.append((o, v))
        conv['ATO_PARTES'] = novo
        HF = headers['P. FÍSICA']; kC, kJ = HF.index('CPF'), HF.index('CPF DO CONJUGE')
        for o, v in conv['P. FÍSICA']:
            if v[kJ] != NC and v[kJ] == v[kC]:
                log.add(mat, 'P. FÍSICA', o, 'CPF DO CONJUGE', v[kJ], NC, 'CPF do cônjuge igual ao da pessoa -> NÃO CONSTA')
                v[kJ] = NC

        # -- avisos (entram na planilha, mas merecem conferência)
        if dab == NC:
            atencao.append((mat, 'DATA ABERTURA = NÃO CONSTA', 'a consolidada aceita; confira se a matrícula entrou'))
        if not atos:
            atencao.append((mat, 'matrícula sem atos', 'a consolidada tem casos assim'))
        elif sorted(nums) != list(range(1, len(nums) + 1)):
            atencao.append((mat, 'numeração de atos com lacunas', ', '.join(ids)))
        if nums and max(nums) >= 10:
            atencao.append((mat, 'ato numerado 10 ou mais', 'a consolidada não tinha nenhum; confira após importar'))
        datas = [datetime.datetime.strptime(v[iD], '%d/%m/%Y').date() for _, v in atos if texto_data(v[iD]) == v[iD]]
        if texto_data(dab) == dab:
            d0 = datetime.datetime.strptime(dab, '%d/%m/%Y').date()
            antes = [d for d in datas if (d0 - d).days > 30]
            if antes:
                atencao.append((mat, 'ato com data anterior à abertura da matrícula', f'abertura {dab}; {len(antes)} ato(s) antes'))
        ordenados = sorted(zip(nums, [texto_data(v[iD]) and datetime.datetime.strptime(v[iD], '%d/%m/%Y').date() for _, v in atos]))
        regride = [(a, b) for (na, a), (nb, b) in zip(ordenados, ordenados[1:]) if a and b and (a - b).days > 365]
        if regride:
            atencao.append((mat, 'ato de número maior com data mais de 1 ano anterior', '; '.join(f'{x} > {y}' for x, y in regride)))
        for _, v in atos:
            tx = norm_txt(v[iX])
            if 'sem efeito' in tx or tx in ('cancelado', 'cancelada'):
                atencao.append((mat, 'ato declarado sem efeito/cancelado no próprio texto', '%s-%s' % (v[iT], v[iO])))
            if re.fullmatch(r'(r|av)\s*[-.]?\s*\d+\s*[-.]?\s*(mat\.?\s*)?[\d.]*', tx):
                atencao.append((mat, 'ato sem texto (só o número)', '%s-%s' % (v[iT], v[iO])))
        if atos and mrow[HM.index('ATIVA')] == 'ATIVA':
            ult = norm_txt(atos[max(range(len(atos)), key=lambda i: nums[i])][1][iX])
            if re.search(r'foi remembrad|ficou remembrad|tornando-se um (so|unico)|originou a (nova )?matricula', ult):
                atencao.append((mat, 'último ato indica remembramento/nova matrícula, mas ATIVA', 'confira o status'))
        if texto_data(dab) == dab:
            viz = [d for n_, d in datas_abertura.items() if n_ != int(mat) and abs(n_ - int(mat)) <= 150]
            if len(viz) >= 10:
                viz.sort(); med = viz[len(viz) // 2]
                d0 = datetime.datetime.strptime(dab, '%d/%m/%Y').date()
                if abs((d0 - med).days) > 730:
                    atencao.append((mat, 'DATA ABERTURA fora da sequência das matrículas vizinhas',
                                    f'{dab}; vizinhas por volta de {med.strftime("%m/%Y")}'))
        cpfs = collections.Counter(v[kC] for _, v in conv['P. FÍSICA'] if v[kC] != NC)
        rep_cpf = [c for c, n in cpfs.items() if n > 1]
        if rep_cpf:
            atencao.append((mat, 'mesmo CPF em mais de uma linha de pessoa física', ', '.join(rep_cpf)))

        destino = 'digitadas' if faltam[mat] == 'DIGITADA' else 'importar'
        for aba in SHEETS:
            saida[destino][aba].extend(v for _, v in conv[aba])

    # ---- grava planilhas de importação a partir do modelo
    def gravar(dados, caminho):
        wb = openpyxl.load_workbook(a.modelo)
        for aba in SHEETS:
            ws = wb[aba]
            ncols = len(headers[aba])
            estilos = [copy.copy(ws.cell(row=2, column=j + 1)._style) for j in range(ncols)]
            altura = ws.row_dimensions[2].height
            ws.delete_rows(2, ws.max_row)
            for k in [k for k in ws.row_dimensions if k > 1]:
                del ws.row_dimensions[k]
            for i, vals in enumerate(dados[aba]):
                r = i + 2
                for j, val in enumerate(vals):
                    c = ws.cell(row=r, column=j + 1, value=val)
                    c.data_type = 's'
                    c._style = copy.copy(estilos[j])
                if altura:
                    ws.row_dimensions[r].height = altura
            ultima = max(ws.max_column, ncols)
            ws.auto_filter.ref = f'A1:{get_column_letter(ultima)}{len(dados[aba]) + 1}'
        wb.save(caminho)

    arquivos = []
    if any(saida['importar'][x] for x in SHEETS):
        p = os.path.join(a.saida, 'Matriculas_faltantes_PADRAO_CONSOLIDADA.xlsx'); gravar(saida['importar'], p); arquivos.append(p)
    if any(saida['digitadas'][x] for x in SHEETS):
        p = os.path.join(a.saida, 'Matriculas_DIGITADAS_encontradas_nas_copias_PADRAO_CONSOLIDADA.xlsx'); gravar(saida['digitadas'], p); arquivos.append(p)

    # ---- piloto: 15 matrículas comuns + 1 de cada tipo de aviso + 1 rural
    imp = saida['importar']
    if imp['MATRÍCULA']:
        mats = [r[0] for r in imp['MATRÍCULA']]
        piloto = list(mats[:15])
        incluir = set(mats)
        vistos_motivo = set()
        for m_, motivo, _ in atencao:
            if motivo not in vistos_motivo and m_ in incluir and m_ not in piloto:
                vistos_motivo.add(motivo); piloto.append(m_)
        extras = []
        if imp['T. RURAL']:
            extras.append(imp['T. RURAL'][0][0])
        hf = headers['P. FÍSICA']; ht = headers['T. URBANO']; hm = headers['MATRÍCULA']
        extras += [r[0] for r in imp['P. FÍSICA'] if r[hf.index('CPF')] == NC][:1]
        conhecidos = {'Terreno/Lote', 'Casa', 'Loja', NC}
        extras += [r[0] for r in imp['T. URBANO'] if r[ht.index('TIPO IMOVEL ONR')] not in conhecidos][:2]
        extras += [r[0] for r in imp['MATRÍCULA'] if r[hm.index('ATIVA')] != 'ATIVA'][:1]
        for aba in ('MATRICULA_ATO', 'ATO_PARTES', 'P. JURÍDICA'):
            if imp[aba]:
                extras.append(max(imp[aba], key=lambda r: max(len(x) for x in r[1:] if x != NC) if len(r) > 1 else 0)[0])
        for m_ in extras:
            if m_ not in piloto:
                piloto.append(m_)
        sel = set(piloto)
        dados_piloto = {aba: [r for r in imp[aba] if r[0] in sel] for aba in SHEETS}
        p = os.path.join(a.saida, 'PILOTO_teste_%d_matriculas_PADRAO_CONSOLIDADA.xlsx' % len(sel))
        gravar(dados_piloto, p); arquivos.insert(0, p)

    # ---- pendentes de conferência
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'LEIA'
    ws.append(['MATRICULA', 'LIVRO', 'CATEGORIA NO RELATÓRIO', 'MOTIVO', 'DETALHE'])
    for r in pendentes: ws.append(list(r))
    for aba in SHEETS:
        w = wb.create_sheet(('PEND ' + aba)[:31])
        w.append(['ARQUIVO DE ORIGEM', 'LINHA NA ORIGEM'] + headers[aba])
        for r in pend_linhas[aba]: w.append(r)
    p = os.path.join(a.saida, 'PENDENTES_conferencia_manual_NAO_IMPORTAR.xlsx'); wb.save(p); arquivos.append(p)

    # ---- relatório (contagens só das matrículas gravadas nas planilhas de importação)
    gravadas = {r[0] for d in ('importar', 'digitadas') for r in saida[d]['MATRÍCULA']}
    contagem = collections.Counter()
    for (m_, regra), n in log.contagem.items():
        if m_ in gravadas:
            contagem[regra] += n
    atencao = [x for x in atencao if x[0] in gravadas]
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Resumo'
    tot = lambda dest, aba: len(saida[dest][aba])
    ws.append(['Item', 'Quantidade'])
    ws.append(['Matrículas que faltam no sistema (relatório menos já importadas)', len(faltam)])
    ws.append(['Matrículas faltantes encontradas nas fontes', len(com_fonte)])
    ws.append(['Matrículas na planilha de importação', tot('importar', 'MATRÍCULA')])
    ws.append(['Matrículas DIGITADAS (planilha separada)', tot('digitadas', 'MATRÍCULA')])
    ws.append(['Matrículas pendentes de conferência manual', len(pendentes)])
    ws.append(['Matrículas faltantes sem nenhuma fonte', len(set(faltam) - com_fonte)])
    for aba in SHEETS:
        ws.append([f'Linhas em {aba} (importação)', tot('importar', aba)])
    ws.append([])
    ws.append(['Ajuste aplicado', 'Vezes'])
    for k, v in sorted(contagem.items(), key=lambda x: -x[1]): ws.append([k, v])
    w = wb.create_sheet('Ajustes'); w.append(['MATRICULA', 'ABA', 'ARQUIVO', 'LINHA', 'COLUNA', 'ANTES', 'DEPOIS', 'REGRA'])
    for r in log.ajustes:
        if r[0] in gravadas: w.append([str(x) for x in r])
    w = wb.create_sheet('Atencao'); w.append(['MATRICULA', 'MOTIVO', 'DETALHE'])
    for r in atencao: w.append(list(r))
    w = wb.create_sheet('Sem_fonte'); w.append(['MATRICULA', 'LIVRO', 'CATEGORIA NO RELATÓRIO'])
    for k in sorted(set(faltam) - com_fonte, key=int): w.append([k, livro_de.get(k, ''), faltam[k]])
    p = os.path.join(a.saida, 'Relatorio_ajustes.xlsx'); wb.save(p); arquivos.append(p)

    print('faltam no sistema:', len(faltam), '| com fonte:', len(com_fonte), '| importar:', tot('importar', 'MATRÍCULA'),
          '| digitadas:', tot('digitadas', 'MATRÍCULA'), '| pendentes:', len(pendentes), '| sem fonte:', len(set(faltam) - com_fonte))
    for k, v in sorted(contagem.items(), key=lambda x: -x[1]): print(f'   {v:7}  {k}')
    for p in arquivos: print('gerado:', p)


if __name__ == '__main__':
    main()
