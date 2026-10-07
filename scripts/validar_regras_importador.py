"""Valida uma planilha no padrão de 7 abas contra as regras com que o importador
do sistema rejeita uma matrícula em silêncio (sem log). As regras foram deduzidas
comparando as matrículas que entraram com as que falharam em 3 importações reais
(1.369 + 2.192 + 31 entraram; 57 + 17 + 1 falharam; 0 exceções):
  R1  dois atos da mesma matrícula com a mesma ORDEM (inclui R-n e AV-n com o mesmo n)
  R2  linha de ATO_PARTES cujo ID não corresponde a nenhum ato (TIPO ATO-ORDEM)
  R3  PROTOCOLO do ato que não é "NÃO CONSTA" nem número (pontos aceitos, até 9 dígitos)
  R4  DATA do ato que não existe no calendário (ex.: 29/02/2011)
  R5  DATA ABERTURA fora do padrão (por extenso etc.; precaução, não comprovado)
Uso: python3 scripts/validar_regras_importador.py arquivo.xlsx [...]"""
import sys, re, datetime, collections
sys.dont_write_bytecode = True
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from gerar_planilha_importacao import ler_xlsx, chave_matricula
NC = 'NÃO CONSTA'
def prot_ok(v):
    v = str(v).strip()
    if v == NC: return True
    if not re.fullmatch(r'\d+(\.\d+)*', v): return False
    return len(v.replace('.', '')) <= 9          # maior aceito visto: 9 dígitos; 11+ rejeitado
def data_ok(v):
    v = str(v).strip()
    if v == NC: return True
    m = re.fullmatch(r'(\d{2})/(\d{2})/(\d{4})', v)
    if not m: return False
    try: datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1))); return True
    except ValueError: return False
def abertura_ok(v):
    v = str(v).strip()
    return v == NC or re.fullmatch(r'\d{4,5}', v) is not None or data_ok(v)
def agrupar(wb):
    d = {}
    for sh, rows in wb.items():
        hdr = None
        for rn, r in rows:
            if hdr is None: hdr = {c: v[1] for c, v in r.items()}; continue
            m = chave_matricula(r.get('A'))
            if m is None: continue
            rec = {hdr.get(c, c): v[1] for c, v in r.items()}; rec['_row'] = rn
            d.setdefault(m, {}).setdefault(sh, []).append(rec)
    return d
def validar(d):
    out = collections.defaultdict(list)
    for m, s in d.items():
        atos = s.get('MATRICULA_ATO', [])
        c = collections.Counter(str(a.get('ORDEM')).strip() for a in atos)
        for a in atos:
            if c[str(a.get('ORDEM')).strip()] > 1:
                out[m].append(('R1 ORDEM repetida', 'MATRICULA_ATO', a['_row'], '%s-%s' % (a.get('TIPO ATO'), a.get('ORDEM'))))
            if not prot_ok(a.get('PROTOCOLO')):
                out[m].append(('R3 PROTOCOLO do ato inválido', 'MATRICULA_ATO', a['_row'], a.get('PROTOCOLO')))
            if not data_ok(a.get('DATA')):
                out[m].append(('R4 DATA do ato inválida', 'MATRICULA_ATO', a['_row'], a.get('DATA')))
        ids = {'%s-%s' % (str(a.get('TIPO ATO')).strip(), str(a.get('ORDEM')).strip()) for a in atos}
        for p in s.get('ATO_PARTES', []):
            if str(p.get('ID')).strip() not in ids:
                out[m].append(('R2 ID da parte sem ato', 'ATO_PARTES', p['_row'], p.get('ID')))
        for r in s.get('MATRÍCULA', []):
            if not abertura_ok(r.get('DATA ABERTURA')):
                out[m].append(('R5 DATA ABERTURA fora do padrão', 'MATRÍCULA', r['_row'], r.get('DATA ABERTURA')))
    return out
if __name__ == '__main__':
    for p in sys.argv[1:]:
        d = agrupar(ler_xlsx(p)); v = validar(d)
        print('==', p, '| matrículas:', len(d), '| com problema:', len(v))
        for m in sorted(v, key=lambda x: (len(x), x)):
            print('  ', m, v[m][:4], '(+%d)' % (len(v[m]) - 4) if len(v[m]) > 4 else '')
