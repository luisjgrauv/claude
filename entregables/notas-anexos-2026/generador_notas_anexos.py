# -*- coding: utf-8 -*-
"""Generador de notas y anexos 2026 — Distribuidora Andina S.A. (FICTICIA).

Lee los libros de la fase 2 (entregables/ajuste/, ya recalculados) y el auxiliar de seguros
(insumos/andina/auxiliares_2026.xlsx). No modifica ningún insumo.
Escribe:
  - anexos_andina_2026.xlsx  (datos importados en azul; anexos y controles con fórmulas)
  - notas_andina_2026.md
  - control_notas_anexos.json (cifras de control para el verificador y la reina)
Uso: python3 generador_notas_anexos.py   (desde la raíz del repositorio)
"""
import json
import os
from collections import OrderedDict

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

RAIZ = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
AJ = os.path.join(RAIZ, 'entregables', 'ajuste')
OUT = os.path.join(RAIZ, 'entregables', 'notas-anexos-2026')
AUX26 = os.path.join(RAIZ, 'insumos', 'andina', 'auxiliares_2026.xlsx')

# ---------------------------------------------------------------- lectura
wb25 = openpyxl.load_workbook(os.path.join(AJ, 'ajuste_2025.xlsx'), data_only=True)
wb26 = openpyxl.load_workbook(os.path.join(AJ, 'ajuste_2026.xlsx'), data_only=True)
wbcmp = openpyxl.load_workbook(os.path.join(AJ, 'comparativo_2025_en_moneda_dic2026.xlsx'), data_only=True)
wbaux = openpyxl.load_workbook(AUX26, data_only=True)


def header_row(ws, first='Código'):
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if r[0] == first:
            return i
    raise ValueError(ws.title)


def bss(ws):
    h = header_row(ws)
    d = OrderedDict()
    for r in ws.iter_rows(min_row=h + 1, values_only=True):
        if r[0] and isinstance(r[0], str) and r[0][0].isdigit():
            d[r[0]] = dict(cuenta=r[1], rubro=r[2], clase=r[3], ini_h=r[4], ini=r[5], fin_h=r[6], fin=r[7])
    return d


B26 = bss(wb26['BSS ajustado'])          # moneda dic-2026
B25 = bss(wb25['BSS ajustado'])          # moneda dic-2025
ws = wb26['Coeficientes']
K = ws['E3'].value                       # coef. dic-2025 -> dic-2026
IDX_DIC26 = ws['C4'].value
K25 = wb25['Coeficientes']['E3'].value   # coef. dic-2024 -> dic-2025
IDX_DIC25 = wb25['Coeficientes']['C4'].value
coef26 = {r[0].month: r[3] for r in ws.iter_rows(values_only=True) if hasattr(r[0], 'year') and r[0].year == 2026}
coef25 = {r[0].month: r[3] for r in wb25['Coeficientes'].iter_rows(values_only=True)
          if hasattr(r[0], 'year') and r[0].year == 2025}
estado26 = {r[0].month: r[2] for r in ws.iter_rows(values_only=True) if hasattr(r[0], 'year') and r[0].year == 2026}
IDX_DIC24 = IDX_DIC25 / K25
# índice dic-2023 (para la inflación acumulada de 3 años)
IDX_DIC23 = None
with open(os.path.join(RAIZ, 'insumos', 'indices', 'indices_facpce.csv'), encoding='utf-8') as f:
    for line in f:
        if line.startswith('2023-12;'):
            IDX_DIC23 = float(line.split(';')[1])


def bienes(ws, anio):
    h = header_row(ws)
    hdr = [c for c in next(ws.iter_rows(min_row=h, max_row=h, values_only=True))]
    col = {n: i for i, n in enumerate(hdr)}
    filas = []
    for r in ws.iter_rows(min_row=h + 1, values_only=True):
        if not r[0] or r[0] == 'TOTAL' or not isinstance(r[0], str) or r[2] is None:
            continue
        filas.append(dict(
            cod=r[0], desc=r[1], cta=r[2], funcion=r[5], fecha=r[6], vida=r[9],
            vo_ini=r[col['VO inicio (aj.)']], altas=r[col['Altas (aj.)']], bajas=r[col['Bajas VO (aj.)']],
            vo_fin=r[col['VO cierre (aj.)']], aa_ini=r[col['AA inicio (aj.)']], aa_bajas=r[col['Bajas AA (aj.)']],
            amort=r[col['Amortización %d (aj.)' % anio]], aa_fin=r[col['AA cierre (aj.)']],
            vo_ini_h=r[col['VO inicio 31/12/%d (hist.)' % (anio - 1)]], altas_h=r[col['Altas (hist.)']],
            bajas_h=r[col['Bajas VO (hist.)']], amort_h=r[col['Amortización %d (hist.)' % anio]]))
    return filas


BU26 = bienes(wb26['Bienes de uso'], 2026)
BU25 = bienes(wb25['Bienes de uso'], 2025)

# Seguros: auxiliar (descripción) + ajuste (importes)
ws = wbaux['Seguros']
h = header_row(ws, 'Póliza')
POL = OrderedDict()
for r in ws.iter_rows(min_row=h + 1, values_only=True):
    if r[0] and str(r[0]).startswith('POL-'):
        POL[r[0]] = dict(aseg=r[1], ramo=r[2], cob=r[3], bienes=r[4], desde=r[5], hasta=r[6],
                         pago=r[7], prima_h=r[8], funcion=r[16])
ws = wb26['Seguros']
h = header_row(ws, 'Póliza')
for r in ws.iter_rows(min_row=h + 1, values_only=True):
    if r[0] and str(r[0]).startswith('POL-'):
        POL[r[0]].update(cuenta=r[3], coef_pago=r[9], prima_aj=r[17], gasto_aj=r[19], nodev_aj=r[20], nodev_h=r[14])

# Comparativo (ER y ESP 2025 en moneda dic-2026) para controles cruzados
def tabla(ws):
    d = {}
    for r in ws.iter_rows(values_only=True):
        if r[0] and isinstance(r[1], (int, float)):
            d[r[0].strip()] = r[1]
    return d


ER25C = tabla(wbcmp['ER 2025'])
ESP25C = tabla(wbcmp['ESP 2025'])
PN26 = {r[0]: r[5] for r in wb26['Patrimonio neto'].iter_rows(values_only=True) if r[0]}
CTRL26 = {r[0]: r[1] for r in wb26['Controles'].iter_rows(values_only=True) if r[0]}

# ---------------------------------------------------------------- cálculos (Python)
def s(code, cual='fin'):
    return B26[code][cual]


def r25(code):
    """Resultado 2025 (moneda dic-2026), signo BSS (deudor +)."""
    return B25[code]['fin'] * K


RUBROS_BU = OrderedDict([
    ('1.6.01', ('Terrenos', None)),
    ('1.6.02', ('Edificios', '1.6.03')),
    ('1.6.04', ('Rodados — camiones', '1.6.05')),
    ('1.6.06', ('Rodados — vehículos ejecutivos', '1.6.07')),
    ('1.6.08', ('Rodados — camionetas de reparto', '1.6.09')),
])


def agrupa(filas, k=1.0):
    g = OrderedDict()
    for cta in RUBROS_BU:
        f = [x for x in filas if x['cta'] == cta]
        g[cta] = {c: sum(x[c] for x in f) * k for c in
                  ('vo_ini', 'altas', 'bajas', 'vo_fin', 'aa_ini', 'aa_bajas', 'amort', 'aa_fin')}
        g[cta]['neto'] = g[cta]['vo_fin'] - g[cta]['aa_fin']
    return g


G26 = agrupa(BU26)
G25 = agrupa(BU25, K)

c = {}
c['caja'] = s('1.1.01'); c['banco'] = s('1.1.02')
c['caja25'] = s('1.1.01', 'ini'); c['banco25'] = s('1.1.02', 'ini')
c['pf'] = s('1.2.01'); c['pf25'] = s('1.2.01', 'ini')
c['deud'] = s('1.3.01'); c['prev'] = s('1.3.02'); c['deud25'] = s('1.3.01', 'ini'); c['prev25'] = s('1.3.02', 'ini')
c['anti'] = s('1.4.01'); c['iva'] = s('1.4.02'); c['seg'] = s('1.4.03')
c['anti25'] = s('1.4.01', 'ini'); c['iva25'] = s('1.4.02', 'ini'); c['seg25'] = s('1.4.03', 'ini')
c['merc'] = s('1.5.01'); c['merc25'] = s('1.5.01', 'ini')
c['prov'] = -s('2.1.01'); c['prov25'] = -s('2.1.01', 'ini')
c['sue'] = -s('2.2.01'); c['sue25'] = -s('2.2.01', 'ini')
c['ig'] = -s('2.3.01'); c['iibb'] = -s('2.3.02'); c['ig25'] = -s('2.3.01', 'ini'); c['iibb25'] = -s('2.3.02', 'ini')
c['prest'] = -s('2.4.01'); c['prest25'] = -s('2.4.01', 'ini')
c['prest_c'] = 150_000_000.0; c['prest_nc'] = c['prest'] - c['prest_c']
c['prest_c25'] = 150_000_000.0 * K; c['prest_nc25'] = c['prest25'] - c['prest_c25']
c['bu'] = sum(g['neto'] for g in G26.values()); c['bu25'] = sum(g['neto'] for g in G25.values())

# Previsión para deudores incobrables (monetaria). Aplicaciones: Dato faltante de fechas -> supuesto lineal mensual.
APL26_H, APL25_H = 21_000_000.0, 16_500_000.0
prev = {}
prev['ini26'] = -s('1.3.02', 'ini'); prev['aum26'] = s('5.2.05')
prev['apl26'] = APL26_H / 12 * sum(coef26.values())
prev['fin26'] = -s('1.3.02'); prev['rec26'] = prev['fin26'] - prev['ini26'] - prev['aum26'] + prev['apl26']
prev['ini25'] = -B25['1.3.02']['ini'] * K; prev['aum25'] = r25('5.2.05')
prev['apl25'] = APL25_H / 12 * sum(coef25.values()) * K
prev['fin25'] = -B25['1.3.02']['fin'] * K; prev['rec25'] = prev['fin25'] - prev['ini25'] - prev['aum25'] + prev['apl25']

# CMV
def merc(wb):
    d = {}
    for r in wb['Mercaderías'].iter_rows(values_only=True):
        if r[0] in ('Total existencia inicial', 'Total compras', 'Total existencia final'):
            d[r[0]] = (r[1], r[3])
    return d


M26, M25 = merc(wb26), merc(wb25)
cmv = dict(ei26=M26['Total existencia inicial'][1], co26=M26['Total compras'][1], ef26=M26['Total existencia final'][1],
           ei25=M25['Total existencia inicial'][1] * K, co25=M25['Total compras'][1] * K,
           ef25=M25['Total existencia final'][1] * K, co26_h=M26['Total compras'][0], co25_h=M25['Total compras'][0])
cmv['cmv26'] = cmv['ei26'] + cmv['co26'] - cmv['ef26']
cmv['cmv25'] = cmv['ei25'] + cmv['co25'] - cmv['ef25']

# Gastos por naturaleza y función (art. 64 LGS)
GASTOS = [  # (naturaleza, cuenta comercialización, cuenta administración, cuenta financiación)
    ('Sueldos y cargas sociales', ['5.2.01'], ['5.3.01'], []),
    ('Combustibles y fletes', ['5.2.02'], [], []),
    ('Mantenimiento de flota', ['5.2.03'], [], []),
    ('Servicios y mantenimiento', [], ['5.3.03'], []),
    ('Honorarios profesionales', [], ['5.3.02'], []),
    ('Impuesto sobre los ingresos brutos', ['5.2.04'], [], []),
    ('Impuestos, tasas y contribuciones', [], ['5.3.04'], []),
    ('Deudores incobrables', ['5.2.05'], [], []),
    ('Seguros', ['5.2.06'], ['5.3.05'], []),
    ('Amortización de bienes de uso', ['5.2.07', '5.2.08', '5.2.09'], ['5.3.06', '5.3.07'], []),
    ('Intereses préstamo bancario', [], [], ['5.4.01']),
]
gas = []
for nat, cc, ca, cf in GASTOS:
    gas.append(dict(nat=nat, cc=cc, ca=ca, cf=cf,
                    com=sum(s(x) for x in cc), adm=sum(s(x) for x in ca), fin=sum(s(x) for x in cf),
                    com25=sum(r25(x) for x in cc), adm25=sum(r25(x) for x in ca), fin25=sum(r25(x) for x in cf)))
for g in gas:
    g['tot'] = g['com'] + g['adm'] + g['fin']; g['tot25'] = g['com25'] + g['adm25'] + g['fin25']
tot_com = sum(g['com'] for g in gas); tot_adm = sum(g['adm'] for g in gas); tot_fin = sum(g['fin'] for g in gas)
tot_com25 = sum(g['com25'] for g in gas); tot_adm25 = sum(g['adm25'] for g in gas); tot_fin25 = sum(g['fin25'] for g in gas)

# Resultados y ER (para IG y controles)
res26 = -sum(v['fin'] for k, v in B26.items() if v['clase'] == 'R')
res25 = -sum(r25(k) for k, v in B25.items() if v['clase'] == 'R')
ig26 = s('5.5.01'); ig25 = r25('5.5.01')
rai26 = res26 + ig26; rai25 = res25 + ig25
TASA = 0.35
rai26_h = -sum(v['fin_h'] for k, v in B26.items() if v['clase'] == 'R') + B26['5.5.01']['fin_h']
rai25_h = -sum(v['fin_h'] for k, v in B25.items() if v['clase'] == 'R') + B25['5.5.01']['fin_h']
recpam26 = -s('4.9.01'); recpam25 = -r25('4.9.01')

# Totales de rubro del ESP desde el BSS ajustado
def rubro_tot(rubro, cual):
    return sum(v[cual] for v in B26.values() if v['rubro'] == rubro)


ESP_RUBROS = ['Caja y bancos', 'Inversiones', 'Créditos por ventas', 'Otros créditos', 'Bienes de cambio',
              'Bienes de uso', 'Deudas comerciales', 'Remuneraciones y cargas sociales', 'Cargas fiscales', 'Préstamos']
esp = {r: (rubro_tot(r, 'fin'), rubro_tot(r, 'ini')) for r in ESP_RUBROS}
tot_act26 = sum(v['fin'] for v in B26.values() if v['clase'].startswith('A'))
tot_act25 = sum(v['ini'] for v in B26.values() if v['clase'].startswith('A'))
tot_pas26 = -sum(v['fin'] for v in B26.values() if v['clase'].startswith('P-'))
tot_pas25 = -sum(v['ini'] for v in B26.values() if v['clase'].startswith('P-'))
pn26 = tot_act26 - tot_pas26; pn25 = tot_act25 - tot_pas25
pmn26 = sum(v['fin'] for v in B26.values() if v['clase'] in ('A-M', 'P-M'))
pmn25 = sum(v['ini'] for v in B26.values() if v['clase'] in ('A-M', 'P-M'))
bu_h = sum(v['fin_h'] for v in B26.values() if v['rubro'] == 'Bienes de uso')

# Reserva legal (cifras de la hoja Patrimonio neto, fase 2)
rl = dict(base=PN26['Base: capital suscripto + ajuste de capital (aj.)'],
          tope=PN26['Tope 20 % de la base'], reserva=PN26['Reserva legal reexpresada (aj.)'],
          faltante=PN26['Faltante para el tope (si es positivo)'],
          omit25=PN26['Asignación omitida en la asamblea 28/04/2026: 5 % del resultado 2025 ajustado, en moneda dic-2026'],
          max26=PN26['Asignación máxima sobre el resultado 2026: 5 % del resultado 2026 ajustado'],
          prop=PN26['Asignación propuesta para la asamblea 2027 = mínimo entre faltante y (omitida 2025 + 5 % de 2026) (NO contabilizada)'])
div26 = -PN26['Dividendos en efectivo — Asamblea 28/04/2026 (pago 30/04/2026), sobre resultado 2025']
rna26 = PN26['RNA al 31/12/2026 después del resultado']

# ---------------------------------------------------------------- Excel
AZUL = Font(color='1F4E9E')
NEG = Font(bold=True)
TIT = Font(bold=True, size=12)
GRIS = PatternFill('solid', fgColor='E7E6E6')
FMT = '#,##0;(#,##0);"-"'
FMT6 = '0.000000'
FDATE = 'DD/MM/YYYY'
fino = Side(style='thin', color='999999')
SUP = Border(top=fino)

xl = openpyxl.Workbook()


def hoja(nombre, titulo, sub):
    w = xl.create_sheet(nombre)
    w['A1'] = titulo; w['A1'].font = TIT
    w['A2'] = sub; w['A2'].font = Font(italic=True, size=9)
    return w


def enc(w, fila, textos):
    for j, t in enumerate(textos, start=1):
        cc = w.cell(row=fila, column=j, value=t)
        cc.font = NEG; cc.fill = GRIS; cc.alignment = Alignment(wrap_text=True, vertical='top')


def put(w, fila, col, val, fmt=FMT, font=None):
    cc = w.cell(row=fila, column=col, value=val)
    if isinstance(val, (int, float)) or (isinstance(val, str) and val.startswith('=')):
        cc.number_format = fmt
    if font:
        cc.font = font
    return cc


def anchos(w, lst):
    for j, a in enumerate(lst, start=1):
        w.column_dimensions[get_column_letter(j)].width = a


xl.remove(xl.active)

# --- Hoja Notas (índice)
w = hoja('Índice', 'Distribuidora Andina S.A. (FICTICIA) — Anexos a los estados contables al 31/12/2026',
         'Moneda homogénea de diciembre de 2026 (índice de dic-2026 ESTIMADO). Pesos sin decimales. Azul = dato importado; negro = fórmula.')
filas_idx = [
    ('Anexo I', 'Bienes de uso (2026 y comparativo 2025)'),
    ('Anexo II', 'Previsiones'),
    ('Anexo III', 'Costo de mercaderías vendidas'),
    ('Anexo IV', 'Gastos por naturaleza y función (art. 64, inc. I, ap. b, LGS)'),
    ('Anexo V', 'Seguros contratados'),
    ('Controles', 'Conciliación de anexos con los BSS ajustados y con el comparativo 2025'),
    ('Datos BSS', 'BSS ajustado 2026 (saldos 31/12/2025 y 31/12/2026 en moneda dic-2026) y resultados 2025 reexpresados'),
    ('Datos BU', 'Bienes de uso bien por bien (2026 y 2025 en moneda dic-2026)'),
    ('Fuentes', 'entregables/ajuste/ajuste_2025.xlsx, ajuste_2026.xlsx, comparativo_2025_en_moneda_dic2026.xlsx; insumos/andina/auxiliares_2026.xlsx (hoja Seguros)'),
    ('Coeficiente', 'Coef. dic-2025 → dic-2026 = %.6f (dic-2026 estimado)' % K),
]
for i, (a, b) in enumerate(filas_idx, start=4):
    w.cell(row=i, column=1, value=a).font = NEG
    w.cell(row=i, column=2, value=b)
w['A15'] = 'Coeficiente dic-2025 → dic-2026'; put(w, 15, 2, K, FMT6, AZUL)
anchos(w, [32, 110])
KREF = "'Índice'!$B$15"

# --- Datos BSS
w = hoja('Datos BSS', 'Datos importados — BSS ajustado', 'Deudor (+) / acreedor (−). Col. D y E: ajuste_2026.xlsx, hoja BSS ajustado. '
         'Col. F: ajuste_2025.xlsx (moneda dic-2025). Col. G = F × coeficiente.')
enc(w, 3, ['Código', 'Cuenta', 'Rubro', 'Saldo 31/12/2025 (moneda dic-2026)', 'Saldo 31/12/2026 (moneda dic-2026)',
           'Saldo final 2025 (moneda dic-2025)', 'Saldo final 2025 (moneda dic-2026)'])
ROW = {}
for i, (code, v) in enumerate(B26.items(), start=4):
    ROW[code] = i
    w.cell(row=i, column=1, value=code); w.cell(row=i, column=2, value=v['cuenta']); w.cell(row=i, column=3, value=v['rubro'])
    put(w, i, 4, v['ini'], font=AZUL); put(w, i, 5, v['fin'], font=AZUL)
    v25 = B25.get(code)
    put(w, i, 6, v25['fin'] if v25 else 0, font=AZUL)
    put(w, i, 7, '=F%d*%s' % (i, KREF))
last = 3 + len(B26)
put(w, last + 1, 2, 'Total (debe ser 0)'); w.cell(row=last + 1, column=2).font = NEG
for col in 'DEFG':
    w['%s%d' % (col, last + 1)] = '=ROUND(SUM(%s4:%s%d),0)' % (col, col, last); w['%s%d' % (col, last + 1)].number_format = FMT
anchos(w, [9, 52, 32, 20, 20, 20, 20])


def D(code, col='E'):
    return "'Datos BSS'!%s%d" % (col, ROW[code])


# --- Datos BU
w = hoja('Datos BU', 'Datos importados — Bienes de uso bien por bien', 'Hoja "Bienes de uso" de ajuste_2026.xlsx (moneda dic-2026) '
         'y de ajuste_2025.xlsx (moneda dic-2025; se lleva a dic-2026 en el Anexo I con el coeficiente).')
cols_bu = ['Ejercicio', 'Código', 'Descripción', 'Cuenta VO', 'Función', 'Fecha de alta', 'Vida útil (años)', 'VO inicio', 'Altas',
           'Bajas', 'VO cierre', 'AA inicio', 'Bajas AA', 'Amortización', 'AA cierre']
enc(w, 3, cols_bu)
i = 4
for anio, filas in ((2026, BU26), (2025, BU25)):
    for x in filas:
        vals = [anio, x['cod'], x['desc'], x['cta'], x['funcion'], x['fecha'], x['vida'], x['vo_ini'], x['altas'], x['bajas'],
                x['vo_fin'], x['aa_ini'], x['aa_bajas'], x['amort'], x['aa_fin']]
        for j, v in enumerate(vals, start=1):
            cc = w.cell(row=i, column=j, value=v)
            if j >= 8:
                cc.number_format = FMT; cc.font = AZUL
            if j == 6:
                cc.number_format = FDATE
        i += 1
BU_LAST = i - 1
anchos(w, [9, 10, 58, 10, 16, 12, 9] + [16] * 8)

# --- Anexo I Bienes de uso
w = hoja('Anexo I BU', 'Anexo I — Bienes de uso al 31/12/2026, comparativo con el 31/12/2025',
         'Moneda dic-2026. Método: costo reexpresado bien por bien (coef. desde el mes de alta); amortización lineal, año de alta completo, '
         'año de baja sin amortización. Fuente: hoja Datos BU (SUMIFS).')
hdr = ['Rubro', 'Cuenta', 'VO al inicio', 'Altas', 'Bajas', 'VO al cierre', 'AA al inicio', 'Bajas AA', 'Amortización del ejercicio',
       'AA al cierre', 'Valor residual al cierre', 'Vida útil (años)']
RANGO = "'Datos BU'!$%s$4:$%s$%d"


def bloque_bu(fila0, anio, factor):
    enc(w, fila0, hdr)
    w.cell(row=fila0 - 1, column=1, value='Ejercicio %d' % anio).font = NEG
    vidas = {'1.6.01': '—', '1.6.02': 50, '1.6.04': 10, '1.6.06': 5, '1.6.08': 5}
    r = fila0 + 1
    for cta, (nombre, _) in RUBROS_BU.items():
        w.cell(row=r, column=1, value=nombre); w.cell(row=r, column=2, value=cta)
        for j, colsrc in zip(range(3, 11), 'HIJKLMNO'):
            if colsrc in 'K':
                f = '=C{0}+D{0}-E{0}'.format(r)
            elif colsrc == 'O':
                f = '=G{0}-H{0}+I{0}'.format(r)
            else:
                f = "=SUMIFS(%s,%s,%d,%s,\"%s\")%s" % (RANGO % (colsrc, colsrc, BU_LAST), RANGO % ('A', 'A', BU_LAST), anio,
                                                        RANGO % ('D', 'D', BU_LAST), cta, factor)
            put(w, r, j, f)
        put(w, r, 11, '=F{0}-J{0}'.format(r))
        w.cell(row=r, column=12, value=vidas[cta])
        r += 1
    w.cell(row=r, column=1, value='Total').font = NEG
    for j in range(3, 12):
        cl = get_column_letter(j)
        cc = put(w, r, j, '=SUM(%s%d:%s%d)' % (cl, fila0 + 1, cl, r - 1)); cc.font = NEG; cc.border = SUP
    return r


T26 = bloque_bu(5, 2026, '')
T25 = bloque_bu(T26 + 4, 2025, '*' + KREF)
n = T25 + 2
notas_bu = [
    'Bajas 2026: camión C01 (alta 05/2018) vendido el 15/06/2026. Precio $28.000.000 + IVA. Resultado reexpresado: pérdida de $13.750.256 '
    '(cuenta 4.3.01); el histórico era ganancia de $27.525.000.',
    'Altas 2026: camión C11 (alta 15/06/2026) $175.000.000 nominales ($196.538.478 en moneda dic-2026). Altas 2025: camión C10 y vehículo E05.',
    'Las amortizaciones se imputan a comercialización (nave de depósito, camiones y camionetas) y administración (sede y vehículos ejecutivos): ver Anexo IV.',
    'Dato faltante: prueba de valor recuperable (RT 54). No se recibió información para estimar el valor de uso ni el valor neto de realización.',
    'Bienes totalmente amortizados en uso al 31/12/2026: E01, E02 y U01.',
]
for t in notas_bu:
    w.cell(row=n, column=1, value=t).font = Font(size=9); n += 1
anchos(w, [34, 9] + [17] * 9 + [10])

# --- Anexo II Previsiones
w = hoja('Anexo II Prev', 'Anexo II — Previsiones', 'Moneda dic-2026. Previsión para deudores incobrables (regularizadora de créditos por ventas; partida monetaria).')
enc(w, 4, ['Concepto', 'Saldo al inicio', 'Aumentos (cargo a resultados)', 'Aplicaciones', 'RECPAM de la previsión',
           'Saldo al cierre', 'Aplicaciones (hist.)', 'Suma de coeficientes mensuales'])
w['A3'] = 'Deducidas del activo'; w['A3'].font = NEG
sum26 = sum(coef26.values()); sum25 = sum(coef25.values())
w['A5'] = 'Para deudores incobrables — 2026'
put(w, 5, 2, '=-%s' % D('1.3.02', 'D')); put(w, 5, 3, '=%s' % D('5.2.05'))
put(w, 5, 7, APL26_H, font=AZUL); put(w, 5, 8, sum26, FMT6, AZUL)
put(w, 5, 4, '=G5/12*H5'); put(w, 5, 6, '=-%s' % D('1.3.02')); put(w, 5, 5, '=F5-B5-C5+D5')
w['A6'] = 'Para deudores incobrables — 2025 (comparativo)'
put(w, 6, 2, '=%f*%s' % (-B25['1.3.02']['ini'], KREF)); w['B6'].font = AZUL
put(w, 6, 3, '=%s' % D('5.2.05', 'G'))
put(w, 6, 7, APL25_H, font=AZUL); put(w, 6, 8, sum25, FMT6, AZUL)
put(w, 6, 4, '=G6/12*H6*%s' % KREF); put(w, 6, 6, '=-%s' % D('1.3.02', 'G')); put(w, 6, 5, '=F6-B6-C6+D6')
w['A8'] = ('Supuesto (S): las aplicaciones se distribuyen en forma lineal mensual (Dato faltante: fechas de las aplicaciones). '
           'Aplicaciones (aj.) = aplicaciones históricas / 12 × Σ coeficientes de los meses del ejercicio.')
w['A9'] = ('RECPAM de la previsión: efecto de la inflación sobre una partida monetaria regularizadora del activo (el saldo pierde '
           'poder adquisitivo y se reduce en moneda de cierre). Se presenta para conciliar inicio y cierre; ya integra el RECPAM del estado de resultados.')
w['A10'] = 'Sin otras previsiones (no se informaron contingencias). Dato faltante: confirmación de juicios y reclamos al cierre.'
for a in ('A8', 'A9', 'A10'):
    w[a].font = Font(size=9)
anchos(w, [46, 17, 19, 17, 17, 17, 17, 14])

# --- Anexo III CMV
w = hoja('Anexo III CMV', 'Anexo III — Costo de mercaderías vendidas', 'Moneda dic-2026. FIFO por mes de compra (capas reexpresadas con el coeficiente del mes de compra).')
enc(w, 4, ['Concepto', '2026', '2025 (comparativo)', '2025 (moneda dic-2025)'])
put(w, 5, 1, 'Existencia al inicio'); put(w, 5, 2, cmv['ei26'], font=AZUL); put(w, 5, 4, M25['Total existencia inicial'][1], font=AZUL)
put(w, 6, 1, 'Más: compras del ejercicio'); put(w, 6, 2, cmv['co26'], font=AZUL); put(w, 6, 4, M25['Total compras'][1], font=AZUL)
put(w, 7, 1, 'Menos: existencia al cierre'); put(w, 7, 2, '=-%s' % D('1.5.01')); put(w, 7, 4, -M25['Total existencia final'][1], font=AZUL)
put(w, 8, 1, 'Costo de mercaderías vendidas'); w['A8'].font = NEG
for col in 'BD':
    cc = put(w, 8, ord(col) - 64, '=SUM(%s5:%s7)' % (col, col)); cc.font = NEG; cc.border = SUP
for rr in range(5, 9):
    put(w, rr, 3, '=D%d*%s' % (rr, KREF))
w['C8'].font = NEG
w['A10'] = 'Compras históricas (nominales) 2026'; put(w, 10, 2, cmv['co26_h'], font=AZUL)
w['A11'] = 'Compras históricas (nominales) 2025'; put(w, 11, 2, cmv['co25_h'], font=AZUL)
w['A13'] = ('Existencia inicial 2026 = existencia final 2025 reexpresada (coef. dic-2025 → dic-2026). Supuesto de la fase 2: '
            'sin prueba de VNR / costo de reposición (Dato faltante).')
w['A13'].font = Font(size=9)
anchos(w, [46, 20, 20, 20])

# --- Anexo IV Gastos
w = hoja('Anexo IV Gastos', 'Anexo IV — Información requerida por el art. 64, inc. I, ap. b), Ley 19.550 (LGS)',
         'Moneda dic-2026. Gastos por naturaleza y por función. El CMV se expone en el Anexo III; el RECPAM no es un gasto y no se incluye.')
enc(w, 4, ['Rubro (naturaleza)', 'Gastos de comercialización', 'Gastos de administración', 'Gastos de financiación',
           'Total 2026', 'Total 2025 (comparativo)', 'Cuentas'])


def sumcod(codes, col='E'):
    return '=' + '+'.join(D(x, col) for x in codes) if codes else 0


r = 5
for g in gas:
    w.cell(row=r, column=1, value=g['nat'])
    put(w, r, 2, sumcod(g['cc'])); put(w, r, 3, sumcod(g['ca'])); put(w, r, 4, sumcod(g['cf']))
    put(w, r, 5, '=SUM(B{0}:D{0})'.format(r))
    allc = g['cc'] + g['ca'] + g['cf']
    put(w, r, 6, sumcod(allc, 'G'))
    w.cell(row=r, column=7, value=', '.join(allc))
    r += 1
GT = r
w.cell(row=r, column=1, value='Total 2026').font = NEG
for j in range(2, 7):
    cl = get_column_letter(j)
    cc = put(w, r, j, '=SUM(%s5:%s%d)' % (cl, cl, r - 1)); cc.font = NEG; cc.border = SUP
r += 1
w.cell(row=r, column=1, value='Total 2025 (comparativo)').font = NEG
lst = lambda key: [x for g in gas for x in g[key]]
put(w, r, 2, sumcod(lst('cc'), 'G')); put(w, r, 3, sumcod(lst('ca'), 'G')); put(w, r, 4, sumcod(lst('cf'), 'G'))
put(w, r, 6, '=SUM(B{0}:D{0})'.format(r))
GT25 = r
anchos(w, [40, 20, 20, 20, 18, 20, 30])

# --- Anexo V Seguros
w = hoja('Anexo V Seguros', 'Anexo V — Seguros contratados al 31/12/2026',
         'Fuente: insumos/andina/auxiliares_2026.xlsx (hoja Seguros) y ajuste_2026.xlsx (hoja Seguros). Importes ajustados en moneda dic-2026.')
enc(w, 4, ['Póliza', 'Aseguradora', 'Ramo', 'Cobertura', 'Bienes asegurados', 'Vigencia desde', 'Vigencia hasta', 'Función',
           'Prima (nominal)', 'Prima (moneda dic-2026)', 'Gasto 2026 (moneda dic-2026)', 'No devengado al 31/12/2026 (moneda dic-2026)',
           'Suma asegurada'])
r = 5
for k, p in POL.items():
    vals = [k, p['aseg'], p['ramo'], p['cob'], p['bienes'], p['desde'], p['hasta'], p['funcion']]
    for j, v in enumerate(vals, start=1):
        cc = w.cell(row=r, column=j, value=v)
        if j in (6, 7):
            cc.number_format = FDATE
    put(w, r, 9, p['prima_h'], font=AZUL); put(w, r, 10, p['prima_aj'], font=AZUL)
    put(w, r, 11, p['gasto_aj'], font=AZUL); put(w, r, 12, p['nodev_aj'], font=AZUL)
    w.cell(row=r, column=13, value='Dato faltante')
    r += 1
ST = r
w.cell(row=r, column=1, value='Total').font = NEG
for j in range(9, 13):
    cl = get_column_letter(j)
    cc = put(w, r, j, '=SUM(%s5:%s%d)' % (cl, cl, r - 1)); cc.font = NEG; cc.border = SUP
w.cell(row=r + 2, column=1, value=('El gasto 2026 incluye el devengamiento de las pólizas anteriores (vigencias 2025-2026). El no devengado integra '
                                   '"Otros créditos — Seguros pagados por adelantado". Dato faltante: sumas aseguradas y franquicias; no se puede '
                                   'evaluar si la cobertura es suficiente respecto del valor de los bienes.')).font = Font(size=9)
anchos(w, [9, 24, 20, 36, 44, 12, 12, 16, 15, 15, 15, 17, 14])

# --- Controles
w = hoja('Controles', 'Controles — anexos contra BSS ajustados y comparativo 2025', 'OK si |diferencia| < 1 peso.')
enc(w, 3, ['Control', 'Anexo', 'BSS / fuente', 'Diferencia', 'Resultado'])
A1 = "'Anexo I BU'!"; A3 = "'Anexo III CMV'!"; A4 = "'Anexo IV Gastos'!"; A2 = "'Anexo II Prev'!"; A5 = "'Anexo V Seguros'!"
bu_codes = ['1.6.0%d' % i for i in range(1, 10)]
ctr = [
    ('Anexo I: residual 2026 = rubro Bienes de uso 31/12/2026', A1 + 'K%d' % T26, '=' + '+'.join(D(x) for x in bu_codes)),
    ('Anexo I: residual 2025 = rubro Bienes de uso 31/12/2025', A1 + 'K%d' % T25, '=' + '+'.join(D(x, 'D') for x in bu_codes)),
    ('Anexo I: VO inicio 2026 = VO cierre 2025 (moneda dic-2026)', A1 + 'C%d' % T26, '=%sF%d' % (A1, T25)),
    ('Anexo I: AA inicio 2026 = AA cierre 2025 (moneda dic-2026)', A1 + 'G%d' % T26, '=%sJ%d' % (A1, T25)),
    ('Anexo I: VO cierre 2026 = cuentas VO del BSS', A1 + 'F%d' % T26, '=' + '+'.join(D(x) for x in ['1.6.01', '1.6.02', '1.6.04', '1.6.06', '1.6.08'])),
    ('Anexo I: amortización 2026 = cuentas de amortización del ER', A1 + 'I%d' % T26, '=' + '+'.join(D(x) for x in ['5.2.07', '5.2.08', '5.2.09', '5.3.06', '5.3.07'])),
    ('Anexo I: amortización 2025 = cuentas de amortización del ER 2025', A1 + 'I%d' % T25, '=' + '+'.join(D(x, 'G') for x in ['5.2.07', '5.2.08', '5.2.09', '5.3.06', '5.3.07'])),
    ('Anexo II: saldo final 2026 = previsión del BSS', A2 + 'F5', '=-' + D('1.3.02')),
    ('Anexo II: aumentos 2026 = cargo por incobrables del ER', A2 + 'C5', '=' + D('5.2.05')),
    ('Anexo III: CMV 2026 = cuenta 5.1.01', A3 + 'B8', '=' + D('5.1.01')),
    ('Anexo III: CMV 2025 = cuenta 5.1.01 2025 (moneda dic-2026)', A3 + 'C8', '=' + D('5.1.01', 'G')),
    ('Anexo III: existencia inicial 2026 = bienes de cambio 31/12/2025', A3 + 'B5', '=' + D('1.5.01', 'D')),
    ('Anexo III: existencia final 2025 (dic-26) = existencia inicial 2026', '=-' + A3 + 'C7', '=' + A3 + 'B5'),
    ('Anexo III: compras históricas 2026 = 13.115.000.000 (cifra de control)', A3 + 'B10', 13115000000),
    ('Anexo IV: total comercialización 2026 = rubro del ER', A4 + 'B%d' % GT, '=' + '+'.join(D(k) for k, v in B26.items() if v['rubro'] == 'Gastos de comercialización')),
    ('Anexo IV: total administración 2026 = rubro del ER', A4 + 'C%d' % GT, '=' + '+'.join(D(k) for k, v in B26.items() if v['rubro'] == 'Gastos de administración')),
    ('Anexo IV: financiación 2026 = intereses del préstamo', A4 + 'D%d' % GT, '=' + D('5.4.01')),
    ('Anexo IV: comercialización 2025 = ER 2025 comparativo', A4 + 'B%d' % GT25, -ER25C['Gastos de comercialización']),
    ('Anexo IV: administración 2025 = ER 2025 comparativo', A4 + 'C%d' % GT25, -ER25C['Gastos de administración']),
    ('Anexo IV: financiación 2025 = ER 2025 comparativo (intereses)', A4 + 'D%d' % GT25, -ER25C['Intereses préstamo bancario']),
    ('Anexo V: gasto de seguros 2026 = cuentas 5.2.06 + 5.3.05', A5 + 'K%d' % ST, '=%s+%s' % (D('5.2.06'), D('5.3.05'))),
    ('Anexo V: no devengado 31/12/2026 = cuenta 1.4.03', A5 + 'L%d' % ST, '=' + D('1.4.03')),
    ('Anexo V: primas nominales 2026 = 184.680.000 (datos adicionales 2026)', A5 + 'I%d' % ST, 184680000),
    ('Datos BSS: saldos 31/12/2026 cuadran', "'Datos BSS'!E%d" % (last + 1), 0),
    ('Datos BSS: saldos 31/12/2025 cuadran', "'Datos BSS'!D%d" % (last + 1), 0),
    ('Bienes de uso 2025 (anexo) = ESP 2025 comparativo', A1 + 'K%d' % T25, ESP25C['Bienes de uso']),
]
for i, (t, a, b) in enumerate(ctr, start=4):
    w.cell(row=i, column=1, value=t)
    put(w, i, 2, a if str(a).startswith('=') else '=' + a)
    put(w, i, 3, b)
    if isinstance(b, (int, float)):
        w.cell(row=i, column=3).font = AZUL
    put(w, i, 4, '=B{0}-C{0}'.format(i), '#,##0.00;(#,##0.00);"-"')
    w.cell(row=i, column=5, value='=IF(ABS(D{0})<1,"OK","ERROR")'.format(i))
CL = 3 + len(ctr)
w.cell(row=CL + 1, column=1, value='Resultado global').font = NEG
w.cell(row=CL + 1, column=5, value='=IF(COUNTIF(E4:E%d,"OK")=%d,"OK","ERROR")' % (CL, len(ctr))).font = NEG
anchos(w, [70, 20, 20, 14, 10])

orden = ['Índice', 'Anexo I BU', 'Anexo II Prev', 'Anexo III CMV', 'Anexo IV Gastos', 'Anexo V Seguros', 'Controles', 'Datos BSS', 'Datos BU']
xl._sheets = [xl[n] for n in orden]
os.makedirs(OUT, exist_ok=True)
XLSX = os.path.join(OUT, 'anexos_andina_2026.xlsx')
xl.save(XLSX)

# ---------------------------------------------------------------- Notas (Markdown)
def m(x):
    """Pesos sin decimales, separador de miles con punto, negativos entre paréntesis."""
    v = round(x)
    t = '{:,.0f}'.format(abs(v)).replace(',', '.')
    return '(%s)' % t if v < 0 else t


def pct(x, d=2):
    return ('{:.%df}' % d).format(x * 100).replace('.', ',') + ' %'


def c6(x):
    return '{:.6f}'.format(x).replace('.', ',')


infl25 = K25 - 1; infl26 = K - 1; infl3 = IDX_DIC26 / IDX_DIC23 - 1
N = []
P = N.append
P('# Distribuidora Andina S.A. (FICTICIA) — Notas a los estados contables al 31/12/2026')
P('')
P('Ejercicio económico iniciado el 01/01/2026 y finalizado el 31/12/2026, presentado en forma comparativa con el ejercicio 2025.')
P('Cifras en pesos de **diciembre de 2026**, sin decimales. Datos ficticios, de práctica.')
P('')
P('> **Advertencia sobre las cifras:** los índices de noviembre y diciembre de 2026 son **estimados**, no oficiales. '
  'Todas las cifras en moneda de dic-2026 deben recalcularse cuando se publiquen los índices oficiales. Ver la Nota 1.3.')
P('')
P('Los anexos I a V están en `anexos_andina_2026.xlsx`. La tabla de control final concilia estas notas con los BSS ajustados.')
P('')
P('## Nota 1 — Bases de preparación')
P('')
P('### 1.1 Normas contables aplicadas')
P('- Los estados contables se prepararon según las normas contables profesionales argentinas de la FACPCE.')
P('- **Norma aplicada:** RT 54, Norma Unificada Argentina de Contabilidad (NUA), texto ordenado por la RT 59. Es de aplicación obligatoria '
  'para ejercicios iniciados desde el 01/01/2025. La fecha de vigencia la fija cada consejo profesional; la jurisdicción del ente es CABA (CPCECABA).')
P('- La RT 54 derogó, entre otras, las RT 6, 8, 9, 17 y 41. Incorpora el procedimiento de reexpresión que antes estaba en la RT 6.')
P('- **Ley 19.550 General de Sociedades (LGS):** arts. 62 a 65 (documentación contable), art. 64 (información del estado de resultados: Anexo IV) y art. 70 (reserva legal: Nota 9).')
P('- La denominación "Bienes de uso" se mantiene por coherencia con el plan de cuentas. En la RT 54 el rubro se denomina "Propiedades, planta y equipo".')
P('- **Duda:** confirmar (a) el número de la resolución del CPCECABA que adoptó la RT 54 y la RT 59 y (b) la numeración exacta de los párrafos de la RT 54 sobre unidad de medida, antes de citarlos en la versión final.')
P('')
P('### 1.2 Unidad de medida (moneda homogénea)')
P('- Los estados contables se expresan en moneda homogénea de diciembre de 2026. Se reconocen los efectos de la inflación según el procedimiento de reexpresión '
  'de la RT 54, que tomó el de la RT 6, y la Res. JG FACPCE 539/18 (modif. Res. JG 553/19 y 584/21).')
P('- **Índice:** serie FACPCE de la Res. JG 539/18. Es el IPC Nacional del INDEC (base dic-2016 = 100) desde ene-2017, empalmado con el IPIM para los meses anteriores. '
  'Fuente: `insumos/indices/indices_facpce.csv`, consultado el 22/11/2026.')
P('- **Contexto de inflación:** la variación acumulada de los últimos tres años (dic-2023 a dic-2026) es de %s. Supera la pauta cuantitativa del 100 %% acumulado en tres años (RT 6, modif. RT 39, retomada por la RT 54; contexto declarado por la Res. JG 539/18).' % pct(infl3, 1))
P('')
P('| Concepto | dic-2024 | dic-2025 | dic-2026 |')
P('|---|---:|---:|---:|')
P('| Índice FACPCE (Res. JG 539/18) | {} | {} | {} (estimado) |'.format(
    '{:,.4f}'.format(IDX_DIC24).replace(',', 'X').replace('.', ',').replace('X', '.'),
    '{:,.4f}'.format(IDX_DIC25).replace(',', 'X').replace('.', ',').replace('X', '.'),
    '{:,.4f}'.format(IDX_DIC26).replace(',', 'X').replace('.', ',').replace('X', '.')))
P('| Coeficiente del ejercicio (dic-previo → dic) | — | %s | %s |' % (c6(K25), c6(K)))
P('| Inflación anual | — | %s (oficial; INDEC publicó 31,5 %%) | %s (estimada) |' % (pct(infl25), pct(infl26)))
P('')
P('- **Coeficiente de reexpresión del comparativo:** cifras de 2025 en moneda de dic-2025 × %s.' % c6(K))
P('- **Método:**')
P('  - Las partidas monetarias se exponen a su valor nominal.')
P('  - Las partidas no monetarias se reexpresan desde su fecha de origen: bienes de uso bien por bien, mercaderías por capas FIFO, seguros por fecha de pago y patrimonio neto por componente.')
P('  - Los resultados se reexpresan con el coeficiente de su mes de devengamiento. Los que consumen activos no monetarios (CMV, amortizaciones, seguros y venta de bienes de uso) se reexpresan desde su auxiliar.')
P('  - El RECPAM se determina por diferencia y se controla por el método directo. La diferencia entre ambos es menor a $1.')
P('- **RECPAM:** pérdida de $%s en 2026 y de $%s en 2025 (moneda dic-2026). Es pérdida porque la posición monetaria neta es activa: $' % (m(-recpam26), m(-recpam25))
  + '%s al 31/12/2026 y $%s al 31/12/2025.' % (m(pmn26), m(pmn25)))
P('')
P('### 1.3 Estimaciones y cifras sujetas a revisión')
P('- **Estimación — índices de nov-2026 y dic-2026.** Son una estimación propia: la media geométrica de las variaciones mensuales de ago a oct-2026 (1,9213 %). No son datos oficiales.')
P('  - El índice de nov-2026 se publica el 16/12/2026. El de dic-2026, en enero de 2027.')
P('  - Sensibilidad: si nov y dic repitieran el 2,17 % de octubre, el índice de dic-2026 sería alrededor de 0,5 % mayor. Todas las cifras en moneda de dic-2026 cambiarían aproximadamente en esa proporción.')
P('  - Recomendación: regenerar el ajuste y estas notas con los índices oficiales antes de emitir los estados contables.')
P('- Otras estimaciones: vidas útiles, previsión para deudores incobrables e impuesto a las ganancias.')
P('')
P('### 1.4 Información comparativa')
P('- Se presentan comparativos del ejercicio 2025 para el ESP, el estado de resultados, el EEPN, las notas y los anexos. Están reexpresados a moneda de dic-2026.')
P('- **Limitación:** 2025 es el primer ejercicio en que se aplicó el ajuste por inflación en esta preparación.')
P('  - Los resultados no asignados al 31/12/2024 se determinaron por diferencia entre activos y pasivos reexpresados y los demás componentes del PN.')
P('  - El comparativo depende de los mismos supuestos de origen que los saldos de 2026: fechas de alta, capas de mercaderías y fechas de pago de las primas.')
P('- El ESP al 31/12/2024 está en `comparativo_2025_en_moneda_dic2026.xlsx`. No se presenta en estas notas.')
P('')
P('## Nota 2 — Principales criterios de medición')
P('')
P('| Rubro | Criterio |')
P('|---|---|')
P('| Caja y bancos | Valor nominal. |')
P('| Inversiones (plazo fijo) | Valor nominal más intereses devengados al cierre. Vencimiento original menor a 90 días: se considera equivalente de efectivo para el EFE. |')
P('| Créditos por ventas | Valor nominal, neto de la previsión para deudores incobrables. No se segregan componentes financieros implícitos: Dato faltante sobre plazos de cobro; se asume que no son significativos. |')
P('| Otros créditos | Anticipos e IVA a valor nominal. Seguros pagados por adelantado: parte no devengada de las primas, reexpresada desde el mes de pago. |')
P('| Bienes de cambio | Costo de adquisición reexpresado, método FIFO por capas mensuales de compra. Dato faltante: prueba de valor neto de realización o costo de reposición. |')
P('| Bienes de uso | Costo de adquisición reexpresado desde el mes de alta, menos amortización acumulada. Amortización lineal: año de alta completo y año de baja sin amortización. Vidas útiles: edificios 50 años, camiones 10 años, vehículos ejecutivos y camionetas 5 años. Los terrenos no se amortizan. Dato faltante: prueba de valor recuperable. |')
P('| Deudas (comerciales, remuneraciones, cargas fiscales y préstamos) | Valor nominal más intereses devengados. El préstamo se separa en porción corriente y no corriente. |')
P('| Impuesto a las ganancias | Método del impuesto corriente (impuesto determinado del ejercicio), según el supuesto simplificador del insumo. Ver la Duda de la Nota 8. |')
P('| Patrimonio neto | Capital a su valor nominal. La reexpresión se expone en "Ajuste de capital". La reserva legal y los resultados no asignados se reexpresan desde su fecha de origen. |')
P('| Resultados | Reexpresados desde el mes de devengamiento. Los resultados por consumo de activos no monetarios se reexpresan desde el auxiliar de esos activos. |')
P('')

def tab(titulo, filas, nota=None):
    P('| %s | 31/12/2026 | 31/12/2025 |' % titulo)
    P('|---|---:|---:|')
    for f in filas:
        if f[0].startswith('**'):
            P('| %s | **%s** | **%s** |' % (f[0], m(f[1]), m(f[2])))
        else:
            P('| %s | %s | %s |' % (f[0], m(f[1]), m(f[2])))
    if nota:
        P('')
        P(nota)
    P('')


P('## Nota 3 — Caja y bancos')
P('')
tab('Concepto', [('Caja', c['caja'], c['caja25']), ('Banco Ficticio — cuenta corriente', c['banco'], c['banco25']),
                 ('**Total**', c['caja'] + c['banco'], c['caja25'] + c['banco25'])])
P('## Nota 4 — Inversiones')
P('')
tab('Concepto', [('Plazo fijo (vencimiento original < 90 días)', c['pf'], c['pf25']), ('**Total corriente**', c['pf'], c['pf25'])],
    'Los intereses ganados en el ejercicio fueron $%s en 2026 y $%s en 2025, ambos en moneda de dic-2026 (cuenta 4.2.01).' % (m(-s('4.2.01')), m(-r25('4.2.01'))))
P('## Nota 5 — Créditos por ventas')
P('')
tab('Concepto', [('Deudores por ventas', c['deud'], c['deud25']), ('Menos: previsión para deudores incobrables (Anexo II)', c['prev'], c['prev25']),
                 ('**Total corriente**', c['deud'] + c['prev'], c['deud25'] + c['prev25'])],
    'Dato faltante: antigüedad de los créditos y plazos de cobro. Por esa razón no se presenta la apertura por vencimiento.')
P('## Nota 6 — Otros créditos')
P('')
tab('Concepto', [('Anticipos de impuesto a las ganancias', c['anti'], c['anti25']), ('IVA saldo técnico a favor', c['iva'], c['iva25']),
                 ('Seguros pagados por adelantado (Anexo V)', c['seg'], c['seg25']),
                 ('**Total corriente**', c['anti'] + c['iva'] + c['seg'], c['anti25'] + c['iva25'] + c['seg25'])],
    '**Duda (presentación):** los anticipos de impuesto a las ganancias se exponen sin compensar con la deuda del impuesto (Nota 7), igual que en el BSS y en el comparativo. '
    'La compensación procede si existe el derecho legal de compensar y la intención de cancelar por el neto. '
    'Si se compensa, otros créditos y cargas fiscales bajan $%s en 2026 y $%s en 2025. El criterio debe coincidir con el de `eecc-2026`.' % (m(c['anti']), m(c['anti25'])))
P('## Nota 7 — Bienes de cambio, bienes de uso y deudas')
P('')
P('### 7.1 Bienes de cambio')
P('')
tab('Concepto', [('Mercaderías de reventa (insumos), FIFO', c['merc'], c['merc25']), ('**Total corriente**', c['merc'], c['merc25'])],
    'La existencia al 31/12/2026 corresponde a compras de noviembre ($%s nominales) y diciembre de 2026 ($%s nominales). El costo de ventas está en el Anexo III.'
    % (m(413120000), m(1211880000)))
P('### 7.2 Bienes de uso (detalle en el Anexo I y en `insumos/andina/auxiliares_2026.xlsx`, hoja Bienes de uso)')
P('')
P('| Rubro | VO 31/12/2026 | AA 31/12/2026 | Neto 31/12/2026 | Neto 31/12/2025 |')
P('|---|---:|---:|---:|---:|')
for cta, (nombre, _) in RUBROS_BU.items():
    P('| %s | %s | %s | %s | %s |' % (nombre, m(G26[cta]['vo_fin']), m(-G26[cta]['aa_fin']), m(G26[cta]['neto']), m(G25[cta]['neto'])))
P('| **Total no corriente** | **%s** | **%s** | **%s** | **%s** |' % (
    m(sum(g['vo_fin'] for g in G26.values())), m(-sum(g['aa_fin'] for g in G26.values())), m(c['bu']), m(c['bu25'])))
P('')
P('Composición al 31/12/2026, según el auxiliar:')
P('- **Inmuebles (2):** INM-01, sede administrativa (alta 03/2016), e INM-02, planta / centro de distribución (alta 07/2019). Cada uno se registra con terreno y edificio por separado.')
P('- **Camiones (10):** C02 a C11.')
P('  - Baja: C01, vendido el 15/06/2026 por $28.000.000 + IVA. Su resultado reexpresado es una **pérdida de $%s**: el valor residual reexpresado desde may-2018 supera el precio de venta reexpresado. En valores históricos había sido una ganancia de $27.525.000.' % m(-s('4.3.01')))
P('  - Alta: C11, adquirido el 15/06/2026 por $175.000.000 nominales.')
P('- **Vehículos ejecutivos (5):** E01 a E05. E01 y E02 están totalmente amortizados al cierre.')
P('- **Camionetas de reparto (2):** U01, totalmente amortizada, y U02.')
P('- **Amortizaciones del ejercicio:** $%s en 2026 y $%s en 2025. Su imputación por función está en el Anexo IV.' % (m(sum(g['amort'] for g in G26.values())), m(sum(g['amort'] for g in G25.values()))))
P('- **Dato faltante — valor recuperable:** el valor residual reexpresado ($%s) multiplica por %s el valor residual histórico ($%s).'
  % (m(c['bu']), ('{:.1f}'.format(c['bu'] / bu_h)).replace('.', ','), m(bu_h)))
P('  - Los inmuebles se reexpresan con coeficientes de ×153 (mar-2016) y ×58 (jul-2019).')
P('  - La RT 54 exige comparar con el valor recuperable cuando hay indicios de deterioro. No se recibieron tasaciones ni proyecciones, por lo que no se pudo hacer la prueba.')
P('  - Si el valor recuperable fuera menor, correspondería una desvalorización con cargo a resultados.')
P('')
P('### 7.3 Deudas comerciales')
P('')
tab('Concepto', [('Proveedores', c['prov'], c['prov25']), ('**Total corriente**', c['prov'], c['prov25'])])
P('### 7.4 Remuneraciones y cargas sociales')
P('')
tab('Concepto', [('Sueldos y cargas sociales a pagar', c['sue'], c['sue25']), ('**Total corriente**', c['sue'], c['sue25'])])
P('### 7.5 Cargas fiscales')
P('')
tab('Concepto', [('Impuesto a las ganancias a pagar (Nota 8)', c['ig'], c['ig25']), ('Impuesto sobre los ingresos brutos a pagar', c['iibb'], c['iibb25']),
                 ('**Total corriente**', c['ig'] + c['iibb'], c['ig25'] + c['iibb25'])])
P('### 7.6 Préstamos')
P('')
tab('Concepto', [('Préstamo bancario — porción corriente', c['prest_c'], c['prest_c25']), ('Préstamo bancario — porción no corriente', c['prest_nc'], c['prest_nc25']),
                 ('**Total**', c['prest'], c['prest25'])],
    '- Préstamo bancario a tasa fija. En 2026 se amortizaron $150.000.000 de capital, en cuotas trimestrales de $37.500.000. Saldo nominal al 31/12/2026: $300.000.000.\n'
    '- Intereses devengados: $%s en 2026 y $%s en 2025, en moneda de dic-2026.\n'
    '- **Supuesto (S):** la porción corriente es de $150.000.000 nominales por año. Al 31/12/2026: $150.000.000 vencen en 2027 y $150.000.000 en 2028.\n'
    '- **Dato faltante:** cronograma contractual, tasa y garantías.' % (m(s('5.4.01')), m(r25('5.4.01'))))
P('## Nota 8 — Impuesto a las ganancias')
P('')
P('**Conclusión:** el cargo del ejercicio ($%s) no equivale al 35 %% del resultado contable ajustado. ' % m(ig26)
  + 'El impuesto se determinó sobre el resultado histórico, con el supuesto simplificador del insumo. La diferencia es el efecto del ajuste por inflación contable sobre la base del cálculo.')
P('')
P('| Concepto | 2026 | 2025 |')
P('|---|---:|---:|')
P('| Resultado antes del impuesto (moneda dic-2026) | %s | %s |' % (m(rai26), m(rai25)))
P('| Tasa | 35 % | 35 % |')
P('| Impuesto a la tasa sobre el resultado contable ajustado | %s | %s |' % (m(rai26 * TASA), m(rai25 * TASA)))
P('| Efecto del ajuste por inflación contable no reconocido en la base del impuesto (resultado histórico − ajustado, × 35 %%) | %s | %s |'
  % (m(ig26 - rai26 * TASA), m(ig25 - rai25 * TASA)))
P('| **Cargo por impuesto a las ganancias (estado de resultados)** | **%s** | **%s** |' % (m(ig26), m(ig25)))
P('| Tasa efectiva sobre el resultado ajustado | %s | %s |' % (pct(ig26 / rai26, 1), pct(ig25 / rai25, 1)))
P('')
P('Control:')
P('- Resultado histórico antes del impuesto 2026: $%s × 35 %% = $%s, igual al cargo.' % (m(rai26_h), m(rai26_h * TASA)))
P('- Resultado histórico antes del impuesto 2025: $%s × 35 %% = $%s en moneda dic-2025, que reexpresado da $%s.' % (m(rai25_h), m(rai25_h * TASA), m(rai25_h * TASA * K)))
P('')
P('Composición del pasivo al 31/12/2026: impuesto determinado $%s, expuesto en Cargas fiscales. Los anticipos pagados en 2026 ($%s) están en Otros créditos (ver la Duda de la Nota 6).'
  % (m(c['ig']), m(c['anti'])))
P('')
P('**Dudas y limitaciones (requieren validación profesional):**')
P('1. **Impuesto diferido.** La RT 54 requiere, como regla general, el método del impuesto diferido. La reexpresión de bienes de uso, bienes de cambio y seguros genera diferencias temporarias con sus valores fiscales.')
P('   - Esas diferencias darían lugar a un pasivo por impuesto diferido que no se reconoce, por el supuesto simplificador del insumo.')
P('   - Hay que confirmar si el ente puede usar alguna opción de la RT 54 para entes pequeños o medianos. Si no puede, hay que calcular el diferido.')
P('2. **Tasa.** El 35 % es un supuesto del insumo. La ley del impuesto (Ley 27.630) establece una escala de alícuotas del 25 %, 30 % y 35 % por tramos de ganancia neta imponible, con montos que se actualizan cada año. Hay que verificar los tramos vigentes para 2026.')
P('3. **Ajuste por inflación impositivo** (Título VI de la ley del impuesto). No se modeló. El impuesto real puede diferir del cargo expuesto.')
P('')
P('## Nota 9 — Patrimonio neto, reserva legal y restricciones a la distribución de resultados')
P('')
P('### 9.1 Capital')
P('- Capital suscripto e integrado: $20.000.000 nominales, integrado el 14/03/2014.')
P('- Ajuste de capital: $%s. Total reexpresado: $%s.' % (m(-s('3.1.02')), m(-s('3.1.01') - s('3.1.02'))))
P('- **Dato faltante:** clases de acciones, valor nominal y votos (datos adicionales 2026).')
P('')
P('### 9.2 Distribución aprobada en el ejercicio')
P('- La asamblea ordinaria del 28/04/2026 aprobó dividendos en efectivo por $300.000.000 nominales sobre el resultado 2025. Se pagaron el 30/04/2026.')
P('- Reexpresados a moneda de dic-2026 suman $%s. La asamblea no asignó importes a la reserva legal.' % m(div26))
P('')
P('### 9.3 Reserva legal (art. 70 LGS)')
P('- El art. 70 de la LGS obliga a destinar a la reserva legal al menos el 5 % de las ganancias realizadas y líquidas de cada ejercicio, hasta alcanzar el 20 % del capital social.')
P('- En valores nominales, la reserva ($4.000.000) alcanzó en 04/2022 el 20 % del capital suscripto nominal ($20.000.000).')
P('- **Duda (societaria):** no está definido si la base del 20 % debe incluir el ajuste de capital.')
P('  - Si la base es capital + ajuste de capital, en moneda homogénea la reserva queda por debajo del tope.')
P('  - Esa es la interpretación propuesta en la fase 2, que siguen, entre otros, el CPCE de Córdoba y la IPJ de Córdoba (RG 95/T).')
P('  - Para CABA hay que validar el criterio de la IGJ con un asesor societario. La RG IGJ 7/2015 fue reemplazada por la RG IGJ 15/2024: verificar qué dice la norma vigente.')
P('')
P('| Concepto (moneda dic-2026) | Importe |')
P('|---|---:|')
P('| Base: capital suscripto + ajuste de capital | %s |' % m(rl['base']))
P('| Tope: 20 %% de la base | %s |' % m(rl['tope']))
P('| Reserva legal reexpresada al 31/12/2026 | %s |' % m(rl['reserva']))
P('| Faltante para alcanzar el tope | %s |' % m(rl['faltante']))
P('| 5 %% del resultado 2025 ajustado, no asignado por la asamblea del 28/04/2026 | %s |' % m(rl['omit25']))
P('| 5 %% del resultado 2026 ajustado | %s |' % m(rl['max26']))
P('| **Asignación mínima propuesta a la asamblea que trate los EECC 2026** | **%s** |' % m(rl['prop']))
P('')
P('- **No se contabilizó ninguna asignación.** Asignar la reserva es una decisión de la asamblea. Es una reclasificación dentro del PN y no cambia su total.')
P('- **Restricción a la distribución:** con el criterio propuesto, de los resultados no asignados al 31/12/2026 ($%s) deben asignarse a reserva legal al menos $%s antes de distribuir dividendos. Esa cifra es el mínimo legal; el art. 70 no impide asignar más.'
  % (m(rna26), m(rl['prop'])))
P('- **Riesgo a validar:** la distribución de $300.000.000 de dividendos sin asignar el 5 % del resultado 2025 ajustado. Si se adopta la base con ajuste, la asamblea podría tener que regularizarlo.')
P('- Los dividendos solo pueden distribuirse de ganancias realizadas y líquidas, según un balance confeccionado conforme a la ley y aprobado por la asamblea (art. 68 LGS).')
P('')
P('## Nota 10 — Seguros contratados')
P('')
P('- Al 31/12/2026 hay 10 pólizas vigentes con Aseguradora Ficticia S.A. Cubren automotores, incendio, robo, responsabilidad civil y riesgos de la naturaleza. El detalle de bienes, coberturas y vigencias está en el Anexo V.')
P('- Las primas renovadas en 2026 suman $%s nominales ($%s en moneda dic-2026).' % (m(sum(p['prima_h'] for p in POL.values())), m(sum(p['prima_aj'] for p in POL.values()))))
P('- Gasto de seguros 2026: $%s. Comercialización $%s; administración $%s.' % (m(s('5.2.06') + s('5.3.05')), m(s('5.2.06')), m(s('5.3.05'))))
P('- Parte no devengada al cierre: $%s (Nota 6).' % m(c['seg']))
P('- **Dato faltante:** sumas aseguradas y franquicias. Sin esos datos no se puede evaluar si la cobertura es suficiente frente al valor de los bienes.')
P('')
P('## Nota 11 — Hechos posteriores al cierre')
P('')
P('- Entre el 31/12/2026 y la fecha de emisión no hay hechos posteriores informados que requieran ajustar o revelar estos estados contables.')
P('- **Dato faltante:** fecha de emisión y confirmación de la gerencia a esa fecha.')
P('- **Hecho conocido después del cierre que afecta las estimaciones:** la publicación de los índices oficiales de noviembre (16/12/2026) y diciembre de 2026 (enero de 2027). '
  'Si ocurre antes de la emisión, corresponde recalcular las cifras en moneda de dic-2026 (Nota 1.3).')
P('')

# Tabla de control
ctrl_rows = [
    ('Caja y bancos', c['caja'] + c['banco'], esp['Caja y bancos'][0], c['caja25'] + c['banco25'], esp['Caja y bancos'][1]),
    ('Inversiones', c['pf'], esp['Inversiones'][0], c['pf25'], esp['Inversiones'][1]),
    ('Créditos por ventas', c['deud'] + c['prev'], esp['Créditos por ventas'][0], c['deud25'] + c['prev25'], esp['Créditos por ventas'][1]),
    ('Otros créditos', c['anti'] + c['iva'] + c['seg'], esp['Otros créditos'][0], c['anti25'] + c['iva25'] + c['seg25'], esp['Otros créditos'][1]),
    ('Bienes de cambio', c['merc'], esp['Bienes de cambio'][0], c['merc25'], esp['Bienes de cambio'][1]),
    ('Bienes de uso (Anexo I)', c['bu'], esp['Bienes de uso'][0], c['bu25'], esp['Bienes de uso'][1]),
    ('Deudas comerciales', c['prov'], -esp['Deudas comerciales'][0], c['prov25'], -esp['Deudas comerciales'][1]),
    ('Remuneraciones y cargas sociales', c['sue'], -esp['Remuneraciones y cargas sociales'][0], c['sue25'], -esp['Remuneraciones y cargas sociales'][1]),
    ('Cargas fiscales', c['ig'] + c['iibb'], -esp['Cargas fiscales'][0], c['ig25'] + c['iibb25'], -esp['Cargas fiscales'][1]),
    ('Préstamos (corriente + no corriente)', c['prest_c'] + c['prest_nc'], -esp['Préstamos'][0], c['prest_c25'] + c['prest_nc25'], -esp['Préstamos'][1]),
]
act_n26 = sum(r[1] for r in ctrl_rows[:6]); act_n25 = sum(r[3] for r in ctrl_rows[:6])
pas_n26 = sum(r[1] for r in ctrl_rows[6:]); pas_n25 = sum(r[3] for r in ctrl_rows[6:])
ctrl_rows += [
    ('**Total activo**', act_n26, tot_act26, act_n25, tot_act25),
    ('**Total pasivo**', pas_n26, tot_pas26, pas_n25, tot_pas25),
    ('**Patrimonio neto (activo − pasivo)**', act_n26 - pas_n26, PN26['Total patrimonio neto al 31/12/2026'], act_n25 - pas_n25, PN26['Total patrimonio neto al 31/12/2025 (en moneda de dic-2026)']),
]
er_rows = [
    ('Costo de mercaderías vendidas (Anexo III)', cmv['cmv26'], s('5.1.01'), cmv['cmv25'], -ER25C['Costo de mercaderías vendidas']),
    ('Gastos de comercialización (Anexo IV)', tot_com, sum(v['fin'] for v in B26.values() if v['rubro'] == 'Gastos de comercialización'), tot_com25, -ER25C['Gastos de comercialización']),
    ('Gastos de administración (Anexo IV)', tot_adm, sum(v['fin'] for v in B26.values() if v['rubro'] == 'Gastos de administración'), tot_adm25, -ER25C['Gastos de administración']),
    ('Intereses del préstamo (Anexo IV)', tot_fin, s('5.4.01'), tot_fin25, -ER25C['Intereses préstamo bancario']),
    ('Cargo por incobrables = aumentos de la previsión (Anexo II)', prev['aum26'], s('5.2.05'), prev['aum25'], r25('5.2.05')),
    ('Amortizaciones (Anexo I) = cuentas de amortización', sum(g['amort'] for g in G26.values()), sum(s(x) for x in ['5.2.07', '5.2.08', '5.2.09', '5.3.06', '5.3.07']),
     sum(g['amort'] for g in G25.values()), sum(r25(x) for x in ['5.2.07', '5.2.08', '5.2.09', '5.3.06', '5.3.07'])),
    ('Impuesto a las ganancias (Nota 8)', ig26, s('5.5.01'), ig25, -ER25C['Impuesto a las ganancias']),
    ('**Resultado del ejercicio**', res26, 496161444, res25, 611508366),
]
otros = [
    ('Compras históricas 2026 (cifra de control del encargo eecc: $13.115.000.000)', cmv['co26_h'], 13115000000),
    ('Anexo V: gasto de seguros 2026 = 5.2.06 + 5.3.05', sum(p['gasto_aj'] for p in POL.values()), s('5.2.06') + s('5.3.05')),
    ('Anexo V: no devengado 31/12/2026 = 1.4.03', sum(p['nodev_aj'] for p in POL.values()), s('1.4.03')),
    ('Previsión: saldo final 2026 (Anexo II) = 1.3.02', prev['fin26'], -s('1.3.02')),
    ('Nota 8: impuesto a la tasa + efecto ajuste = cargo 2026', rai26 * TASA + (ig26 - rai26 * TASA), ig26),
    ('Nota 8: resultado histórico antes de IG 2026 × 35 % = cargo', rai26_h * TASA, ig26),
    ('Nota 9: RNA 31/12/2026 = RNA inicio − dividendos + resultado', PN26['RNA al 31/12/2025 (aj.): cierre 2025 ajustado (incluye resultado 2025) × coef. dic-25 → dic-26'] - div26 + res26, rna26),
]
P('## Tabla de control — notas y anexos contra los BSS ajustados')
P('')
P('Fuente de la columna BSS: `entregables/ajuste/ajuste_2026.xlsx`, hoja BSS ajustado. El 31/12/2025 sale del saldo inicial en moneda dic-2026. '
  'El estado de resultados 2025 sale de `comparativo_2025_en_moneda_dic2026.xlsx`. Diferencia aceptada: menor a $1.')
P('')
P('### Estado de situación patrimonial')
P('')
P('| Rubro | Nota 31/12/2026 | BSS 31/12/2026 | Dif. | Nota 31/12/2025 | BSS 31/12/2025 | Dif. | Resultado |')
P('|---|---:|---:|---:|---:|---:|---:|---|')
CTRL_OK = True
CTRL_LIST = []


def fila_ctrl(t, a, b, a2, b2):
    global CTRL_OK
    ok = abs(a - b) < 1 and abs(a2 - b2) < 1
    CTRL_OK &= ok
    CTRL_LIST.append(dict(control=t.replace('*', ''), nota_2026=a, bss_2026=b, nota_2025=a2, bss_2025=b2, ok=ok))
    P('| %s | %s | %s | %s | %s | %s | %s | %s |' % (t, m(a), m(b), m(a - b), m(a2), m(b2), m(a2 - b2), 'OK' if ok else 'ERROR'))


for r_ in ctrl_rows:
    fila_ctrl(*r_)
P('')
P('### Estado de resultados (egresos con signo positivo; resultado del ejercicio: ganancia positiva)')
P('')
P('| Concepto | Nota / anexo 2026 | BSS 2026 | Dif. | Nota / anexo 2025 | Comparativo 2025 | Dif. | Resultado |')
P('|---|---:|---:|---:|---:|---:|---:|---|')
for r_ in er_rows:
    fila_ctrl(*r_)
P('')
P('### Otros controles')
P('')
P('| Control | Nota / anexo | Fuente | Dif. | Resultado |')
P('|---|---:|---:|---:|---|')
for t, a, b in otros:
    ok = abs(a - b) < 1
    CTRL_OK &= ok
    CTRL_LIST.append(dict(control=t, nota_2026=a, bss_2026=b, ok=ok))
    P('| %s | %s | %s | %s | %s |' % (t, m(a), m(b), m(a - b), 'OK' if ok else 'ERROR'))
P('')
P('**Resultado global de los controles: %s.** La hoja Controles de `anexos_andina_2026.xlsx` repite los cruces con fórmulas.' % ('OK' if CTRL_OK else 'ERROR'))
P('')
P('**Cruce con `eecc-2026`: pendiente.** La sesión eecc-2026 trabaja en paralelo y su entregable no estaba integrado al cerrar estas notas. Lo hace la reina con esta misma tabla.')
P('')
P('## Supuestos, Dudas y datos faltantes (resumen)')
P('')
P('| Tipo | Tema | Nota |')
P('|---|---|---|')
for t in [
    ('Estimación', 'Índices de nov y dic-2026 no oficiales: afectan todas las cifras en moneda de dic-2026', '1.3'),
    ('Duda', 'Número de la resolución de adopción de la RT 54 / RT 59 en el CPCECABA y párrafos de la RT 54 sobre unidad de medida', '1.1'),
    ('Duda', 'Base de la reserva legal (art. 70 LGS) con o sin ajuste de capital: faltante de $%s y asignación mínima propuesta de $%s' % (m(rl['faltante']), m(rl['prop'])), '9.3'),
    ('Duda', 'Impuesto diferido no reconocido (supuesto del insumo) y escala de alícuotas de la Ley 27.630', '8'),
    ('Duda', 'Compensación de los anticipos de IG con la deuda del impuesto (criterio a unificar con eecc-2026)', '6'),
    ('Dato faltante', 'Prueba de valor recuperable de bienes de uso y de VNR de bienes de cambio', '2, 7.1, 7.2'),
    ('Dato faltante', 'Fechas de aplicación de la previsión para incobrables (supuesto lineal mensual)', 'Anexo II'),
    ('Dato faltante', 'Cronograma, tasa y garantías del préstamo', '7.6'),
    ('Dato faltante', 'Sumas aseguradas y franquicias', '10'),
    ('Dato faltante', 'Antigüedad de los créditos; acciones (clases y votos); fecha de emisión y hechos posteriores', '5, 9.1, 11'),
]:
    P('| %s | %s | %s |' % t)
P('')
P('---')
P('Generado con `entregables/notas-anexos-2026/generador_notas_anexos.py` y verificado con `verificar_notas_anexos.py` (Python + LibreOffice).')
MD = os.path.join(OUT, 'notas_andina_2026.md')
with open(MD, 'w', encoding='utf-8') as f:
    f.write('\n'.join(N) + '\n')

json.dump(dict(ok=CTRL_OK, controles=CTRL_LIST,
               totales=dict(activo_2026=tot_act26, pasivo_2026=tot_pas26, pn_2026=pn26, activo_2025=tot_act25, pasivo_2025=tot_pas25,
                            pn_2025=pn25, resultado_2026=res26, resultado_2025=res25, bienes_uso_2026=c['bu'], bienes_uso_2025=c['bu25'],
                            cmv_2026=cmv['cmv26'], cmv_2025=cmv['cmv25'], gastos_com_2026=tot_com, gastos_adm_2026=tot_adm,
                            ig_2026=ig26, recpam_2026=recpam26, reserva_faltante=rl['faltante'], reserva_propuesta=rl['prop'])),
          open(os.path.join(OUT, 'control_notas_anexos.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('Controles de las notas:', 'OK' if CTRL_OK else 'ERROR')
for x in CTRL_LIST:
    if not x['ok']:
        print('ERROR', x)
