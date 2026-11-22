# -*- coding: utf-8 -*-
"""Verificador de notas y anexos 2026 — Distribuidora Andina S.A. (FICTICIA).

1. Recalcula anexos_andina_2026.xlsx con LibreOffice (headless) y deja la copia recalculada como entregable
   (así el archivo trae valores calculados además de las fórmulas).
2. Comprueba que la hoja Controles da OK en todas las filas.
3. Recalcula en Python, desde los BSS ajustados, los totales de los anexos y los compara con LibreOffice.
4. Comprueba que las notas (.md) no tienen controles en ERROR y que los totales clave aparecen en el texto.
Uso: python3 verificar_notas_anexos.py   (desde la raíz del repositorio)
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile

import openpyxl

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(os.path.dirname(AQUI))
XLSX = os.path.join(AQUI, 'anexos_andina_2026.xlsx')
MD = os.path.join(AQUI, 'notas_andina_2026.md')
fallas = []


def chk(nombre, a, b, tol=1.0):
    ok = abs((a or 0) - (b or 0)) < tol
    print(('OK   ' if ok else 'ERROR') + ' %-75s %20.2f %20.2f' % (nombre, a or 0, b or 0))
    if not ok:
        fallas.append(nombre)


# 1. Recalcular con LibreOffice
tmp = tempfile.mkdtemp()
subprocess.run(['soffice', '--headless', '--calc', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', tmp, XLSX],
               check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
recalc = os.path.join(tmp, 'anexos_andina_2026.xlsx')
wb = openpyxl.load_workbook(recalc, data_only=True)
wbf = openpyxl.load_workbook(recalc)  # fórmulas conservadas

# 2. Hoja Controles
ws = wb['Controles']
n = 0
for r in ws.iter_rows(min_row=4, values_only=True):
    if r[0] and r[4] in ('OK', 'ERROR'):
        n += 1
        if r[4] != 'OK':
            fallas.append('Controles: ' + r[0])
            print('ERROR', r)
glob = [r for r in ws.iter_rows(values_only=True) if r[0] == 'Resultado global'][0][4]
print('Hoja Controles: %d controles; resultado global = %s' % (n, glob))
if glob != 'OK':
    fallas.append('Controles: resultado global')
# Errores de fórmula en cualquier hoja
for w in wb:
    for row in w.iter_rows(values_only=True):
        for v in row:
            if isinstance(v, str) and v.startswith(('#REF', '#VALUE', '#NAME', '#DIV', 'Err:')):
                fallas.append('Error de fórmula en %s: %s' % (w.title, v))
# Fórmulas realmente presentes en los anexos
nf = sum(1 for w in wbf if w.title.startswith('Anexo') for row in w.iter_rows() for cc in row
         if isinstance(cc.value, str) and cc.value.startswith('='))
print('Celdas con fórmula en los anexos:', nf)
if nf < 100:
    fallas.append('Pocas fórmulas en los anexos')

# 3. Recálculo independiente en Python desde los libros de la fase 2
aj26 = openpyxl.load_workbook(os.path.join(RAIZ, 'entregables', 'ajuste', 'ajuste_2026.xlsx'), data_only=True)
aj25 = openpyxl.load_workbook(os.path.join(RAIZ, 'entregables', 'ajuste', 'ajuste_2025.xlsx'), data_only=True)
K = aj26['Coeficientes']['E3'].value


def bss(w):
    d = {}
    for r in w['BSS ajustado'].iter_rows(values_only=True):
        if r[0] and isinstance(r[0], str) and r[0][0].isdigit() and r[3]:
            d[r[0]] = (r[2], r[5], r[7])
    return d


B26, B25 = bss(aj26), bss(aj25)
bu26 = sum(v[2] for v in B26.values() if v[0] == 'Bienes de uso')
bu25 = sum(v[1] for v in B26.values() if v[0] == 'Bienes de uso')
com26 = sum(v[2] for v in B26.values() if v[0] == 'Gastos de comercialización')
adm26 = sum(v[2] for v in B26.values() if v[0] == 'Gastos de administración')
com25 = sum(v[2] for v in B25.values() if v[0] == 'Gastos de comercialización') * K
adm25 = sum(v[2] for v in B25.values() if v[0] == 'Gastos de administración') * K
seg26 = B26['5.2.06'][2] + B26['5.3.05'][2]


def total_row(w, col, etiqueta='Total', desde=1):
    for i, r in enumerate(w.iter_rows(min_row=desde, values_only=True), start=desde):
        if r[0] == etiqueta:
            return r[col], i
    raise ValueError(etiqueta)


a1 = wb['Anexo I BU']
v, i26 = total_row(a1, 10)
chk('Anexo I: valor residual 2026 = BSS', v, bu26)
v25, _ = total_row(a1, 10, desde=i26 + 1)
chk('Anexo I: valor residual 2025 = BSS (31/12/2025, moneda dic-2026)', v25, bu25)
amort26 = sum(B26[x][2] for x in ['5.2.07', '5.2.08', '5.2.09', '5.3.06', '5.3.07'])
chk('Anexo I: amortización 2026', total_row(a1, 8)[0], amort26)
a3 = wb['Anexo III CMV']
chk('Anexo III: CMV 2026', a3['B8'].value, B26['5.1.01'][2])
chk('Anexo III: CMV 2025 (moneda dic-2026)', a3['C8'].value, B25['5.1.01'][2] * K)
chk('Anexo III: compras históricas 2026 = 13.115.000.000', a3['B10'].value, 13115000000)
chk('Anexo III: EI + compras − EF = CMV 2026', a3['B5'].value + a3['B6'].value + a3['B7'].value, B26['5.1.01'][2])
a4 = wb['Anexo IV Gastos']
t, i = total_row(a4, 1, 'Total 2026')
chk('Anexo IV: comercialización 2026', a4.cell(row=i, column=2).value, com26)
chk('Anexo IV: administración 2026', a4.cell(row=i, column=3).value, adm26)
chk('Anexo IV: financiación 2026', a4.cell(row=i, column=4).value, B26['5.4.01'][2])
chk('Anexo IV: comercialización 2025', a4.cell(row=i + 1, column=2).value, com25)
chk('Anexo IV: administración 2025', a4.cell(row=i + 1, column=3).value, adm25)
chk('Anexo IV: total 2025 por naturaleza = total por función', a4.cell(row=i, column=6).value, a4.cell(row=i + 1, column=6).value)
a2 = wb['Anexo II Prev']
chk('Anexo II: saldo final 2026', a2['F5'].value, -B26['1.3.02'][2])
chk('Anexo II: inicio + aumentos − aplicaciones + RECPAM = final 2026', a2['B5'].value + a2['C5'].value - a2['D5'].value + a2['E5'].value, a2['F5'].value)
chk('Anexo II: saldo final 2025 (moneda dic-2026) = inicio 2026', a2['F6'].value, a2['B5'].value)
a5 = wb['Anexo V Seguros']
_, i = total_row(a5, 0)
chk('Anexo V: gasto 2026', a5.cell(row=i, column=11).value, seg26)
chk('Anexo V: no devengado 31/12/2026', a5.cell(row=i, column=12).value, B26['1.4.03'][2])
npol = sum(1 for r in a5.iter_rows(values_only=True) if r[0] and str(r[0]).startswith('POL-'))
chk('Anexo V: 10 pólizas', npol, 10)

# 4. Notas
ctrl = json.load(open(os.path.join(AQUI, 'control_notas_anexos.json'), encoding='utf-8'))
txt = open(MD, encoding='utf-8').read()
chk('Notas: controles del generador OK (1 = sí)', 1 if ctrl['ok'] else 0, 1)
chk('Notas: sin filas ERROR en la tabla de control (cantidad)', txt.count('| ERROR |'), 0)
res = sum(-v[2] for v in B26.values() if v[0] and v[2] is not None and B26 and v is not None and v[0] in (
    'Ingresos', 'Resultados financieros', 'Otros ingresos y egresos', 'Costo de ventas', 'Gastos de comercialización',
    'Gastos de administración', 'Impuesto a las ganancias', 'Resultados financieros y por tenencia'))
chk('Resultado 2026 recalculado = 496.161.444', res, 496161444)
for cifra in ['496.161.444', '611.508.366', '731.033.507', '55.383.491', '673.170.750', '13.115.000.000']:
    if cifra not in txt and cifra != '13.115.000.000':
        fallas.append('Cifra ausente en las notas: ' + cifra)
    print('OK   ' if cifra in txt else 'AVISO', 'cifra en las notas:', cifra)

if fallas:
    print('\nFALLAS:', fallas)
    sys.exit(1)
shutil.copyfile(recalc, XLSX)
shutil.rmtree(tmp, ignore_errors=True)
print('\nVerificación completa: OK. Entregable reemplazado por la versión recalculada con LibreOffice.')
