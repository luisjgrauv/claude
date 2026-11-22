"""Verifica eecc_andina_2026.xlsx.

1. Recalcula el libro con LibreOffice (headless) en un directorio temporal.
2. Lee la hoja Controles: todos deben dar OK.
3. Recalcula en Python, sin fórmulas, desde los libros de la fase 2: totales del ESP,
   resultado y EFE. Compara con lo que muestra el libro recalculado.
Uso: python3 verificar_eecc.py
"""
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl

BASE = Path(__file__).resolve().parent
AJ = BASE.parent / "ajuste"
LIBRO = BASE / "eecc_andina_2026.xlsx"
TOL = 0.5

tmp = Path(tempfile.mkdtemp())
subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(tmp), str(LIBRO)],
               check=True, capture_output=True)
wb = openpyxl.load_workbook(tmp / LIBRO.name, data_only=True)

fallas = []
wc = wb["Controles"]
n = 0
for r in wc.iter_rows(min_row=5, values_only=True):
    if r[0] == "Resultado global":
        if r[3] != "OK":
            fallas.append("Resultado global de Controles distinto de OK")
        break
    n += 1
    if r[3] != "OK" or abs(r[1] or 0) >= TOL:
        fallas.append(f"Control: {r[0]} = {r[1]}")
print(f"Controles del libro: {n} revisados")

# ---- recálculo independiente
a26 = openpyxl.load_workbook(AJ / "ajuste_2026.xlsx", data_only=True)
bss = {}
for r in a26["BSS ajustado"].iter_rows(min_row=4, values_only=True):
    if not r[0]:
        break
    bss[r[0]] = (r[2], r[3], r[5], r[7])  # rubro, clase, ini, fin


def rub(nombre, k):
    return sum(v[k] for v in bss.values() if v[0] == nombre)


ini, fin = 2, 3
anti = bss["1.4.01"]
activo26 = sum(v[fin] for c, v in bss.items() if v[1].startswith("A")) - anti[fin]
pasivo26 = -sum(v[fin] for v in bss.values() if v[1].startswith("P-")) - anti[fin]
pn26 = activo26 - pasivo26
res26 = -sum(v[fin] for v in bss.values() if v[1] == "R")
efe_var = sum(bss[c][fin] - bss[c][ini] for c in ("1.1.01", "1.1.02", "1.2.01"))

esp = {r[0]: r[1] for r in wb["ESP"].iter_rows(min_row=5, values_only=True) if r[0]}
er = {r[0]: r[1] for r in wb["ER"].iter_rows(min_row=5, values_only=True) if r[0]}
efe = {r[0]: r for r in wb["EFE"].iter_rows(min_row=5, values_only=True) if r[0]}


def cmp(nombre, a, b):
    ok = abs(a - b) < TOL
    print(f"{'OK ' if ok else 'ERR'} {nombre}: libro {a:,.2f} | Python {b:,.2f}")
    if not ok:
        fallas.append(nombre)


cmp("Total activo 31/12/2026", esp["TOTAL DEL ACTIVO"], activo26)
cmp("Total pasivo 31/12/2026", esp["TOTAL DEL PASIVO"], pasivo26)
cmp("PN 31/12/2026", esp["PATRIMONIO NETO (según estado respectivo)"], pn26)
cmp("Resultado 2026", er["RESULTADO DEL EJERCICIO — GANANCIA"], res26)
cmp("Resultado 2026 vs cifra de la reina (496.161.444)", er["RESULTADO DEL EJERCICIO — GANANCIA"], 496161444.443603)
cmp("PN 31/12/2026 vs cifra de la reina (8.026.382.100)", esp["PATRIMONIO NETO (según estado respectivo)"], 8026382099.55735)
cmp("Variación de efectivo 2026 (EFE)", efe["(Disminución) aumento neto del efectivo"][1], efe_var)
cmp("Préstamo corriente 2026", esp["Préstamos (porción corriente)"], 150000000)
cmp("Préstamo no corriente 2026", esp["Préstamos (porción no corriente)"], 150000000)

# EFE 2026 por causas, recalculado
coef_jun = a26["Bienes de uso"]["I19"].value
recm = [(r[11], r[12]) for r in a26["RECPAM"].iter_rows(min_row=17, max_row=28, values_only=True)]
k = a26["Controles"]["B32"].value
c0 = sum(bss[c][ini] for c in ("1.1.01", "1.1.02", "1.2.01")) / k
c1 = sum(bss[c][fin] for c in ("1.1.01", "1.1.02", "1.2.01"))
rec_ef = sum(-(c0 + (c1 - c0) * i / 12) * inf * cf for i, (inf, cf) in enumerate(recm))
prest = 37.5e6 * (recm[2][1] + recm[5][1] + recm[8][1] + recm[11][1])
g = -bss["2.4.01"][ini] + bss["2.4.01"][fin] - prest
div = bss["3.3.01"][fin] - bss["3.3.01"][ini]  # dividendos (deudor en RNA => +)
amort = sum(bss[c][fin] for c in ("5.2.07", "5.2.08", "5.2.09", "5.3.06", "5.3.07"))
wc_op = 0
for nombre in ("Créditos por ventas", "Otros créditos", "Bienes de cambio"):
    wc_op += rub(nombre, ini) - rub(nombre, fin)
for nombre in ("Deudas comerciales", "Remuneraciones y cargas sociales", "Cargas fiscales"):
    wc_op += -rub(nombre, fin) + rub(nombre, ini)
oper = res26 + amort + bss["4.3.01"][fin] - rec_ef - g + wc_op
inv = -175e6 * coef_jun + 28e6 * coef_jun
fina = -prest - div
cmp("EFE 2026 operativas", efe["Flujo neto de efectivo generado por las actividades operativas"][1], oper)
cmp("EFE 2026 inversión", efe["Flujo neto de efectivo utilizado en las actividades de inversión"][1], inv)
cmp("EFE 2026 financiación", efe["Flujo neto de efectivo utilizado en las actividades de financiación"][1], fina)
cmp("EFE 2026 suma de causas = variación", oper + inv + fina + rec_ef, efe_var)

print("RESULTADO:", "OK" if not fallas else f"ERROR ({len(fallas)})")
for f in fallas:
    print("  -", f)
sys.exit(1 if fallas else 0)
