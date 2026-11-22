"""Genera entregables/eecc-2026/eecc_andina_2026.xlsx.

Estados contables básicos de Distribuidora Andina S.A. (FICTICIA) al 31/12/2026,
comparativos con 2025, en moneda de dic-2026.

Insumos (valores importados, recalculados por la fase 2):
- entregables/ajuste/ajuste_2026.xlsx: BSS ajustado, RECPAM (inflación y coeficientes mensuales).
- entregables/ajuste/ajuste_2025.xlsx: RECPAM (inflación y coeficientes mensuales 2025).
- entregables/ajuste/comparativo_2025_en_moneda_dic2026.xlsx: datos 2025 y estados de control.

Uso: python3 generador_eecc.py
"""
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
AJ = BASE.parent / "ajuste"
OUT = BASE / "eecc_andina_2026.xlsx"

NUM = '#,##0;(#,##0);"-"'
COEF = "0.000000"
PCT = "0.0000%"
AZUL = Font(color="0000CC")
NEG = Font(bold=True)
TIT = Font(bold=True, size=12)
GRIS = PatternFill("solid", fgColor="EEEEEE")
LINEA = Border(top=Side(style="thin"))
EMPRESA = "Distribuidora Andina S.A. (FICTICIA)"

# ---------------------------------------------------------------- insumos
a26 = openpyxl.load_workbook(AJ / "ajuste_2026.xlsx", data_only=True)
a25 = openpyxl.load_workbook(AJ / "ajuste_2025.xlsx", data_only=True)
comp = openpyxl.load_workbook(AJ / "comparativo_2025_en_moneda_dic2026.xlsx", data_only=True)


def filas_bss(ws, desde, cols):
    out = []
    for r in ws.iter_rows(min_row=desde, values_only=True):
        if not r[0]:
            break
        out.append([r[c] for c in cols])
    return out


bss26 = filas_bss(a26["BSS ajustado"], 4, [0, 1, 2, 3, 5, 7])  # cód, cuenta, rubro, clase, ini aj, fin aj
bss25 = filas_bss(comp["Datos 2025 (dic-2025)"], 5, [0, 1, 2, 5, 6])  # cód, cuenta, rubro, 31/12/24, 31/12/25 (dic-26)


def recpam_mensual(ws):
    out = []
    for r in ws.iter_rows(min_row=17, max_row=28, values_only=True):
        out.append((r[0], r[11], r[12]))  # mes, inflación, coef mes -> cierre
    return out


rec26 = recpam_mensual(a26["RECPAM"])
rec25 = recpam_mensual(a25["RECPAM"])
coef_25_26 = a26["Controles"]["B32"].value
coef_24_25 = a25["Coeficientes"]["E3"].value

bu = a26["Bienes de uso"]
coef_c11 = bu["I19"].value  # alta C11 jun-2026 -> dic-2026
coef_c01_baja = bu["AD9"].value  # baja C01 jun-2026 -> dic-2026
coef_c10 = bu["I18"].value  # alta C10 may-2025 -> dic-2026
coef_e05 = bu["I24"].value  # alta E05 sep-2025 -> dic-2026

# Cifras de control de la fase 2 (importadas)
f2 = {
    "res26": a26["Controles"]["B26"].value,
    "pn26": a26["Controles"]["B28"].value,
    "pn25": a26["Controles"]["B29"].value,
    "act26": a26["Controles"]["B30"].value,
    "res25": comp["Controles"]["B16"].value,
}
esp25 = {}
for r in comp["ESP 2025"].iter_rows(min_row=5, values_only=True):
    if r[0] and r[1] is not None:
        esp25[r[0]] = r[1]

wb = openpyxl.Workbook()
wb.remove(wb.active)


def hoja(nombre, titulo, sub=None, anchos=(58, 20, 20)):
    ws = wb.create_sheet(nombre)
    ws["A1"] = f"{EMPRESA} — {titulo}"
    ws["A1"].font = TIT
    if sub:
        ws["A2"] = sub
        ws["A2"].font = Font(italic=True, size=9)
    for i, w in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.sheet_view.showGridLines = False
    return ws


def put(ws, ref, v, fmt=None, font=None, fill=None, border=None):
    c = ws[ref]
    c.value = v
    if fmt:
        c.number_format = fmt
    if font:
        c.font = font
    if fill:
        c.fill = fill
    if border:
        c.border = border
    return c


# Orden de hojas: estados primero
for n in ["Supuestos", "Controles", "ESP", "ER", "EEPN", "EFE", "BSS 2026", "BSS 2025", "Datos EFE", "Fase 2"]:
    pass

# ---------------------------------------------------------------- BSS 2026
wsb = hoja("BSS 2026", "BSS ajustado — saldos al 31/12/2025 y al 31/12/2026 en moneda de dic-2026",
           "Valores importados de entregables/ajuste/ajuste_2026.xlsx, hoja 'BSS ajustado' (recalculada con LibreOffice). Deudor (+) / acreedor (−). Azul = insumo.",
           (10, 62, 34, 8, 20, 20))
hdr = ["Código", "Cuenta", "Rubro", "Clase", "Saldo 31/12/2025 (aj.)", "Saldo 31/12/2026 (aj.)"]
for j, h in enumerate(hdr, 1):
    put(wsb, f"{get_column_letter(j)}4", h, font=NEG, fill=GRIS)
r = 5
for fila in bss26:
    for j, v in enumerate(fila, 1):
        put(wsb, f"{get_column_letter(j)}{r}", v, NUM if j >= 5 else None, AZUL)
    r += 1
B26_LAST = r - 1
put(wsb, f"B{r}", "Total (debe ser 0)", font=NEG)
put(wsb, f"E{r}", f"=SUM(E5:E{B26_LAST})", NUM, NEG)
put(wsb, f"F{r}", f"=SUM(F5:F{B26_LAST})", NUM, NEG)
B26_TOT = r

# ---------------------------------------------------------------- BSS 2025
ws5 = hoja("BSS 2025", "Ejercicio 2025 ajustado — saldos al 31/12/2024 y al 31/12/2025 en moneda de dic-2026",
           "Valores importados de entregables/ajuste/comparativo_2025_en_moneda_dic2026.xlsx, hoja 'Datos 2025 (dic-2025)', columnas F y G. Deudor (+) / acreedor (−).",
           (10, 62, 34, 20, 20))
hdr = ["Código", "Cuenta", "Rubro", "Saldo 31/12/2024 (moneda dic-2026)", "Saldo 31/12/2025 (moneda dic-2026)"]
for j, h in enumerate(hdr, 1):
    put(ws5, f"{get_column_letter(j)}4", h, font=NEG, fill=GRIS)
r = 5
for fila in bss25:
    for j, v in enumerate(fila, 1):
        put(ws5, f"{get_column_letter(j)}{r}", v, NUM if j >= 4 else None, AZUL)
    r += 1
B25_LAST = r - 1
put(ws5, f"B{r}", "Total (debe ser 0)", font=NEG)
put(ws5, f"D{r}", f"=SUM(D5:D{B25_LAST})", NUM, NEG)
put(ws5, f"E{r}", f"=SUM(E5:E{B25_LAST})", NUM, NEG)
B25_TOT = r


# Helpers de fórmulas -------------------------------------------------
def s26(col, codigo=None, rubro=None):
    """col: 'ini' (31/12/2025) o 'fin' (31/12/2026)."""
    c = "E" if col == "ini" else "F"
    rng = f"'BSS 2026'!${c}$5:${c}${B26_LAST}"
    if codigo:
        return f"SUMIFS({rng},'BSS 2026'!$A$5:$A${B26_LAST},\"{codigo}\")"
    return f"SUMIFS({rng},'BSS 2026'!$C$5:$C${B26_LAST},\"{rubro}\")"


def s25(col, codigo=None, rubro=None):
    """col: 'ini' (31/12/2024) o 'fin' (31/12/2025)."""
    c = "D" if col == "ini" else "E"
    rng = f"'BSS 2025'!${c}$5:${c}${B25_LAST}"
    if codigo:
        return f"SUMIFS({rng},'BSS 2025'!$A$5:$A${B25_LAST},\"{codigo}\")"
    return f"SUMIFS({rng},'BSS 2025'!$C$5:$C${B25_LAST},\"{rubro}\")"


def suma(fn, col, codigos):
    return "+".join(fn(col, codigo=c) for c in codigos)


EFECTIVO = ["1.1.01", "1.1.02", "1.2.01"]
AMORT = ["5.2.07", "5.2.08", "5.2.09", "5.3.06", "5.3.07"]

# ---------------------------------------------------------------- Datos EFE
wd = hoja("Datos EFE", "Datos para el estado de flujo de efectivo (moneda de dic-2026)",
          "Azul = insumo (con su fuente). Negro = fórmula.", (62, 18, 18, 18, 18, 18, 18))
put(wd, "A4", "Coeficientes", font=NEG, fill=GRIS)
put(wd, "A5", "Coef. dic-2025 → dic-2026 (ajuste_2026.xlsx, Controles!B32; dic-2026 ESTIMADO)")
put(wd, "B5", coef_25_26, COEF, AZUL)
put(wd, "A6", "Coef. dic-2024 → dic-2025 (ajuste_2025.xlsx, Coeficientes!E3)")
put(wd, "B6", coef_24_25, COEF, AZUL)

put(wd, "A8", "Actividades de inversión (fechas y coeficientes del auxiliar de bienes de uso de la fase 2)", font=NEG, fill=GRIS)
put(wd, "B8", "Nominal", font=NEG, fill=GRIS)
put(wd, "C8", "Coef. a dic-2026", font=NEG, fill=GRIS)
put(wd, "D8", "Moneda dic-2026", font=NEG, fill=GRIS)
inv = [
    (9, "2026 — Compra camión C11 (15/06/2026), sin IVA", 175000000, coef_c11),
    (10, "2026 — Cobro venta camión C01 (15/06/2026), sin IVA", 28000000, coef_c01_baja),
    (11, "2025 — Compra camión C10 (06/05/2025), sin IVA", 138000000, coef_c10),
    (12, "2025 — Compra vehículo ejecutivo E05 (15/09/2025), sin IVA", 62100000, coef_e05),
]
for rr, txt, nom, cf in inv:
    put(wd, f"A{rr}", txt)
    put(wd, f"B{rr}", nom, NUM, AZUL)
    put(wd, f"C{rr}", cf, COEF, AZUL)
    put(wd, f"D{rr}", f"=B{rr}*C{rr}", NUM)

put(wd, "A14", "Préstamo: amortización de capital $37.500.000 por trimestre (mar, jun, sep, dic) — supuesto de la fase 1 (Dato faltante: cronograma contractual)", font=NEG, fill=GRIS)
put(wd, "B15", "Nominal", font=NEG)
put(wd, "C15", "Coef. mes → cierre del año", font=NEG)
put(wd, "D15", "Moneda dic-2026", font=NEG)
put(wd, "E15", "Año", font=NEG)
rr = 16
idx = {2: "mar", 5: "jun", 8: "sep", 11: "dic"}
for anio, rec, conv in [(2026, rec26, "1"), (2025, rec25, "$B$5")]:
    for m, nombre in idx.items():
        put(wd, f"A{rr}", f"Cuota de capital {nombre}-{anio}")
        put(wd, f"B{rr}", 37500000, NUM, AZUL)
        put(wd, f"C{rr}", rec[m][2], COEF, AZUL)
        put(wd, f"D{rr}", f"=B{rr}*C{rr}*{conv}", NUM)
        put(wd, f"E{rr}", anio)
        rr += 1
PREST26 = "SUMIFS('Datos EFE'!$D$16:$D$23,'Datos EFE'!$E$16:$E$23,2026)"
PREST25 = "SUMIFS('Datos EFE'!$D$16:$D$23,'Datos EFE'!$E$16:$E$23,2025)"

# RECPAM generado por el efectivo (estimación)
put(wd, "A26", "RECPAM generado por el efectivo y equivalentes — ESTIMACIÓN", font=NEG, fill=GRIS)
put(wd, "A27", "Dato faltante: saldos mensuales de efectivo. Supuesto: saldo nominal lineal entre el inicio y el cierre. RECPAM del mes = −saldo al inicio del mes × inflación del mes × coef. mes → cierre. Inflación y coeficientes: hoja RECPAM de ajuste_2026.xlsx y ajuste_2025.xlsx.")
wd["A27"].alignment = Alignment(wrap_text=True)
wd.row_dimensions[27].height = 45
hdrs = ["Mes", "Inflación del mes", "Coef. mes → cierre del año", "Efectivo nominal al inicio del mes", "RECPAM del mes (moneda cierre del año)", "RECPAM en moneda dic-2026"]
bloques = {}
fila = 29
for anio, rec in [(2026, rec26), (2025, rec25)]:
    put(wd, f"A{fila}", f"Ejercicio {anio}", font=NEG)
    if anio == 2026:
        ini = f"=({suma(s26, 'ini', EFECTIVO)})/$B$5"
        fin = f"={suma(s26, 'fin', EFECTIVO)}"
        conv = "1"
    else:
        ini = f"=({suma(s25, 'ini', EFECTIVO)})/$B$5/$B$6"
        fin = f"=({suma(s25, 'fin', EFECTIVO)})/$B$5"
        conv = "$B$5"
    put(wd, f"A{fila+1}", "Efectivo nominal al inicio del ejercicio")
    put(wd, f"B{fila+1}", ini, NUM)
    put(wd, f"A{fila+2}", "Efectivo nominal al cierre del ejercicio")
    put(wd, f"B{fila+2}", fin, NUM)
    h = fila + 3
    for j, t in enumerate(hdrs, 1):
        put(wd, f"{get_column_letter(j)}{h}", t, font=NEG, fill=GRIS)
    p = h + 1
    for k, (mes, infl, cf) in enumerate(rec):
        q = p + k
        put(wd, f"A{q}", mes, "mmm-yyyy")
        put(wd, f"B{q}", infl, PCT, AZUL)
        put(wd, f"C{q}", cf, COEF, AZUL)
        put(wd, f"D{q}", f"=$B${fila+1}+($B${fila+2}-$B${fila+1})*{k}/12", NUM)
        put(wd, f"E{q}", f"=-D{q}*B{q}*C{q}", NUM)
        put(wd, f"F{q}", f"=E{q}*{conv}", NUM)
    t = p + 12
    put(wd, f"A{t}", f"RECPAM generado por el efectivo {anio} (estimado)", font=NEG)
    put(wd, f"F{t}", f"=SUM(F{p}:F{p+11})", NUM, NEG)
    bloques[anio] = t
    fila = t + 3
RECEF26 = f"'Datos EFE'!$F${bloques[2026]}"
RECEF25 = f"'Datos EFE'!$F${bloques[2025]}"

# ---------------------------------------------------------------- ESP
we = hoja("ESP", "Estado de situación patrimonial al 31/12/2026, comparativo con el 31/12/2025",
          "En pesos de moneda homogénea de diciembre de 2026 (índices de nov y dic-2026 estimados). RT 54 FACPCE (t.o. RT 59).")
put(we, "A4", "Rubro", font=NEG, fill=GRIS)
put(we, "B4", "31/12/2026", font=NEG, fill=GRIS)
put(we, "C4", "31/12/2025", font=NEG, fill=GRIS)
for c in "BC":
    we[f"{c}4"].alignment = Alignment(horizontal="right")

esp = {}


def linea(ws, row, texto, f26, f25, bold=False, borde=False):
    put(ws, f"A{row}", texto, font=NEG if bold else None)
    put(ws, f"B{row}", f26, NUM, NEG if bold else None, border=LINEA if borde else None)
    put(ws, f"C{row}", f25, NUM, NEG if bold else None, border=LINEA if borde else None)


r = 5
put(we, f"A{r}", "ACTIVO", font=NEG); r += 1
put(we, f"A{r}", "Activo corriente", font=NEG); r += 1
items_ac = [
    ("Caja y bancos", f"={s26('fin', rubro='Caja y bancos')}", f"={s26('ini', rubro='Caja y bancos')}"),
    ("Inversiones (plazo fijo)", f"={s26('fin', rubro='Inversiones')}", f"={s26('ini', rubro='Inversiones')}"),
    ("Créditos por ventas", f"={s26('fin', rubro='Créditos por ventas')}", f"={s26('ini', rubro='Créditos por ventas')}"),
    ("Otros créditos (sin anticipos de impuesto a las ganancias, compensados en Cargas fiscales)",
     f"={s26('fin', rubro='Otros créditos')}-{s26('fin', codigo='1.4.01')}",
     f"={s26('ini', rubro='Otros créditos')}-{s26('ini', codigo='1.4.01')}"),
    ("Bienes de cambio", f"={s26('fin', rubro='Bienes de cambio')}", f"={s26('ini', rubro='Bienes de cambio')}"),
]
ini_ac = r
for t, a, b in items_ac:
    linea(we, r, t, a, b); esp[t.split(" (")[0]] = r; r += 1
linea(we, r, "Total del activo corriente", f"=SUM(B{ini_ac}:B{r-1})", f"=SUM(C{ini_ac}:C{r-1})", True, True)
TAC = r; r += 2
put(we, f"A{r}", "Activo no corriente", font=NEG); r += 1
linea(we, r, "Bienes de uso", f"={s26('fin', rubro='Bienes de uso')}", f"={s26('ini', rubro='Bienes de uso')}")
esp["Bienes de uso"] = r; r += 1
linea(we, r, "Total del activo no corriente", f"=B{r-1}", f"=C{r-1}", True, True)
TANC = r; r += 1
linea(we, r, "TOTAL DEL ACTIVO", f"=B{TAC}+B{TANC}", f"=C{TAC}+C{TANC}", True, True)
TA = r; r += 2

put(we, f"A{r}", "PASIVO", font=NEG); r += 1
put(we, f"A{r}", "Pasivo corriente", font=NEG); r += 1
ini_pc = r
linea(we, r, "Deudas comerciales", f"=-{s26('fin', rubro='Deudas comerciales')}", f"=-{s26('ini', rubro='Deudas comerciales')}")
esp["Deudas comerciales"] = r; r += 1
linea(we, r, "Remuneraciones y cargas sociales", f"=-{s26('fin', rubro='Remuneraciones y cargas sociales')}", f"=-{s26('ini', rubro='Remuneraciones y cargas sociales')}")
esp["Remuneraciones y cargas sociales"] = r; r += 1
linea(we, r, "Cargas fiscales (impuesto a las ganancias neto de anticipos)",
      f"=-{s26('fin', rubro='Cargas fiscales')}-{s26('fin', codigo='1.4.01')}",
      f"=-{s26('ini', rubro='Cargas fiscales')}-{s26('ini', codigo='1.4.01')}")
esp["Cargas fiscales"] = r; r += 1
# Préstamo: porción corriente $150.000.000 nominales por año (supuesto de la fase 1); 2025 reexpresado
linea(we, r, "Préstamos (porción corriente)",
      f"=MIN(150000000,-{s26('fin', rubro='Préstamos')})",
      f"=MIN(150000000*'Datos EFE'!$B$5,-{s26('ini', rubro='Préstamos')})")
esp["Préstamos corriente"] = r; r += 1
linea(we, r, "Total del pasivo corriente", f"=SUM(B{ini_pc}:B{r-1})", f"=SUM(C{ini_pc}:C{r-1})", True, True)
TPC = r; r += 2
put(we, f"A{r}", "Pasivo no corriente", font=NEG); r += 1
linea(we, r, "Préstamos (porción no corriente)",
      f"=-{s26('fin', rubro='Préstamos')}-B{esp['Préstamos corriente']}",
      f"=-{s26('ini', rubro='Préstamos')}-C{esp['Préstamos corriente']}")
esp["Préstamos no corriente"] = r; r += 1
linea(we, r, "Total del pasivo no corriente", f"=B{r-1}", f"=C{r-1}", True, True)
TPNC = r; r += 1
linea(we, r, "TOTAL DEL PASIVO", f"=B{TPC}+B{TPNC}", f"=C{TPC}+C{TPNC}", True, True)
TP = r; r += 2
linea(we, r, "PATRIMONIO NETO (según estado respectivo)", "=EEPN!G12", "=EEPN!G9", True)
PN = r; r += 1
linea(we, r, "TOTAL DEL PASIVO Y PATRIMONIO NETO", f"=B{TP}+B{PN}", f"=C{TP}+C{PN}", True, True)
TPPN = r; r += 2
put(we, f"A{r}", "Las notas y anexos (a cargo de la obrera notas-anexos-2026) forman parte de este estado.", font=Font(italic=True, size=9))

# ---------------------------------------------------------------- ER
wr = hoja("ER", "Estado de resultados por el ejercicio finalizado el 31/12/2026, comparativo con 2025",
          "En pesos de moneda homogénea de diciembre de 2026. Ganancia (+) / pérdida (−).")
put(wr, "A4", "Concepto", font=NEG, fill=GRIS)
put(wr, "B4", "2026", font=NEG, fill=GRIS)
put(wr, "C4", "2025", font=NEG, fill=GRIS)
er = {}
r = 5
linea(wr, r, "Ventas netas de insumos", f"=-{s26('fin', rubro='Ingresos')}", f"=-{s25('fin', rubro='Ingresos')}"); er["ventas"] = r; r += 1
linea(wr, r, "Costo de las mercaderías vendidas", f"=-{s26('fin', rubro='Costo de ventas')}", f"=-{s25('fin', rubro='Costo de ventas')}"); er["cmv"] = r; r += 1
linea(wr, r, "Resultado bruto", f"=B{r-2}+B{r-1}", f"=C{r-2}+C{r-1}", True, True); er["bruto"] = r; r += 2
linea(wr, r, "Gastos de comercialización", f"=-{s26('fin', rubro='Gastos de comercialización')}", f"=-{s25('fin', rubro='Gastos de comercialización')}"); er["com"] = r; r += 1
linea(wr, r, "Gastos de administración", f"=-{s26('fin', rubro='Gastos de administración')}", f"=-{s25('fin', rubro='Gastos de administración')}"); er["adm"] = r; r += 1
linea(wr, r, "Otros ingresos y egresos (resultado por venta de bienes de uso)", f"=-{s26('fin', rubro='Otros ingresos y egresos')}", f"=-{s25('fin', rubro='Otros ingresos y egresos')}"); er["otros"] = r; r += 1
put(wr, f"A{r}", "Resultados financieros y por tenencia (incluido el RECPAM)", font=NEG); r += 1
ini_rf = r
linea(wr, r, "   Generados por activos — intereses ganados de plazo fijo", f"=-{s26('fin', codigo='4.2.01')}", f"=-{s25('fin', codigo='4.2.01')}"); r += 1
linea(wr, r, "   Generados por pasivos — intereses del préstamo bancario", f"=-{s26('fin', codigo='5.4.01')}", f"=-{s25('fin', codigo='5.4.01')}"); r += 1
linea(wr, r, "   RECPAM (resultado por exposición al cambio en el poder adquisitivo de la moneda)", f"=-{s26('fin', codigo='4.9.01')}", f"=-{s25('fin', codigo='4.9.01')}"); er["recpam"] = r; r += 1
linea(wr, r, "Total resultados financieros y por tenencia", f"=SUM(B{ini_rf}:B{r-1})", f"=SUM(C{ini_rf}:C{r-1})", True, True); er["rf"] = r; r += 2
linea(wr, r, "Resultado antes del impuesto a las ganancias",
      f"=B{er['bruto']}+B{er['com']}+B{er['adm']}+B{er['otros']}+B{er['rf']}",
      f"=C{er['bruto']}+C{er['com']}+C{er['adm']}+C{er['otros']}+C{er['rf']}", True, True); er["rai"] = r; r += 1
linea(wr, r, "Impuesto a las ganancias", f"=-{s26('fin', rubro='Impuesto a las ganancias')}", f"=-{s25('fin', rubro='Impuesto a las ganancias')}"); er["ig"] = r; r += 1
linea(wr, r, "RESULTADO DEL EJERCICIO — GANANCIA", f"=B{r-2}+B{r-1}", f"=C{r-2}+C{r-1}", True, True); er["res"] = r; r += 2
put(wr, f"A{r}", "El detalle de gastos por naturaleza y función va en el anexo de gastos (obrera notas-anexos-2026).", font=Font(italic=True, size=9))
RES26 = f"ER!$B${er['res']}"
RES25 = f"ER!$C${er['res']}"

# ---------------------------------------------------------------- EEPN
wp = hoja("EEPN", "Estado de evolución del patrimonio neto por el ejercicio finalizado el 31/12/2026, comparativo con 2025",
          "En pesos de moneda homogénea de diciembre de 2026. Capital suscripto a valor nominal; su reexpresión, en Ajuste de capital.",
          (54, 18, 20, 20, 18, 20, 20))
cols = ["Concepto", "Capital suscripto", "Ajuste de capital", "Total aportes de los propietarios", "Reserva legal", "Resultados no asignados", "Total del patrimonio neto"]
for j, h in enumerate(cols, 1):
    c = put(wp, f"{get_column_letter(j)}4", h, font=NEG, fill=GRIS)
    c.alignment = Alignment(wrap_text=True, horizontal="center")
wp.row_dimensions[4].height = 45


def fila_pn(row, texto, b, c, e, f, bold=False, borde=False):
    put(wp, f"A{row}", texto, font=NEG if bold else None)
    for col, v in zip("BCEF", (b, c, e, f)):
        put(wp, f"{col}{row}", v, NUM, NEG if bold else None, border=LINEA if borde else None)
    put(wp, f"D{row}", f"=B{row}+C{row}", NUM, NEG if bold else None, border=LINEA if borde else None)
    put(wp, f"G{row}", f"=D{row}+E{row}+F{row}", NUM, NEG if bold else None, border=LINEA if borde else None)


cap_nom = f"-{s26('ini', codigo='3.1.01')}"
fila_pn(5, "Saldos al 31/12/2024",
        f"={cap_nom}",
        f"=-({s25('ini', codigo='3.1.01')}+{s25('ini', codigo='3.1.02')})-B5",
        f"=-{s25('ini', codigo='3.2.01')}",
        f"=-{s25('ini', codigo='3.3.01')}", True)
fila_pn(6, "Asamblea ordinaria del 28/04/2025 — dividendos en efectivo", 0, 0, 0,
        f"=-{s25('fin', codigo='3.3.01')}+{s25('ini', codigo='3.3.01')}")
fila_pn(7, "Resultado del ejercicio 2025 — ganancia", 0, 0, 0, f"={RES25}")
fila_pn(8, "", None, None, None, None)
for col in "BCDEFG":
    wp[f"{col}8"].value = None
fila_pn(9, "Saldos al 31/12/2025", "=SUM(B5:B7)", "=SUM(C5:C7)", "=SUM(E5:E7)", "=SUM(F5:F7)", True, True)
fila_pn(10, "Asamblea ordinaria del 28/04/2026 — dividendos en efectivo", 0, 0, 0,
        f"=-{s26('fin', codigo='3.3.01')}+{s26('ini', codigo='3.3.01')}")
fila_pn(11, "Resultado del ejercicio 2026 — ganancia", 0, 0, 0, f"={RES26}")
fila_pn(12, "Saldos al 31/12/2026", "=SUM(B9:B11)", "=SUM(C9:C11)", "=SUM(E9:E11)", "=SUM(F9:F11)", True, True)
put(wp, "A14", "Duda — Reserva legal (art. 70 LGS): en moneda homogénea queda por debajo del 20 % de capital + ajuste de capital (faltante $731,0 M al 31/12/2026). No se contabiliza asignación: requiere decisión de asamblea. Ver hoja Supuestos.", font=Font(italic=True, size=9))
wp["A14"].alignment = Alignment(wrap_text=True)
wp.merge_cells("A14:G14")
wp.row_dimensions[14].height = 30

# ---------------------------------------------------------------- EFE
wf = hoja("EFE", "Estado de flujo de efectivo por el ejercicio finalizado el 31/12/2026, comparativo con 2025 — método indirecto",
          "En pesos de moneda homogénea de diciembre de 2026. Efectivo = caja + bancos + plazo fijo. Aumento (+) / disminución (−).")
put(wf, "A4", "Concepto", font=NEG, fill=GRIS)
put(wf, "B4", "2026", font=NEG, fill=GRIS)
put(wf, "C4", "2025", font=NEG, fill=GRIS)
ef = {}
r = 5
put(wf, f"A{r}", "VARIACIÓN DEL EFECTIVO", font=NEG); r += 1
linea(wf, r, "Efectivo al inicio del ejercicio (reexpresado)", f"={suma(s26, 'ini', EFECTIVO)}", f"={suma(s25, 'ini', EFECTIVO)}"); ef["ini"] = r; r += 1
linea(wf, r, "Efectivo al cierre del ejercicio", f"={suma(s26, 'fin', EFECTIVO)}", f"={suma(s25, 'fin', EFECTIVO)}"); ef["fin"] = r; r += 1
linea(wf, r, "(Disminución) aumento neto del efectivo", f"=B{r-1}-B{r-2}", f"=C{r-1}-C{r-2}", True, True); ef["var"] = r; r += 2

put(wf, f"A{r}", "CAUSAS DE LAS VARIACIONES DEL EFECTIVO", font=NEG); r += 1
put(wf, f"A{r}", "Actividades operativas", font=NEG); r += 1
ini_op = r
linea(wf, r, "Resultado del ejercicio", f"={RES26}", f"={RES25}"); r += 1
put(wf, f"A{r}", "Ajustes para arribar al flujo neto de efectivo de las actividades operativas:", font=Font(italic=True)); r += 1
linea(wf, r, "   Amortizaciones de bienes de uso", f"={suma(s26, 'fin', AMORT)}", f"={suma(s25, 'fin', AMORT)}"); r += 1
linea(wf, r, "   Resultado por venta de bienes de uso", f"={s26('fin', codigo='4.3.01')}", f"={s25('fin', codigo='4.3.01')}"); r += 1
linea(wf, r, "   RECPAM generado por el efectivo (se expone por separado)", f"=-{RECEF26}", f"=-{RECEF25}"); r += 1
linea(wf, r, "   RECPAM del préstamo bancario (partida no operativa)",
      f"=-((-{s26('ini', rubro='Préstamos')})-(-{s26('fin', rubro='Préstamos')})-{PREST26})",
      f"=-((-{s25('ini', rubro='Préstamos')})-(-{s25('fin', rubro='Préstamos')})-{PREST25})"); ef["recprest"] = r; r += 1
put(wf, f"A{r}", "Cambios en activos y pasivos operativos (en moneda homogénea):", font=Font(italic=True)); r += 1
cambios = [
    ("   (Aumento) disminución de créditos por ventas", "A", "Créditos por ventas", None),
    ("   (Aumento) disminución de otros créditos", "A", "Otros créditos", "neto"),
    ("   (Aumento) disminución de bienes de cambio", "A", "Bienes de cambio", None),
    ("   Aumento (disminución) de deudas comerciales", "P", "Deudas comerciales", None),
    ("   Aumento (disminución) de remuneraciones y cargas sociales", "P", "Remuneraciones y cargas sociales", None),
    ("   Aumento (disminución) de cargas fiscales (neto de anticipos)", "P", "Cargas fiscales", "neto"),
]
for t, tipo, rub, neto in cambios:
    if tipo == "A":
        x26 = f"=({s26('ini', rubro=rub)})-({s26('fin', rubro=rub)})"
        x25 = f"=({s25('ini', rubro=rub)})-({s25('fin', rubro=rub)})"
        if neto:
            x26 += f"-({s26('ini', codigo='1.4.01')})+({s26('fin', codigo='1.4.01')})"
            x25 += f"-({s25('ini', codigo='1.4.01')})+({s25('fin', codigo='1.4.01')})"
    else:
        x26 = f"=-({s26('fin', rubro=rub)})+({s26('ini', rubro=rub)})"
        x25 = f"=-({s25('fin', rubro=rub)})+({s25('ini', rubro=rub)})"
        if neto:
            x26 += f"-({s26('fin', codigo='1.4.01')})+({s26('ini', codigo='1.4.01')})"
            x25 += f"-({s25('fin', codigo='1.4.01')})+({s25('ini', codigo='1.4.01')})"
    linea(wf, r, t, x26, x25); r += 1
linea(wf, r, "Flujo neto de efectivo generado por las actividades operativas", f"=SUM(B{ini_op}:B{r-1})", f"=SUM(C{ini_op}:C{r-1})", True, True); ef["op"] = r; r += 2

put(wf, f"A{r}", "Actividades de inversión", font=NEG); r += 1
ini_inv = r
linea(wf, r, "Pagos por compras de bienes de uso", "=-'Datos EFE'!D9", "=-('Datos EFE'!D11+'Datos EFE'!D12)"); r += 1
linea(wf, r, "Cobros por ventas de bienes de uso", "='Datos EFE'!D10", "=0"); r += 1
linea(wf, r, "Flujo neto de efectivo utilizado en las actividades de inversión", f"=SUM(B{ini_inv}:B{r-1})", f"=SUM(C{ini_inv}:C{r-1})", True, True); ef["inv"] = r; r += 2

put(wf, f"A{r}", "Actividades de financiación", font=NEG); r += 1
ini_fin = r
linea(wf, r, "Pagos de préstamos (capital)", f"=-{PREST26}", f"=-{PREST25}"); r += 1
linea(wf, r, "Pago de dividendos en efectivo", "=EEPN!G10", "=EEPN!G6"); r += 1
linea(wf, r, "Flujo neto de efectivo utilizado en las actividades de financiación", f"=SUM(B{ini_fin}:B{r-1})", f"=SUM(C{ini_fin}:C{r-1})", True, True); ef["fin_act"] = r; r += 2

linea(wf, r, "Resultados financieros y por tenencia generados por el efectivo (RECPAM) — estimado", f"={RECEF26}", f"={RECEF25}", True); ef["recef"] = r; r += 2
linea(wf, r, "(Disminución) aumento neto del efectivo", f"=B{ef['op']}+B{ef['inv']}+B{ef['fin_act']}+B{ef['recef']}",
      f"=C{ef['op']}+C{ef['inv']}+C{ef['fin_act']}+C{ef['recef']}", True, True); ef["tot"] = r; r += 2
notas_efe = [
    "Intereses pagados y cobrados: clasificados en actividades operativas (incluidos en el resultado).",
    "Impuesto a las ganancias pagado: incluido en actividades operativas (variación de cargas fiscales neta de anticipos).",
    "El IVA de la compra y la venta de bienes de uso se incluye en la variación de otros créditos (operativas).",
    "RECPAM del efectivo: estimación (Dato faltante: saldos mensuales de efectivo). Ver hoja Datos EFE.",
    "Conciliación efectivo — ESP: caja y bancos + inversiones (plazo fijo, considerado equivalente de efectivo).",
]
for n in notas_efe:
    put(wf, f"A{r}", n, font=Font(italic=True, size=9)); r += 1

# ---------------------------------------------------------------- Fase 2 (valores de control)
w2 = hoja("Fase 2", "Cifras de control importadas de la fase 2 (ajuste por inflación)",
          "Valores importados de entregables/ajuste/ (recalculados con LibreOffice). Se usan solo en la hoja Controles.", (70, 22, 22))
put(w2, "A4", "Concepto", font=NEG, fill=GRIS)
put(w2, "B4", "Importe", font=NEG, fill=GRIS)
put(w2, "C4", "Fuente", font=NEG, fill=GRIS)
f2rows = [
    ("Resultado 2026 ajustado", f2["res26"], "ajuste_2026.xlsx Controles!B26"),
    ("PN al 31/12/2026", f2["pn26"], "ajuste_2026.xlsx Controles!B28"),
    ("PN al 31/12/2025 (moneda dic-2026)", f2["pn25"], "ajuste_2026.xlsx Controles!B29"),
    ("Total activo al 31/12/2026 (bruto, sin compensar anticipos)", f2["act26"], "ajuste_2026.xlsx Controles!B30"),
    ("Resultado 2025 (moneda dic-2026)", f2["res25"], "comparativo Controles!B16"),
]
F2 = {}
r = 5
for t, v, fu in f2rows:
    put(w2, f"A{r}", t); put(w2, f"B{r}", v, NUM, AZUL); put(w2, f"C{r}", fu); F2[t] = r; r += 1
r += 1
put(w2, f"A{r}", "ESP al 31/12/2025 del archivo comparativo (rubros)", font=NEG, fill=GRIS); r += 1
mapa25 = [
    ("Caja y bancos", "Caja y bancos"), ("Inversiones", "Inversiones"), ("Créditos por ventas", "Créditos por ventas"),
    ("Otros créditos", "Otros créditos"), ("Bienes de cambio", "Bienes de cambio"), ("Bienes de uso", "Bienes de uso"),
    ("Deudas comerciales", "Deudas comerciales"), ("Remuneraciones y cargas sociales", "Remuneraciones y cargas sociales"),
    ("Cargas fiscales", "Cargas fiscales"),
    ("Préstamos corriente", "Préstamos (porción corriente) — supuesto $150.000.000 nominal por año"),
    ("Préstamos no corriente", "Préstamos (porción no corriente)"),
]
F2ESP = {}
for clave, rot in mapa25:
    put(w2, f"A{r}", rot); put(w2, f"B{r}", esp25[rot], NUM, AZUL); put(w2, f"C{r}", "comparativo, hoja ESP 2025")
    F2ESP[clave] = r; r += 1

# ---------------------------------------------------------------- Controles
wc = hoja("Controles", "Controles de los criterios de aceptación",
          "OK si |valor| < tolerancia ($0,50). Todos los valores deben ser 0.", (92, 18, 12, 10))
for j, h in enumerate(["Control", "Valor", "Tolerancia", "Resultado"], 1):
    put(wc, f"{get_column_letter(j)}4", h, font=NEG, fill=GRIS)
anti25 = s26("ini", codigo="1.4.01")
reclas = {"Otros créditos": f"+{anti25}", "Cargas fiscales": f"+{anti25}"}
dif25 = "+".join(
    f"ABS(ESP!C{esp[k]}{reclas.get(k, '')}-'Fase 2'!B{F2ESP[k]})" for k, _ in mapa25
)
ctrls = [
    ("ESP 31/12/2026: activo − (pasivo + PN)", f"=ESP!B{TA}-ESP!B{TPPN}"),
    ("ESP 31/12/2025: activo − (pasivo + PN)", f"=ESP!C{TA}-ESP!C{TPPN}"),
    ("ER 2026 − variación de RNA del EEPN (RNA cierre − RNA inicio − dividendos)", f"={RES26}-(EEPN!F12-EEPN!F9-EEPN!F10)"),
    ("ER 2025 − variación de RNA del EEPN", f"={RES25}-(EEPN!F9-EEPN!F5-EEPN!F6)"),
    ("EFE 2026: variación por causas − variación del efectivo en el ESP", f"=EFE!B{ef['tot']}-((ESP!B{esp['Caja y bancos']}+ESP!B{esp['Inversiones']})-(ESP!C{esp['Caja y bancos']}+ESP!C{esp['Inversiones']}))"),
    ("EFE 2025: variación por causas − variación del efectivo (BSS 2025)", f"=EFE!C{ef['tot']}-EFE!C{ef['var']}"),
    ("EFE: efectivo al cierre 2025 del EFE − efectivo al inicio 2026", f"=EFE!C{ef['fin']}-EFE!B{ef['ini']}"),
    ("ER 2026: resultado − (−Σ cuentas de resultado del BSS 2026)", f"={RES26}+SUMIFS('BSS 2026'!$F$5:$F${B26_LAST},'BSS 2026'!$D$5:$D${B26_LAST},\"R\")"),
    ("ER 2026 = resultado de la fase 2 (ajuste_2026.xlsx)", f"={RES26}-'Fase 2'!B{F2['Resultado 2026 ajustado']}"),
    ("ER 2025 = resultado del comparativo de la fase 2", f"={RES25}-'Fase 2'!B{F2['Resultado 2025 (moneda dic-2026)']}"),
    ("PN 31/12/2026 = fase 2", f"=ESP!B{PN}-'Fase 2'!B{F2['PN al 31/12/2026']}"),
    ("PN 31/12/2025 = fase 2 (y = PN inicial 2026)", f"=ESP!C{PN}-'Fase 2'!B{F2['PN al 31/12/2025 (moneda dic-2026)']}"),
    ("ESP 31/12/2025 = comparativo de la fase 2, rubro por rubro (Σ|dif.|, neto de la reclasificación de anticipos de IG)", f"={dif25}"),
    ("Activo 31/12/2026 = fase 2 menos anticipos de IG compensados", f"=ESP!B{TA}-('Fase 2'!B{F2['Total activo al 31/12/2026 (bruto, sin compensar anticipos)']}-{s26('fin', codigo='1.4.01')})"),
    ("EEPN 31/12/2025 = saldos iniciales del BSS 2026 por componente (Σ|dif.|)",
     f"=ABS(EEPN!B9+{s26('ini', codigo='3.1.01')})+ABS(EEPN!C9+{s26('ini', codigo='3.1.02')})+ABS(EEPN!E9+{s26('ini', codigo='3.2.01')})+ABS(EEPN!F9+{s26('ini', codigo='3.3.01')})"),
    ("EEPN 31/12/2026: RNA = BSS 2026 pre-cierre + resultado", f"=EEPN!F12-(-{s26('fin', codigo='3.3.01')}+{RES26})"),
    ("BSS 2026 cuadra (saldos 31/12/2025 y 31/12/2026)", f"=ABS('BSS 2026'!E{B26_TOT})+ABS('BSS 2026'!F{B26_TOT})"),
    ("BSS 2025 cuadra (saldos 31/12/2024 y 31/12/2025)", f"=ABS('BSS 2025'!D{B25_TOT})+ABS('BSS 2025'!E{B25_TOT})"),
    ("Préstamo corriente 31/12/2026 = $150.000.000", f"=ESP!B{esp['Préstamos corriente']}-150000000"),
]
r = 5
for t, fml in ctrls:
    put(wc, f"A{r}", t)
    put(wc, f"B{r}", fml, '#,##0.00;(#,##0.00);"0"')
    put(wc, f"C{r}", 0.5)
    put(wc, f"D{r}", f'=IF(ABS(B{r})<C{r},"OK","ERROR")')
    r += 1
put(wc, f"A{r}", "Resultado global", font=NEG)
put(wc, f"D{r}", f'=IF(COUNTIF(D5:D{r-1},"OK")={len(ctrls)},"OK","ERROR")', font=NEG)
CTRL_GLOBAL = r
r += 2
put(wc, f"A{r}", "Cifras principales", font=NEG, fill=GRIS); r += 1
for t, fml in [
    ("Total del activo al 31/12/2026", f"=ESP!B{TA}"), ("Total del pasivo al 31/12/2026", f"=ESP!B{TP}"),
    ("Patrimonio neto al 31/12/2026", f"=ESP!B{PN}"), ("Resultado del ejercicio 2026", f"={RES26}"),
    ("RECPAM 2026", f"=ER!B{er['recpam']}"), ("Total del activo al 31/12/2025", f"=ESP!C{TA}"),
    ("Patrimonio neto al 31/12/2025", f"=ESP!C{PN}"), ("Resultado del ejercicio 2025", f"={RES25}"),
]:
    put(wc, f"A{r}", t); put(wc, f"B{r}", fml, NUM); r += 1

# ---------------------------------------------------------------- Supuestos
wsu = hoja("Supuestos", "Norma, supuestos, dudas y datos faltantes", None, (150,))
textos = [
    ("NORMA", True),
    ("Normas contables profesionales argentinas: RT 54 FACPCE — Norma Unificada Argentina de Contabilidad (NUA), texto ordenado por la RT 59 (incorpora los cambios aprobados al 01/12/2025).", False),
    ("Vigencia: obligatoria para ejercicios iniciados a partir del 01/01/2025 (t.o. RT 59). El ejercicio 2026 (01/01/2026–31/12/2026) está alcanzado.", False),
    ("RT 8 y RT 9: la RT 54 deja sin efecto la RT 8 (Normas generales de exposición contable) y la RT 6 (texto ordenado publicado por CPCECABA). Fuentes de CPCE Salta indican que también deroga totalmente la RT 9, junto con las RT 11, 17, 41, 42 y 48, y parcialmente la RT 18 y la RT 21. Conclusión: la presentación se rige por la RT 54.", False),
    ("Duda: la derogación de la RT 9 y los números de párrafo de la RT 54 sobre presentación no se verificaron sobre el texto oficial de la FACPCE. La red de esta sesión bloqueó la descarga de los PDF. Sí hay coincidencia entre fuentes de varios consejos (CPCECABA, CPCE Salta, CPCE Santa Fe).", False),
    ("Res. JG FACPCE 660/26: dispensa transitoria y limitada a la presentación (no al reconocimiento ni a la medición), solo para el primer ejercicio iniciado desde el 01/01/2025 en que se adopta la RT 54. No es una prórroga. No se aplica: 2026 no es el primer ejercicio de aplicación (Supuesto).", False),
    ("Fuentes consultadas: consejo.org.ar (micrositio RT 54 y PDF del t.o. RT 59); consejosalta.org.ar («Vientos de cambio: la nueva RT 54»; «Presentación del EEPN y del EFE según la NUA»); cpcesfe2.org.ar; facpce.org.ar (Res. JG 660/26). Consulta: 22/11/2026.", False),
    ("Unidad de medida: moneda homogénea de diciembre de 2026 (RT 54, procedimiento de reexpresión que proviene de la RT 6; serie FACPCE Res. JG 539/18). Estimación: los índices de noviembre y diciembre de 2026 no son oficiales. Hay que regenerar cuando se publiquen.", False),
    ("Exposición del RECPAM: en el ER, dentro de «Resultados financieros y por tenencia (incluido el RECPAM)», en un renglón propio. En el EFE solo se expone el RECPAM generado por el efectivo, como causa separada de las tres actividades (Interpretación FACPCE 2). Según la guía del CPCECABA, la RT 54 no cambió los métodos de presentación del EFE.", False),
    ("SUPUESTOS", True),
    ("(S) Insumos: BSS ajustados de la fase 2, integrados y verificados por la reina (27 controles OK). Este libro los importa como valores y no los modifica.", False),
    ("(S) Efectivo = caja + bancos + plazo fijo, según el encargo. El plazo fijo se considera equivalente de efectivo (vencimiento corto, sin riesgo significativo de cambio de valor). Dato faltante: el plazo original de cada colocación.", False),
    ("(S) Impuesto a las ganancias: el pasivo se presenta neto de los anticipos del mismo período fiscal (2026: 673.170.750 − 350.000.000; 2025: neto de los anticipos 2025). Hay derecho legal de compensación y se cancela por el neto. Por eso el ESP 31/12/2025 difiere del comparativo de la fase 2, que lo presenta en bruto, en otros créditos y cargas fiscales por $196.840.462. El total del PN no cambia.", False),
    ("(S) Préstamo: porción corriente de $150.000.000 nominales por año (supuesto de la fase 1). Al 31/12/2025 se reexpresa a moneda de dic-2026. Dato faltante: cronograma contractual; se asumen cuotas trimestrales de $37,5 M en mar, jun, sep y dic.", False),
    ("(S) EFE en moneda homogénea: los flujos de inversión y financiación se reexpresan con el coeficiente de su mes. Los dividendos se reexpresan desde la asamblea. El RECPAM del préstamo es una partida no operativa y se ajusta en operativas. El RECPAM del resto de las partidas monetarias operativas queda dentro de sus variaciones.", False),
    ("(S) Intereses cobrados y pagados e impuesto a las ganancias pagado: en actividades operativas (alternativa admitida).", False),
    ("Estimación: RECPAM generado por el efectivo. Dato faltante: saldos mensuales de efectivo. Se estima con saldo nominal lineal entre inicio y cierre (hoja Datos EFE). No afecta la variación total del efectivo; solo la distribución entre operativas y el renglón de RECPAM.", False),
    ("(S) Ajuste de capital: el capital suscripto se expone a valor nominal ($20.000.000) y su reexpresión en Ajuste de capital, también para el 31/12/2024.", False),
    ("(S) Formato: importes en pesos sin decimales (formato de celda), con separador de miles. Las fórmulas usan SUMIFS, compatibles con LibreOffice 24.2 para poder recalcular y verificar.", False),
    ("DUDAS Y DATOS FALTANTES", True),
    ("Duda — Reserva legal (art. 70 LGS): con base capital + ajuste de capital, la reserva reexpresada ($220,0 M) queda por debajo del 20 % ($951,0 M). Faltante: $731,0 M al 31/12/2026. La asamblea del 28/04/2026 no asignó el 5 % del resultado 2025 ajustado ($30,6 M) y distribuyó dividendos. No se contabiliza ninguna asignación. Propuesta de la fase 2: informarlo en nota y que la asamblea 2027 asigne hasta $55,4 M. Validar con el asesor societario y el criterio de la IGJ.", False),
    ("Dato faltante — Valor recuperable de los bienes de uso ($5.631,4 M al 31/12/2026) y VNR de los bienes de cambio ($1.632,9 M): no hay datos para la prueba. Se asume que el valor reexpresado no supera el recuperable.", False),
    ("Estimación — Índices FACPCE de nov y dic-2026: afectan todas las cifras en moneda de dic-2026, comparativos incluidos.", False),
    ("Dato faltante — Cronograma contractual del préstamo (ver supuestos).", False),
    ("Duda — Res. JG 660/26: confirmar si la sociedad aplicó la RT 54 en 2025. Si 2026 fuera su primer ejercicio de aplicación, podría usar la dispensa de presentación.", False),
    ("Fuera de alcance: notas y anexos (obrera notas-anexos-2026).", False),
]
r = 3
for t, b in textos:
    c = put(wsu, f"A{r}", t, font=NEG if b else None, fill=GRIS if b else None)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    if not b:
        wsu.row_dimensions[r].height = 15 * max(1, len(t) // 140 + 1)
    r += 1

# Orden final
orden = ["Supuestos", "Controles", "ESP", "ER", "EEPN", "EFE", "BSS 2026", "BSS 2025", "Datos EFE", "Fase 2"]
wb._sheets = [wb[n] for n in orden]
for ws in wb:
    ws.sheet_properties.pageSetUpPr = None
wb.save(OUT)
print("OK", OUT, "| controles: fila global", CTRL_GLOBAL)
