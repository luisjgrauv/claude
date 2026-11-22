# -*- coding: utf-8 -*-
"""
Generador del ajuste por inflación de Distribuidora Andina S.A. (FICTICIA).

Produce, con fórmulas vinculadas (recalculadas luego con LibreOffice):
  entregables/ajuste/ajuste_2025.xlsx                         (moneda de dic-2025)
  entregables/ajuste/ajuste_2026.xlsx                         (moneda de dic-2026)
  entregables/ajuste/comparativo_2025_en_moneda_dic2026.xlsx  (2025 en moneda de dic-2026)

Uso:
  python3 entregables/ajuste/generador_ajuste.py
El script genera 2025, lo recalcula con LibreOffice, importa sus saldos de cierre
para abrir 2026 y para el comparativo, y recalcula todos los libros.
No modifica ningún insumo (solo lee insumos/).
"""
import csv
import datetime as dt
import os
import shutil
import subprocess
import tempfile

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

RAIZ = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
INS = os.path.join(RAIZ, "insumos")
OUT = os.path.join(RAIZ, "entregables", "ajuste")

# ----------------------------------------------------------------------------- estilos
F_IN = Font(color="0000CC")              # insumo (dato importado de los auxiliares / BSS)
F_B = Font(bold=True)
F_T = Font(bold=True, size=12)
F_H = Font(bold=True, color="FFFFFF")
FILL_H = PatternFill("solid", fgColor="1F4E78")
FILL_TOT = PatternFill("solid", fgColor="DDEBF7")
FILL_CTRL = PatternFill("solid", fgColor="E2EFDA")
NUM = '#,##0;[Red]-#,##0'
COEF = '0.000000'
PCT = '0.0000%'
FECHA = 'DD/MM/YYYY'
MES = 'MMM-YYYY'
THIN = Side(style="thin", color="999999")

EMPRESA = "Distribuidora Andina S.A. (FICTICIA)"


def hdr(ws, row, titles, col=1):
    for i, t in enumerate(titles):
        c = ws.cell(row=row, column=col + i, value=t)
        c.font = F_H
        c.fill = FILL_H
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[row].height = 45


def put(ws, ref, value, fmt=None, font=None, fill=None):
    c = ws[ref]
    c.value = value
    if fmt:
        c.number_format = fmt
    if font:
        c.font = font
    if fill:
        c.fill = fill
    return c


def titulo(ws, t, sub):
    put(ws, "A1", f"{EMPRESA} — {t}", font=F_T)
    put(ws, "A2", sub)


def widths(ws, w):
    for i, x in enumerate(w, start=1):
        ws.column_dimensions[L(i)].width = x


def mes1(d):
    return dt.datetime(d.year, d.month, 1)


# ----------------------------------------------------------------------------- lectura de insumos
def leer_indices():
    filas = []
    with open(os.path.join(INS, "indices", "indices_facpce.csv"), encoding="utf-8") as fh:
        for r in csv.DictReader(fh, delimiter=";"):
            y, m = map(int, r["mes"].split("-"))
            filas.append((dt.datetime(y, m, 1), float(r["indice"]), r["estado"]))
    return filas


def leer_bss(anio):
    filas = []
    with open(os.path.join(INS, "andina", f"bss_{anio}.csv"), encoding="utf-8") as fh:
        rd = csv.reader(fh, delimiter=";")
        next(rd)
        for r in rd:
            filas.append(dict(cod=r[0], cuenta=r[1], rubro=r[2], si=float(r[3]), d=float(r[4]),
                              c=float(r[5]), sf=float(r[6])))
    return filas


def clase(cod):
    if cod.startswith("1."):
        return "A-NM" if cod in ("1.4.03",) or cod.startswith("1.5") or cod.startswith("1.6") else "A-M"
    if cod.startswith("2."):
        return "P-M"
    if cod.startswith("3."):
        return "PN"
    return "R"


CTA_VO = {"Terrenos": ("1.6.01", None), "Edificios": ("1.6.02", "1.6.03"),
          "Rodados - camiones": ("1.6.04", "1.6.05"), "Rodados - vehículos ejecutivos": ("1.6.06", "1.6.07"),
          "Rodados - camionetas de reparto": ("1.6.08", "1.6.09")}


def cta_amort(cuenta_bss, funcion):
    if cuenta_bss == "Terrenos":
        return None
    if cuenta_bss == "Edificios":
        return "5.3.06" if funcion == "Administración" else "5.2.07"
    return {"Rodados - camiones": "5.2.08", "Rodados - vehículos ejecutivos": "5.3.07",
            "Rodados - camionetas de reparto": "5.2.09"}[cuenta_bss]


def leer_bu(anio):
    if anio == 2025:
        ws = openpyxl.load_workbook(os.path.join(INS, "andina", "auxiliares_2025.xlsx"), data_only=True)["Bienes de uso"]
    else:
        ws = openpyxl.load_workbook(os.path.join(INS, "andina", "auxiliares_2026.xlsx"), data_only=True)["Bienes de uso"]
    out = []
    for r in ws.iter_rows(min_row=5, values_only=True):
        if r[0] in (None, "TOTAL"):
            break
        if anio == 2025:
            cod, tipo, desc, dom, f, vo, vu, vo_ini, aa_ini, am, aa_fin, vrc, cta, fun = r[:14]
            alta = vo if vo_ini == 0 else 0
            out.append(dict(cod=cod, desc=f"{tipo} — {desc}", fecha=f, vu=vu, vo_ini=vo_ini, alta=alta, baja_vo=0,
                            aa_ini=aa_ini, baja_aa=0, am=am, cta=cta, fun=fun, precio=0, fbaja=None))
        else:
            (cod, tipo, desc, dom, f, vo, vu, vo_ini, aa_ini, alta, fbaja, baja_vo, baja_aa, am,
             _vo_fin, _aa_fin, _vrc, precio, _res, cta, fun) = r[:21]
            out.append(dict(cod=cod, desc=f"{tipo} — {desc}", fecha=f, vu=vu, vo_ini=vo_ini, alta=alta or 0,
                            baja_vo=baja_vo or 0, aa_ini=aa_ini, baja_aa=baja_aa or 0, am=am, cta=cta, fun=fun,
                            precio=precio or 0, fbaja=fbaja if isinstance(fbaja, dt.datetime) else None))
    return out


def leer_seguros(anio):
    arch = "auxiliares_2025.xlsx" if anio == 2025 else "auxiliares_2026.xlsx"
    ws = openpyxl.load_workbook(os.path.join(INS, "andina", arch), data_only=True)["Seguros"]
    out = []
    for r in ws.iter_rows(min_row=5, values_only=True):
        if r[0] in (None, "TOTAL"):
            break
        pol, _a, ramo, cob, bienes, vd, vh, fpago, prima, prima_ant, nd_ini, dev_ant, meses = r[:13]
        fun = r[16]
        fpago_ant = dt.datetime(fpago.year - 1, fpago.month, fpago.day)
        out.append(dict(pol=pol, desc=f"{ramo} — {cob}", fun=fun, fpago_ant=fpago_ant, nd_ini=nd_ini,
                        fpago=fpago, prima=prima, meses=meses))
    return out


def leer_merc(anio):
    arch = "auxiliares_complementarios_2025.xlsx" if anio == 2025 else "auxiliares_2026.xlsx"
    ws = openpyxl.load_workbook(os.path.join(INS, "andina", arch))["Mercaderías"]
    compras = [(ws.cell(r, 1).value, ws.cell(r, 3).value) for r in range(6, 18)]
    ei = [(ws.cell(r, 1).value, ws.cell(r, 2).value) for r in (22, 23)]
    ef = [(ws.cell(r, 1).value, ws.cell(r, 2).value) for r in (28, 29)]
    matriz = [(ws.cell(r, 1).value, [ws.cell(r, c).value for c in range(2, 14)]) for r in range(34, 47)]
    return dict(compras=compras, ei=ei, ef=ef, matriz=matriz)


def leer_resultados(anio):
    arch = "auxiliares_complementarios_2025.xlsx" if anio == 2025 else "auxiliares_2026.xlsx"
    ws = openpyxl.load_workbook(os.path.join(INS, "andina", arch))["Resultados mensuales"]
    out = []
    for r in ws.iter_rows(min_row=5, values_only=True):
        if r[0] is None or not str(r[0])[0] in "45":
            break
        out.append(dict(cod=r[0], cuenta=r[1], rubro=r[2], signo=r[3], meses=list(r[4:16])))
    return out


# ----------------------------------------------------------------------------- PN: historia (insumo)
CAPITAL = (dt.datetime(2014, 3, 14), 20000000)
RESERVA = [(dt.datetime(2015, 4, 28), 92500), (dt.datetime(2016, 4, 28), 162000), (dt.datetime(2017, 4, 28), 235500),
           (dt.datetime(2018, 4, 27), 349000), (dt.datetime(2019, 4, 26), 521000), (dt.datetime(2020, 4, 28), 817500),
           (dt.datetime(2021, 4, 28), 1195000), (dt.datetime(2022, 4, 28), 627500)]
RESULTADOS_HIST = {2014: 1850000, 2015: 3240000, 2016: 4710000, 2017: 6980000, 2018: 10420000, 2019: 16350000,
                   2020: 23900000, 2021: 41600000, 2022: 120000000, 2023: 330000000, 2024: 702703500}
DIVIDENDOS_HIST = [(dt.datetime(2019, 4, 30), 4000000), (dt.datetime(2020, 4, 30), 6000000),
                   (dt.datetime(2021, 4, 30), 10000000), (dt.datetime(2022, 4, 29), 18000000),
                   (dt.datetime(2023, 4, 28), 30000000), (dt.datetime(2024, 4, 30), 60000000)]
DIVIDENDO = {2025: (dt.datetime(2025, 4, 28), 120000000, "Asamblea 28/04/2025 (pago 30/04/2025), sobre resultado 2024"),
             2026: (dt.datetime(2026, 4, 28), 300000000, "Asamblea 28/04/2026 (pago 30/04/2026), sobre resultado 2025")}

COEF_LOOK = "INDEX(Coeficientes!$D$7:$D$400,MATCH({m},Coeficientes!$A$7:$A$400,0))"
IDX_LOOK = "INDEX(Coeficientes!$B$7:$B$400,MATCH({m},Coeficientes!$A$7:$A$400,0))"


def coef_f(cell_mes):
    return COEF_LOOK.format(m=cell_mes)


# ============================================================================= libro anual
def libro_anual(anio, inicio=None):
    """inicio: dict cod -> saldo de cierre 2025 ajustado en moneda dic-2025 (solo para 2026)."""
    cierre = dt.datetime(anio, 12, 1)
    previo = dt.datetime(anio - 1, 12, 1)
    idx = leer_indices()
    bss = leer_bss(anio)
    bu = leer_bu(anio)
    seg = leer_seguros(anio)
    merc = leer_merc(anio)
    res = leer_resultados(anio)
    wb = openpyxl.Workbook()

    # ------------------------------------------------------------------ Notas (se completa al final)
    wsN = wb.active
    wsN.title = "Notas"

    # ------------------------------------------------------------------ Coeficientes
    ws = wb.create_sheet("Coeficientes")
    titulo(ws, f"Índices y coeficientes de reexpresión a moneda de dic-{anio}",
           "Fuente: insumos/indices/indices_facpce.csv (Res. JG FACPCE 539/18). Coeficiente = índice del mes de cierre / índice del mes de origen.")
    put(ws, "A3", "Mes de cierre (moneda de medición)")
    put(ws, "C3", cierre, MES, F_IN)
    put(ws, "A4", "Índice del mes de cierre")
    put(ws, "C4", "=" + IDX_LOOK.format(m="C3"), '#,##0.0000')
    put(ws, "D3", "Coef. dic-previo → cierre")
    put(ws, "E3", f"=C4/{IDX_LOOK.format(m='E4')}", COEF, F_B)
    put(ws, "D4", "Mes previo")
    put(ws, "E4", previo, MES, F_IN)
    hdr(ws, 6, ["Mes", "Índice FACPCE", "Estado", f"Coef. a dic-{anio}", "Variación mensual"])
    r0 = 7
    filas = [x for x in idx if x[0] <= cierre]
    for i, (m, v, e) in enumerate(filas):
        r = r0 + i
        put(ws, f"A{r}", m, MES, F_IN)
        put(ws, f"B{r}", v, '#,##0.0000', F_IN)
        put(ws, f"C{r}", e, font=F_IN)
        put(ws, f"D{r}", f"=$C$4/B{r}", COEF)
        if i > 0:
            put(ws, f"E{r}", f"=B{r}/B{r-1}-1", PCT)
    widths(ws, [14, 16, 12, 16, 16])
    ws.freeze_panes = "A7"

    # ------------------------------------------------------------------ BSS histórico (insumo)
    ws = wb.create_sheet("BSS histórico")
    titulo(ws, f"BSS {anio} en valores históricos (insumo)", f"Copia de insumos/andina/bss_{anio}.csv. Deudor (+) / acreedor (−).")
    hdr(ws, 4, ["Código", "Cuenta", "Rubro", "Clase", f"Saldo inicial 31/12/{anio-1}", f"Débitos {anio}",
                f"Créditos {anio}", f"Saldo final 31/12/{anio}"])
    fila_bss = {}
    for i, b in enumerate(bss):
        r = 5 + i
        fila_bss[b["cod"]] = r
        for col, v in zip("ABC", (b["cod"], b["cuenta"], b["rubro"])):
            put(ws, f"{col}{r}", v, font=F_IN)
        put(ws, f"D{r}", clase(b["cod"]))
        put(ws, f"E{r}", b["si"], NUM, F_IN)
        put(ws, f"F{r}", b["d"], NUM, F_IN)
        put(ws, f"G{r}", b["c"], NUM, F_IN)
        put(ws, f"H{r}", f"=E{r}+F{r}-G{r}", NUM)
    rt = 5 + len(bss)
    put(ws, f"B{rt}", "Totales", font=F_B)
    for col in "EFGH":
        put(ws, f"{col}{rt}", f"=SUM({col}5:{col}{rt-1})", NUM, F_B, FILL_TOT)
    put(ws, f"B{rt+1}", "Control: saldo final calculado vs. insumo (debe ser 0)")
    put(ws, f"H{rt+1}", "=" + "+".join(f"ABS(H{fila_bss[b['cod']]}-{b['sf']})" for b in bss), NUM, fill=FILL_CTRL)
    widths(ws, [9, 46, 26, 7, 18, 18, 18, 18])
    ws.freeze_panes = "C5"
    BH = lambda cod, col: f"'BSS histórico'!{col}{fila_bss[cod]}"  # noqa: E731

    # ------------------------------------------------------------------ Bienes de uso
    ws = wb.create_sheet("Bienes de uso")
    titulo(ws, f"Bienes de uso reexpresados a moneda de dic-{anio}",
           "Bien por bien. Coeficiente desde el mes de alta. Valor de origen, amortización acumulada y del ejercicio con el mismo coeficiente (método de la RT 6: reexpresión del costo y de su amortización).")
    cols = ["Código", "Descripción", "Cuenta VO", "Cuenta AA", "Cuenta amort. (resultado)", "Función",
            "Fecha de origen", "Mes de origen", f"Coef. a dic-{anio}", "Vida útil (años)",
            f"VO inicio 31/12/{anio-1} (hist.)", "Altas (hist.)", "Bajas VO (hist.)", f"VO cierre (hist.)",
            f"AA inicio (hist.)", "Bajas AA (hist.)", f"Amortización {anio} (hist.)", "AA cierre (hist.)",
            "VO inicio (aj.)", "Altas (aj.)", "Bajas VO (aj.)", "VO cierre (aj.)", "AA inicio (aj.)",
            "Bajas AA (aj.)", f"Amortización {anio} (aj.)", "AA cierre (aj.)", "VRC cierre (aj.)",
            "Precio de venta (hist.)", "Fecha de baja", "Coef. mes de baja", "Resultado venta (aj.)"]
    hdr(ws, 4, cols)
    r1 = 5
    for i, b in enumerate(bu):
        r = r1 + i
        vo_c, aa_c = CTA_VO[b["cta"]]
        put(ws, f"A{r}", b["cod"], font=F_IN)
        put(ws, f"B{r}", b["desc"], font=F_IN)
        put(ws, f"C{r}", vo_c)
        put(ws, f"D{r}", aa_c or "")
        put(ws, f"E{r}", cta_amort(b["cta"], b["fun"]) or "")
        put(ws, f"F{r}", b["fun"], font=F_IN)
        put(ws, f"G{r}", b["fecha"], FECHA, F_IN)
        put(ws, f"H{r}", f"=DATE(YEAR(G{r}),MONTH(G{r}),1)", MES)
        put(ws, f"I{r}", "=" + coef_f(f"H{r}"), COEF)
        put(ws, f"J{r}", b["vu"], font=F_IN)
        put(ws, f"K{r}", b["vo_ini"], NUM, F_IN)
        put(ws, f"L{r}", b["alta"], NUM, F_IN)
        put(ws, f"M{r}", b["baja_vo"], NUM, F_IN)
        put(ws, f"N{r}", f"=K{r}+L{r}-M{r}", NUM)
        put(ws, f"O{r}", b["aa_ini"], NUM, F_IN)
        put(ws, f"P{r}", b["baja_aa"], NUM, F_IN)
        put(ws, f"Q{r}", b["am"], NUM, F_IN)
        put(ws, f"R{r}", f"=O{r}-P{r}+Q{r}", NUM)
        for src, dst in zip("KLMOPQ", "STUWXY"):
            put(ws, f"{dst}{r}", f"={src}{r}*$I{r}", NUM)
        put(ws, f"V{r}", f"=S{r}+T{r}-U{r}", NUM)
        put(ws, f"Z{r}", f"=W{r}-X{r}+Y{r}", NUM)
        put(ws, f"AA{r}", f"=V{r}-Z{r}", NUM)
        put(ws, f"AB{r}", b["precio"], NUM, F_IN)
        if b["fbaja"]:
            put(ws, f"AC{r}", b["fbaja"], FECHA, F_IN)
            put(ws, f"AD{r}", "=" + coef_f(f"DATE(YEAR(AC{r}),MONTH(AC{r}),1)"), COEF)
            put(ws, f"AE{r}", f"=AB{r}*AD{r}-(U{r}-X{r})", NUM)
        else:
            put(ws, f"AE{r}", 0, NUM)
    rn = r1 + len(bu) - 1
    rt = rn + 1
    put(ws, f"A{rt}", "TOTAL", font=F_B)
    for c in range(11, 32):
        col = L(c)
        if col in ("AC", "AD"):
            continue
        put(ws, f"{col}{rt}", f"=SUM({col}{r1}:{col}{rn})", NUM, F_B, FILL_TOT)
    put(ws, f"A{rt+1}", "Control: VO + altas − bajas − AA y amortización recalculada (aj.) — ver hoja Controles")
    widths(ws, [9, 40, 8, 8, 9, 15, 11, 10, 10, 6] + [15] * 21)
    ws.freeze_panes = "C5"
    BU = dict(r1=r1, rn=rn, rt=rt)

    def bu_sum(col, crit_col, crit):
        return f"SUMIF('Bienes de uso'!${crit_col}${r1}:${crit_col}${rn},\"{crit}\",'Bienes de uso'!{col}${r1}:{col}${rn})"

    # ------------------------------------------------------------------ Seguros
    ws = wb.create_sheet("Seguros")
    titulo(ws, f"Seguros pagados por adelantado y gasto de seguros reexpresados a moneda de dic-{anio}",
           "Coeficiente desde el mes de pago de la prima. El devengado conserva el coeficiente de la prima que lo origina. "
           "Supuesto (S): la póliza anterior se pagó el mismo día del año previo.")
    hdr(ws, 4, ["Póliza", "Ramo / cobertura", "Función", "Cuenta gasto", "Pago póliza anterior", "Coef. pago anterior",
                f"No devengado 31/12/{anio-1} (hist.)", "Devengado de póliza anterior (hist.)", "Pago póliza actual",
                "Coef. pago actual", "Prima actual (hist.)", "Meses devengados", "Devengado póliza actual (hist.)",
                f"Gasto {anio} (hist.)", f"No devengado 31/12/{anio} (hist.)", "No devengado inicio (aj.)",
                "Devengado póliza anterior (aj.)", "Prima actual (aj.)", "Devengado póliza actual (aj.)",
                f"Gasto {anio} (aj.)", "No devengado cierre (aj.)"])
    s1 = 5
    for i, s in enumerate(seg):
        r = s1 + i
        put(ws, f"A{r}", s["pol"], font=F_IN)
        put(ws, f"B{r}", s["desc"], font=F_IN)
        put(ws, f"C{r}", s["fun"], font=F_IN)
        put(ws, f"D{r}", "5.3.05" if s["fun"] == "Administración" else "5.2.06")
        put(ws, f"E{r}", s["fpago_ant"], FECHA, F_IN)
        put(ws, f"F{r}", "=" + coef_f(f"DATE(YEAR(E{r}),MONTH(E{r}),1)"), COEF)
        put(ws, f"G{r}", s["nd_ini"], NUM, F_IN)
        put(ws, f"H{r}", f"=G{r}", NUM)
        put(ws, f"I{r}", s["fpago"], FECHA, F_IN)
        put(ws, f"J{r}", "=" + coef_f(f"DATE(YEAR(I{r}),MONTH(I{r}),1)"), COEF)
        put(ws, f"K{r}", s["prima"], NUM, F_IN)
        put(ws, f"L{r}", s["meses"], font=F_IN)
        put(ws, f"M{r}", f"=K{r}*L{r}/12", NUM)
        put(ws, f"N{r}", f"=H{r}+M{r}", NUM)
        put(ws, f"O{r}", f"=K{r}-M{r}", NUM)
        put(ws, f"P{r}", f"=G{r}*F{r}", NUM)
        put(ws, f"Q{r}", f"=H{r}*F{r}", NUM)
        put(ws, f"R{r}", f"=K{r}*J{r}", NUM)
        put(ws, f"S{r}", f"=M{r}*J{r}", NUM)
        put(ws, f"T{r}", f"=Q{r}+S{r}", NUM)
        put(ws, f"U{r}", f"=O{r}*J{r}", NUM)
    sn = s1 + len(seg) - 1
    st = sn + 1
    put(ws, f"A{st}", "TOTAL", font=F_B)
    for col in "GHKMNOPQRSTU":
        put(ws, f"{col}{st}", f"=SUM({col}{s1}:{col}{sn})", NUM, F_B, FILL_TOT)
    widths(ws, [8, 38, 15, 8, 11, 10] + [15] * 15)
    ws.freeze_panes = "C5"
    SG = dict(s1=s1, sn=sn, st=st)

    def seg_sum(col, cta):
        return f"SUMIF(Seguros!$D${s1}:$D${sn},\"{cta}\",Seguros!{col}${s1}:{col}${sn})"

    # ------------------------------------------------------------------ Mercaderías
    ws = wb.create_sheet("Mercaderías")
    titulo(ws, f"Mercaderías de reventa reexpresadas a moneda de dic-{anio} (costo FIFO por mes de compra)",
           "Cada capa conserva su mes de compra. CMV reexpresado por mes de compra (no por mes de venta). "
           "Supuesto: costo de reposición no inferior al costo reexpresado (sin prueba de VNR en esta fase).")
    hdr(ws, 4, ["A. Existencia inicial por mes de compra", f"Importe 31/12/{anio-1} (hist.)", "Coef.", "Importe (aj.)"])
    for i, (m, v) in enumerate(merc["ei"]):
        r = 5 + i
        put(ws, f"A{r}", m, MES, F_IN)
        put(ws, f"B{r}", v, NUM, F_IN)
        put(ws, f"C{r}", "=" + coef_f(f"A{r}"), COEF)
        put(ws, f"D{r}", f"=B{r}*C{r}", NUM)
    put(ws, "A7", "Total existencia inicial", font=F_B)
    put(ws, "B7", "=SUM(B5:B6)", NUM, F_B, FILL_TOT)
    put(ws, "D7", "=SUM(D5:D6)", NUM, F_B, FILL_TOT)

    hdr(ws, 9, ["B. Compras por mes", "Compras (hist.)", "Coef.", "Compras (aj.)"])
    for i, (m, v) in enumerate(merc["compras"]):
        r = 10 + i
        put(ws, f"A{r}", m, MES, F_IN)
        put(ws, f"B{r}", v, NUM, F_IN)
        put(ws, f"C{r}", "=" + coef_f(f"A{r}"), COEF)
        put(ws, f"D{r}", f"=B{r}*C{r}", NUM)
    put(ws, "A22", "Total compras", font=F_B)
    put(ws, "B22", "=SUM(B10:B21)", NUM, F_B, FILL_TOT)
    put(ws, "D22", "=SUM(D10:D21)", NUM, F_B, FILL_TOT)

    hdr(ws, 24, ["C. Existencia final por mes de compra", f"Importe 31/12/{anio} (hist.)", "Coef.", "Importe (aj.)"])
    for i, (m, v) in enumerate(merc["ef"]):
        r = 25 + i
        put(ws, f"A{r}", m, MES, F_IN)
        put(ws, f"B{r}", v, NUM, F_IN)
        put(ws, f"C{r}", "=" + coef_f(f"A{r}"), COEF)
        put(ws, f"D{r}", f"=B{r}*C{r}", NUM)
    put(ws, "A27", "Total existencia final", font=F_B)
    put(ws, "B27", "=SUM(B25:B26)", NUM, F_B, FILL_TOT)
    put(ws, "D27", "=SUM(D25:D26)", NUM, F_B, FILL_TOT)

    meses_txt = [dt.date(anio, m, 1).strftime("%m-%Y") for m in range(1, 13)]
    hdr(ws, 29, ["D. CMV por mes de compra (filas) y mes de venta (columnas)"] + [f"Venta {t} (hist.)" for t in meses_txt]
        + ["Total (hist.)", "Coef. mes de compra", "CMV (aj.)"])
    for i, (m, vals) in enumerate(merc["matriz"]):
        r = 30 + i
        put(ws, f"A{r}", m, MES, F_IN)
        for j, v in enumerate(vals):
            put(ws, f"{L(2+j)}{r}", v, NUM, F_IN)
        put(ws, f"N{r}", f"=SUM(B{r}:M{r})", NUM)
        put(ws, f"O{r}", "=" + coef_f(f"A{r}"), COEF)
        put(ws, f"P{r}", f"=N{r}*O{r}", NUM)
    put(ws, "A43", "Total CMV (hist.)", font=F_B)
    for c in range(2, 15):
        put(ws, f"{L(c)}43", f"=SUM({L(c)}30:{L(c)}42)", NUM, F_B, FILL_TOT)
    put(ws, "P43", "=SUM(P30:P42)", NUM, F_B, FILL_TOT)
    put(ws, "A44", "CMV (aj.) por mes de venta", font=F_B)
    for c in range(2, 14):
        put(ws, f"{L(c)}44", f"=SUMPRODUCT({L(c)}30:{L(c)}42,$O$30:$O$42)", NUM, F_B)
    put(ws, "N44", "=SUM(B44:M44)", NUM, F_B)
    put(ws, "A46", "Control (aj.): EI + compras − CMV − EF (debe ser 0)", font=F_B)
    put(ws, "D46", "=D7+D22-P43-D27", NUM, fill=FILL_CTRL)
    put(ws, "A47", "Control (hist.): EI + compras − CMV − EF (debe ser 0)", font=F_B)
    put(ws, "D47", "=B7+B22-N43-B27", NUM, fill=FILL_CTRL)
    widths(ws, [44] + [15] * 15)

    # ------------------------------------------------------------------ Resultados (mensuales)
    ws = wb.create_sheet("Resultados")
    titulo(ws, f"Resultados {anio} reexpresados a moneda de dic-{anio}",
           "Partidas que se devengan en el mes (ventas, gastos monetarios, intereses, impuestos): importe del mes × coeficiente del mes. "
           "CMV, amortizaciones, seguros y resultado por venta de bienes de uso: se toman de su auxiliar (coeficiente de la partida que se consume).")
    hdr(ws, 4, ["Código", "Cuenta", "Rubro", "Signo"] + [f"{t} (hist.)" for t in meses_txt]
        + ["Total (hist.)", "BSS (hist.)", "Dif.", "Método", "Total (aj.)", "Ajuste (aj. − hist.)"])
    put(ws, "B5", "Mes", font=F_B)
    put(ws, "B6", "Coeficiente del mes", font=F_B)
    for j in range(12):
        put(ws, f"{L(5+j)}5", dt.datetime(anio, j + 1, 1), MES)
        put(ws, f"{L(5+j)}6", "=" + coef_f(f"{L(5+j)}5"), COEF)
    fila_res = {}
    ESPECIAL = {"5.1.01": "Mercaderías (CMV por mes de compra)", "5.2.06": "Seguros (por mes de pago)",
                "5.3.05": "Seguros (por mes de pago)", "5.2.07": "Bienes de uso (por mes de alta)",
                "5.2.08": "Bienes de uso (por mes de alta)", "5.2.09": "Bienes de uso (por mes de alta)",
                "5.3.06": "Bienes de uso (por mes de alta)", "5.3.07": "Bienes de uso (por mes de alta)",
                "4.3.01": "Bienes de uso (precio × coef. mes de venta − VRC aj.)"}
    for i, x in enumerate(res):
        r = 7 + i
        fila_res[x["cod"]] = r
        put(ws, f"A{r}", x["cod"], font=F_IN)
        put(ws, f"B{r}", x["cuenta"], font=F_IN)
        put(ws, f"C{r}", x["rubro"], font=F_IN)
        put(ws, f"D{r}", x["signo"], font=F_IN)
        for j, v in enumerate(x["meses"]):
            put(ws, f"{L(5+j)}{r}", v, NUM, F_IN)
        put(ws, f"Q{r}", f"=SUM(E{r}:P{r})", NUM)
        put(ws, f"R{r}", f"=ABS({BH(x['cod'], 'H')})", NUM)
        put(ws, f"S{r}", f"=Q{r}-R{r}", NUM)
        cod = x["cod"]
        if cod == "5.1.01":
            f = "=Mercaderías!P43"
        elif cod in ("5.2.06", "5.3.05"):
            f = "=" + seg_sum("T", cod)
        elif cod in ("5.2.07", "5.2.08", "5.2.09", "5.3.06", "5.3.07"):
            f = "=" + bu_sum("Y", "E", cod)
        elif cod == "4.3.01":
            f = f"='Bienes de uso'!AE{BU['rt']}"
        else:
            f = f"=SUMPRODUCT(E{r}:P{r},$E$6:$P$6)"
        put(ws, f"T{r}", ESPECIAL.get(cod, "Mensual × coef. del mes"))
        put(ws, f"U{r}", f, NUM)
        put(ws, f"V{r}", f"=U{r}-Q{r}", NUM)
    rlast = 7 + len(res) - 1
    rr = rlast + 2
    put(ws, f"B{rr}", "Resultado antes del RECPAM", font=F_B)
    put(ws, f"Q{rr}", f"=SUMPRODUCT($D$7:$D${rlast},Q7:Q{rlast})", NUM, F_B, FILL_TOT)
    put(ws, f"U{rr}", f"=SUMPRODUCT($D$7:$D${rlast},U7:U{rlast})", NUM, F_B, FILL_TOT)
    put(ws, f"V{rr}", f"=U{rr}-Q{rr}", NUM, F_B, FILL_TOT)
    put(ws, f"B{rr+1}", "RECPAM (hoja RECPAM, por diferencia)", font=F_B)
    put(ws, f"U{rr+1}", "=RECPAM!C12", NUM, F_B)
    put(ws, f"B{rr+2}", "Resultado del ejercicio ajustado", font=F_B)
    put(ws, f"Q{rr+2}", f"=Q{rr}", NUM, F_B, FILL_TOT)
    put(ws, f"U{rr+2}", f"=U{rr}+U{rr+1}", NUM, F_B, FILL_TOT)
    # Compras (dato para el método directo)
    rc = rr + 4
    put(ws, f"B{rc}", "Compras de mercaderías (débitos 1.5.01) — dato para el RECPAM directo", font=F_IN)
    for j in range(12):
        put(ws, f"{L(5+j)}{rc}", f"=Mercaderías!B{10+j}", NUM)
    put(ws, f"Q{rc}", f"=SUM(E{rc}:P{rc})", NUM)
    widths(ws, [9, 42, 24, 6] + [14] * 12 + [16, 16, 8, 34, 16, 16])
    ws.freeze_panes = "E7"
    RES = dict(rlast=rlast, rr=rr)

    # ------------------------------------------------------------------ Patrimonio neto
    ws = wb.create_sheet("Patrimonio neto")
    titulo(ws, f"Patrimonio neto reexpresado a moneda de dic-{anio} (por componente y fecha de origen)",
           "Capital y ajuste de capital por separado. Reserva legal por fecha de constitución. "
           + ("RNA al inicio por diferencia (RT 6 / Res. JG 539/18, primer ejercicio de aplicación)." if anio == 2025
              else "RNA al inicio = cierre 2025 ajustado × coeficiente dic-2025 → dic-2026."))
    hdr(ws, 4, ["Componente", "Fecha de origen", "Mes", "Coef.", "Importe (hist.)", "Importe (aj.)", "Ajuste"])
    put(ws, "A5", "Capital suscripto e integrado", font=F_B)
    put(ws, "B5", CAPITAL[0], FECHA, F_IN)
    put(ws, "C5", "=DATE(YEAR(B5),MONTH(B5),1)", MES)
    put(ws, "D5", "=" + coef_f("C5"), COEF)
    put(ws, "E5", CAPITAL[1], NUM, F_IN)
    put(ws, "F5", "=E5*D5", NUM)
    put(ws, "G5", "=F5-E5", NUM)
    put(ws, "A6", "  de lo cual: Capital suscripto (3.1.01, nominal)")
    put(ws, "F6", "=E5", NUM)
    put(ws, "A7", "  de lo cual: Ajuste de capital (3.1.02)")
    put(ws, "F7", "=G5", NUM)
    put(ws, "A9", "Reserva legal (art. 70 LGS) — capas por asamblea", font=F_B)
    for i, (f, v) in enumerate(RESERVA):
        r = 10 + i
        put(ws, f"A{r}", f"Constitución asamblea {f.year}")
        put(ws, f"B{r}", f, FECHA, F_IN)
        put(ws, f"C{r}", f"=DATE(YEAR(B{r}),MONTH(B{r}),1)", MES)
        put(ws, f"D{r}", "=" + coef_f(f"C{r}"), COEF)
        put(ws, f"E{r}", v, NUM, F_IN)
        put(ws, f"F{r}", f"=E{r}*D{r}", NUM)
        put(ws, f"G{r}", f"=F{r}-E{r}", NUM)
    put(ws, "A18", "Total reserva legal", font=F_B)
    for col in "EFG":
        put(ws, f"{col}18", f"=SUM({col}10:{col}17)", NUM, F_B, FILL_TOT)

    put(ws, "A20", "Resultados no asignados", font=F_B)
    if anio == 2025:
        put(ws, "A21", "RNA al 31/12/2024 (aj.): por diferencia = activo aj. − pasivo − capital aj. − reserva aj.")
        put(ws, "E21", f"=-{BH('3.3.01', 'E')}", NUM)
        put(ws, "F21", "=-'BSS ajustado'!E{rna}", NUM)  # se completa abajo con la fila real
    else:
        put(ws, "A21", "RNA al 31/12/2025 (aj.): cierre 2025 ajustado (incluye resultado 2025) × coef. dic-25 → dic-26")
        put(ws, "E21", f"=-{BH('3.3.01', 'E')}", NUM)
        put(ws, "F21", "=-Inicio!F{rna}", NUM)
    put(ws, "G21", "=F21-E21", NUM)
    fd, vd, td = DIVIDENDO[anio]
    put(ws, "A22", f"Dividendos en efectivo — {td}")
    put(ws, "B22", fd, FECHA, F_IN)
    put(ws, "C22", "=DATE(YEAR(B22),MONTH(B22),1)", MES)
    put(ws, "D22", "=" + coef_f("C22"), COEF)
    put(ws, "E22", -vd, NUM, F_IN)
    put(ws, "F22", "=E22*D22", NUM)
    put(ws, "G22", "=F22-E22", NUM)
    put(ws, "A23", f"RNA al cierre antes del resultado {anio}", font=F_B)
    put(ws, "E23", "=E21+E22", NUM, F_B, FILL_TOT)
    put(ws, "F23", "=F21+F22", NUM, F_B, FILL_TOT)
    put(ws, "A24", f"Resultado del ejercicio {anio} (incluye RECPAM)")
    put(ws, "E24", f"=Resultados!Q{RES['rr']+2}", NUM)
    put(ws, "F24", f"=Resultados!U{RES['rr']+2}", NUM)
    put(ws, "A25", f"RNA al 31/12/{anio} después del resultado", font=F_B)
    put(ws, "E25", "=E23+E24", NUM, F_B, FILL_TOT)
    put(ws, "F25", "=F23+F24", NUM, F_B, FILL_TOT)
    put(ws, "A27", f"Total patrimonio neto al 31/12/{anio}", font=F_B)
    put(ws, "E27", "=E5+E18+E25", NUM, F_B, FILL_TOT)
    put(ws, "F27", "=F5+F18+F25", NUM, F_B, FILL_TOT)
    put(ws, "A28", f"Total patrimonio neto al 31/12/{anio-1} (en moneda de dic-{anio})", font=F_B)
    put(ws, "E28", "=E5+E18+E21" if anio == 2025 else f"=E5+E18+E21", NUM, F_B)
    put(ws, "F28", "=F5+F18+F21", NUM, F_B)

    # Reserva legal: análisis art. 70 LGS (Duda)
    put(ws, "A31", "Duda — Reserva legal (art. 70 LGS): análisis en moneda homogénea", font=F_B)
    put(ws, "A32", "Base: capital suscripto + ajuste de capital (aj.)")
    put(ws, "F32", "=F5", NUM)
    put(ws, "A33", "Tope 20 % de la base")
    put(ws, "F33", "=0.2*F32", NUM)
    put(ws, "A34", "Reserva legal reexpresada (aj.)")
    put(ws, "F34", "=F18", NUM)
    put(ws, "A35", "Faltante para el tope (si es positivo)")
    put(ws, "F35", "=MAX(0,F33-F34)", NUM, F_B)
    if anio == 2025:
        put(ws, "A36", "Asignación máxima del ejercicio: 5 % del resultado 2025 ajustado (a decidir en la asamblea de 2026)")
        put(ws, "F36", "=0.05*MAX(0,F24)", NUM)
        put(ws, "A37", "Asignación propuesta = mínimo entre faltante y 5 % del resultado (NO contabilizada; ver Notas)")
        put(ws, "F37", "=MIN(F35,F36)", NUM, F_B)
    else:
        put(ws, "A36", "Asignación omitida en la asamblea 28/04/2026: 5 % del resultado 2025 ajustado, en moneda dic-2026")
        put(ws, "F36", "=0.05*Inicio!H3", NUM)
        put(ws, "A37", "Asignación máxima sobre el resultado 2026: 5 % del resultado 2026 ajustado")
        put(ws, "F37", "=0.05*MAX(0,F24)", NUM)
        put(ws, "A38", "Asignación propuesta para la asamblea 2027 = mínimo entre faltante y (omitida 2025 + 5 % de 2026) (NO contabilizada)")
        put(ws, "F38", "=MIN(F35,F36+F37)", NUM, F_B)

    if anio == 2025:
        # Informativo: RNA por capas reexpresadas vs. RNA por diferencia
        put(ws, "A41", "Informativo — RNA al 31/12/2024 reexpresado por capas (resultado de cada ejercicio, reserva y dividendos por fecha)", font=F_B)
        hdr(ws, 42, ["Movimiento", "Fecha", "Mes", "Coef.", "Importe (hist.)", "Importe (aj.)"])
        movs = [(f"Resultado {y}", dt.datetime(y, 12, 31), v) for y, v in RESULTADOS_HIST.items()]
        movs += [(f"Reserva legal (asamblea {f.year})", f, -v) for f, v in RESERVA]
        movs += [(f"Dividendos pagados {f.year}", f, -v) for f, v in DIVIDENDOS_HIST]
        for i, (t, f, v) in enumerate(movs):
            r = 43 + i
            put(ws, f"A{r}", t)
            put(ws, f"B{r}", f, FECHA, F_IN)
            put(ws, f"C{r}", f"=DATE(YEAR(B{r}),MONTH(B{r}),1)", MES)
            put(ws, f"D{r}", "=" + coef_f(f"C{r}"), COEF)
            put(ws, f"E{r}", v, NUM, F_IN)
            put(ws, f"F{r}", f"=E{r}*D{r}", NUM)
        rl = 43 + len(movs)
        put(ws, f"A{rl}", "RNA 31/12/2024 por capas", font=F_B)
        put(ws, f"E{rl}", f"=SUM(E43:E{rl-1})", NUM, F_B, FILL_TOT)
        put(ws, f"F{rl}", f"=SUM(F43:F{rl-1})", NUM, F_B, FILL_TOT)
        put(ws, f"A{rl+1}", "Control: capas (hist.) vs. BSS (debe ser 0)")
        put(ws, f"E{rl+1}", f"=E{rl}-E21", NUM, fill=FILL_CTRL)
        put(ws, f"A{rl+2}", "Diferencia RNA por diferencia − RNA por capas (aj.): RECPAM y reexpresiones de ejercicios anteriores no reconocidos en los resultados históricos (informativo, no se contabiliza)")
        put(ws, f"F{rl+2}", f"=F21-F{rl}", NUM, F_B)
    widths(ws, [70, 12, 10, 10, 18, 18, 18])

    # ------------------------------------------------------------------ Inicio (solo 2026)
    if anio == 2026:
        wsI = wb.create_sheet("Inicio")
        titulo(wsI, "Saldos iniciales 2026 = cierre 2025 ajustado × coeficiente dic-2025 → dic-2026",
               "Columna D: valores importados de ajuste_2025.xlsx, hoja 'BSS ajustado', saldo de cierre ajustado (moneda dic-2025), con el resultado 2025 (incl. RECPAM) cerrado contra RNA.")
        put(wsI, "G3", "Resultado 2025 ajustado (moneda dic-2026)")
        hdr(wsI, 4, ["Código", "Cuenta", "Clase", "Cierre 2025 aj. (moneda dic-2025)", "Coef. dic-25 → dic-26",
                     "Inicio 2026 (moneda dic-2026)"])
        fila_ini = {}
        codigos = [k for k in inicio if k != "_resultado"]
        for i, cod in enumerate(codigos):
            r = 5 + i
            fila_ini[cod] = r
            put(wsI, f"A{r}", cod, font=F_IN)
            put(wsI, f"B{r}", inicio_nombres.get(cod, ""), font=F_IN)
            put(wsI, f"C{r}", clase(cod))
            put(wsI, f"D{r}", inicio[cod], NUM, F_IN)
            put(wsI, f"E{r}", "=Coeficientes!$E$3", COEF)
            put(wsI, f"F{r}", f"=D{r}*E{r}", NUM)
        # Capital suscripto se mantiene nominal; el ajuste de capital absorbe toda la reexpresión
        a, b = fila_ini["3.1.01"], fila_ini["3.1.02"]
        put(wsI, f"F{a}", f"=D{a}", NUM)
        put(wsI, f"F{b}", f"=(D{a}+D{b})*E{b}-D{a}", NUM)
        put(wsI, f"G{a}", "Capital suscripto: valor nominal (no se reexpresa)")
        put(wsI, f"G{b}", "Ajuste de capital: (capital + ajuste) × coef. − capital nominal")
        ri = 5 + len(codigos)
        put(wsI, f"B{ri}", "Total (debe ser 0)", font=F_B)
        put(wsI, f"D{ri}", f"=SUM(D5:D{ri-1})", NUM, F_B, FILL_CTRL)
        put(wsI, f"F{ri}", f"=SUM(F5:F{ri-1})", NUM, F_B, FILL_CTRL)
        put(wsI, "H3", f"={inicio['_resultado']}*Coeficientes!$E$3", NUM, F_IN)
        put(wsI, f"B{ri+1}", "PN inicial 2026 (aj.)", font=F_B)
        put(wsI, f"F{ri+1}", "=-(" + "+".join(f"F{fila_ini[c]}" for c in codigos if c.startswith("3.")) + ")", NUM, F_B)
        put(wsI, f"D{ri+1}", "=-(" + "+".join(f"D{fila_ini[c]}" for c in codigos if c.startswith("3.")) + ")", NUM, F_B)
        widths(wsI, [9, 46, 7, 20, 12, 20, 34, 18])
        wb["Patrimonio neto"]["F21"].value = f"=-Inicio!F{fila_ini['3.3.01']}"

    # ------------------------------------------------------------------ BSS ajustado
    ws = wb.create_sheet("BSS ajustado")
    titulo(ws, f"BSS ajustado — saldos al 31/12/{anio-1} y al 31/12/{anio} en moneda de dic-{anio}",
           "Deudor (+) / acreedor (−). Monetarias: nominal (inicio × coef. dic-previo → cierre). No monetarias: desde su auxiliar. "
           "RECPAM: cuenta de cuadre (método por diferencia, RT 6).")
    hdr(ws, 4, ["Código", "Cuenta", "Rubro", "Clase", f"Saldo inicial (hist.)", f"Saldo inicial (aj., moneda dic-{anio})",
                "Saldo final (hist.)", "Saldo final (aj.)", "Ajuste final (aj. − hist.)", "Origen del saldo ajustado"])
    cuentas = []
    for b in bss:
        cuentas.append((b["cod"], b["cuenta"], b["rubro"]))
        if b["cod"] == "3.1.01":
            cuentas.append(("3.1.02", "Ajuste de capital", "Patrimonio neto"))
    cuentas.append(("4.9.01", "RECPAM (resultado por exposición al cambio en el poder adquisitivo de la moneda)",
                    "Resultados financieros y por tenencia"))
    fila_aj = {}
    for i, (cod, nom, rub) in enumerate(cuentas):
        fila_aj[cod] = 5 + i
    rFin = 5 + len(cuentas) - 1

    VO_C = {"1.6.01", "1.6.02", "1.6.04", "1.6.06", "1.6.08"}
    AA_C = {"1.6.03", "1.6.05", "1.6.07", "1.6.09"}
    for cod, nom, rub in cuentas:
        r = fila_aj[cod]
        put(ws, f"A{r}", cod)
        put(ws, f"B{r}", nom)
        put(ws, f"C{r}", rub)
        put(ws, f"D{r}", clase(cod))
        if cod in fila_bss:
            put(ws, f"E{r}", f"={BH(cod, 'E')}", NUM)
            put(ws, f"G{r}", f"={BH(cod, 'H')}", NUM)
        else:
            put(ws, f"E{r}", 0, NUM)
            put(ws, f"G{r}", 0, NUM)
        cl = clase(cod)
        # ----- saldo inicial ajustado
        if anio == 2026 and cod in fila_ini:
            fi, origen_i = f"=Inicio!F{fila_ini[cod]}", "Inicio (cierre 2025 aj. × coef.)"
        elif anio == 2026:
            fi, origen_i = "=0", ""
        elif cl in ("A-M", "P-M"):
            fi, origen_i = f"=E{r}*Coeficientes!$E$3", ""
        elif cod == "1.4.03":
            fi = f"=Seguros!P{st}"
        elif cod == "1.5.01":
            fi = "=Mercaderías!D7"
        elif cod in VO_C:
            fi = "=" + bu_sum("S", "C", cod)
        elif cod in AA_C:
            fi = "=-" + bu_sum("W", "D", cod)
        elif cod == "3.1.01":
            fi = "=-'Patrimonio neto'!F6"
        elif cod == "3.1.02":
            fi = "=-'Patrimonio neto'!F7"
        elif cod == "3.2.01":
            fi = "=-'Patrimonio neto'!F18"
        elif cod == "3.3.01":
            fi = f"=-SUM(F5:F{r-1})-SUM(F{r+1}:F{rFin})"   # RNA inicial por diferencia
        else:
            fi = "=0"
        put(ws, f"F{r}", fi, NUM)
        # ----- saldo final ajustado
        if cl in ("A-M", "P-M"):
            ff, org = f"=G{r}", "Monetaria: valor nominal"
        elif cod == "1.4.03":
            ff, org = f"=Seguros!U{st}", "Seguros (no devengado × coef. mes de pago)"
        elif cod == "1.5.01":
            ff, org = "=Mercaderías!D27", "Mercaderías (capas FIFO × coef. mes de compra)"
        elif cod in VO_C:
            ff, org = "=" + bu_sum("V", "C", cod), "Bienes de uso (VO × coef. mes de alta)"
        elif cod in AA_C:
            ff, org = "=-" + bu_sum("Z", "D", cod), "Bienes de uso (AA × coef. mes de alta)"
        elif cod == "3.1.01":
            ff, org = "=-'Patrimonio neto'!F6", "Capital nominal"
        elif cod == "3.1.02":
            ff, org = "=-'Patrimonio neto'!F7", "Capital × coef. mar-2014 − nominal"
        elif cod == "3.2.01":
            ff, org = "=-'Patrimonio neto'!F18", "Capas por fecha de asamblea"
        elif cod == "3.3.01":
            ff, org = "=-'Patrimonio neto'!F23", ("RNA inicio por diferencia − dividendos aj." if anio == 2025
                                                   else "RNA inicio (Inicio) − dividendos aj.")
        elif cod == "4.9.01":
            ff, org = f"=-SUM(H5:H{r-1})", "Cuadre: RECPAM por diferencia"
        else:  # resultados
            rr_ = fila_res[cod]
            ff, org = f"=-Resultados!D{rr_}*Resultados!U{rr_}", "Resultados (hoja Resultados)"
        put(ws, f"H{r}", ff, NUM)
        put(ws, f"I{r}", f"=H{r}-G{r}", NUM)
        put(ws, f"J{r}", org)
    rT = rFin + 1
    put(ws, f"B{rT}", "Total (debe ser 0)", font=F_B)
    for col in "EFGHI":
        put(ws, f"{col}{rT}", f"=SUM({col}5:{col}{rFin})", NUM, F_B, FILL_CTRL)
    put(ws, f"B{rT+1}", "Sumas de saldos deudores", font=F_B)
    put(ws, f"B{rT+2}", "Sumas de saldos acreedores", font=F_B)
    for col in "EFGH":
        put(ws, f"{col}{rT+1}", f'=SUMIF({col}5:{col}{rFin},">0")', NUM, F_B)
        put(ws, f"{col}{rT+2}", f'=-SUMIF({col}5:{col}{rFin},"<0")', NUM, F_B)
    # Totales ESP
    put(ws, f"B{rT+4}", "Total activo", font=F_B)
    put(ws, f"B{rT+5}", "Total pasivo", font=F_B)
    put(ws, f"B{rT+6}", "Patrimonio neto (activo − pasivo)", font=F_B)
    put(ws, f"B{rT+7}", "Posición monetaria neta (activos − pasivos monetarios)", font=F_B)
    for col in "EFGH":
        put(ws, f"{col}{rT+4}", f'=SUMIF($D$5:$D${rFin},"A-*",{col}5:{col}{rFin})', NUM, F_B)
        put(ws, f"{col}{rT+5}", f'=-SUMIF($D$5:$D${rFin},"P-*",{col}5:{col}{rFin})', NUM, F_B)
        put(ws, f"{col}{rT+6}", f"={col}{rT+4}-{col}{rT+5}", NUM, F_B, FILL_TOT)
        put(ws, f"{col}{rT+7}", f'=SUMIF($D$5:$D${rFin},"A-M",{col}5:{col}{rFin})+SUMIF($D$5:$D${rFin},"P-M",{col}5:{col}{rFin})', NUM)
    widths(ws, [9, 50, 26, 7, 17, 19, 17, 19, 17, 42])
    ws.freeze_panes = "C5"
    if anio == 2025:
        wb["Patrimonio neto"]["F21"].value = f"=-'BSS ajustado'!F{fila_aj['3.3.01']}"
    AJ = dict(rFin=rFin, rT=rT, fila=fila_aj)

    # ------------------------------------------------------------------ RECPAM
    ws = wb.create_sheet("RECPAM")
    titulo(ws, f"RECPAM {anio} en moneda de dic-{anio}: método por diferencia y método directo (control)",
           "Por diferencia: resultado total del ejercicio (variación del PN ajustado + dividendos) − resultados reexpresados. "
           "Directo: posición monetaria neta (PMN) al inicio de cada mes × inflación del mes, llevado a moneda de cierre.")
    put(ws, "A4", "A. Método por diferencia", font=F_B)
    put(ws, "A5", f"PN ajustado al 31/12/{anio} (activo aj. − pasivo)")
    put(ws, "C5", f"='BSS ajustado'!H{rT+6}", NUM)
    put(ws, "A6", f"PN ajustado al 31/12/{anio-1} en moneda de dic-{anio}")
    put(ws, "C6", f"='BSS ajustado'!F{rT+6}", NUM)
    put(ws, "A7", "Dividendos en efectivo (aj.)")
    put(ws, "C7", "=-'Patrimonio neto'!F22", NUM)
    put(ws, "A8", "Aportes de los propietarios (aj.)")
    put(ws, "C8", 0, NUM, F_IN)
    put(ws, "A9", "Resultado total del ejercicio ajustado = PN cierre − PN inicio + dividendos − aportes")
    put(ws, "C9", "=C5-C6+C7-C8", NUM, F_B)
    put(ws, "A10", "Resultados reexpresados antes del RECPAM (hoja Resultados)")
    put(ws, "C10", f"=Resultados!U{RES['rr']}", NUM)
    put(ws, "A12", "RECPAM por diferencia (ganancia + / pérdida −)", font=F_B)
    put(ws, "C12", "=C9-C10", NUM, F_B, FILL_TOT)
    put(ws, "A13", "Control: igual al cuadre del BSS ajustado (cuenta 4.9.01, con signo invertido) — debe ser 0")
    put(ws, "C13", f"=C12+'BSS ajustado'!H{fila_aj['4.9.01']}", NUM, fill=FILL_CTRL)

    put(ws, "A15", "B. Método directo (control de razonabilidad)", font=F_B)
    hdr(ws, 16, ["Mes", "PMN al inicio del mes (hist.)", "Ingresos monetarios (ventas, intereses, otros)",
                 "Gastos monetarios (sueldos, fletes, IIBB, incobrables, honorarios, servicios, tasas, intereses, IG)",
                 "Compras de mercaderías", "Altas de bienes de uso", "Cobro por venta de bienes de uso (sin IVA)",
                 "Primas de seguros pagadas", "Dividendos pagados", "Flujo monetario neto del mes", "PMN al cierre del mes (hist.)",
                 "Inflación del mes", "Coef. mes → cierre", "RECPAM del mes (moneda de cierre)"])
    ing_m = [c for c in fila_res if c.startswith("4.") and c != "4.3.01"]
    gas_m = [c for c in fila_res if c.startswith("5.") and c not in ESPECIAL]
    put(ws, "A30", "Total", font=F_B)
    for j in range(12):
        r = 17 + j
        colR = L(5 + j)
        put(ws, f"A{r}", dt.datetime(anio, j + 1, 1), MES)
        if j == 0:
            put(ws, f"B{r}", f"='BSS ajustado'!E{rT+7}", NUM)
        else:
            put(ws, f"B{r}", f"=K{r-1}", NUM)
        put(ws, f"C{r}", "=" + "+".join(f"Resultados!{colR}{fila_res[c]}" for c in ing_m), NUM)
        put(ws, f"D{r}", "=-(" + "+".join(f"Resultados!{colR}{fila_res[c]}" for c in gas_m) + ")", NUM)
        put(ws, f"E{r}", f"=-Mercaderías!B{10+j}", NUM)
        put(ws, f"F{r}", f"=-SUMPRODUCT(('Bienes de uso'!$H${r1}:$H${rn}=A{r})*'Bienes de uso'!$L${r1}:$L${rn})", NUM)
        put(ws, f"G{r}", f"=SUMPRODUCT((DATE(YEAR('Bienes de uso'!$AC${r1}:$AC${rn}+0),MONTH('Bienes de uso'!$AC${r1}:$AC${rn}+0),1)=A{r})*('Bienes de uso'!$AC${r1}:$AC${rn}<>\"\")*'Bienes de uso'!$AB${r1}:$AB${rn})", NUM)
        put(ws, f"H{r}", f"=-SUMPRODUCT((DATE(YEAR(Seguros!$I${s1}:$I${sn}),MONTH(Seguros!$I${s1}:$I${sn}),1)=A{r})*Seguros!$K${s1}:$K${sn})", NUM)
        put(ws, f"I{r}", f"=IF('Patrimonio neto'!$C$22=A{r},'Patrimonio neto'!$E$22,0)", NUM)
        put(ws, f"J{r}", f"=SUM(C{r}:I{r})", NUM)
        put(ws, f"K{r}", f"=B{r}+J{r}", NUM)
        put(ws, f"L{r}", f"={IDX_LOOK.format(m=f'A{r}')}/{IDX_LOOK.format(m=f'EDATE(A{r},-1)')}-1", PCT)
        put(ws, f"M{r}", "=" + coef_f(f"A{r}"), COEF)
        put(ws, f"N{r}", f"=-B{r}*L{r}*M{r}", NUM)
    for col in "CDEFGHIJN":
        put(ws, f"{col}30", f"=SUM({col}17:{col}28)", NUM, F_B, FILL_TOT)
    put(ws, "A32", "Control: PMN al cierre (método directo) − PMN al cierre según BSS (debe ser 0)")
    put(ws, "C32", f"=K28-'BSS ajustado'!G{rT+7}", NUM, fill=FILL_CTRL)
    put(ws, "A33", "RECPAM método directo", font=F_B)
    put(ws, "C33", "=N30", NUM, F_B)
    put(ws, "A34", "Diferencia: por diferencia − directo", font=F_B)
    put(ws, "C34", "=C12-C33", NUM, F_B, FILL_CTRL)
    put(ws, "A35", "Diferencia relativa sobre el RECPAM")
    put(ws, "C35", "=IF(C12=0,0,C34/ABS(C12))", PCT)
    put(ws, "A36", ("Explicación: con flujos monetarios mensuales tomados de los mismos auxiliares y el mismo índice mensual, "
                    "el método directo es algebraicamente igual al método por diferencia "
                    "(Σ PMNₘ₋₁ × inflaciónₘ × coefₘ = PMN cierre − PMN inicio × coef. − Σ flujosₘ × coefₘ). "
                    "La diferencia residual solo puede venir de redondeos de los auxiliares o de partidas no monetarias "
                    "cuyo coeficiente no coincide con el mes del flujo monetario."))
    put(ws, "A38", "C. Estimación simplificada (orden de magnitud)", font=F_B)
    put(ws, "A39", "PMN promedio simple (13 saldos: inicio y fin de cada mes)")
    put(ws, "C39", "=(B17+SUM(K17:K28))/13", NUM)
    put(ws, "A40", "Inflación del ejercicio (coef. dic-previo → cierre − 1)")
    put(ws, "C40", "=Coeficientes!E3-1", PCT)
    put(ws, "A41", "RECPAM estimado = −PMN promedio × inflación del ejercicio")
    put(ws, "C41", "=-C39*C40", NUM, F_B)
    put(ws, "A42", "Diferencia con el método por diferencia (explicada por la distribución de la inflación y de la PMN dentro del año)")
    put(ws, "C42", "=C12-C41", NUM)
    widths(ws, [16, 18, 18, 22, 18, 16, 16, 16, 14, 18, 18, 12, 12, 18])
    ws.column_dimensions["A"].width = 60
    ws["A36"].alignment = Alignment(wrap_text=True)
    ws.merge_cells("A36:N36")
    ws.row_dimensions[36].height = 48

    # ------------------------------------------------------------------ Conciliación
    ws = wb.create_sheet("Conciliación")
    titulo(ws, "Conciliación auxiliares ajustados vs. BSS ajustado", "Diferencia esperada: 0 en todas las filas.")
    hdr(ws, 4, ["Concepto", "Cuenta", "Auxiliar (aj.)", "BSS ajustado", "Diferencia", "Auxiliar (hist.)", "BSS (hist.)", "Diferencia hist."])
    filas_c = []
    for cod in sorted(VO_C):
        filas_c.append((f"Bienes de uso VO cierre", cod, "=" + bu_sum("V", "C", cod), "=" + bu_sum("N", "C", cod)))
    for cod in sorted(AA_C):
        filas_c.append((f"Bienes de uso AA cierre", cod, "=-" + bu_sum("Z", "D", cod), "=-" + bu_sum("R", "D", cod)))
    filas_c.append(("Seguros no devengados cierre", "1.4.03", f"=Seguros!U{st}", f"=Seguros!O{st}"))
    filas_c.append(("Mercaderías existencia final", "1.5.01", "=Mercaderías!D27", "=Mercaderías!B27"))
    for cod in ("5.2.07", "5.2.08", "5.2.09", "5.3.06", "5.3.07"):
        filas_c.append(("Amortización del ejercicio", cod, "=" + bu_sum("Y", "E", cod), "=" + bu_sum("Q", "E", cod)))
    for cod in ("5.2.06", "5.3.05"):
        filas_c.append(("Gasto de seguros", cod, "=" + seg_sum("T", cod), "=" + seg_sum("N", cod)))
    filas_c.append(("CMV", "5.1.01", "=Mercaderías!P43", "=Mercaderías!N43"))
    filas_c.append(("Capital + ajuste de capital", "3.1.01+3.1.02", "='Patrimonio neto'!F5", "='Patrimonio neto'!E5"))
    filas_c.append(("Reserva legal", "3.2.01", "='Patrimonio neto'!F18", "='Patrimonio neto'!E18"))
    filas_c.append(("RNA cierre (antes del resultado)", "3.3.01", "='Patrimonio neto'!F23", "='Patrimonio neto'!E23"))
    for i, (t, cod, fa, fh) in enumerate(filas_c):
        r = 5 + i
        put(ws, f"A{r}", t)
        put(ws, f"B{r}", cod)
        put(ws, f"C{r}", fa, NUM)
        put(ws, f"F{r}", fh, NUM)
        if cod == "3.1.01+3.1.02":
            put(ws, f"D{r}", f"=-('BSS ajustado'!H{fila_aj['3.1.01']}+'BSS ajustado'!H{fila_aj['3.1.02']})", NUM)
            put(ws, f"G{r}", f"=-'BSS histórico'!H{fila_bss['3.1.01']}", NUM)
        elif cod.startswith("3."):
            put(ws, f"D{r}", f"=-'BSS ajustado'!H{fila_aj[cod]}", NUM)
            put(ws, f"G{r}", f"=-'BSS histórico'!H{fila_bss[cod]}", NUM)
        else:
            put(ws, f"D{r}", f"='BSS ajustado'!H{fila_aj[cod]}" if cod.startswith("1.") else f"=ABS('BSS ajustado'!H{fila_aj[cod]})", NUM)
            put(ws, f"G{r}", f"='BSS histórico'!H{fila_bss[cod]}" if cod.startswith("1.") else f"=ABS('BSS histórico'!H{fila_bss[cod]})", NUM)
        put(ws, f"E{r}", f"=C{r}-D{r}", NUM, fill=FILL_CTRL)
        put(ws, f"H{r}", f"=F{r}-G{r}", NUM, fill=FILL_CTRL)
    rc_end = 5 + len(filas_c) - 1
    put(ws, f"A{rc_end+1}", "Suma de diferencias absolutas", font=F_B)
    put(ws, f"E{rc_end+1}", f"=SUMPRODUCT(ABS(E5:E{rc_end}))", NUM, F_B, FILL_CTRL)
    put(ws, f"H{rc_end+1}", f"=SUMPRODUCT(ABS(H5:H{rc_end}))", NUM, F_B, FILL_CTRL)
    extra = rc_end + 3
    if anio == 2026:
        put(ws, f"A{extra}", "Saldos iniciales: auxiliares (desde el origen, moneda dic-2026) vs. Inicio (cierre 2025 aj. × coef.)", font=F_B)
        hdr(ws, extra + 1, ["Concepto", "Cuenta", "Auxiliar inicio (aj.)", "Inicio", "Diferencia"])
        fi_c = []
        for cod in sorted(VO_C):
            fi_c.append(("VO inicio", cod, "=" + bu_sum("S", "C", cod), f"=Inicio!F{fila_ini[cod]}"))
        for cod in sorted(AA_C):
            fi_c.append(("AA inicio", cod, "=-" + bu_sum("W", "D", cod), f"=Inicio!F{fila_ini[cod]}"))
        fi_c.append(("Seguros no devengados inicio", "1.4.03", f"=Seguros!P{st}", f"=Inicio!F{fila_ini['1.4.03']}"))
        fi_c.append(("Mercaderías existencia inicial", "1.5.01", "=Mercaderías!D7", f"=Inicio!F{fila_ini['1.5.01']}"))
        fi_c.append(("Capital + ajuste de capital", "3.1", "='Patrimonio neto'!F5", f"=-(Inicio!F{fila_ini['3.1.01']}+Inicio!F{fila_ini['3.1.02']})"))
        fi_c.append(("Reserva legal", "3.2.01", "='Patrimonio neto'!F18", f"=-Inicio!F{fila_ini['3.2.01']}"))
        for i, (t, cod, fa, fb) in enumerate(fi_c):
            r = extra + 2 + i
            put(ws, f"A{r}", t)
            put(ws, f"B{r}", cod)
            put(ws, f"C{r}", fa, NUM)
            put(ws, f"D{r}", fb, NUM)
            put(ws, f"E{r}", f"=C{r}-D{r}", NUM, fill=FILL_CTRL)
        rx = extra + 2 + len(fi_c)
        put(ws, f"A{rx}", "Suma de diferencias absolutas (tolerancia: redondeo de los valores importados, < $1)", font=F_B)
        put(ws, f"E{rx}", f"=SUMPRODUCT(ABS(E{extra+2}:E{rx-1}))", NUM, F_B, FILL_CTRL)
        CONC_INI = f"E{rx}"
    widths(ws, [52, 14, 18, 18, 14, 18, 18, 14])
    CONC = f"E{rc_end+1}"
    CONC_H = f"H{rc_end+1}"

    # ------------------------------------------------------------------ Controles
    ws = wb.create_sheet("Controles")
    titulo(ws, "Controles de los criterios de aceptación", "OK si |valor| < tolerancia.")
    hdr(ws, 4, ["Control", "Valor", "Tolerancia", "Resultado"])
    ctr = [
        ("BSS histórico: total de saldos finales = 0", f"='BSS histórico'!H{5+len(bss)}", 0.5),
        ("BSS histórico: saldos calculados = insumo", f"='BSS histórico'!H{6+len(bss)}", 0.5),
        ("BSS ajustado: saldos iniciales cuadran (total = 0)", f"='BSS ajustado'!F{rT}", 0.5),
        ("BSS ajustado: saldos finales cuadran (total = 0)", f"='BSS ajustado'!H{rT}", 0.5),
        ("Auxiliares ajustados = BSS ajustado (Σ|dif.|)", f"=Conciliación!{CONC}", 0.5),
        ("Auxiliares históricos = BSS histórico (Σ|dif.|)", f"=Conciliación!{CONC_H}", 0.5),
        ("Resultados mensuales = BSS (Σ|dif.|)", f"=SUMPRODUCT(ABS(Resultados!S7:S{RES['rlast']}))", 0.5),
        ("Mercaderías ajustadas: EI + compras − CMV − EF = 0", "=Mercaderías!D46", 0.5),
        ("RECPAM por diferencia = cuadre del BSS ajustado", "=RECPAM!C13", 0.5),
        ("Método directo: PMN de cierre = BSS", "=RECPAM!C32", 0.5),
        ("RECPAM: diferencia por diferencia − directo", "=RECPAM!C34", 1.0),
        ("PN ajustado (PN hoja) = activo − pasivo ajustados", f"='Patrimonio neto'!F27-'BSS ajustado'!H{rT+6}", 0.5),
        ("Resultado histórico de la hoja = cifra de control", f"=Resultados!Q{RES['rr']}-{926354325 if anio == 2025 else 1250174250}", 0.5),
    ]
    if anio == 2025:
        ctr.append(("PN: capas de RNA (hist.) = BSS", f"='Patrimonio neto'!E{43+len(RESULTADOS_HIST)+len(RESERVA)+len(DIVIDENDOS_HIST)+1}", 0.5))
    else:
        ctr.append(("Inicio 2026: cierre 2025 aj. importado cuadra (total = 0)", f"=Inicio!F{5+len([k for k in inicio if k!='_resultado'])}", 0.5))
        ctr.append(("Saldos iniciales: auxiliares desde el origen = Inicio (Σ|dif.|)", f"=Conciliación!{CONC_INI}", 1.0))
    for i, (t, f, tol) in enumerate(ctr):
        r = 5 + i
        put(ws, f"A{r}", t)
        put(ws, f"B{r}", f, '#,##0.00')
        put(ws, f"C{r}", tol)
        put(ws, f"D{r}", f'=IF(ABS(B{r})<C{r},"OK","ERROR")', fill=FILL_CTRL)
    rk = 5 + len(ctr)
    put(ws, f"A{rk}", "Resultado global", font=F_B)
    put(ws, f"D{rk}", f'=IF(COUNTIF(D5:D{rk-1},"OK")={len(ctr)},"OK","ERROR")', font=F_B, fill=FILL_CTRL)
    put(ws, f"A{rk+2}", "Resumen", font=F_B)
    resumen = [
        ("Resultado del ejercicio histórico", f"=Resultados!Q{RES['rr']}"),
        ("Resultado antes del RECPAM (aj.)", f"=Resultados!U{RES['rr']}"),
        ("RECPAM (aj.)", "=RECPAM!C12"),
        ("Resultado del ejercicio ajustado", f"=Resultados!U{RES['rr']+2}"),
        ("RECPAM método directo", "=RECPAM!C33"),
        (f"PN ajustado al 31/12/{anio}", f"='BSS ajustado'!H{rT+6}"),
        (f"PN ajustado al 31/12/{anio-1} (moneda dic-{anio})", f"='BSS ajustado'!F{rT+6}"),
        (f"Total activo ajustado al 31/12/{anio}", f"='BSS ajustado'!H{rT+4}"),
        (f"Total pasivo al 31/12/{anio}", f"='BSS ajustado'!H{rT+5}"),
        ("Coeficiente dic-previo → cierre", "=Coeficientes!E3"),
    ]
    for i, (t, f) in enumerate(resumen):
        r = rk + 3 + i
        put(ws, f"A{r}", t)
        put(ws, f"B{r}", f, COEF if "Coef" in t else NUM)
    widths(ws, [66, 20, 11, 11])

    # ------------------------------------------------------------------ Notas
    notas_hoja(wsN, anio)
    # orden de hojas
    orden = ["Notas", "Controles", "Coeficientes", "BSS histórico", "Bienes de uso", "Seguros", "Mercaderías",
             "Patrimonio neto", "Resultados", "RECPAM", "BSS ajustado", "Conciliación"] + (["Inicio"] if anio == 2026 else [])
    wb._sheets = [wb[n] for n in orden]
    ruta = os.path.join(OUT, f"ajuste_{anio}.xlsx")
    wb.save(ruta)
    return ruta, dict(AJ=AJ, RES=RES, fila_res=fila_res, cuentas=cuentas)


inicio_nombres = {}


def notas_hoja(ws, anio):
    titulo(ws, f"Ajuste por inflación {anio} — método, norma y supuestos",
           f"Moneda de medición: dic-{anio}. Valores en pesos. Azul = insumo; negro = fórmula.")
    lineas = [
        "NORMA",
        "RT 6 FACPCE (Estados contables en moneda homogénea), modificada por RT 19, 27 y 39 y por la Res. JG 539/18 (modif. Res. JG 553/19 y 584/21). "
        "La RT 54 (NUA, vigente para ejercicios iniciados desde el 01/01/2025) incorpora el procedimiento de reexpresión de la RT 6; "
        "Duda: confirmar la numeración exacta de los párrafos de la RT 54 (t.o. RT 59) sobre unidad de medida.",
        "Índice: serie FACPCE Res. JG 539/18 (IPC Nacional INDEC desde ene-2017; IPIM empalmado antes). Ver insumos/indices/README.md.",
        "MÉTODO",
        "1. Partidas monetarias (caja, bancos, plazo fijo, créditos, anticipos e IVA, deudas, préstamo): valor nominal al cierre. Saldos iniciales × coef. dic-previo → cierre.",
        "2. Bienes de uso: bien por bien; VO, amortización acumulada y del ejercicio × coef. del mes de alta. Bajas: VO y AA de la baja con el mismo coeficiente.",
        "3. Seguros pagados por adelantado: no devengado × coef. del mes de pago de la prima. Gasto: devengado de cada prima × su coef. de pago.",
        "4. Mercaderías (FIFO): existencia por capas × coef. del mes de compra; CMV por mes de compra (matriz FIFO) × coef. del mes de compra.",
        "5. Patrimonio neto: capital nominal + ajuste de capital (capital × coef. mar-2014 − nominal); reserva legal por fecha de cada asamblea; dividendos × coef. del mes de la asamblea.",
        "   RNA al inicio: " + ("por diferencia entre activo y pasivo reexpresados y los demás componentes del PN (RT 6, primer ejercicio de aplicación del ajuste en este juego de EECC)."
                                if anio == 2025 else "cierre 2025 ajustado (incluido el resultado 2025 con RECPAM) × coef. dic-2025 → dic-2026."),
        "6. Resultados: mes a mes × coef. del mes de devengamiento; los que consumen partidas no monetarias (CMV, amortizaciones, seguros, resultado por venta de bienes de uso) desde su auxiliar.",
        "7. RECPAM: por diferencia (cuenta de cuadre del BSS ajustado). Control: método directo mensual (PMN × inflación del mes) y estimación simplificada.",
        "SUPUESTOS",
        "(S) Coeficiente con el índice del mes de origen (sin prorrateo diario).",
        "(S) Póliza anterior pagada el mismo día del año previo (fecha no informada en el auxiliar).",
        "(S) Sin prueba de valor recuperable / VNR de bienes de cambio y bienes de uso en esta fase (se asume que el valor reexpresado no supera el recuperable).",
        "(S) Impuesto a las ganancias imputado en diciembre (monetario); sin impuesto diferido (supuesto simplificador del insumo).",
        "(S) Dividendos reexpresados desde el mes de la asamblea (abril; el pago es en el mismo mes).",
        "(S) Fórmulas compatibles con LibreOffice 24.2 (INDEX/MATCH y SUMPRODUCT en lugar de XLOOKUP/LET) para poder recalcular y verificar en este entorno.",
        "DUDAS / DATOS FALTANTES",
        "Duda — Reserva legal (art. 70 LGS): ver hoja 'Patrimonio neto', bloque 'Duda'. Propuesta: base = capital suscripto + ajuste de capital (interpretación habitual; validar con asesor societario / criterio IGJ). "
        "En moneda homogénea la reserva queda por debajo del 20 %; la asignación faltante NO se contabiliza (requiere decisión de asamblea; es una reclasificación dentro del PN sin efecto en el total). "
        "Se informa el importe para la próxima asamblea y para nota a los EECC.",
        "Dato faltante — Cronograma contractual del préstamo: se mantiene el supuesto de amortización trimestral lineal (afecta solo la distribución mensual de intereses).",
    ]
    if anio == 2026:
        lineas.append("Estimación — Índices de nov-2026 y dic-2026 estimados (no oficiales). Todos los importes en moneda de dic-2026 deben recalcularse cuando se publiquen (16/12/2026 y ene-2027).")
    for i, t in enumerate(lineas):
        c = put(ws, f"A{4+i}", t, font=F_B if t.isupper() else None)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width = 160


# ============================================================================= recalcular
def recalcular(ruta):
    tmp = tempfile.mkdtemp()
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", tmp, ruta],
                   check=True, capture_output=True)
    shutil.move(os.path.join(tmp, os.path.basename(ruta)), ruta)
    shutil.rmtree(tmp, ignore_errors=True)


def leer_valores(ruta, hoja):
    return openpyxl.load_workbook(ruta, data_only=True)[hoja]


# ============================================================================= comparativo
def comparativo(r25, meta25):
    ws25 = leer_valores(r25, "BSS ajustado")
    AJ = meta25["AJ"]
    filas = []
    for cod, nom, rub in meta25["cuentas"]:
        r = AJ["fila"][cod]
        filas.append((cod, nom, rub, ws25[f"F{r}"].value, ws25[f"H{r}"].value))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Datos 2025 (dic-2025)"
    titulo(ws, "Ejercicio 2025 ajustado en moneda de dic-2025 (insumo)",
           "Valores importados de ajuste_2025.xlsx, hoja 'BSS ajustado' (columnas F y H, recalculadas con LibreOffice). Deudor (+) / acreedor (−).")
    hdr(ws, 4, ["Código", "Cuenta", "Rubro", "Saldo 31/12/2024 (aj. dic-2025)", "Saldo 31/12/2025 (aj. dic-2025)",
                "Saldo 31/12/2024 (moneda dic-2026)", "Saldo 31/12/2025 (moneda dic-2026)"])
    fila = {}
    for i, (cod, nom, rub, si, sf) in enumerate(filas):
        r = 5 + i
        fila[cod] = r
        put(ws, f"A{r}", cod, font=F_IN)
        put(ws, f"B{r}", nom, font=F_IN)
        put(ws, f"C{r}", rub, font=F_IN)
        put(ws, f"D{r}", si, NUM, F_IN)
        put(ws, f"E{r}", sf, NUM, F_IN)
        put(ws, f"F{r}", f"=D{r}*Coeficiente!$C$6", NUM)
        put(ws, f"G{r}", f"=E{r}*Coeficiente!$C$6", NUM)
    rt = 5 + len(filas)
    put(ws, f"B{rt}", "Total (debe ser 0)", font=F_B)
    for col in "DEFG":
        put(ws, f"{col}{rt}", f"=SUM({col}5:{col}{rt-1})", NUM, F_B, FILL_CTRL)
    widths(ws, [9, 50, 30, 20, 20, 20, 20])
    D = "'Datos 2025 (dic-2025)'"

    wc = wb.create_sheet("Coeficiente")
    titulo(wc, "Coeficiente dic-2025 → dic-2026", "Fuente: insumos/indices/indices_facpce.csv. Dic-2026 es ESTIMADO (no oficial).")
    idx = {m: v for m, v, e in leer_indices()}
    put(wc, "A3", "Índice dic-2024 (oficial)")
    put(wc, "C3", idx[dt.datetime(2024, 12, 1)], '#,##0.0000', F_IN)
    put(wc, "A4", "Índice dic-2025 (oficial)")
    put(wc, "C4", idx[dt.datetime(2025, 12, 1)], '#,##0.0000', F_IN)
    put(wc, "A5", "Índice dic-2026 (estimado)")
    put(wc, "C5", idx[dt.datetime(2026, 12, 1)], '#,##0.0000', F_IN)
    put(wc, "A6", "Coeficiente dic-2025 → dic-2026", font=F_B)
    put(wc, "C6", "=C5/C4", COEF, F_B)
    widths(wc, [40, 4, 16])

    def s(col, cods, neg=False):
        expr = "+".join(f"{D}!{col}{fila[c]}" for c in cods)
        return f"=-({expr})" if neg else f"={expr}"

    def sumif_rub(col, rub, neg=False):
        f = f"SUMIF({D}!$C$5:$C${rt-1},\"{rub}\",{D}!{col}$5:{col}${rt-1})"
        return f"=-{f}" if neg else f"={f}"

    def res_sum(col):
        a = f"{D}!$A$5:$A${rt-1}"
        return f"SUMPRODUCT((LEFT({a},1)=\"4\")+(LEFT({a},1)=\"5\"),{D}!{col}$5:{col}${rt-1})"

    # ESP
    we = wb.create_sheet("ESP 2025")
    titulo(we, "Estado de situación patrimonial al 31/12/2025 y al 31/12/2024 — en moneda de dic-2026",
           "Cifras ajustadas por inflación (RT 6), reexpresadas a moneda de dic-2026 para su uso como comparativo. Estado de trabajo (la presentación formal es de la fase 3).")
    hdr(we, 4, ["Rubro", "31/12/2025", "31/12/2024"])
    activo_c = ["Caja y bancos", "Inversiones", "Créditos por ventas", "Otros créditos", "Bienes de cambio"]
    r = 5
    put(we, f"A{r}", "ACTIVO", font=F_B); r += 1
    put(we, f"A{r}", "Activo corriente", font=F_B); r += 1
    ini_ac = r
    for rub in activo_c:
        put(we, f"A{r}", rub)
        put(we, f"B{r}", sumif_rub("G", rub), NUM)
        put(we, f"C{r}", sumif_rub("F", rub), NUM)
        r += 1
    put(we, f"A{r}", "Total activo corriente", font=F_B)
    put(we, f"B{r}", f"=SUM(B{ini_ac}:B{r-1})", NUM, F_B, FILL_TOT)
    put(we, f"C{r}", f"=SUM(C{ini_ac}:C{r-1})", NUM, F_B, FILL_TOT)
    tac = r; r += 1
    put(we, f"A{r}", "Activo no corriente", font=F_B); r += 1
    put(we, f"A{r}", "Bienes de uso")
    put(we, f"B{r}", sumif_rub("G", "Bienes de uso"), NUM)
    put(we, f"C{r}", sumif_rub("F", "Bienes de uso"), NUM)
    bu_r = r; r += 1
    put(we, f"A{r}", "Total activo no corriente", font=F_B)
    put(we, f"B{r}", f"=B{bu_r}", NUM, F_B, FILL_TOT)
    put(we, f"C{r}", f"=C{bu_r}", NUM, F_B, FILL_TOT)
    tanc = r; r += 1
    put(we, f"A{r}", "TOTAL ACTIVO", font=F_B)
    put(we, f"B{r}", f"=B{tac}+B{tanc}", NUM, F_B, FILL_TOT)
    put(we, f"C{r}", f"=C{tac}+C{tanc}", NUM, F_B, FILL_TOT)
    ta = r; r += 2
    put(we, f"A{r}", "PASIVO", font=F_B); r += 1
    put(we, f"A{r}", "Pasivo corriente", font=F_B); r += 1
    ini_pc = r
    for rub in ["Deudas comerciales", "Remuneraciones y cargas sociales", "Cargas fiscales"]:
        put(we, f"A{r}", rub)
        put(we, f"B{r}", sumif_rub("G", rub, True), NUM)
        put(we, f"C{r}", sumif_rub("F", rub, True), NUM)
        r += 1
    put(we, f"A{r}", "Préstamos (porción corriente) — supuesto $150.000.000 nominal por año")
    put(we, f"B{r}", "=MIN(150000000*Coeficiente!$C$5/Coeficiente!$C$4," + sumif_rub("G", "Préstamos", True)[1:] + ")", NUM)
    put(we, f"C{r}", "=MIN(150000000*Coeficiente!$C$5/Coeficiente!$C$3," + sumif_rub("F", "Préstamos", True)[1:] + ")", NUM)
    pc_prest = r; r += 1
    put(we, f"A{r}", "Total pasivo corriente", font=F_B)
    put(we, f"B{r}", f"=SUM(B{ini_pc}:B{r-1})", NUM, F_B, FILL_TOT)
    put(we, f"C{r}", f"=SUM(C{ini_pc}:C{r-1})", NUM, F_B, FILL_TOT)
    tpc = r; r += 1
    put(we, f"A{r}", "Pasivo no corriente", font=F_B); r += 1
    put(we, f"A{r}", "Préstamos (porción no corriente)")
    put(we, f"B{r}", sumif_rub("G", "Préstamos", True) + f"-B{pc_prest}", NUM)
    put(we, f"C{r}", sumif_rub("F", "Préstamos", True) + f"-C{pc_prest}", NUM)
    pnc = r; r += 1
    put(we, f"A{r}", "Total pasivo no corriente", font=F_B)
    put(we, f"B{r}", f"=B{pnc}", NUM, F_B, FILL_TOT)
    put(we, f"C{r}", f"=C{pnc}", NUM, F_B, FILL_TOT)
    tpnc = r; r += 1
    put(we, f"A{r}", "TOTAL PASIVO", font=F_B)
    put(we, f"B{r}", f"=B{tpc}+B{tpnc}", NUM, F_B, FILL_TOT)
    put(we, f"C{r}", f"=C{tpc}+C{tpnc}", NUM, F_B, FILL_TOT)
    tp = r; r += 2
    put(we, f"A{r}", "PATRIMONIO NETO (según EEPN)", font=F_B)
    put(we, f"B{r}", "=EEPN!F9", NUM, F_B, FILL_TOT)
    put(we, f"C{r}", "=EEPN!F5", NUM, F_B, FILL_TOT)
    tpn = r; r += 1
    put(we, f"A{r}", "TOTAL PASIVO + PATRIMONIO NETO", font=F_B)
    put(we, f"B{r}", f"=B{tp}+B{tpn}", NUM, F_B, FILL_TOT)
    put(we, f"C{r}", f"=C{tp}+C{tpn}", NUM, F_B, FILL_TOT)
    r += 1
    put(we, f"A{r}", "Control: activo − (pasivo + PN) (debe ser 0)")
    put(we, f"B{r}", f"=B{ta}-B{r-1}", NUM, fill=FILL_CTRL)
    put(we, f"C{r}", f"=C{ta}-C{r-1}", NUM, fill=FILL_CTRL)
    esp_ctrl = r
    widths(we, [62, 20, 20])
    # ER
    wr = wb.create_sheet("ER 2025")
    titulo(wr, "Estado de resultados 2025 — en moneda de dic-2026", "Ajustado por inflación (RT 6) y reexpresado con el coeficiente dic-2025 → dic-2026. Ingreso (+) / egreso (−).")
    hdr(wr, 4, ["Concepto", "2025 (moneda dic-2026)", "2025 (moneda dic-2025)"])
    lineas = [
        ("Ventas de insumos", ["4.1.01"]),
        ("Costo de mercaderías vendidas", ["5.1.01"]),
        ("RESULTADO BRUTO", "sum:5:6"),
        ("Gastos de comercialización", [c for c, *_ in meta25["cuentas"] if c.startswith("5.2.")]),
        ("Gastos de administración", [c for c, *_ in meta25["cuentas"] if c.startswith("5.3.")]),
        ("Intereses ganados plazo fijo", ["4.2.01"]),
        ("Intereses préstamo bancario", ["5.4.01"]),
        ("RECPAM", ["4.9.01"]),
        ("Resultados financieros y por tenencia (incl. RECPAM)", "sum:10:12"),
        ("RESULTADO ANTES DEL IMPUESTO A LAS GANANCIAS", "tot"),
        ("Impuesto a las ganancias", ["5.5.01"]),
        ("RESULTADO DEL EJERCICIO", "fin"),
    ]
    for i, (t, cods) in enumerate(lineas):
        r = 5 + i
        put(wr, f"A{r}", t, font=F_B if t.isupper() or "(incl." in t else None)
        if isinstance(cods, list):
            put(wr, f"B{r}", s("G", cods, True), NUM)
            put(wr, f"C{r}", s("E", cods, True), NUM)
        elif cods.startswith("sum"):
            a, b = cods.split(":")[1:]
            if t.startswith("RESULTADO BRUTO"):
                put(wr, f"B{r}", "=B5+B6", NUM, F_B, FILL_TOT)
                put(wr, f"C{r}", "=C5+C6", NUM, F_B, FILL_TOT)
            else:
                put(wr, f"B{r}", "=SUM(B10:B12)", NUM, F_B)
                put(wr, f"C{r}", "=SUM(C10:C12)", NUM, F_B)
        elif cods == "tot":
            put(wr, f"B{r}", "=B7+B8+B9+B13", NUM, F_B, FILL_TOT)
            put(wr, f"C{r}", "=C7+C8+C9+C13", NUM, F_B, FILL_TOT)
        elif cods == "fin":
            put(wr, f"B{r}", "=B14+B15", NUM, F_B, FILL_TOT)
            put(wr, f"C{r}", "=C14+C15", NUM, F_B, FILL_TOT)
    put(wr, "A18", "Control: resultado = − Σ cuentas de resultado (debe ser 0)")
    put(wr, "B18", "=B16+" + res_sum("G"), NUM, fill=FILL_CTRL)
    put(wr, "A19", "Control: resultado ER = resultado EEPN (debe ser 0)")
    put(wr, "B19", "=B16-EEPN!E8", NUM, fill=FILL_CTRL)
    widths(wr, [58, 22, 22])

    # EEPN
    wp = wb.create_sheet("EEPN")
    titulo(wp, "Estado de evolución del patrimonio neto 2025 — en moneda de dic-2026",
           "Capital y ajuste de capital por separado. Dividendos reexpresados desde la asamblea (abr-2025).")
    hdr(wp, 4, ["Concepto", "Capital suscripto", "Ajuste de capital", "Reserva legal", "Resultados no asignados", "Total"])
    put(wp, "A5", "Saldos al 31/12/2024")
    put(wp, "A6", "Dividendos en efectivo (asamblea 28/04/2025)")
    put(wp, "A7", "Subtotal")
    put(wp, "A8", "Resultado del ejercicio 2025")
    put(wp, "A9", "Saldos al 31/12/2025", font=F_B)
    for col, cod in zip("BCDE", ("3.1.01", "3.1.02", "3.2.01", "3.3.01")):
        if cod == "3.1.01":   # capital suscripto: nominal
            put(wp, "B5", f"=-{D}!D{fila[cod]}", NUM)
            put(wp, "B7", f"=-{D}!E{fila[cod]}", NUM)
        elif cod == "3.1.02":  # ajuste de capital: absorbe la reexpresión del capital nominal
            put(wp, "C5", f"=-({D}!F{fila['3.1.01']}+{D}!F{fila[cod]})-B5", NUM)
            put(wp, "C7", f"=-({D}!G{fila['3.1.01']}+{D}!G{fila[cod]})-B7", NUM)
        else:
            put(wp, f"{col}5", f"=-{D}!F{fila[cod]}", NUM)
            put(wp, f"{col}7", f"=-{D}!G{fila[cod]}", NUM)
        put(wp, f"{col}6", f"={col}7-{col}5", NUM)
    put(wp, "E8", "=-" + res_sum("G"), NUM)
    for col in "BCD":
        put(wp, f"{col}8", 0, NUM)
    for col in "BCDE":
        put(wp, f"{col}9", f"={col}7+{col}8", NUM, F_B, FILL_TOT)
    for rr in range(5, 10):
        put(wp, f"F{rr}", f"=SUM(B{rr}:E{rr})", NUM, F_B if rr == 9 else None)
    put(wp, "A11", "Control: PN cierre EEPN = activo − pasivo del ESP (debe ser 0)")
    put(wp, "F11", f"=F9-('ESP 2025'!B{ta}-'ESP 2025'!B{tp})", NUM, fill=FILL_CTRL)
    put(wp, "A12", "Control: dividendos = 120.000.000 × coef. abr-2025 → dic-2026 (debe ser 0)")
    put(wp, "F12", f"=F6+120000000*Coeficiente!C5/{idx[dt.datetime(2025, 4, 1)]}", NUM, fill=FILL_CTRL)
    widths(wp, [48, 18, 18, 18, 22, 20])
    put(wp, "A14", "Duda — Reserva legal: en moneda homogénea queda por debajo del 20 % de (capital + ajuste de capital). Ver ajuste_2025.xlsx / ajuste_2026.xlsx, hoja 'Patrimonio neto'. No se contabiliza asignación adicional.")

    # Controles
    wk = wb.create_sheet("Controles", 0)
    titulo(wk, "Comparativo 2025 en moneda de dic-2026 — controles", "OK si |valor| < tolerancia.")
    hdr(wk, 4, ["Control", "Valor", "Tolerancia", "Resultado"])
    cc = [("Datos importados cuadran (total = 0, 31/12/2025)", f"={D}!G{rt}", 0.5),
          ("Datos importados cuadran (total = 0, 31/12/2024)", f"={D}!F{rt}", 0.5),
          ("ESP 31/12/2025: activo = pasivo + PN", f"='ESP 2025'!B{esp_ctrl}", 0.5),
          ("ESP 31/12/2024: activo = pasivo + PN", f"='ESP 2025'!C{esp_ctrl}", 0.5),
          ("ER: resultado = Σ cuentas de resultado", "='ER 2025'!B18", 0.5),
          ("ER = EEPN (resultado)", "='ER 2025'!B19", 0.5),
          ("EEPN = ESP (PN cierre)", "=EEPN!F11", 0.5),
          ("Dividendos reexpresados", "=EEPN!F12", 0.5)]
    for i, (t, f, tol) in enumerate(cc):
        r = 5 + i
        put(wk, f"A{r}", t)
        put(wk, f"B{r}", f, '#,##0.00')
        put(wk, f"C{r}", tol)
        put(wk, f"D{r}", f'=IF(ABS(B{r})<C{r},"OK","ERROR")', fill=FILL_CTRL)
    rk = 5 + len(cc)
    put(wk, f"A{rk}", "Resultado global", font=F_B)
    put(wk, f"D{rk}", f'=IF(COUNTIF(D5:D{rk-1},"OK")={len(cc)},"OK","ERROR")', font=F_B, fill=FILL_CTRL)
    put(wk, f"A{rk+2}", "PN al 31/12/2025 en moneda de dic-2026 (debe coincidir con el PN inicial de ajuste_2026.xlsx)")
    put(wk, f"B{rk+2}", "=EEPN!F9", NUM)
    put(wk, f"A{rk+3}", "Resultado 2025 en moneda de dic-2026")
    put(wk, f"B{rk+3}", "='ER 2025'!B16", NUM)
    put(wk, f"A{rk+4}", "Aviso: el coeficiente usa el índice de dic-2026 ESTIMADO. Recalcular con el índice oficial.")
    widths(wk, [80, 20, 11, 11])
    ruta = os.path.join(OUT, "comparativo_2025_en_moneda_dic2026.xlsx")
    wb.save(ruta)
    return ruta


# ============================================================================= main
def main():
    os.makedirs(OUT, exist_ok=True)
    r25, m25 = libro_anual(2025)
    recalcular(r25)
    ws = leer_valores(r25, "BSS ajustado")
    AJ = m25["AJ"]
    ini = {}
    res_total = 0.0
    for cod, nom, rub in m25["cuentas"]:
        v = ws[f"H{AJ['fila'][cod]}"].value
        inicio_nombres[cod] = nom
        if cod[0] in "45":
            res_total += v
        else:
            ini[cod] = v
    # el resultado 2025 (incl. RECPAM) se cierra contra RNA
    ini["3.3.01"] += res_total
    ini["_resultado"] = -res_total
    r26, m26 = libro_anual(2026, inicio=ini)
    recalcular(r26)
    rc = comparativo(r25, m25)
    recalcular(rc)
    print("Generados:", r25, r26, rc, sep="\n  ")


if __name__ == "__main__":
    main()
