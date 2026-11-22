# -*- coding: utf-8 -*-
"""
Verificación de los criterios de aceptación del encargo ajuste-inflacion.

1) Recalcula copias de los tres libros con LibreOffice (headless) y lee los valores.
2) Recalcula en Python puro, sin usar las fórmulas de los libros, los importes clave
   (bienes de uso, seguros, mercaderías, PN, resultados, RECPAM por diferencia y directo)
   y los compara con los valores de LibreOffice.
3) Controla los criterios de aceptación y la ausencia de importes fijos en las columnas ajustadas.
Uso: python3 entregables/ajuste/verificar_ajuste.py
"""
import datetime as dt
import os
import shutil
import subprocess
import sys
import tempfile

import openpyxl

sys.path.insert(0, os.path.dirname(__file__))
import generador_ajuste as G  # noqa: E402  (solo para reutilizar los lectores de insumos)

OUT = G.OUT
errores, oks = [], []


def ok(cond, msg):
    (oks if cond else errores).append(msg)


def cerca(a, b, tol=1.0):
    return a is not None and b is not None and abs(a - b) <= tol


def recalc_copia(nombre, tmp):
    src = os.path.join(OUT, nombre)
    dst_dir = os.path.join(tmp, "in")
    os.makedirs(dst_dir, exist_ok=True)
    shutil.copy(src, dst_dir)
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir",
                    os.path.join(tmp, "out"), os.path.join(dst_dir, nombre)], check=True, capture_output=True)
    return openpyxl.load_workbook(os.path.join(tmp, "out", nombre), data_only=True)


IDX = {m: v for m, v, e in G.leer_indices()}


def coef(fecha, cierre):
    return IDX[dt.datetime(cierre, 12, 1)] / IDX[dt.datetime(fecha.year, fecha.month, 1)]


def calculo_independiente(anio, inicio_pn=None):
    """Devuelve dict con los importes ajustados calculados en Python."""
    bu = G.leer_bu(anio)
    seg = G.leer_seguros(anio)
    merc = G.leer_merc(anio)
    res = G.leer_resultados(anio)
    bss = {b["cod"]: b for b in G.leer_bss(anio)}
    k = lambda f: coef(f, anio)  # noqa: E731
    out = {}
    # bienes de uso
    vo = {}
    aa = {}
    am = {}
    vo_i = {}
    aa_i = {}
    res_venta = 0.0
    for b in bu:
        c = k(b["fecha"])
        vc, ac = G.CTA_VO[b["cta"]]
        vo[vc] = vo.get(vc, 0) + (b["vo_ini"] + b["alta"] - b["baja_vo"]) * c
        vo_i[vc] = vo_i.get(vc, 0) + b["vo_ini"] * c
        if ac:
            aa[ac] = aa.get(ac, 0) + (b["aa_ini"] - b["baja_aa"] + b["am"]) * c
            aa_i[ac] = aa_i.get(ac, 0) + b["aa_ini"] * c
            ca = G.cta_amort(b["cta"], b["fun"])
            am[ca] = am.get(ca, 0) + b["am"] * c
        if b["fbaja"]:
            res_venta += b["precio"] * k(b["fbaja"]) - (b["baja_vo"] - b["baja_aa"]) * c
    out.update(vo=vo, aa=aa, am=am, vo_i=vo_i, aa_i=aa_i, res_venta=res_venta)
    # seguros
    nd = sum((s["prima"] - s["prima"] * s["meses"] / 12) * k(s["fpago"]) for s in seg)
    nd_i = sum(s["nd_ini"] * k(s["fpago_ant"]) for s in seg)
    gasto = {}
    for s in seg:
        cta = "5.3.05" if s["fun"] == "Administración" else "5.2.06"
        gasto[cta] = gasto.get(cta, 0) + s["nd_ini"] * k(s["fpago_ant"]) + s["prima"] * s["meses"] / 12 * k(s["fpago"])
    out.update(seg_nd=nd, seg_nd_i=nd_i, seg_gasto=gasto)
    # mercaderías
    out["ef"] = sum(v * k(m) for m, v in merc["ef"])
    out["ei"] = sum(v * k(m) for m, v in merc["ei"])
    out["cmv"] = sum(sum(vals) * k(m) for m, vals in merc["matriz"])
    # resultados
    coefm = [k(dt.datetime(anio, j + 1, 1)) for j in range(12)]
    especiales = {"5.1.01": out["cmv"], "4.3.01": res_venta, **am, **gasto}
    r_aj = 0.0
    for x in res:
        v = especiales[x["cod"]] if x["cod"] in especiales else sum(a * b for a, b in zip(x["meses"], coefm))
        r_aj += x["signo"] * v
    out["res_antes_recpam"] = r_aj
    # PN
    cap = G.CAPITAL[1] * k(G.CAPITAL[0])
    resv = sum(v * k(f) for f, v in G.RESERVA)
    fd, vd, _ = G.DIVIDENDO[anio]
    div = vd * k(fd)
    # monetarias
    mon = [c for c in bss if G.clase(c) in ("A-M", "P-M")]
    pmn_i = sum(bss[c]["si"] for c in mon)
    pmn_f = sum(bss[c]["sf"] for c in mon)
    kprev = IDX[dt.datetime(anio, 12, 1)] / IDX[dt.datetime(anio - 1, 12, 1)]
    nm_i = sum(vo_i.values()) - sum(aa_i.values()) + nd_i + out["ei"]
    nm_f = sum(vo.values()) - sum(aa.values()) + nd + out["ef"]
    pn_i = pmn_i * kprev + nm_i
    pn_f = pmn_f + nm_f
    if anio == 2025:
        rna_i = pn_i - cap - resv
    else:
        rna_i = inicio_pn - cap - resv
        pn_i = inicio_pn
    out.update(cap=cap, resv=resv, div=div, pn_i=pn_i, pn_f=pn_f, rna_i=rna_i)
    resultado_total = pn_f - pn_i + div
    out["recpam_dif"] = resultado_total - r_aj
    out["resultado_aj"] = resultado_total
    # RECPAM directo (flujos monetarios mensuales)
    pmn = pmn_i
    rec = 0.0
    for j in range(12):
        m = dt.datetime(anio, j + 1, 1)
        flujo = 0.0
        for x in res:
            if x["cod"] in ("5.1.01", "4.3.01") or x["cod"] in am or x["cod"] in gasto:
                continue
            flujo += x["signo"] * x["meses"][j]
        flujo -= merc["compras"][j][1]
        flujo -= sum(b["alta"] for b in bu if G.mes1(b["fecha"]) == m)
        flujo += sum(b["precio"] for b in bu if b["fbaja"] and G.mes1(b["fbaja"]) == m)
        flujo -= sum(s["prima"] for s in seg if G.mes1(s["fpago"]) == m)
        flujo -= vd if G.mes1(fd) == m else 0
        prev = dt.datetime(anio - 1, 12, 1) if j == 0 else dt.datetime(anio, j, 1)
        infl = IDX[m] / IDX[prev] - 1
        rec += -pmn * infl * IDX[dt.datetime(anio, 12, 1)] / IDX[m]
        pmn += flujo
    out["recpam_dir"] = rec
    out["pmn_dir_fin"] = pmn
    out["pmn_f"] = pmn_f
    return out


def valor_controles(wb):
    ws = wb["Controles"]
    d = {}
    for r in ws.iter_rows(min_row=5, values_only=True):
        if r[0] and r[3] in ("OK", "ERROR"):
            d[r[0]] = (r[1], r[3])
        elif r[0] and r[1] is not None:
            d[r[0]] = (r[1], None)
    return d


def sin_importes_fijos(wb_f, hoja, columnas, filas):
    """Las columnas ajustadas deben ser fórmulas (no constantes numéricas)."""
    ws = wb_f[hoja]
    malas = []
    for col in columnas:
        for r in filas:
            v = ws[f"{col}{r}"].value
            if isinstance(v, (int, float)) and v != 0:
                malas.append(f"{hoja}!{col}{r}")
    return malas


def main():
    tmp = tempfile.mkdtemp()
    try:
        w25 = recalc_copia("ajuste_2025.xlsx", os.path.join(tmp, "a"))
        w26 = recalc_copia("ajuste_2026.xlsx", os.path.join(tmp, "b"))
        wc = recalc_copia("comparativo_2025_en_moneda_dic2026.xlsx", os.path.join(tmp, "c"))
    finally:
        pass
    f25 = openpyxl.load_workbook(os.path.join(OUT, "ajuste_2025.xlsx"))
    f26 = openpyxl.load_workbook(os.path.join(OUT, "ajuste_2026.xlsx"))

    # errores de fórmula
    for nombre, wb in (("2025", w25), ("2026", w26), ("comparativo", wc)):
        n = sum(1 for ws in wb for row in ws.iter_rows() for c in row
                if isinstance(c.value, str) and (c.value.startswith("#") or c.value.startswith("Err:")))
        ok(n == 0, f"{nombre}: sin errores de fórmula tras recalcular con LibreOffice ({n})")
        k = valor_controles(wb)
        ok(k.get("Resultado global", (None, None))[1] == "OK", f"{nombre}: hoja Controles = OK")

    # cálculo independiente 2025
    c25 = calculo_independiente(2025)
    k25 = valor_controles(w25)
    ok(cerca(k25["RECPAM (aj.)"][0], c25["recpam_dif"]), f"2025 RECPAM por diferencia: libro {k25['RECPAM (aj.)'][0]:,.0f} vs Python {c25['recpam_dif']:,.0f}")
    ok(cerca(k25["RECPAM método directo"][0], c25["recpam_dir"]), f"2025 RECPAM directo: libro {k25['RECPAM método directo'][0]:,.0f} vs Python {c25['recpam_dir']:,.0f}")
    ok(cerca(c25["recpam_dif"], c25["recpam_dir"], 1.0), "2025 RECPAM: diferencia − directo < $1 (Python)")
    ok(cerca(c25["pmn_dir_fin"], c25["pmn_f"]), "2025 PMN de cierre por flujos = BSS (Python)")
    ok(cerca(k25["Resultado antes del RECPAM (aj.)"][0], c25["res_antes_recpam"]), "2025 resultado antes del RECPAM: libro = Python")
    ok(cerca(k25["Resultado del ejercicio ajustado"][0], c25["resultado_aj"]), "2025 resultado ajustado: libro = Python")
    ok(cerca(k25["PN ajustado al 31/12/2025"][0], c25["pn_f"]), "2025 PN ajustado de cierre: libro = Python")
    ok(cerca(k25["PN ajustado al 31/12/2024 (moneda dic-2025)"][0], c25["pn_i"]), "2025 PN ajustado de inicio: libro = Python")

    # criterio: PN 2025 × coef = PN inicial 2026
    k26 = valor_controles(w26)
    coef2526 = IDX[dt.datetime(2026, 12, 1)] / IDX[dt.datetime(2025, 12, 1)]
    pn25 = k25["PN ajustado al 31/12/2025"][0]
    pn26i = k26["PN ajustado al 31/12/2025 (moneda dic-2026)"][0]
    ok(cerca(pn25 * coef2526, pn26i, 0.5), f"PN 31/12/2025 aj. × {coef2526:.6f} = PN inicial 2026: {pn25*coef2526:,.2f} vs {pn26i:,.2f}")
    pnc = valor_controles(wc)["PN al 31/12/2025 en moneda de dic-2026 (debe coincidir con el PN inicial de ajuste_2026.xlsx)"][0]
    ok(cerca(pnc, pn26i, 0.5), f"Comparativo: PN 31/12/2025 en moneda dic-2026 = PN inicial 2026 ({pnc:,.2f})")
    resc = valor_controles(wc)["Resultado 2025 en moneda de dic-2026"][0]
    ok(cerca(resc, k25["Resultado del ejercicio ajustado"][0] * coef2526, 0.5), "Comparativo: resultado 2025 = resultado aj. × coef.")

    # cálculo independiente 2026
    c26 = calculo_independiente(2026, inicio_pn=c25["pn_f"] * coef2526)
    ok(cerca(k26["RECPAM (aj.)"][0], c26["recpam_dif"]), f"2026 RECPAM por diferencia: libro {k26['RECPAM (aj.)'][0]:,.0f} vs Python {c26['recpam_dif']:,.0f}")
    ok(cerca(k26["RECPAM método directo"][0], c26["recpam_dir"]), f"2026 RECPAM directo: libro {k26['RECPAM método directo'][0]:,.0f} vs Python {c26['recpam_dir']:,.0f}")
    ok(cerca(c26["pmn_dir_fin"], c26["pmn_f"]), "2026 PMN de cierre por flujos = BSS (Python)")
    ok(cerca(k26["Resultado antes del RECPAM (aj.)"][0], c26["res_antes_recpam"]), "2026 resultado antes del RECPAM: libro = Python")
    ok(cerca(k26["Resultado del ejercicio ajustado"][0], c26["resultado_aj"]), "2026 resultado ajustado: libro = Python")
    ok(cerca(k26["PN ajustado al 31/12/2026"][0], c26["pn_f"]), "2026 PN ajustado de cierre: libro = Python")
    # saldos iniciales 2026 desde el origen = cierre 2025 × coef (no monetarias)
    ok(cerca(sum(c26["vo_i"].values()), sum(c25["vo"].values()) * coef2526), "2026 VO inicial (origen) = VO cierre 2025 aj. × coef.")
    ok(cerca(c26["ei"], c25["ef"] * coef2526), "2026 mercaderías iniciales = cierre 2025 aj. × coef.")
    ok(cerca(c26["seg_nd_i"], c25["seg_nd"] * coef2526), "2026 seguros iniciales = cierre 2025 aj. × coef.")

    # importes fijos en columnas ajustadas
    malas = []
    for wbf in (f25, f26):
        n_bu = wbf["Bienes de uso"].max_row
        malas += sin_importes_fijos(wbf, "Bienes de uso", ["I", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA"], range(5, n_bu))
        malas += sin_importes_fijos(wbf, "Seguros", list("FJMNOPQRSTU"), range(5, 16))
        malas += sin_importes_fijos(wbf, "Mercaderías", ["C", "D"], list(range(5, 8)) + list(range(10, 23)) + list(range(25, 28)))
        malas += sin_importes_fijos(wbf, "Resultados", ["U", "V"], range(7, 31))
        malas += sin_importes_fijos(wbf, "BSS ajustado", ["F", "H", "I"], range(5, 60))
        malas += sin_importes_fijos(wbf, "Patrimonio neto", ["D", "F", "G"], range(5, 39))
    ok(not malas, f"Columnas ajustadas sin importes fijos ({len(malas)}: {malas[:5]})")

    print("=" * 80)
    for m in oks:
        print("OK     ", m)
    for m in errores:
        print("ERROR  ", m)
    print("=" * 80)
    print(f"{len(oks)} OK · {len(errores)} errores")
    print("\nResumen:")
    for y, c in (("2025", c25), ("2026", c26)):
        print(f"  {y}: resultado antes RECPAM {c['res_antes_recpam']:,.0f} · RECPAM {c['recpam_dif']:,.0f} "
              f"(directo {c['recpam_dir']:,.0f}) · resultado aj. {c['resultado_aj']:,.0f} · PN cierre {c['pn_f']:,.0f}")
    shutil.rmtree(tmp, ignore_errors=True)
    sys.exit(1 if errores else 0)


if __name__ == "__main__":
    main()
