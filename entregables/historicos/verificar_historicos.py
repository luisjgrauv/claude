# -*- coding: utf-8 -*-
"""
Verificación de los criterios de aceptación del encargo insumos-historicos.
1) Recalcula todos los libros con LibreOffice (soffice --headless --convert-to csv, todas las hojas).
2) Controla diferencias de conciliación, cuadre del BSS y fechas de origen.
Uso: python3 entregables/historicos/verificar_historicos.py [directorio_temporal]
"""
import csv
import datetime as dt
import os
import subprocess
import sys
import tempfile

import openpyxl

RAIZ = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
LIBROS = ["insumos/andina/auxiliares_complementarios_2025.xlsx", "insumos/andina/bss_2026.xlsx",
          "insumos/andina/auxiliares_2026.xlsx", "entregables/historicos/eecc_historicos_2025_2026.xlsx"]
FILTRO = "csv:Text - txt - csv (StarCalc):59,34,76,1,,0,false,true,false,false,false,-1"
errores, oks = [], []


def ok(cond, msg):
    (oks if cond else errores).append(msg)


def num(x):
    try:
        return float(x.replace(",", "."))
    except (ValueError, AttributeError):
        return None


def convertir(tmp):
    hojas = {}
    for lib in LIBROS:
        subprocess.run(["soffice", "--headless", "--convert-to", FILTRO, "--outdir", tmp, os.path.join(RAIZ, lib)],
                       check=True, capture_output=True)
    for f in os.listdir(tmp.encode()):
        nombre = f.decode("latin-1") if not _utf8(f) else f.decode()
        with open(os.path.join(tmp.encode(), f), encoding="utf-8") as fh:
            hojas[nombre[:-4]] = list(csv.reader(fh, delimiter=";"))
    return hojas


def _utf8(b):
    try:
        b.decode("utf-8")
        return True
    except UnicodeDecodeError:
        return False


def buscar(hojas, libro, hoja):
    for k, v in hojas.items():
        if k.startswith(libro + "-") and k.split("-", 1)[1].startswith(hoja[:6]):
            return v
    raise KeyError((libro, hoja))


def col_dif(filas, encabezado_contiene="Diferencia"):
    """Devuelve todos los valores numéricos de las columnas cuyo encabezado contiene 'Diferencia'."""
    vals, idx = [], []
    for f in filas:
        if any(encabezado_contiene in c for c in f):
            idx = [i for i, c in enumerate(f) if encabezado_contiene in c]
            continue
        for i in idx:
            if i < len(f) and f[i] != "":
                vals.append((f[0], num(f[i])))
    return vals


def main():
    tmp = sys.argv[1] if len(sys.argv) > 1 else tempfile.mkdtemp()
    os.makedirs(tmp, exist_ok=True)
    hojas = convertir(tmp)

    # 1. Conciliaciones (2025 y 2026) y resultados mensuales
    for libro in ("auxiliares_complementarios_2025", "auxiliares_2026"):
        for hoja in ("Conciliación BSS", "Resultados mensuales"):
            difs = col_dif(buscar(hojas, libro, hoja))
            malos = [d for d in difs if d[1] is None or abs(d[1]) > 0.0001]
            ok(len(difs) > 0 and not malos, f"{libro} / {hoja}: {len(difs)} diferencias, todas 0" + (f" — FALLAN {malos}" if malos else ""))
        merc = buscar(hojas, libro, "Mercaderías")
        ctrl = [f for f in merc if f and f[0].startswith("Control")][0]
        ok(all(num(x) == 0 for x in ctrl[1:] if x != ""), f"{libro} / Mercaderías: control CMV por mes de compra = 0")

    # Conciliación original 2025 (archivo no modificado)
    wb = openpyxl.load_workbook(os.path.join(RAIZ, "insumos/andina/auxiliares_2025.xlsx"))
    difs = [r[3] for r in wb["Conciliación BSS"].iter_rows(min_row=5, values_only=True) if isinstance(r[3], (int, float))]
    difs += [r[6] for r in wb["Conciliación BSS"].iter_rows(min_row=5, values_only=True) if isinstance(r[6], (int, float))]
    ok(all(d == 0 for d in difs), f"auxiliares_2025 (original) / Conciliación BSS: {len(difs)} diferencias, todas 0")

    # 2. BSS 2026
    bss = buscar(hojas, "bss_2026", "BSS")
    tot = [f for f in bss if f and f[0] == "TOTAL"][0]
    si, de, cr, sf, sd, sa = (num(x) for x in tot[3:9])
    ok(si == 0 and sf == 0, f"BSS 2026 (LibreOffice): saldos iniciales {si:,.0f} y finales {sf:,.0f}")
    ok(de == cr, f"BSS 2026 (LibreOffice): débitos {de:,.0f} = créditos {cr:,.0f}")
    ok(sd == sa, f"BSS 2026 (LibreOffice): saldos deudores {sd:,.0f} = acreedores {sa:,.0f}")
    for f in bss:
        if f and f[0].startswith("Control"):
            ok(all(num(x) == 0 for x in f[1:] if x != ""), f"BSS 2026: {f[0]} = 0")
    # Recalculo independiente en Python desde el CSV entregado
    with open(os.path.join(RAIZ, "insumos/andina/bss_2026.csv"), encoding="utf-8") as fh:
        filas = list(csv.DictReader(fh, delimiter=";"))
    s = lambda c: sum(int(f[c]) for f in filas)
    ok(all(int(f["Saldo inicial 31/12/2025"]) + int(f["Débitos 2026"]) - int(f["Créditos 2026"]) == int(f["Saldo final 31/12/2026"]) for f in filas),
       "BSS 2026 (Python, CSV): saldo inicial + débitos − créditos = saldo final en todas las cuentas")
    ok(s("Débitos 2026") == s("Créditos 2026") == de, "BSS 2026 (Python, CSV): totales iguales a los de LibreOffice")
    ok(s("Saldo inicial 31/12/2025") == 0 and s("Saldo final 31/12/2026") == 0, "BSS 2026 (Python, CSV): saldos iniciales y finales suman 0")
    # Saldo inicial 2026 = saldo final 2025 después del cierre
    with open(os.path.join(RAIZ, "insumos/andina/bss_2025.csv"), encoding="utf-8") as fh:
        f25 = {f["Código"]: int(f["Saldo final 31/12/2025"]) for f in csv.DictReader(fh, delimiter=";")}
    res25 = sum(v for k, v in f25.items() if k[0] in "45")
    si26 = {f["Código"]: int(f["Saldo inicial 31/12/2025"]) for f in filas}
    esperado = {k: (v if k[0] in "123" else 0) for k, v in f25.items()}
    esperado["3.3.01"] += res25
    ok(all(si26[k] == v for k, v in esperado.items()) and all(v == 0 for k, v in si26.items() if k not in esperado),
       f"BSS 2026: saldo inicial = saldo final 2025 después del cierre (RNA inicial {si26['3.3.01']:,.0f})")

    # 3. Bienes y pólizas al 31/12/2026
    bu = buscar(hojas, "auxiliares_2026", "Bienes de uso")
    cuenta = {f[0]: num(f[1]) for f in bu if f and f[0] in ("Camión", "Vehículo ejecutivo", "Camioneta de reparto", "Inmuebles")}
    ok(cuenta == {"Camión": 10, "Vehículo ejecutivo": 5, "Camioneta de reparto": 2, "Inmuebles": 2}, f"Existencias físicas 31/12/2026: {cuenta}")
    seg = buscar(hojas, "auxiliares_2026", "Seguros")
    ok(sum(1 for f in seg if f and f[0].startswith("POL-")) == 10, "Pólizas vigentes 31/12/2026: 10")

    # 4. Fechas de origen de partidas no monetarias (sobre los .xlsx)
    w26 = openpyxl.load_workbook(os.path.join(RAIZ, "insumos/andina/auxiliares_2026.xlsx"))
    w25c = openpyxl.load_workbook(os.path.join(RAIZ, "insumos/andina/auxiliares_complementarios_2025.xlsx"))
    es_f = lambda v: isinstance(v, (dt.date, dt.datetime))
    filas_bu = [r for r in w26["Bienes de uso"].iter_rows(min_row=5, values_only=True) if r[0] and r[0] != "TOTAL" and r[19]]
    ok(all(es_f(r[4]) for r in filas_bu), f"Bienes de uso 2026: {len(filas_bu)} bienes con fecha de origen")
    filas_pol = [r for r in w26["Seguros"].iter_rows(min_row=5, values_only=True) if r[0] and str(r[0]).startswith("POL-")]
    ok(all(es_f(r[7]) for r in filas_pol), "Seguros 2026: 10 pólizas con fecha de pago (origen)")
    for w, n in ((w25c, "2025"), (w26, "2026")):
        ws = w["Mercaderías"]
        capas = [r for r in ws.iter_rows(min_row=5, values_only=True) if r[0] is not None and isinstance(r[1], (int, float)) and r[2] is None]
        ok(len(capas) > 0 and all(es_f(r[0]) for r in capas), f"Mercaderías {n}: {len(capas)} capas/filas FIFO con mes de compra")
        pn = w["Patrimonio neto"]
        sinf = [r for r in pn.iter_rows(min_row=5, values_only=True)
                if isinstance(r[0], int) and r[0] > 2000 and not es_f(r[1])]
        ok(not sinf, f"Patrimonio neto {n}: reserva legal y resultados por año con fecha de origen")

    # 5. EECC
    esp = buscar(hojas, "eecc_historicos_2025_2026", "ESP")
    er = buscar(hojas, "eecc_historicos_2025_2026", "ER")
    for f in esp + er:
        if f and f[0].startswith("Control"):
            ok(all(num(x) == 0 for x in f[1:] if x != ""), f"EECC: {f[0]} = 0")
    res = [f for f in er if f and f[0] == "RESULTADO DEL EJERCICIO"][0]
    ta = [f for f in esp if f and f[0] == "TOTAL ACTIVO"][0]
    print("Resultado 2025 / 2026:", res[1], res[2])
    print("Total activo 31/12/2024 / 2025 / 2026:", ta[1:4])

    # 6. Originales 2025 sin cambios
    git = subprocess.run(["git", "status", "--porcelain", "--", "insumos/andina/bss_2025.xlsx", "insumos/andina/bss_2025.csv",
                          "insumos/andina/auxiliares_2025.xlsx", "insumos/andina/datos_adicionales.md"],
                         cwd=RAIZ, capture_output=True, text=True).stdout.strip()
    ok(git == "", "Archivos 2025 originales sin modificar (git status limpio)")

    for m in oks:
        print("OK   ", m)
    for m in errores:
        print("ERROR", m)
    print(f"\n{len(oks)} controles OK, {len(errores)} errores")
    sys.exit(1 if errores else 0)


if __name__ == "__main__":
    main()
