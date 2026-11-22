# -*- coding: utf-8 -*-
"""
Generador de la base histórica nominal de Distribuidora Andina S.A. (FICTICIA), 2025 y 2026.

Lee (sin modificar) insumos/andina/bss_2025.xlsx y auxiliares_2025.xlsx, y genera:
  - insumos/andina/auxiliares_complementarios_2025.xlsx
  - insumos/andina/bss_2026.xlsx y bss_2026.csv
  - insumos/andina/auxiliares_2026.xlsx
  - entregables/historicos/eecc_historicos_2025_2026.xlsx
  - entregables/historicos/control_generador.json (cifras de control para datos_adicionales_2026.md)

Valores históricos nominales, en pesos enteros. Uso: python3 entregables/historicos/generador_historicos.py
"""
import csv
import datetime as dt
import json
import os

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L

D = dt.date
RAIZ = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
INS = os.path.join(RAIZ, "insumos", "andina")
ENT = os.path.join(RAIZ, "entregables", "historicos")
EMPRESA = "Distribuidora Andina S.A. (FICTICIA)"
LEYENDA = "Valores históricos nominales (sin reexpresar por inflación). Importes en pesos."
MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
G = [1.02 ** i for i in range(12)]  # inflación nominal de referencia ~2 % mensual

# --------------------------------------------------------------------------------------
# Utilidades
# --------------------------------------------------------------------------------------

def reparte(total, pesos, unidad=100_000):
    """Reparte total según pesos, redondeando a 'unidad'; el último mes absorbe el redondeo."""
    s = sum(pesos)
    out = [int(total * p / s / unidad + 0.5) * unidad for p in pesos[:-1]]
    out.append(total - sum(out))
    return out


def habil_atras(d):
    while d.weekday() >= 5:
        d -= dt.timedelta(days=1)
    return d


def fecha_asamblea(y):
    return habil_atras(D(y, 4, 28))


def fecha_pago_div(y):
    return habil_atras(D(y, 4, 30))


# --------------------------------------------------------------------------------------
# Lectura de insumos 2025 (solo lectura)
# --------------------------------------------------------------------------------------

def leer_bss2025():
    wb = openpyxl.load_workbook(os.path.join(INS, "bss_2025.xlsx"))
    ws = wb["BSS"]
    cuentas = []
    for r in ws.iter_rows(min_row=4, values_only=True):
        if r[0] in (None, "TOTAL", "Código", "Póliza"):
            continue
        cuentas.append(dict(cod=r[0], cta=r[1], rubro=r[2], si=r[3], deb=r[4], cred=r[5], sf=r[6]))
    return cuentas


def leer_aux2025():
    wb = openpyxl.load_workbook(os.path.join(INS, "auxiliares_2025.xlsx"))
    bienes, polizas = [], []
    for r in wb["Bienes de uso"].iter_rows(min_row=4, values_only=True):
        if r[0] in (None, "TOTAL", "Código", "Póliza"):
            continue
        bienes.append(dict(cod=r[0], tipo=r[1], comp=r[2], ubic=r[3], fecha=r[4].date(), vo=r[5], vida=r[6],
                           vo24=r[7], aa24=r[8], am25=r[9], aa25=r[10], vrc25=r[11], cuenta=r[12], func=r[13]))
    for r in wb["Seguros"].iter_rows(min_row=4, values_only=True):
        if r[0] in (None, "TOTAL", "Código", "Póliza"):
            continue
        polizas.append(dict(pol=r[0], aseg=r[1], ramo=r[2], cob=r[3], bienes=r[4], desde=r[5].date(),
                            hasta=r[6].date(), pago=r[7].date(), prima=r[8], prima_ant=r[9], nd_ini=r[10],
                            dev_ant=r[11], meses=r[12], dev_act=r[13], gasto=r[14], nd_fin=r[15], func=r[16]))
    return bienes, polizas


BSS25 = leer_bss2025()
B25 = {c["cod"]: c for c in BSS25}
BIENES25, POL25 = leer_aux2025()

# --------------------------------------------------------------------------------------
# Lógica replicada del generador original
# --------------------------------------------------------------------------------------

def amort_anual(b, anio):
    """Lineal, año de alta completo, año de baja sin amortización, sin valor residual."""
    if b["vida"] == 0:
        return 0
    alta = b["fecha"].year
    baja = b.get("baja").year if b.get("baja") else None
    if anio < alta or (baja is not None and anio >= baja):
        return 0
    if anio - alta + 1 > b["vida"]:
        return 0
    return b["vo"] // b["vida"] if b["vo"] % b["vida"] == 0 else round(b["vo"] / b["vida"])


def amort_acum(b, hasta_anio):
    return sum(amort_anual(b, a) for a in range(b["fecha"].year, hasta_anio + 1))


# Control: la lógica replicada reproduce el auxiliar 2025.
for b in BIENES25:
    assert amort_anual(b, 2025) == b["am25"], b["cod"]
    assert amort_acum(b, 2024) == b["aa24"], b["cod"]


def seguro_mensual(p_desde, prima_ant, prima, mes):
    """Devengamiento mensual lineal: póliza anterior hasta el inicio de vigencia, luego la nueva."""
    return prima_ant // 12 if mes < p_desde.month else prima // 12


for p in POL25:
    assert sum(seguro_mensual(p["desde"], p["prima_ant"], p["prima"], m) for m in range(1, 13)) == p["gasto"]

# --------------------------------------------------------------------------------------
# Mercaderías (FIFO mensual)
# --------------------------------------------------------------------------------------

def fifo(capas_ini, compras, cmv, anio):
    """capas_ini: lista [(fecha_mes, importe)]. Devuelve consumo {(origen, mes): importe} y capas finales."""
    capas = [list(c) for c in capas_ini]
    consumo = {}
    for m in range(12):
        capas.append([D(anio, m + 1, 1), compras[m]])
        resto = cmv[m]
        while resto > 0:
            o, imp = capas[0]
            t = min(imp, resto)
            consumo[(o, m)] = consumo.get((o, m), 0) + t
            resto -= t
            capas[0][1] -= t
            if capas[0][1] == 0:
                capas.pop(0)
    return consumo, [tuple(c) for c in capas if c[1] > 0]


# 2025
VENTAS25 = reparte(14_000_000_000, G)
CMV25 = [v * 7 // 10 for v in VENTAS25]  # margen bruto 30 % constante (9.800 / 14.000)
VAR_EX25 = reparte(350_000_000, [1] * 12)  # existencia crece de 900 M a 1.250 M
COMPRAS25 = [c + d for c, d in zip(CMV25, VAR_EX25)]
CAPAS_INI25 = [(D(2024, 11, 1), 160_000_000), (D(2024, 12, 1), 740_000_000)]  # supuesto
assert sum(CMV25) == B25["5.1.01"]["deb"] and sum(COMPRAS25) == B25["1.5.01"]["deb"]
assert sum(i for _, i in CAPAS_INI25) == B25["1.5.01"]["si"]
CONS25, CAPAS_FIN25 = fifo(CAPAS_INI25, COMPRAS25, CMV25, 2025)
assert sum(i for _, i in CAPAS_FIN25) == B25["1.5.01"]["sf"]

# 2026
VENTAS26 = reparte(18_200_000_000, G)  # +30 % nominal
CMV26 = [v * 7 // 10 for v in VENTAS26]
EXFIN26 = 1_625_000_000
VAR_EX26 = reparte(EXFIN26 - 1_250_000_000, [1] * 12)
COMPRAS26 = [c + d for c, d in zip(CMV26, VAR_EX26)]
CONS26, CAPAS_FIN26 = fifo(CAPAS_FIN25, COMPRAS26, CMV26, 2026)
assert sum(i for _, i in CAPAS_FIN26) == EXFIN26

# --------------------------------------------------------------------------------------
# Patrimonio neto: historia 2014-2026 (supuestos verosímiles, cuadrados contra el BSS)
# --------------------------------------------------------------------------------------
CAPITAL = 20_000_000
FECHA_CAPITAL = D(2014, 3, 14)  # supuesto de día; el mes (03/2014) es dato
RESULTADOS = {2014: 1_850_000, 2015: 3_240_000, 2016: 4_710_000, 2017: 6_980_000, 2018: 10_420_000,
              2019: 16_350_000, 2020: 23_900_000, 2021: 41_600_000, 2022: 120_000_000, 2023: 330_000_000}
# Dividendos: (año de la asamblea) -> importe; se imputan al resultado del ejercicio inmediato anterior.
DIVIDENDOS = {2019: 4_000_000, 2020: 6_000_000, 2021: 10_000_000, 2022: 18_000_000, 2023: 30_000_000,
              2024: 60_000_000, 2025: 120_000_000, 2026: 300_000_000}
RESERVA = {}
acum = 0
for y in range(2014, 2024):
    r = min(round(RESULTADOS[y] * 0.05), 4_000_000 - acum)  # art. 70 LGS: 5 % hasta el 20 % del capital
    if r > 0:
        RESERVA[y] = r
        acum += r
assert acum == B25["3.2.01"]["si"] * -1
RNA24 = -B25["3.3.01"]["si"]
remanentes = sum(RESULTADOS[y] - RESERVA.get(y, 0) - DIVIDENDOS.get(y + 1, 0) for y in range(2014, 2024))
RESULTADOS[2024] = RNA24 - remanentes  # resultado 2024 = residuo para cuadrar los RNA al 31/12/2024
assert DIVIDENDOS[2025] == B25["3.3.01"]["deb"]
RES25 = -sum(c["sf"] for c in BSS25 if c["cod"][0] in "45")
RESULTADOS[2025] = RES25

# --------------------------------------------------------------------------------------
# Bienes de uso 2026
# --------------------------------------------------------------------------------------
FECHA_C11 = D(2026, 6, 15)
FECHA_BAJA_C01 = D(2026, 6, 15)
PRECIO_C01 = 28_000_000
VO_C11 = 175_000_000
BIENES26 = []
for b in BIENES25:
    nb = dict(b)
    nb["baja"] = FECHA_BAJA_C01 if b["cod"] == "C01" else None
    BIENES26.append(nb)
idx = [i for i, b in enumerate(BIENES26) if b["cod"] == "C10"][0]
BIENES26.insert(idx + 1, dict(cod="C11", tipo="Camión", comp="Camión pesado 26 t con caja térmica", ubic="AJ012WX",
                              fecha=FECHA_C11, vo=VO_C11, vida=10, cuenta="Rodados - camiones",
                              func="Comercialización", baja=None))
for b in BIENES26:
    b["vo25"] = b["vo"] if b["fecha"].year <= 2025 else 0
    b["aa25"] = amort_acum(b, 2025) if b["fecha"].year <= 2025 else 0
    b["alta26"] = b["vo"] if b["fecha"].year == 2026 else 0
    b["baja_vo"] = b["vo"] if b["baja"] else 0
    b["baja_aa"] = b["aa25"] if b["baja"] else 0
    b["am26"] = amort_anual(b, 2026)
    b["vo26"] = b["vo25"] + b["alta26"] - b["baja_vo"]
    b["aa26"] = b["aa25"] - b["baja_aa"] + b["am26"]
    b["precio"] = PRECIO_C01 if b["baja"] else 0
    b["res_venta"] = b["precio"] - (b["baja_vo"] - b["baja_aa"]) if b["baja"] else 0
ACTIVOS26 = [b for b in BIENES26 if not b["baja"]]
assert sum(1 for b in ACTIVOS26 if b["tipo"] == "Camión") == 10
assert sum(1 for b in ACTIVOS26 if b["tipo"] == "Vehículo ejecutivo") == 5
assert sum(1 for b in ACTIVOS26 if b["tipo"] == "Camioneta de reparto") == 2
assert len({b["ubic"] for b in ACTIVOS26 if b["cod"].startswith("INM")}) == 2

CTA_AMORT = {  # cuenta patrimonial -> (acumulada, [(función, cuenta de resultado)])
    "Edificios": "Amortización acumulada edificios",
    "Rodados - camiones": "Amortización acumulada rodados - camiones",
    "Rodados - vehículos ejecutivos": "Amortización acumulada rodados - vehículos ejecutivos",
    "Rodados - camionetas de reparto": "Amortización acumulada rodados - camionetas de reparto",
}
CTA_GASTO_AM = {("Edificios", "Comercialización"): "5.2.07", ("Edificios", "Administración"): "5.3.06",
                ("Rodados - camiones", "Comercialización"): "5.2.08",
                ("Rodados - camionetas de reparto", "Comercialización"): "5.2.09",
                ("Rodados - vehículos ejecutivos", "Administración"): "5.3.07"}


def amort_por_cuenta(bienes, campo):
    out = {}
    for b in bienes:
        if b[campo]:
            k = CTA_GASTO_AM[(b["cuenta"], b["func"])]
            out[k] = out.get(k, 0) + b[campo]
    return out


AM25 = amort_por_cuenta(BIENES25, "am25")
for k, v in AM25.items():
    assert B25[k]["deb"] == v
AM26 = amort_por_cuenta(BIENES26, "am26")

# --------------------------------------------------------------------------------------
# Seguros 2026: renovación con igual vigencia y primas +35 % nominal
# --------------------------------------------------------------------------------------
POL26 = []
for p in POL25:
    q = dict(p)
    q["desde"] = D(p["desde"].year + 1, p["desde"].month, p["desde"].day)
    q["hasta"] = D(p["hasta"].year + 1, p["hasta"].month, p["hasta"].day)
    q["pago"] = q["desde"]
    q["prima_ant"] = p["prima"]
    q["prima"] = p["prima"] * 135 // 100
    assert q["prima"] % 12 == 0
    q["nd_ini"] = p["nd_fin"]
    q["meses"] = 12 - q["desde"].month + 1
    q["dev_ant"] = (q["desde"].month - 1) * (q["prima_ant"] // 12)
    assert q["dev_ant"] == q["nd_ini"], q["pol"]
    q["dev_act"] = q["meses"] * (q["prima"] // 12)
    q["gasto"] = q["dev_ant"] + q["dev_act"]
    q["nd_fin"] = q["prima"] - q["dev_act"]
    POL26.append(q)


def seguros_mensuales(pols):
    out = {"Comercialización": [0] * 12, "Administración": [0] * 12}
    for p in pols:
        for m in range(12):
            out[p["func"]][m] += seguro_mensual(p["desde"], p["prima_ant"], p["prima"], m + 1)
    return out


SEG25 = seguros_mensuales(POL25)
assert sum(SEG25["Comercialización"]) == B25["5.2.06"]["deb"] and sum(SEG25["Administración"]) == B25["5.3.05"]["deb"]
SEG26 = seguros_mensuales(POL26)

# --------------------------------------------------------------------------------------
# Resultados mensuales
# --------------------------------------------------------------------------------------

def saldos_prestamo(ini):
    """Amortización de capital trimestral de 37,5 M (mar, jun, sep, dic); saldo vigente en cada mes."""
    return [ini - 37_500_000 * (m // 3) for m in range(12)]


def resultados_mensuales(anio):
    if anio == 2025:
        tot = {k: B25[k]["deb"] if k[0] == "5" else B25[k]["cred"] for k in B25 if k[0] in "45"}
        ventas, cmv, seg, am, prest = VENTAS25, CMV25, SEG25, AM25, saldos_prestamo(600_000_000)
    else:
        tot = dict(TOT_RES26)
        ventas, cmv, seg, am, prest = VENTAS26, CMV26, SEG26, AM26, saldos_prestamo(450_000_000)
    m = {}
    m["4.1.01"] = ventas
    m["5.1.01"] = cmv
    m["5.2.04"] = [v * 3 // 100 for v in ventas]  # IIBB 3 % de las ventas
    m["5.2.06"] = seg["Comercialización"]
    m["5.3.05"] = seg["Administración"]
    for k, v in am.items():
        m[k] = reparte(v, [1] * 12, 1)
    m["5.4.01"] = reparte(tot["5.4.01"], prest, 1_000)
    m["5.5.01"] = [0] * 11 + [tot["5.5.01"]]  # impuesto determinado al cierre
    if "4.3.01" in tot:
        m["4.3.01"] = [0] * 5 + [tot["4.3.01"]] + [0] * 6  # venta del C01 en junio
    for k in tot:
        if k not in m:
            m[k] = reparte(tot[k], G, 100_000 if tot[k] >= 30_000_000 else 1_000)
    for k in tot:
        assert sum(m[k]) == tot[k], (anio, k)
    return m


RM25 = resultados_mensuales(2025)

# --------------------------------------------------------------------------------------
# 2026: asientos resumidos del ejercicio -> BSS
# --------------------------------------------------------------------------------------
CUENTAS26 = [dict(cod=c["cod"], cta=c["cta"], rubro=c["rubro"]) for c in BSS25]
pos = [i for i, c in enumerate(CUENTAS26) if c["cod"] == "4.2.01"][0]
CUENTAS26.insert(pos + 1, dict(cod="4.3.01", cta="Resultado por venta de bienes de uso", rubro="Otros ingresos y egresos"))
SI26 = {}
for c in BSS25:
    SI26[c["cod"]] = c["sf"] if c["cod"][0] in "123" else 0
SI26["3.3.01"] = B25["3.3.01"]["sf"] - RES25  # cierre de resultados 2025 contra RNA
SI26["4.3.01"] = 0
assert sum(SI26.values()) == 0

DEB = {c["cod"]: 0 for c in CUENTAS26}
CRE = {c["cod"]: 0 for c in CUENTAS26}
ASIENTOS = []


def asiento(desc, *mov):
    """mov: (código, débito, crédito)."""
    td = sum(d for _, d, _ in mov)
    tc = sum(c for _, _, c in mov)
    assert td == tc, (desc, td, tc)
    for cod, d, c in mov:
        DEB[cod] += d
        CRE[cod] += c
    ASIENTOS.append((desc, mov))


BANCO = "1.1.02"
V26 = sum(VENTAS26)
P26 = sum(COMPRAS26)
GASTOS26 = {"5.2.01": 1_170_000_000, "5.2.02": 494_000_000, "5.2.03": 156_000_000, "5.2.05": 39_000_000,
            "5.3.01": 624_000_000, "5.3.02": 91_000_000, "5.3.03": 110_500_000, "5.3.04": 45_500_000}
INT_PF26 = 72_000_000
INT_PREST26 = 105_000_000
DEUDORES_FIN26 = 2_275_000_000
APLIC_PREV26 = 21_000_000
PROV_FIN26 = 1_495_000_000
SUELDOS_FIN26 = 169_000_000
IIBB_FIN26 = 54_600_000
IVA_FIN26 = 20_000_000
ANTICIPOS26 = 350_000_000

asiento("Ventas 2026 (IVA 21 %)", ("1.3.01", V26 * 121 // 100, 0), ("4.1.01", 0, V26), ("1.4.02", 0, V26 * 21 // 100))
asiento("Aplicación de la previsión a créditos incobrables", ("1.3.02", APLIC_PREV26, 0), ("1.3.01", 0, APLIC_PREV26))
asiento("Cargo por deudores incobrables", ("5.2.05", GASTOS26["5.2.05"], 0), ("1.3.02", 0, GASTOS26["5.2.05"]))
cobranzas = SI26["1.3.01"] + DEB["1.3.01"] - CRE["1.3.01"] - DEUDORES_FIN26
asiento("Cobranzas de clientes", (BANCO, cobranzas, 0), ("1.3.01", 0, cobranzas))
asiento("Compras de mercaderías (IVA 21 %)", ("1.5.01", P26, 0), ("1.4.02", P26 * 21 // 100, 0), ("2.1.01", 0, P26 * 121 // 100))
pagos_prov = -SI26["2.1.01"] + CRE["2.1.01"] - PROV_FIN26
asiento("Pagos a proveedores", ("2.1.01", pagos_prov, 0), (BANCO, 0, pagos_prov))
asiento("Costo de mercaderías vendidas (FIFO)", ("5.1.01", sum(CMV26), 0), ("1.5.01", 0, sum(CMV26)))
sueldos = GASTOS26["5.2.01"] + GASTOS26["5.3.01"]
asiento("Devengamiento de sueldos y cargas sociales", ("5.2.01", GASTOS26["5.2.01"], 0), ("5.3.01", GASTOS26["5.3.01"], 0), ("2.2.01", 0, sueldos))
pago_sueldos = -SI26["2.2.01"] + sueldos - SUELDOS_FIN26
asiento("Pago de sueldos y cargas sociales", ("2.2.01", pago_sueldos, 0), (BANCO, 0, pago_sueldos))
iibb = V26 * 3 // 100
asiento("Impuesto sobre los ingresos brutos devengado (3 %)", ("5.2.04", iibb, 0), ("2.3.02", 0, iibb))
pago_iibb = -SI26["2.3.02"] + iibb - IIBB_FIN26
asiento("Pago de ingresos brutos", ("2.3.02", pago_iibb, 0), (BANCO, 0, pago_iibb))
serv = sum(GASTOS26[k] for k in ("5.2.02", "5.2.03", "5.3.02", "5.3.03"))
asiento("Servicios con IVA 21 % pagados por banco", *[(k, GASTOS26[k], 0) for k in ("5.2.02", "5.2.03", "5.3.02", "5.3.03")],
        ("1.4.02", serv * 21 // 100, 0), (BANCO, 0, serv * 121 // 100))
asiento("Impuestos, tasas y contribuciones", ("5.3.04", GASTOS26["5.3.04"], 0), (BANCO, 0, GASTOS26["5.3.04"]))
primas26 = sum(p["prima"] for p in POL26)
asiento("Pago de primas de seguros (renovación de 10 pólizas)", ("1.4.03", primas26, 0), (BANCO, 0, primas26))
asiento("Devengamiento de seguros", ("5.2.06", sum(SEG26["Comercialización"]), 0), ("5.3.05", sum(SEG26["Administración"]), 0),
        ("1.4.03", 0, sum(SEG26["Comercialización"]) + sum(SEG26["Administración"])))
COD_ACUM = {"5.2.07": "1.6.03", "5.3.06": "1.6.03", "5.2.08": "1.6.05", "5.2.09": "1.6.09", "5.3.07": "1.6.07"}
mov_am = [(k, v, 0) for k, v in sorted(AM26.items())]
acum_cr = {}
for k, v in AM26.items():
    acum_cr[COD_ACUM[k]] = acum_cr.get(COD_ACUM[k], 0) + v
mov_am += [(k, 0, v) for k, v in sorted(acum_cr.items())]
asiento("Amortizaciones 2026", *mov_am)
c01 = [b for b in BIENES26 if b["cod"] == "C01"][0]
asiento("Venta del camión C01 (IVA 21 %)", (BANCO, PRECIO_C01 * 121 // 100, 0), ("1.6.05", c01["baja_aa"], 0),
        ("1.6.04", 0, c01["baja_vo"]), ("1.4.02", 0, PRECIO_C01 * 21 // 100), ("4.3.01", 0, c01["res_venta"]))
asiento("Compra del camión C11 (IVA 10,5 %)", ("1.6.04", VO_C11, 0), ("1.4.02", VO_C11 * 105 // 1000, 0),
        (BANCO, 0, VO_C11 * 1105 // 1000))
pago_iva = IVA_FIN26 - SI26["1.4.02"] - DEB["1.4.02"] + CRE["1.4.02"]
asiento("Pagos de IVA (saldo técnico mensual)", ("1.4.02", pago_iva, 0), (BANCO, 0, pago_iva))
asiento("Constitución de plazos fijos", ("1.2.01", 400_000_000, 0), (BANCO, 0, 400_000_000))
asiento("Rescate de plazos fijos", (BANCO, 350_000_000, 0), ("1.2.01", 0, 350_000_000))
asiento("Intereses cobrados de plazos fijos", (BANCO, INT_PF26, 0), ("4.2.01", 0, INT_PF26))
asiento("Amortización de capital del préstamo", ("2.4.01", 150_000_000, 0), (BANCO, 0, 150_000_000))
asiento("Intereses pagados del préstamo", ("5.4.01", INT_PREST26, 0), (BANCO, 0, INT_PREST26))
ig25 = -SI26["2.3.01"]
ant25 = SI26["1.4.01"]
asiento("Cancelación del impuesto a las ganancias 2025 (neto de anticipos)", ("2.3.01", ig25, 0),
        ("1.4.01", 0, ant25), (BANCO, 0, ig25 - ant25))
asiento("Anticipos de impuesto a las ganancias 2026", ("1.4.01", ANTICIPOS26, 0), (BANCO, 0, ANTICIPOS26))
asiento("Dividendos en efectivo (Asamblea de abril de 2026)", ("3.3.01", DIVIDENDOS[2026], 0), (BANCO, 0, DIVIDENDOS[2026]))
asiento("Extracciones de banco para caja", ("1.1.01", 41_600_000, 0), (BANCO, 0, 41_600_000))
asiento("Depósitos de caja en banco", (BANCO, 39_600_000, 0), ("1.1.01", 0, 39_600_000))
rai26 = sum(CRE[k] - DEB[k] for k in DEB if k[0] in "45")
IG26 = round(rai26 * 0.35)
asiento("Impuesto a las ganancias 2026 (35 % del resultado contable)", ("5.5.01", IG26, 0), ("2.3.01", 0, IG26))
RAI26 = rai26
RES26 = RAI26 - IG26

SF26 = {k: SI26[k] + DEB[k] - CRE[k] for k in SI26}
assert sum(DEB.values()) == sum(CRE.values())
assert sum(SF26.values()) == 0
assert SF26[BANCO] > 0, SF26[BANCO]
assert SF26["1.3.01"] == DEUDORES_FIN26 and SF26["1.5.01"] == EXFIN26 and SF26["1.4.02"] == IVA_FIN26
TOT_RES26 = {k: DEB[k] if k[0] == "5" else CRE[k] for k in DEB if k[0] in "45"}
RM26 = resultados_mensuales(2026)
RESULTADOS[2026] = RES26

# Control de bienes de uso y seguros contra el BSS 2026
def suma(bienes, cuenta, campo):
    return sum(b[campo] for b in bienes if b["cuenta"] == cuenta)


for cta, acum in CTA_AMORT.items():
    cod = [c["cod"] for c in CUENTAS26 if c["cta"] == cta][0]
    coda = [c["cod"] for c in CUENTAS26 if c["cta"] == acum][0]
    assert suma(BIENES26, cta, "vo26") == SF26[cod]
    assert -suma(BIENES26, cta, "aa26") == SF26[coda]
assert suma(BIENES26, "Terrenos", "vo26") == SF26["1.6.01"]
assert sum(p["nd_fin"] for p in POL26) == SF26["1.4.03"]

# --------------------------------------------------------------------------------------
# Escritura de libros Excel
# --------------------------------------------------------------------------------------
NUM = "#,##0;-#,##0;0"
FEC = "dd/mm/yyyy"
MES = "mm/yyyy"
NEG = Font(bold=True)
CAB = PatternFill("solid", fgColor="DDEBF7")
TOT = PatternFill("solid", fgColor="F2F2F2")


class Hoja:
    def __init__(self, wb, nombre, titulo, subtitulo, primera=False):
        self.ws = wb.active if primera else wb.create_sheet(nombre)
        self.ws.title = nombre
        self.ws["A1"] = titulo
        self.ws["A1"].font = Font(bold=True, size=12)
        self.ws["A2"] = subtitulo
        self.ws["A2"].alignment = Alignment(wrap_text=False)
        self.r = 4

    def fila(self, valores, estilo=None, formatos=None):
        for j, v in enumerate(valores, start=1):
            c = self.ws.cell(row=self.r, column=j, value=v)
            if isinstance(v, (dt.date, dt.datetime)):
                c.number_format = (formatos or {}).get(j, FEC)
            elif isinstance(v, (int, float)) or (isinstance(v, str) and v.startswith("=")):
                c.number_format = (formatos or {}).get(j, NUM)
            if estilo == "cab":
                c.font = NEG
                c.fill = CAB
                c.alignment = Alignment(wrap_text=True, vertical="center")
            elif estilo == "tot":
                c.font = NEG
                c.fill = TOT
            elif estilo == "sec":
                c.font = Font(bold=True, italic=True)
        self.r += 1
        return self.r - 1

    def salto(self, n=1):
        self.r += n

    def anchos(self, anchos):
        for j, a in enumerate(anchos, start=1):
            self.ws.column_dimensions[L(j)].width = a


def hoja_notas(wb, titulo, lineas):
    h = Hoja(wb, "Notas y supuestos", titulo, "Supuestos adoptados (S), datos del encargo (D) y Dudas.")
    for t in lineas:
        h.fila([t])
    h.anchos([160])


# ---------- Mercaderías ----------
def hoja_mercaderias(wb, anio, capas_ini, compras, cmv, consumo, capas_fin, ref):
    h = Hoja(wb, "Mercaderías", f"{EMPRESA} — Auxiliar de mercaderías de reventa (insumos) {anio}",
             f"{LEYENDA} Costo FIFO. Fecha de origen = mes de compra. Rotación aproximada: 1,3 meses.")
    h.fila(["A. Movimiento mensual"], "sec")
    h.fila(["Mes", "Existencia inicial", "Compras", "CMV", "Existencia final"], "cab")
    r0 = h.r
    for m in range(12):
        ei = None if m == 0 else f"=E{h.r - 1}"
        h.fila([D(anio, m + 1, 1), ei, compras[m], cmv[m], f"=B{h.r}+C{h.r}-D{h.r}"], formatos={1: MES})
    rt = h.fila(["TOTAL", f"=B{r0}", f"=SUM(C{r0}:C{h.r - 1})", f"=SUM(D{r0}:D{h.r - 1})", f"=E{h.r - 1}"], "tot")
    ref["ei"], ref["compras"], ref["cmv"], ref["ef"] = f"B{rt}", f"C{rt}", f"D{rt}", f"E{rt}"
    h.salto()
    h.fila([f"B. Existencia inicial al 31/12/{anio - 1} por mes de compra"], "sec")
    h.fila(["Mes de compra (fecha de origen)", "Importe"], "cab")
    a = h.r
    for o, i in capas_ini:
        h.fila([o, i], formatos={1: MES})
    rt = h.fila(["TOTAL", f"=SUM(B{a}:B{h.r - 1})"], "tot")
    ref["ei_capas"] = f"B{rt}"
    h.ws.cell(row=r0, column=2, value=f"=B{rt}").number_format = NUM
    h.salto()
    h.fila([f"C. Existencia final al 31/12/{anio} por mes de compra"], "sec")
    h.fila(["Mes de compra (fecha de origen)", "Importe"], "cab")
    a = h.r
    for o, i in capas_fin:
        h.fila([o, i], formatos={1: MES})
    rt = h.fila(["TOTAL", f"=SUM(B{a}:B{h.r - 1})"], "tot")
    ref["ef_capas"] = f"B{rt}"
    h.salto()
    h.fila([f"D. CMV {anio} abierto por mes de compra (FIFO): filas = mes de compra; columnas = mes de venta"], "sec")
    h.fila(["Mes de compra"] + [f"{MESES[m]}-{anio}" for m in range(12)] + ["Total"], "cab")
    a = h.r
    origenes = sorted({o for o, _ in consumo})
    for o in origenes:
        h.fila([o] + [consumo.get((o, m), 0) for m in range(12)] + [f"=SUM(B{h.r}:M{h.r})"], formatos={1: MES})
    rt = h.fila(["TOTAL"] + [f"=SUM({L(j)}{a}:{L(j)}{h.r - 1})" for j in range(2, 15)], "tot")
    ref["cmv_matriz"] = f"N{rt}"
    h.fila(["Control vs. sección A"] + [f"={L(j)}{rt}-D{r0 + j - 2}" for j in range(2, 14)] + [f"=N{rt}-{ref['cmv']}"])
    h.anchos([30] + [16] * 13)
    h.ws.freeze_panes = "B4"


# ---------- Patrimonio neto ----------
def hoja_pn(wb, anio, ref):
    h = Hoja(wb, "Patrimonio neto", f"{EMPRESA} — Auxiliar de patrimonio neto al 31/12/{anio}",
             f"{LEYENDA} Fecha de origen de cada componente, base para el ajuste por inflación.")
    h.fila(["A. Capital"], "sec")
    h.fila(["Concepto", "Fecha de origen", "Importe"], "cab")
    rc = h.fila(["Capital suscripto e integrado", FECHA_CAPITAL, CAPITAL])
    ref["capital"] = f"C{rc}"
    h.salto()
    h.fila(["B. Reserva legal (art. 70 LGS: 5 % del resultado hasta el 20 % del capital)"], "sec")
    h.fila(["Ejercicio del resultado", "Fecha de constitución (asamblea)", "Resultado del ejercicio", "Constituido", "Acumulado"], "cab")
    a = h.r
    for y, v in sorted(RESERVA.items()):
        h.fila([y, fecha_asamblea(y + 1), RESULTADOS[y], v, f"=SUM(D{a}:D{h.r})"], formatos={1: "0"})
    rt = h.fila(["TOTAL", None, None, f"=SUM(D{a}:D{h.r - 1})", None], "tot")
    ref["reserva"] = f"D{rt}"
    h.fila([f"Nota: no se constituye reserva sobre resultados posteriores a 2021 porque el saldo nominal ya alcanzó el 20 % del capital. "
            "Duda: base de la reserva después del ajuste (la define la fase 2)."])
    h.salto()
    h.fila(["C. Resultados no asignados por año de origen"], "sec")
    if anio == 2025:
        cab = ["Ejercicio de origen", "Fecha de origen (cierre)", "Resultado del ejercicio", "Reserva legal asignada",
               "Dividendos asignados hasta 31/12/2024", "Fecha de pago de esos dividendos", "Saldo RNA 31/12/2024",
               "Dividendos pagados en 2025", "Fecha de pago", "Resultado 2025", "Saldo RNA 31/12/2025"]
        anios = range(2014, 2026)
    else:
        cab = ["Ejercicio de origen", "Fecha de origen (cierre)", "Resultado del ejercicio", "Reserva legal asignada",
               "Dividendos asignados hasta 31/12/2025", "Fecha de pago de esos dividendos", "Saldo RNA 31/12/2025",
               "Dividendos pagados en 2026", "Fecha de pago", "Resultado 2026", "Saldo RNA 31/12/2026"]
        anios = range(2014, 2027)
    h.fila(cab, "cab")
    a = h.r
    for y in anios:
        res = RESULTADOS[y]
        es_actual = y == anio
        div_prev = DIVIDENDOS.get(y + 1, 0) if y + 1 < anio else 0
        f_prev = fecha_pago_div(y + 1) if div_prev else None
        div_act = DIVIDENDOS.get(y + 1, 0) if y + 1 == anio else 0
        f_act = fecha_pago_div(anio) if div_act else None
        r = h.r
        h.fila([y, D(y, 12, 31), 0 if es_actual else res, RESERVA.get(y, 0), div_prev, f_prev,
                f"=C{r}-D{r}-E{r}", div_act, f_act, res if es_actual else 0, f"=G{r}-H{r}+J{r}"], formatos={1: "0"})
    rt = h.fila(["TOTAL", None] + [f"=SUM({L(j)}{a}:{L(j)}{h.r - 1})" for j in range(3, 6)] + [None] +
                [f"=SUM({L(j)}{a}:{L(j)}{h.r - 1})" for j in (7, 8)] + [None] +
                [f"=SUM({L(j)}{a}:{L(j)}{h.r - 1})" for j in (10, 11)], "tot")
    ref["rna_ini"], ref["div"], ref["res_ej"], ref["rna_fin"] = f"G{rt}", f"H{rt}", f"J{rt}", f"K{rt}"
    h.fila(["Supuesto: cada dividendo se imputa al resultado del ejercicio inmediato anterior a la asamblea. Los resultados 2014-2023 "
            f"son supuestos; el de 2024 es el residuo que cuadra los RNA al 31/12/2024 con el BSS."])
    h.salto()
    h.fila(["D. Historial de dividendos en efectivo"], "sec")
    h.fila(["Fecha de asamblea", "Fecha de pago", "Ejercicio al que se imputan", "Importe"], "cab")
    a = h.r
    for y, v in sorted(DIVIDENDOS.items()):
        if y <= anio:
            h.fila([fecha_asamblea(y), fecha_pago_div(y), y - 1, v], formatos={3: "0"})
    h.fila(["TOTAL", None, None, f"=SUM(D{a}:D{h.r - 1})"], "tot")
    h.anchos([24, 18, 18, 16, 20, 18, 18, 18, 14, 16, 18])


# ---------- Resultados mensuales ----------
def hoja_resultados(wb, anio, rm, compras, totales_bss, cuentas, ref):
    h = Hoja(wb, "Resultados mensuales", f"{EMPRESA} — Resultados mensuales {anio} (base del ajuste de resultados)",
             f"{LEYENDA} Importes en valor absoluto; el signo indica ingreso (+1) o egreso (−1).")
    h.fila(["Código", "Cuenta", "Rubro", "Signo"] + [f"{MESES[m]}-{anio}" for m in range(12)] + ["Total", "BSS", "Diferencia"], "cab")
    a = h.r
    for c in cuentas:
        k = c["cod"]
        if k[0] not in "45":
            continue
        s = 1 if k[0] == "4" else -1
        r = h.r
        h.fila([k, c["cta"], c["rubro"], s] + rm[k] + [f"=SUM(E{r}:P{r})", totales_bss[k], f"=Q{r}-R{r}"])
        ref.setdefault("res", {})[k] = f"Q{r}"
    b = h.r - 1
    rr = h.fila(["", "Resultado del mes", "", ""] + [f"=SUMPRODUCT($D${a}:$D${b},{L(j)}{a}:{L(j)}{b})" for j in range(5, 18)] +
                [sum(totales_bss[k] * (1 if k[0] == "4" else -1) for k in totales_bss), f"=Q{h.r}-R{h.r}"], "tot")
    ref["resultado"] = f"Q{rr}"
    h.salto()
    r = h.r
    h.fila(["1.5.01", "Compras de mercaderías (débitos)", "Bienes de cambio", ""] + compras + [f"=SUM(E{r}:P{r})", sum(compras), f"=Q{r}-R{r}"])
    ref["compras_rm"] = f"Q{r}"
    h.salto()
    h.fila(["Supuestos de distribución mensual: ventas y gastos con crecimiento nominal de ~2 % mensual; CMV = 70 % de las ventas (FIFO); "
            "IIBB = 3 % de las ventas; seguros según devengamiento de cada póliza; amortizaciones en doceavos; intereses del préstamo "
            "proporcionales al saldo de capital (amortización trimestral de 37,5 M); impuesto a las ganancias imputado en diciembre."])
    h.anchos([9, 44, 24, 6] + [15] * 15)
    h.ws.freeze_panes = "E5"


# ---------- Conciliación ----------
def hoja_conciliacion_simple(wb, filas, titulo):
    h = Hoja(wb, "Conciliación BSS", titulo, "Diferencia esperada: 0 en todas las filas. Auxiliar = fórmula sobre las hojas de este libro; BSS = saldo del balance de sumas y saldos.")
    h.fila(["Concepto", "Cuenta BSS", "Auxiliar", "BSS", "Diferencia"], "cab")
    for concepto, cod, f_aux, v_bss in filas:
        h.fila([concepto, cod, f_aux, v_bss, f"=C{h.r}-D{h.r}"])
    h.anchos([60, 12, 18, 18, 14])


# ======================================================================================
# 1) auxiliares_complementarios_2025.xlsx
# ======================================================================================

def notas_comunes(anio):
    return [
        f"(D) {EMPRESA}: distribuidora de insumos. Valores históricos nominales, sin reexpresar. Datos ficticios para práctica.",
        "(D) Mercaderías al costo FIFO. El costo de cada capa conserva su mes de compra como fecha de origen.",
        "(S) Existencia inicial al 31/12/2024 abierta en nov-2024 $160.000.000 y dic-2024 $740.000.000 (rotación ~1,3 meses).",
        "(S) Ventas con crecimiento nominal de ~2 % mensual; margen bruto constante del 30 % (CMV = 70 % de las ventas).",
        "(S) Compras mensuales = CMV del mes + incremento lineal de la existencia (de 900 M a 1.250 M en 2025; de 1.250 M a 1.625 M en 2026).",
        f"(D) Capital $20.000.000 integrado en 03/2014. (S) Día de integración: {FECHA_CAPITAL:%d/%m/%Y}.",
        "(S) Resultados 2014-2023 y dividendos 2019-2024 son supuestos verosímiles; el resultado 2024 es el residuo que cuadra RNA 31/12/2024 = $1.129.753.500.",
        "(S) Asambleas: último día hábil hasta el 28 de abril; pago de dividendos: último día hábil hasta el 30 de abril.",
        "(S) Cada dividendo se imputa al resultado del ejercicio inmediato anterior a la asamblea.",
        "(D) Reserva legal $4.000.000 = 20 % del capital (art. 70 LGS), constituida con el 5 % de los resultados 2014-2021.",
        "Duda: base de la reserva legal (art. 70 LGS) después del ajuste por inflación. La define la fase 2.",
        "(S) Intereses del préstamo distribuidos según saldo de capital, con amortización trimestral de $37.500.000 (mar, jun, sep, dic).",
        "(S) Impuesto a las ganancias imputado en diciembre (determinado al cierre).",
    ]


def gen_aux_comp_2025():
    wb = openpyxl.Workbook()
    hoja_notas_wb = wb.active
    wb.remove(hoja_notas_wb)
    rM, rP, rR = {}, {}, {}
    hoja_mercaderias(wb, 2025, CAPAS_INI25, COMPRAS25, CMV25, CONS25, CAPAS_FIN25, rM)
    hoja_pn(wb, 2025, rP)
    tot25 = {k: (B25[k]["deb"] if k[0] == "5" else B25[k]["cred"]) for k in B25 if k[0] in "45"}
    hoja_resultados(wb, 2025, RM25, COMPRAS25, tot25, BSS25, rR)
    M, P, R = "'Mercaderías'!", "'Patrimonio neto'!", "'Resultados mensuales'!"
    filas = [
        ("Mercaderías: existencia inicial 31/12/2024 (por mes de compra)", "1.5.01", f"={M}{rM['ei_capas']}", B25["1.5.01"]["si"]),
        ("Mercaderías: compras 2025 (débitos)", "1.5.01", f"={M}{rM['compras']}", B25["1.5.01"]["deb"]),
        ("Mercaderías: CMV 2025 (créditos)", "1.5.01", f"={M}{rM['cmv']}", B25["1.5.01"]["cred"]),
        ("Mercaderías: CMV 2025 abierto por mes de compra", "5.1.01", f"={M}{rM['cmv_matriz']}", B25["5.1.01"]["deb"]),
        ("Mercaderías: existencia final 31/12/2025 (movimiento)", "1.5.01", f"={M}{rM['ef']}", B25["1.5.01"]["sf"]),
        ("Mercaderías: existencia final 31/12/2025 (por mes de compra)", "1.5.01", f"={M}{rM['ef_capas']}", B25["1.5.01"]["sf"]),
        ("Capital suscripto", "3.1.01", f"={P}{rP['capital']}", -B25["3.1.01"]["sf"]),
        ("Reserva legal", "3.2.01", f"={P}{rP['reserva']}", -B25["3.2.01"]["sf"]),
        ("Resultados no asignados 31/12/2024", "3.3.01", f"={P}{rP['rna_ini']}", -B25["3.3.01"]["si"]),
        ("Dividendos 2025 (débitos RNA)", "3.3.01", f"={P}{rP['div']}", B25["3.3.01"]["deb"]),
        ("Resultados no asignados 31/12/2025 (sin resultado 2025)", "3.3.01", f"={P}{rP['rna_fin']}-{P}{rP['res_ej']}", -B25["3.3.01"]["sf"]),
        ("Resultado 2025 (auxiliar PN vs. BSS)", "4 y 5", f"={P}{rP['res_ej']}", RES25),
        ("Resultado 2025 (resultados mensuales vs. BSS)", "4 y 5", f"={R}{rR['resultado']}", RES25),
        ("Compras 2025 (resultados mensuales vs. BSS)", "1.5.01", f"={R}{rR['compras_rm']}", B25["1.5.01"]["deb"]),
    ]
    for c in BSS25:
        if c["cod"][0] in "45":
            filas.append((f"Resultados mensuales: {c['cta']}", c["cod"], f"={R}{rR['res'][c['cod']]}", tot25[c["cod"]]))
    hoja_conciliacion_simple(wb, filas, f"{EMPRESA} — Conciliación auxiliares complementarios 2025 vs. BSS 2025")
    hoja_notas(wb, f"{EMPRESA} — Auxiliares complementarios 2025: notas", notas_comunes(2025))
    wb.save(os.path.join(INS, "auxiliares_complementarios_2025.xlsx"))


# ======================================================================================
# 2) bss_2026.xlsx / .csv
# ======================================================================================
CAB_BSS26 = ["Código", "Cuenta", "Rubro", "Saldo inicial 31/12/2025", "Débitos 2026", "Créditos 2026",
             "Saldo final 31/12/2026", "Saldo deudor", "Saldo acreedor"]


def filas_bss26():
    out = []
    for c in CUENTAS26:
        k = c["cod"]
        sf = SF26[k]
        out.append([k, c["cta"], c["rubro"], SI26[k], DEB[k], CRE[k], sf, max(sf, 0), max(-sf, 0)])
    return out


def gen_bss26():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BSS"
    ws["A1"] = f"{EMPRESA} — Balance de sumas y saldos al 31/12/2026"
    ws["A1"].font = Font(bold=True, size=12)
    ws["A2"] = ("Valores históricos nominales (sin reexpresar por inflación). Importes en pesos. Saldos con signo: deudor (+), acreedor (−). "
                "Cuentas de resultado abiertas (pre-cierre). Saldo inicial = saldo final 2025 después del cierre de resultados.")
    for j, v in enumerate(CAB_BSS26, start=1):
        c = ws.cell(row=3, column=j, value=v)
        c.font = NEG
        c.fill = CAB
    filas = filas_bss26()
    for i, f in enumerate(filas, start=4):
        for j, v in enumerate(f, start=1):
            c = ws.cell(row=i, column=j, value=v)
            if j >= 4:
                c.number_format = NUM
    n = 3 + len(filas)
    t = n + 1
    ws.cell(row=t, column=1, value="TOTAL").font = NEG
    for j in range(4, 10):
        c = ws.cell(row=t, column=j, value=f"=SUM({L(j)}4:{L(j)}{n})")
        c.number_format = NUM
        c.font = NEG
    ws.cell(row=t + 2, column=1, value="Control: débitos − créditos")
    ws.cell(row=t + 2, column=5, value=f"=E{t}-F{t}").number_format = NUM
    ws.cell(row=t + 3, column=1, value="Control: saldo final recalculado − saldo final")
    ws.cell(row=t + 3, column=7, value=f"=SUMPRODUCT(D4:D{n}+E4:E{n}-F4:F{n}-G4:G{n})").number_format = NUM
    for j, a in enumerate([9, 52, 30, 18, 18, 18, 18, 18, 18], start=1):
        ws.column_dimensions[L(j)].width = a
    ws.freeze_panes = "D4"
    wb.save(os.path.join(INS, "bss_2026.xlsx"))
    with open(os.path.join(INS, "bss_2026.csv"), "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(CAB_BSS26)
        w.writerows(filas)


# ======================================================================================
# 3) auxiliares_2026.xlsx
# ======================================================================================

def gen_aux26():
    wb = openpyxl.Workbook()
    # Bienes de uso
    h = Hoja(wb, "Bienes de uso", f"{EMPRESA} — Auxiliar de bienes de uso (inmuebles y rodados) al 31/12/2026",
             f"{LEYENDA} Amortización lineal, año de alta completo, año de baja sin amortización, sin valor residual. Terrenos no se amortizan.",
             primera=True)
    cab = ["Código", "Bien / tipo", "Componente / descripción", "Ubicación / dominio", "Fecha de origen", "Valor de origen",
           "Vida útil (años)", "Valor de origen al 31/12/2025", "Amort. acumulada 31/12/2025", "Altas 2026", "Fecha de baja",
           "Bajas 2026 - valor de origen", "Bajas 2026 - amort. acumulada", "Amortización 2026", "Valor de origen al 31/12/2026",
           "Amort. acumulada 31/12/2026", "Valor residual contable 31/12/2026", "Precio de venta (sin IVA)", "Resultado por venta",
           "Cuenta BSS", "Función del cargo"]
    h.fila(cab, "cab")
    a = h.r
    for b in BIENES26:
        r = h.r
        h.fila([b["cod"], b["tipo"], b["comp"], b["ubic"], b["fecha"], b["vo"], b["vida"], b["vo25"], b["aa25"], b["alta26"],
                b["baja"], b["baja_vo"], b["baja_aa"], b["am26"], f"=H{r}+J{r}-L{r}", f"=I{r}-M{r}+N{r}", f"=O{r}-P{r}",
                b["precio"], f"=IF(K{r}=\"\",0,R{r}-(L{r}-M{r}))", b["cuenta"], b["func"]], formatos={7: "0"})
    zb = h.r - 1
    h.fila(["TOTAL", None, None, None, None, f"=SUM(F{a}:F{zb})", None] +
           [f"=SUM({L(j)}{a}:{L(j)}{zb})" for j in (8, 9, 10)] + [None] +
           [f"=SUM({L(j)}{a}:{L(j)}{zb})" for j in range(12, 20)] + [None, None], "tot")
    h.salto()
    h.fila(["Existencias físicas al 31/12/2026 (sin bajas)"], "sec")
    for tipo in ("Camión", "Vehículo ejecutivo", "Camioneta de reparto"):
        h.fila([tipo, f"=COUNTIFS(B{a}:B{zb},\"{tipo}\",K{a}:K{zb},\"\")"], formatos={2: "0"})
    h.fila(["Inmuebles", f"=COUNTIFS(C{a}:C{zb},\"Terreno\")"], formatos={2: "0"})
    h.anchos([10, 24, 40, 30, 12, 16, 8, 16, 16, 16, 12, 16, 16, 16, 16, 16, 16, 16, 16, 34, 18])
    h.ws.freeze_panes = "B4"
    BU = ("'Bienes de uso'!", a, zb)

    # Seguros
    h = Hoja(wb, "Seguros", f"{EMPRESA} — Auxiliar de seguros (pólizas vigentes) al 31/12/2026",
             f"{LEYENDA} Pólizas anuales renovadas con igual vigencia, prima +35 % nominal, pagada al inicio de la vigencia, devengamiento mensual lineal.")
    cab = ["Póliza", "Aseguradora", "Ramo", "Cobertura", "Bienes asegurados", "Vigencia desde", "Vigencia hasta",
           "Fecha de pago (origen)", "Prima póliza 2026", "Prima póliza anterior (2025)", "No devengado al 31/12/2025",
           "Devengado 2026 de póliza anterior", "Meses devengados 2026 póliza actual", "Devengado 2026 póliza actual",
           "Gasto total 2026", "No devengado al 31/12/2026", "Función del cargo"]
    h.fila(cab, "cab")
    sa = h.r
    for p in POL26:
        r = h.r
        h.fila([p["pol"], p["aseg"], p["ramo"], p["cob"], p["bienes"], p["desde"], p["hasta"], p["pago"], p["prima"],
                p["prima_ant"], p["nd_ini"], p["dev_ant"], p["meses"], f"=I{r}/12*M{r}", f"=L{r}+N{r}", f"=I{r}-N{r}", p["func"]],
               formatos={13: "0"})
    sz = h.r - 1
    h.fila(["TOTAL", None, None, None, None, None, None, None] + [f"=SUM({L(j)}{sa}:{L(j)}{sz})" for j in (9, 10, 11, 12)] +
           [None] + [f"=SUM({L(j)}{sa}:{L(j)}{sz})" for j in (14, 15, 16)] + [None], "tot")
    h.anchos([9, 24, 20, 30, 36, 12, 12, 12, 15, 15, 15, 15, 9, 15, 15, 15, 18])
    h.ws.freeze_panes = "B4"

    rM, rP, rR = {}, {}, {}
    hoja_mercaderias(wb, 2026, CAPAS_FIN25, COMPRAS26, CMV26, CONS26, CAPAS_FIN26, rM)
    hoja_pn(wb, 2026, rP)
    hoja_resultados(wb, 2026, RM26, COMPRAS26, TOT_RES26, CUENTAS26, rR)

    # Conciliación
    h = Hoja(wb, "Conciliación BSS", f"{EMPRESA} — Conciliación auxiliares 2026 vs. BSS 2026",
             "Diferencia esperada: 0 en todas las filas. Auxiliar = fórmula sobre las hojas de este libro; BSS = saldos de bss_2026.xlsx.")
    h.fila(["Cuenta", "Auxiliar 31/12/2025", "BSS 31/12/2025", "Diferencia", "Auxiliar 31/12/2026", "BSS 31/12/2026", "Diferencia"], "cab")
    S, A_, Z = BU

    def sumif(col, cta, signo=""):
        return f"={signo}SUMIF({S}$T${A_}:$T${Z},\"{cta}\",{S}{col}${A_}:{col}${Z})"

    cod_de = {c["cta"]: c["cod"] for c in CUENTAS26}
    for cta in ["Terrenos", "Edificios", "Rodados - camiones", "Rodados - vehículos ejecutivos", "Rodados - camionetas de reparto"]:
        h.fila([cta, sumif("H", cta), SI26[cod_de[cta]], f"=B{h.r}-C{h.r}", sumif("O", cta), SF26[cod_de[cta]], f"=E{h.r}-F{h.r}"])
        if cta in CTA_AMORT:
            ac = CTA_AMORT[cta]
            h.fila([ac, sumif("I", cta, "-"), SI26[cod_de[ac]], f"=B{h.r}-C{h.r}", sumif("P", cta, "-"), SF26[cod_de[ac]], f"=E{h.r}-F{h.r}"])
    SG = "'Seguros'!"
    h.fila(["Seguros pagados por adelantado", f"={SG}K{sz + 1}", SI26["1.4.03"], f"=B{h.r}-C{h.r}", f"={SG}P{sz + 1}", SF26["1.4.03"], f"=E{h.r}-F{h.r}"])
    M, P, R = "'Mercaderías'!", "'Patrimonio neto'!", "'Resultados mensuales'!"
    h.fila(["Mercaderías de reventa (insumos)", f"={M}{rM['ei_capas']}", SI26["1.5.01"], f"=B{h.r}-C{h.r}", f"={M}{rM['ef_capas']}", SF26["1.5.01"], f"=E{h.r}-F{h.r}"])
    h.fila(["Capital suscripto", f"=-{P}{rP['capital']}", SI26["3.1.01"], f"=B{h.r}-C{h.r}", f"=-{P}{rP['capital']}", SF26["3.1.01"], f"=E{h.r}-F{h.r}"])
    h.fila(["Reserva legal", f"=-{P}{rP['reserva']}", SI26["3.2.01"], f"=B{h.r}-C{h.r}", f"=-{P}{rP['reserva']}", SF26["3.2.01"], f"=E{h.r}-F{h.r}"])
    h.fila(["Resultados no asignados (31/12/2026 sin resultado 2026)", f"=-{P}{rP['rna_ini']}", SI26["3.3.01"], f"=B{h.r}-C{h.r}",
            f"=-({P}{rP['rna_fin']}-{P}{rP['res_ej']})", SF26["3.3.01"], f"=E{h.r}-F{h.r}"])
    h.salto()
    h.fila(["Movimientos y resultados del ejercicio 2026", None, None, None, "Auxiliar 2026", "BSS 2026", "Diferencia"], "cab")
    BUf = f"{S}N{Z + 1}"
    filas = [
        ("Amortizaciones del ejercicio (resultado)", f"={BUf}", sum(AM26.values())),
        ("Gasto de seguros del ejercicio (resultado)", f"={SG}O{sz + 1}", TOT_RES26["5.2.06"] + TOT_RES26["5.3.05"]),
        ("Altas de bienes de uso (débitos 1.6.04)", f"={S}J{Z + 1}", DEB["1.6.04"]),
        ("Bajas de bienes de uso - valor de origen (créditos 1.6.04)", f"={S}L{Z + 1}", CRE["1.6.04"]),
        ("Bajas de bienes de uso - amort. acumulada (débitos 1.6.05)", f"={S}M{Z + 1}", DEB["1.6.05"]),
        ("Resultado por venta de bienes de uso", f"={S}S{Z + 1}", TOT_RES26["4.3.01"]),
        ("Primas pagadas 2026 (débitos 1.4.03)", f"={SG}I{sz + 1}", DEB["1.4.03"]),
        ("Compras de mercaderías (débitos 1.5.01)", f"={M}{rM['compras']}", DEB["1.5.01"]),
        ("CMV (créditos 1.5.01)", f"={M}{rM['cmv']}", CRE["1.5.01"]),
        ("CMV abierto por mes de compra (5.1.01)", f"={M}{rM['cmv_matriz']}", TOT_RES26["5.1.01"]),
        ("Dividendos 2026 (débitos 3.3.01)", f"={P}{rP['div']}", DEB["3.3.01"]),
        ("Resultado 2026 (auxiliar PN)", f"={P}{rP['res_ej']}", RES26),
        ("Resultado 2026 (resultados mensuales)", f"={R}{rR['resultado']}", RES26),
        ("Compras 2026 (resultados mensuales)", f"={R}{rR['compras_rm']}", DEB["1.5.01"]),
    ]
    for c in CUENTAS26:
        if c["cod"][0] in "45":
            filas.append((f"Resultados mensuales: {c['cta']}", f"={R}{rR['res'][c['cod']]}", TOT_RES26[c["cod"]]))
    for concepto, fa, vb in filas:
        h.fila([concepto, None, None, None, fa, vb, f"=E{h.r}-F{h.r}"])
    h.anchos([60, 18, 18, 12, 18, 18, 12])

    notas = notas_comunes(2026) + [
        f"(D) Asamblea del {fecha_asamblea(2026):%d/%m/%Y}: dividendos en efectivo ${DIVIDENDOS[2026]:,.0f} sobre el resultado 2025, pagados el {fecha_pago_div(2026):%d/%m/%Y}. Reserva legal sin cambios.".replace(",", "."),
        f"(S) Baja por venta del camión C01 el {FECHA_BAJA_C01:%d/%m/%Y}: precio ${PRECIO_C01:,.0f} + IVA 21 %. Año de baja sin amortización.".replace(",", "."),
        f"(S) Alta del camión C11 el {FECHA_C11:%d/%m/%Y}: ${VO_C11:,.0f} + IVA 10,5 %. Vida útil 10 años, año de alta completo.".replace(",", "."),
        "(S) Renovación de las 10 pólizas con igual vigencia y prima +35 % nominal.",
        "(S) Ventas 2026: $18.200.000.000 (+30 %). Gastos operativos +30 %. Intereses de plazo fijo $72.000.000; intereses del préstamo $105.000.000.",
        "(S) Préstamo: amortización de capital $150.000.000 en 2026 (37,5 M por trimestre). Saldo 31/12/2026: $300.000.000.",
        "(S) Impuesto a las ganancias 35 % sobre el resultado contable, sin impuesto diferido. Anticipos 2026: $350.000.000.",
    ]
    hoja_notas(wb, f"{EMPRESA} — Auxiliares 2026: notas", notas)
    wb.save(os.path.join(INS, "auxiliares_2026.xlsx"))


# ======================================================================================
# 4) EECC históricos
# ======================================================================================

def saldos_esp():
    """Saldos (con signo BSS) por cuenta para 31/12/2024, 31/12/2025 y 31/12/2026."""
    s24 = {c["cod"]: c["si"] for c in BSS25}
    s25 = {k: SI26[k] for k in SI26}
    s26 = {k: SF26[k] for k in SF26}
    return s24, s25, s26


def gen_eecc():
    wb = openpyxl.Workbook()
    s24, s25, s26 = saldos_esp()
    h = Hoja(wb, "ESP", f"{EMPRESA} — Estado de situación patrimonial histórico (estado de trabajo)",
             f"{LEYENDA} No es una presentación formal. Préstamo: porción corriente = amortización de capital de los 12 meses siguientes (supuesto: $150.000.000 por año).",
             primera=True)
    h.fila(["Rubro", "31/12/2024", "31/12/2025", "31/12/2026"], "cab")
    rub = {}

    def rubro(nombre, codigos, signo=1):
        r = h.r
        h.fila([nombre] + [signo * sum(s[k] for k in codigos if k in s) for s in (s24, s25, s26)])
        rub[nombre] = r
        return r

    def total(nombre, filas_):
        r = h.fila([nombre] + [("=" + "+".join(f"{c}{x}" for x in filas_)) for c in "BCD"], "tot")
        return r

    h.fila(["ACTIVO"], "sec")
    h.fila(["Activo corriente"], "sec")
    ac = [rubro("Caja y bancos", ["1.1.01", "1.1.02"]), rubro("Inversiones", ["1.2.01"]),
          rubro("Créditos por ventas", ["1.3.01", "1.3.02"]), rubro("Otros créditos", ["1.4.01", "1.4.02", "1.4.03"]),
          rubro("Bienes de cambio", ["1.5.01"])]
    tac = total("Total activo corriente", ac)
    h.fila(["Activo no corriente"], "sec")
    anc = [rubro("Bienes de uso", [f"1.6.0{i}" for i in range(1, 10)])]
    tanc = total("Total activo no corriente", anc)
    ta = total("TOTAL ACTIVO", [tac, tanc])
    h.salto()
    h.fila(["PASIVO"], "sec")
    h.fila(["Pasivo corriente"], "sec")
    pc = [rubro("Deudas comerciales", ["2.1.01"], -1), rubro("Remuneraciones y cargas sociales", ["2.2.01"], -1),
          rubro("Cargas fiscales", ["2.3.01", "2.3.02"], -1)]
    r = h.fila(["Préstamos (corriente)", 150_000_000, 150_000_000, 150_000_000])
    pc.append(r)
    tpc = total("Total pasivo corriente", pc)
    h.fila(["Pasivo no corriente"], "sec")
    r = h.fila(["Préstamos (no corriente)", -s24["2.4.01"] - 150_000_000, -s25["2.4.01"] - 150_000_000, -s26["2.4.01"] - 150_000_000])
    tpnc = total("Total pasivo no corriente", [r])
    tp = total("TOTAL PASIVO", [tpc, tpnc])
    h.salto()
    h.fila(["PATRIMONIO NETO"], "sec")
    pn = [rubro("Capital suscripto", ["3.1.01"], -1), rubro("Reserva legal", ["3.2.01"], -1)]
    res = {2024: RESULTADOS[2024], 2025: RES25, 2026: RES26}
    pn.append(h.fila(["Resultados no asignados de ejercicios anteriores", -s24["3.3.01"] - res[2024], -s25["3.3.01"] - res[2025] + 0, -s26["3.3.01"]]))
    pn.append(h.fila(["Resultado del ejercicio", res[2024], res[2025], res[2026]]))
    tpn = total("TOTAL PATRIMONIO NETO", pn)
    tppn = total("TOTAL PASIVO + PATRIMONIO NETO", [tp, tpn])
    h.fila(["Control: activo − (pasivo + PN)"] + [f"={c}{ta}-{c}{tppn}" for c in "BCD"])
    h.anchos([52, 18, 18, 18])

    # Nota: s25 son saldos iniciales 2026 (RNA ya incluye el resultado 2025). s26 son saldos pre-cierre (RNA sin resultado 2026).
    # ER
    h = Hoja(wb, "ER", f"{EMPRESA} — Estado de resultados histórico (estado de trabajo)", f"{LEYENDA} Ganancia (+) / pérdida (−).")
    h.fila(["Concepto", "2025", "2026"], "cab")
    r25 = {k: -B25[k]["sf"] for k in B25 if k[0] in "45"}
    r26 = {k: -SF26[k] for k in SF26 if k[0] in "45"}

    def linea(nombre, codigos):
        return h.fila([nombre] + [sum(r.get(k, 0) for k in codigos) for r in (r25, r26)])

    v = linea("Ventas netas de insumos", ["4.1.01"])
    c = linea("Costo de mercaderías vendidas", ["5.1.01"])
    rb = h.fila(["Resultado bruto", f"=B{v}+B{c}", f"=C{v}+C{c}"], "tot")
    gc = linea("Gastos de comercialización", [k for k in r26 if k.startswith("5.2")])
    ga = linea("Gastos de administración", [k for k in r26 if k.startswith("5.3")])
    oi = linea("Otros ingresos y egresos (resultado por venta de bienes de uso)", ["4.3.01"])
    rf1 = linea("Resultados financieros: intereses ganados", ["4.2.01"])
    rf2 = linea("Resultados financieros: intereses del préstamo", ["5.4.01"])
    rai = h.fila(["Resultado antes del impuesto a las ganancias"] + [f"={x}{rb}+{x}{gc}+{x}{ga}+{x}{oi}+{x}{rf1}+{x}{rf2}" for x in "BC"], "tot")
    ig = linea("Impuesto a las ganancias", ["5.5.01"])
    rn = h.fila(["RESULTADO DEL EJERCICIO"] + [f"={x}{rai}+{x}{ig}" for x in "BC"], "tot")
    h.fila(["Control: resultado ER − resultado BSS", f"=B{rn}-({RES25})", f"=C{rn}-({RES26})"])
    h.fila(["Control: resultado ER − resultado del ejercicio en ESP", f"=B{rn}-ESP!C{pn[-1]}", f"=C{rn}-ESP!D{pn[-1]}"])
    h.salto()
    h.fila(["Apertura de gastos por cuenta"], "sec")
    h.fila(["Cuenta", "2025", "2026"], "cab")
    for cta in CUENTAS26:
        k = cta["cod"]
        if k[0] in "45":
            h.fila([f"{k} {cta['cta']}", r25.get(k, 0), r26.get(k, 0)])
    h.anchos([62, 18, 18])
    wb.save(os.path.join(ENT, "eecc_historicos_2025_2026.xlsx"))


# ======================================================================================
if __name__ == "__main__":
    os.makedirs(ENT, exist_ok=True)
    gen_aux_comp_2025()
    gen_bss26()
    gen_aux26()
    gen_eecc()
    control = {
        "resultado_2024": RESULTADOS[2024], "resultado_2025": RES25,
        "rai_2026": RAI26, "ig_2026": IG26, "resultado_2026": RES26,
        "ventas_2026": V26, "cmv_2026": sum(CMV26), "compras_2026": P26,
        "amortizaciones_2026": sum(AM26.values()), "seguros_gasto_2026": TOT_RES26["5.2.06"] + TOT_RES26["5.3.05"],
        "seguros_no_devengados_2026": SF26["1.4.03"], "primas_2026": primas26,
        "resultado_venta_c01": c01["res_venta"], "dividendos_2026": DIVIDENDOS[2026],
        "debitos_2026": sum(DEB.values()), "creditos_2026": sum(CRE.values()),
        "banco_fin_2026": SF26[BANCO], "rna_inicial_2026": SI26["3.3.01"],
        "saldos_finales_2026": SF26, "saldos_iniciales_2026": SI26,
        "capas_fin_2025": [(o.isoformat(), i) for o, i in CAPAS_FIN25],
        "capas_fin_2026": [(o.isoformat(), i) for o, i in CAPAS_FIN26],
        "resultados_historicos": RESULTADOS, "reserva": RESERVA, "dividendos": DIVIDENDOS,
        "asientos_2026": [(d, [list(m) for m in mov]) for d, mov in ASIENTOS],
    }
    with open(os.path.join(ENT, "control_generador.json"), "w", encoding="utf-8") as fh:
        json.dump(control, fh, ensure_ascii=False, indent=1, default=str)
    print(json.dumps({k: v for k, v in control.items() if not isinstance(v, (dict, list))}, indent=1, ensure_ascii=False))
