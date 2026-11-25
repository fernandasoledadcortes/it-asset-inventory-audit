# -*- coding: utf-8 -*-
"""Arma el Excel de entrega: elaborado/entrega/Cortes_Fernanda_DataChallenge.xlsx

Lee la base revisada (02_limpiar), el registro de cambios, la completitud y los chequeos, y escribe
un único .xlsx en inglés con portada, resumen, dashboard, análisis, descripción, calidad,
reglas de validación, supuestos, datos limpios, registro de cambios y datos originales.
Las tablas del análisis son FÓRMULAS sobre Clean_Data (se recalculan si cambian los datos o los
supuestos); los textos del resumen usan los mismos cálculos hechos acá en pandas.
Correr desde la raíz del proyecto:
    python3 codigo/03_entregar/crear_excel_entrega.py
"""

# %% 0. Preparación ------------------------------------------------------------------------------
import sys
from datetime import date, datetime
from pathlib import Path

import openpyxl
import pandas as pd

# Si el editor lee este .py con otra codificación (no UTF-8), los acentos del código se rompen
# y la limpieza da otros resultados: mejor frenar acá.
assert "ñ" == "\u00f1", "Este archivo se leyó con una codificación distinta de UTF-8: abrirlo como UTF-8"
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Carpeta de este script: con __file__ si se corre entero; por celdas (VS Code / Jupyter) no
# existe __file__ y se busca la raíz del proyecto desde la carpeta actual.
try:
    AQUI = Path(__file__).resolve().parent
except NameError:
    RAIZ_PROYECTO = next(p for p in [Path.cwd(), *Path.cwd().parents] if (p / "codigo" / "00_comun").is_dir())
    AQUI = RAIZ_PROYECTO / "codigo" / "03_entregar"
sys.path.insert(0, str(AQUI.parent / "00_comun"))
sys.path.insert(0, str(AQUI))
from campos import (APLICA_A, CON_SO, ESTADOS_DONDE_APLICA, NO_APLICA, TIPOS_ACTIVO,  # noqa: E402
                    USUARIO_FINAL, marcar_no_aplica)
from config import FX  # noqa: E402
from config import (ARCHIVO_CRUDO, ARCHIVO_ENTREGA, CIUDAD, ELABORADO, FECHA_CORTE,  # noqa: E402
                    NOMBRE_COMPLETO, REGISTRO_CAMBIOS)
from textos_en import (CHEQUEOS, COLUMNAS, DICC_CRUDO, DICC_DERIVADAS, NOTAS, REGLAS,  # noqa: E402
                       VALORES)

CORTE = pd.Timestamp(FECHA_CORTE)
df = pd.read_pickle(ELABORADO / "activos_ti_2026_revision.pkl")
chequeos = pd.read_csv(ELABORADO / "activos_ti_2026_chequeos.csv")
completitud = pd.read_csv(ELABORADO / "activos_ti_2026_completitud.csv")
registro = pd.read_csv(REGISTRO_CAMBIOS)
faltan = set(chequeos["chequeo"]) - set(CHEQUEOS)
assert not faltan, f"Chequeos sin traducir: {faltan}"
faltan = set(registro["regla"]) - set(REGLAS)
assert not faltan, f"Reglas del registro sin traducir: {faltan}"

# Estilo
FUENTE = "Arial"
AZUL, AZUL_CLARO, GRIS = "1F3864", "D9E2F3", "F2F2F2"
f_normal = Font(name=FUENTE, size=10)
f_negrita = Font(name=FUENTE, size=10, bold=True)
f_titulo = Font(name=FUENTE, size=16, bold=True, color=AZUL)
f_subtitulo = Font(name=FUENTE, size=12, bold=True, color=AZUL)
f_encabezado = Font(name=FUENTE, size=10, bold=True, color="FFFFFF")
f_input = Font(name=FUENTE, size=10, color="0000FF")      # supuestos editables
f_nota = Font(name=FUENTE, size=9, italic=True, color="595959")
relleno_enc = PatternFill("solid", fgColor=AZUL)
relleno_formula = PatternFill("solid", fgColor="E2EFDA")  # columnas calculadas con fórmula
relleno_input = PatternFill("solid", fgColor="FFF2CC")
relleno_na = PatternFill("solid", fgColor=GRIS)
borde = Border(bottom=Side(style="thin", color="BFBFBF"))
envolver = Alignment(wrap_text=True, vertical="top")

wb = openpyxl.Workbook()
wb.remove(wb.active)
HOJAS = ["Cover", "Executive_Summary", "Dashboard", "Analysis", "Data_Description", "Data_Quality",
         "Validation_Rules", "Assumptions", "Clean_Data", "Flags", "Change_Log", "Original_Data"]
ws_ = {h: wb.create_sheet(h) for h in HOJAS}


def celda(ws, ref, valor, fuente=f_normal, **kw):
    c = ws[ref]
    c.value = valor
    c.font = fuente
    for k, v in kw.items():
        setattr(c, k, v)
    return c


def tabla(ws, fila, col, encabezados, filas, anchos=None, formatos=None):
    """Escribe encabezado + filas desde (fila, col). Devuelve la última fila escrita."""
    for j, h in enumerate(encabezados):
        c = ws.cell(fila, col + j, h)
        c.font, c.fill, c.alignment = f_encabezado, relleno_enc, Alignment(wrap_text=True, vertical="center")
    for i, valores in enumerate(filas, start=1):
        for j, v in enumerate(valores):
            c = ws.cell(fila + i, col + j, v)
            c.font, c.border = f_normal, borde
            c.alignment = Alignment(vertical="top", wrap_text=isinstance(v, str) and len(v) > 40)
            if formatos and formatos[j]:
                c.number_format = formatos[j]
    if anchos:
        for j, a in enumerate(anchos):
            ws.column_dimensions[get_column_letter(col + j)].width = a
    return fila + len(filas)


def valor_excel(v):
    if v is None or v is pd.NA or v is pd.NaT:
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if hasattr(v, "item"):  # numpy
        return v.item()
    return v


# %% 1. Clean_Data -----------------------------------------------------------------------------
# Desde la base revisada: "Not applicable" donde el campo no corresponde, columnas y valores en
# inglés, notas por fila traducidas. Después se agregan columnas con fórmula (USD, edad, garantía).
datos = marcar_no_aplica(df)
for col, mapa in VALORES.items():
    datos[col] = datos[col].map(lambda v: mapa.get(v, v))
datos["notas_limpieza"] = df[list(NOTAS)].apply(
    lambda f: "; ".join(n for b, n in NOTAS.items() if f[b]) or None, axis=1)
datos["contradicciones"] = df["contradicciones"].map(
    lambda s: "; ".join(CHEQUEOS[c][0] for c in s.split("; ")) if isinstance(s, str) else None)
datos["dup_valores_descartados"] = datos["dup_valores_descartados"].str.replace("(fila ", "(source row ", regex=False)
# purchase_cost_usd viene calculada de Python; en el Excel se reemplaza por la fórmula (usa Assumptions)
datos = datos.drop(columns="purchase_cost_usd").rename(columns=COLUMNAS)
datos.insert(0, "source_row", datos.index)
datos = datos.reset_index(drop=True)

# Las banderas flag_* (38, casi todas FALSE) no van como columnas en Clean_Data: van a la hoja Flags en
# formato largo (una fila por bandera activada) y en Clean_Data quedan resumidas en cleaning_notes,
# needs_review y reliable_for_*. Se quedan solo las que usan fórmulas del Analysis. En el .pkl siguen todas.
FLAGS_EN_DATOS = ["flag_email_not_corporate", "flag_user_terminated"]
cols_flag = [c for c in datos.columns if c.startswith("flag_") and c not in FLAGS_EN_DATOS]
nota_en = {COLUMNAS.get(k, k): v for k, v in NOTAS.items()}
banderas = (datos.melt(id_vars=["source_row", "asset_tag", "asset_type"], value_vars=cols_flag, var_name="flag")
            .query("value == True").drop(columns="value"))
banderas["note"] = banderas["flag"].map(lambda f: nota_en.get(f) or DICC_DERIVADAS.get(f, ("", ""))[1])
assert banderas["note"].astype(bool).all(), f"Bandera sin nota: {set(banderas.loc[~banderas['note'].astype(bool), 'flag'])}"
banderas = banderas.sort_values(["source_row", "flag"])
datos = datos.drop(columns=cols_flag)
print(f"Flags: {len(cols_flag)} banderas -> {len(banderas)} filas (fuera de Clean_Data)")

# Las columnas con fórmula van al lado de lo que usan (después de los datos del crudo)
FORMULAS = ["purchase_cost_usd", "book_value_usd", "age_years", "warranty_days_left", "refresh_due",
            "os_unsupported", "days_since_seen"]
columnas = list(datos.columns)
pos = columnas.index("notes") + 1
columnas = columnas[:pos] + FORMULAS + columnas[pos:]
LETRA = {c: get_column_letter(i + 1) for i, c in enumerate(columnas)}
N = len(datos)
ULTIMA = N + 1


def rango(col):
    """Rango absoluto de una columna de Clean_Data (sin encabezado)."""
    return f"Clean_Data!${LETRA[col]}$2:${LETRA[col]}${ULTIMA}"


# %% 2. Assumptions ---------------------------------------------------------------------------------
wa = ws_["Assumptions"]
celda(wa, "A1", "Assumptions (editable inputs in blue; all formulas read from here)", f_titulo)
celda(wa, "A3", "Cutoff date (analysis 'today')", f_negrita)
celda(wa, "B3", datetime(FECHA_CORTE.year, FECHA_CORTE.month, FECHA_CORTE.day), f_input,
      number_format="yyyy-mm-dd", fill=relleno_input)
celda(wa, "C3", "Latest last_seen is 2026-09-30 22:34 and the source Book Value formulas use DATE(2026,10,1); "
      "TODAY() would make results change every time the file is opened.", f_nota)
celda(wa, "A4", "Useful life for book value (years)", f_negrita)
celda(wa, "B4", 5, f_input, fill=relleno_input)
celda(wa, "C4", "149 of 175 source book values match 5 years; the source documentation row says 4 years.", f_nota)
celda(wa, "A5", "Not-seen threshold (days)", f_negrita)
celda(wa, "B5", 90, f_input, fill=relleno_input)
celda(wa, "A6", "Warranty 'expiring soon' window (days)", f_negrita)
celda(wa, "B6", 180, f_input, fill=relleno_input)
CUTOFF, VIDA, UMBRAL, VENTANA = "Assumptions!$B$3", "Assumptions!$B$4", "Assumptions!$B$5", "Assumptions!$B$6"

celda(wa, "A8", "FX rates to USD (1 unit of currency = x USD)", f_subtitulo)
fila_fx = 9
tabla(wa, fila_fx, 1, ["Currency", "USD per unit", "Note"],
      [[k, v, "Approximate 2025 average market rate; replace with the official month-end rate for Sep-2026."
        if k != "USD" else "Base currency"] for k, v in FX.items()], anchos=[38, 14, 95])
for i in range(len(FX)):
    wa.cell(fila_fx + 1 + i, 2).font = f_input
    wa.cell(fila_fx + 1 + i, 2).fill = relleno_input
    wa.cell(fila_fx + 1 + i, 2).number_format = "0.0000"
FX_COD = f"Assumptions!$A${fila_fx + 1}:$A${fila_fx + len(FX)}"
FX_VAL = f"Assumptions!$B${fila_fx + 1}:$B${fila_fx + len(FX)}"

REFRESH = {"Laptop": 4, "Desktop": 5, "Mobile Phone": 3, "Tablet": 3, "Monitor": 6, "Printer": 5,
           "Server": 5, "Network Switch": 6, "Firewall": 5, "Virtual Machine": None}
fila_rf = fila_fx + len(FX) + 2
celda(wa, f"A{fila_rf}", "Refresh age by asset type (years) — also the asset-type catalog", f_subtitulo)
tabla(wa, fila_rf + 1, 1, ["Asset type", "Refresh at age (years)", "Note"],
      [[k, v, "Typical lifecycle policy (assumption); replace with company policy." if v else
        "Not applicable: VMs are not refreshed as hardware."] for k, v in REFRESH.items()])
for i in range(len(REFRESH)):
    wa.cell(fila_rf + 2 + i, 2).font = f_input
    wa.cell(fila_rf + 2 + i, 2).fill = relleno_input
RF_TIPO = f"Assumptions!$A${fila_rf + 2}:$A${fila_rf + 1 + len(REFRESH)}"
RF_EDAD = f"Assumptions!$B${fila_rf + 2}:$B${fila_rf + 1 + len(REFRESH)}"

EOL = [("Windows 7", "2020-01-14"), ("Windows 8.1", "2023-01-10"), ("Windows 10", "2025-10-14"),
       ("Windows Server 2008 R2", "2020-01-14"), ("Windows Server 2012 R2", "2023-10-10"),
       ("CentOS 7", "2024-06-30"), ("ESXi 7", "2025-10-02")]
fila_eol = fila_rf + len(REFRESH) + 3
celda(wa, f"A{fila_eol}", "Operating systems out of vendor support at cutoff", f_subtitulo)
tabla(wa, fila_eol + 1, 1, ["os_version", "End of support", "Note"],
      [[o, f, "Vendor end-of-support date (Microsoft / Red Hat / Broadcom lifecycle pages)."] for o, f in EOL])
celda(wa, f"A{fila_eol + len(EOL) + 2}", "Windows Server 2016 ends extended support on 2027-01-12 (3 months after cutoff). "
      "Mobile OS and macOS are not assessed: Apple / Google do not publish fixed end-of-support dates.", f_nota)
EOL_RNG = f"Assumptions!$A${fila_eol + 2}:$A${fila_eol + 1 + len(EOL)}"

ESTADOS_ORDEN = ["In Use", "In Stock", "In Repair", "Missing", "On Order", "Lost/Stolen", "Retired", "Disposed"]
fila_es = fila_eol + len(EOL) + 4
celda(wa, f"A{fila_es}", "Asset status catalog", f_subtitulo)
tabla(wa, fila_es + 1, 1, ["asset_status", "In active inventory?"],
      [[e, "No" if e in ["Lost/Stolen", "Retired", "Disposed", "On Order"] else "Yes"] for e in ESTADOS_ORDEN])
ES_RNG = f"Assumptions!$A${fila_es + 2}:$A${fila_es + 1 + len(ESTADOS_ORDEN)}"
wa.column_dimensions["A"].width = 40

# %% 3. Escribir Clean_Data con fórmulas ------------------------------------------------------------
wc = ws_["Clean_Data"]
for j, c in enumerate(columnas, start=1):
    h = wc.cell(1, j, c)
    h.font, h.fill = f_encabezado, relleno_enc
    if c in FORMULAS:
        h.fill = PatternFill("solid", fgColor="548235")
FORMATO = {"purchase_date": "yyyy-mm-dd", "warranty_end_date": "yyyy-mm-dd", "last_seen": "yyyy-mm-dd hh:mm",
           "malware_reimage_date": "yyyy-mm-dd", "lease_end_date": "yyyy-mm-dd", "audit_month": "yyyy-mm",
           "purchase_cost": "#,##0.00", "book_value": "#,##0.00", "book_value_calc": "#,##0.00",
           "purchase_cost_usd": "$#,##0", "book_value_usd": "$#,##0", "age_years": "0.0"}
L = LETRA
f_gris = Font(name=FUENTE, size=10, color="808080")
for i, fila in enumerate(datos.itertuples(index=False), start=2):
    registro_fila = dict(zip(datos.columns, fila))
    for j, c in enumerate(columnas, start=1):
        if c in FORMULAS:
            continue
        v = valor_excel(registro_fila[c])
        cel = wc.cell(i, j, v)
        cel.font = f_gris if v == NO_APLICA else f_normal
        if c in FORMATO:
            cel.number_format = FORMATO[c]
    f = {
        "purchase_cost_usd": f'=IF(ISNUMBER({L["purchase_cost"]}{i}),IFERROR({L["purchase_cost"]}{i}*INDEX({FX_VAL},MATCH({L["currency"]}{i},{FX_COD},0)),""),"")',
        "book_value_usd": f'=IF(ISNUMBER({L["book_value_calc"]}{i}),IFERROR({L["book_value_calc"]}{i}*INDEX({FX_VAL},MATCH({L["currency"]}{i},{FX_COD},0)),""),"")',
        "age_years": f'=IF(ISNUMBER({L["purchase_date"]}{i}),ROUND(({CUTOFF}-{L["purchase_date"]}{i})/365.25,1),"")',
        "warranty_days_left": f'=IF(ISNUMBER({L["warranty_end_date"]}{i}),{L["warranty_end_date"]}{i}-{CUTOFF},"")',
        "refresh_due": f'=IF(ISNUMBER({L["age_years"]}{i}),IF({L["age_years"]}{i}>=IFERROR(1/(1/INDEX({RF_EDAD},MATCH({L["asset_type"]}{i},{RF_TIPO},0))),999),"Yes","No"),"")',
        "os_unsupported": f'=IF(COUNTIF({EOL_RNG},{L["os_version"]}{i})>0,"Yes","No")',
        "days_since_seen": f'=IF(ISNUMBER({L["last_seen"]}{i}),INT({CUTOFF}-{L["last_seen"]}{i}),"")',
    }
    for c, formula in f.items():
        cel = wc.cell(i, columnas.index(c) + 1, formula)
        cel.fill, cel.font = relleno_formula, f_normal
        if c in FORMATO:
            cel.number_format = FORMATO[c]
for j, c in enumerate(columnas, start=1):
    wc.column_dimensions[get_column_letter(j)].width = min(max(len(c) + 2, 10), 40)
wc.freeze_panes = "C2"
wc.auto_filter.ref = f"A1:{get_column_letter(len(columnas))}{ULTIMA}"
print(f"Clean_Data: {N} filas x {len(columnas)} columnas ({len(FORMULAS)} con fórmula)")

# %% 4. Cálculos en pandas (para los textos y para verificar las fórmulas) ------------------------------
d = df.copy()
d["usd"] = d["purchase_cost"] * d["currency"].map(FX)
d["bv_usd"] = d["book_value_calc"] * d["currency"].map(FX)
d["edad"] = ((CORTE - d["purchase_date"]).dt.days / 365.25).round(1)
d["dias_garantia"] = (d["warranty_end_date"] - CORTE).dt.days
d["dias_sin_ver"] = (CORTE - d["last_seen"]).dt.days
d["refresh"] = d["edad"] >= d["asset_type"].map(REFRESH)
d["so_sin_soporte"] = d["os_version"].isin([o for o, _ in EOL])
act = d[d["en_inventario_activo"]]
K = {
    "registros": len(d), "activos": len(act),
    "gasto_usd": act["usd"].sum(), "gasto_usd_conf": act.loc[act["confiable_costos"], "usd"].sum(),
    "atipicos": int(d["flag_costo_atipico"].sum()),
    "atipicos_usd": (act["purchase_cost_descartado"] * act["currency"].map(FX)).sum(),
    "bv_usd": act.loc[act["confiable_costos"], "bv_usd"].sum(),
    "so_sin_soporte": int(act["so_sin_soporte"].sum()),
    "win10": int((act["os_version"] == "Windows 10").sum()),
    "cifrado_desconocido": int((act["asset_type"].isin(USUARIO_FINAL) & ~act["encrypted"].eq("Yes")).sum()),
    "usuario_final": int(act["asset_type"].isin(USUARIO_FINAL).sum()),
    "no_parcheado": int(act["patch_compliance"].isin(["Non-Compliant", "Pending Reboot"]).sum()),
    "con_parche": int(act["patch_compliance"].notna().sum()),
    "sin_ver_90": int((act["asset_status"].eq("In Use") & act["asset_type"].isin(CON_SO) & (act["dias_sin_ver"] > 90)).sum()),
    "ex_empleado": int((act["flag_user_terminated"]).sum()),
    "perdidos_reportando": int((d["asset_status"].eq("Lost/Stolen") & (d["dias_sin_ver"] <= 30)).sum()),
    "garantia_vencida_en_uso": int((act["asset_status"].eq("In Use") & act["confiable_fechas"] & (act["dias_garantia"] < 0)).sum()),
    "garantia_180": int((act["confiable_fechas"] & act["dias_garantia"].between(0, 180)).sum()),
    "refresh": int((act["refresh"] & act["confiable_fechas"]).sum()),
    "refresh_usd": act.loc[act["refresh"] & act["confiable_fechas"], "usd"].sum(),
    "revisar": int(d["revisar"].sum()), "contradiccion": int(d["n_contradicciones"].gt(0).sum()),
    "conf_inv": int(d["confiable_inventario"].sum()), "conf_cost": int(d["confiable_costos"].sum()),
    "conf_fech": int(d["confiable_fechas"].sum()),
    "cambios": len(registro), "edad_mediana": act["edad"].median(),
    "dup_host": int(d["flag_hostname_conflicto"].sum()),
    "email_otro": int(d["contradicciones"].fillna("").str.contains("email_no_coincide").sum()),
    "leasing": int(d["contradicciones"].fillna("").str.contains("leasing_vencido").sum()),
}
K["excluidas"] = int(registro["regla"].str.startswith("Excluida: duplicado").sum()
                     + registro["regla"].str.startswith("Excluida: mismo").sum())
print({k: (round(v) if isinstance(v, float) else v) for k, v in K.items()})

# %% 5. Analysis (fórmulas sobre Clean_Data) ------------------------------------------------------------
wn = ws_["Analysis"]
celda(wn, "A1", "Analysis — all figures are formulas on Clean_Data and Assumptions", f_titulo)
celda(wn, "A2", "Scope: active inventory (in_active_inventory = TRUE: excludes Retired, Disposed, Lost/Stolen, On Order) "
      "unless stated. Spend uses USD at the Assumptions FX rates and only rows reliable_for_cost = TRUE; "
      "age and warranty use reliable_for_dates = TRUE.", f_nota)
ACT = f'{rango("in_active_inventory")},TRUE'
fila = 4
ANCLAS = {}  # nombre de tabla -> (fila_encabezado, fila_ultima, col_etiqueta, cols de valores)


def bloque(titulo, encabezados, filas, formatos=None, nombre=None):
    global fila
    celda(wn, f"A{fila}", titulo, f_subtitulo)
    ini = fila + 1
    fin = tabla(wn, ini, 1, encabezados, filas, formatos=formatos)
    if nombre:
        ANCLAS[nombre] = (ini, fin)
    fila = fin + 3
    return ini, fin


# 5a. KPIs
kpis = [
    ["Records in clean register", f'=COUNTA({rango("source_row")})', "Rows after excluding 15 duplicate rows (see Change_Log)"],
    ["Active inventory (assets)", f'=COUNTIFS({ACT})', "Excludes Retired, Disposed, Lost/Stolen, On Order"],
    ["Purchase spend of active inventory (USD, reliable cost)", f'=SUMIFS({rango("purchase_cost_usd")},{ACT},{rango("reliable_for_cost")},TRUE)', "Original purchase cost converted at Assumptions FX"],
    ["Book value of active inventory (USD, 5-yr straight line)", f'=SUMIFS({rango("book_value_usd")},{ACT},{rango("reliable_for_cost")},TRUE)', "book_value_calc x FX"],
    ["Median age of active assets (years)", f'=MEDIAN(IF(({rango("in_active_inventory")}=TRUE)*ISNUMBER({rango("age_years")}),{rango("age_years")}))', "Array formula"],
    ["Rows needing review (needs_review)", f'=COUNTIFS({rango("needs_review")},TRUE)', "Cleaning decision to confirm or contradiction"],
    ["Rows with at least one contradiction", f'=COUNTIFS({rango("n_contradictions")},">0")', ""],
]
ini, fin = bloque("Key figures", ["Metric", "Value", "Note"], kpis, formatos=[None, "#,##0", None], nombre="kpi")
for r in (ini + 3, ini + 4):
    wn.cell(r, 2).number_format = "$#,##0"
wn.cell(ini + 5, 2).number_format = "0.0"
from openpyxl.worksheet.formula import ArrayFormula  # noqa: E402
wn.cell(ini + 5, 2).value = ArrayFormula(f"B{ini + 5}", kpis[4][1])
KPI_CELDA = {k: f"Analysis!$B${ini + 1 + i}" for i, k in enumerate(
    ["registros", "activos", "gasto", "bv", "edad", "revisar", "contradiccion"])}

# 5b. Estado por tipo (registro completo)
tipos = d["asset_type"].value_counts().index.tolist()
filas_t = []
for t in tipos:
    filas_t.append([t] + [f'=COUNTIFS({rango("asset_type")},"{t}",{rango("asset_status")},"{e}")' for e in ESTADOS_ORDEN])
ini, fin = bloque("Full register by asset type and status (count of assets)", ["Asset type"] + ESTADOS_ORDEN + ["Total", "Active"],
                  [r + [None, None] for r in filas_t], nombre="tipo_estado")
for k in range(ini + 1, fin + 1):
    wn.cell(k, 10).value = f"=SUM(B{k}:I{k})"
    wn.cell(k, 11).value = f'=COUNTIFS({rango("asset_type")},A{k},{ACT})'
wn.cell(fin + 1, 1, "Total").font = f_negrita
for j in range(2, 12):
    col = get_column_letter(j)
    wn.cell(fin + 1, j, f"=SUM({col}{ini + 1}:{col}{fin})").font = f_negrita
fila += 1

# 5c. Por sitio
sitios = act["location"].value_counts().index.tolist()
filas_s = []
for s in sitios:
    filas_s.append([s, None, None, None, None, None])
ini, fin = bloque("Active inventory by site (spend in USD, reliable cost only)",
                  ["Site", "Country", "Active assets", "Purchase spend (USD)", "Book value (USD)",
                   "Avg laptop cost (USD)", "Spend per asset (USD)"], [r + [None] for r in filas_s],
                  formatos=[None, None, "#,##0", "$#,##0", "$#,##0", "$#,##0", "$#,##0"], nombre="sitio")
for k in range(ini + 1, fin + 1):
    s = f"A{k}"
    wn.cell(k, 2).value = f'=INDEX({rango("country")},MATCH({s},{rango("location")},0))'
    wn.cell(k, 3).value = f'=COUNTIFS({rango("location")},{s},{ACT})'
    wn.cell(k, 4).value = f'=SUMIFS({rango("purchase_cost_usd")},{rango("location")},{s},{ACT},{rango("reliable_for_cost")},TRUE)'
    wn.cell(k, 5).value = f'=SUMIFS({rango("book_value_usd")},{rango("location")},{s},{ACT},{rango("reliable_for_cost")},TRUE)'
    wn.cell(k, 6).value = f'=IFERROR(AVERAGEIFS({rango("purchase_cost_usd")},{rango("location")},{s},{ACT},{rango("reliable_for_cost")},TRUE,{rango("asset_type")},"Laptop"),"")'
    wn.cell(k, 7).value = f'=IFERROR(D{k}/COUNTIFS({rango("location")},{s},{ACT},{rango("reliable_for_cost")},TRUE,{rango("purchase_cost_usd")},">0"),"")'
celda(wn, f"A{fin + 1}", "Fair comparison across sites: convert to one currency (USD), compare like-for-like (same asset type, "
      "e.g. average laptop cost) and per asset, not totals — sites differ in size and mix (Eagan is the data center).", f_nota)
fila += 1

# 5d. Edad
cortes = [(0, 1), (1, 2), (2, 3), (3, 4), (4, 5), (5, 6), (6, 99)]
filas_e = [[f"{a}-{b} yrs" if b < 99 else f"{a}+ yrs",
            f'=COUNTIFS({ACT},{rango("reliable_for_dates")},TRUE,{rango("age_years")},">={a}",{rango("age_years")},"<{b}")']
           for a, b in cortes]
bloque("Age of active assets (years since purchase, reliable dates)", ["Age band", "Assets"], filas_e, nombre="edad")

# 5e. Garantía
bandas = [("Expired", None, 0), ("0-90 days", 0, 91), ("91-180 days", 91, 181), ("181-365 days", 181, 366),
          ("More than 1 year", 366, None)]
filas_g = []
for nombre, a, b in bandas:
    crit = [ACT, f'{rango("reliable_for_dates")},TRUE']
    if a is not None:
        crit.append(f'{rango("warranty_days_left")},">={a}"')
    if b is not None:
        crit.append(f'{rango("warranty_days_left")},"<{b}"')
    filas_g.append([nombre, f'=COUNTIFS({",".join(crit)})'])
filas_g.append(["Unknown (physical, no date)", f'=SUMPRODUCT(({rango("in_active_inventory")}=TRUE)*({rango("asset_type")}<>"Virtual Machine")*({rango("warranty_end_date")}=""))'])
bloque("Warranty status of active assets (days left at cutoff)", ["Warranty", "Assets"], filas_g, nombre="garantia")

# 5f. Refresh por tipo
filas_r = []
for t in [t for t in TIPOS_ACTIVO if REFRESH[t]]:
    filas_r.append([t, f'=INDEX({RF_EDAD},MATCH(A{{k}},{RF_TIPO},0))',
                    f'=COUNTIFS({rango("asset_type")},A{{k}},{ACT})',
                    f'=COUNTIFS({rango("asset_type")},A{{k}},{ACT},{rango("reliable_for_dates")},TRUE,{rango("refresh_due")},"Yes")',
                    f'=SUMIFS({rango("purchase_cost_usd")},{rango("asset_type")},A{{k}},{ACT},{rango("reliable_for_dates")},TRUE,{rango("refresh_due")},"Yes",{rango("reliable_for_cost")},TRUE)'])
ini, fin = bloque("Assets due for refresh (age >= policy age), active, reliable dates",
                  ["Asset type", "Refresh age (yrs)", "Active assets", "Due for refresh", "Replacement budget proxy (USD, original cost)"],
                  [[r[0], None, None, None, None] for r in filas_r], formatos=[None, "0", "#,##0", "#,##0", "$#,##0"], nombre="refresh")
for k, r in zip(range(ini + 1, fin + 1), filas_r):
    for j in range(1, 5):
        wn.cell(k, j + 1).value = r[j].replace("{k}", str(k))
wn.cell(fin + 1, 1, "Total").font = f_negrita
for j in (3, 4, 5):
    col = get_column_letter(j)
    wn.cell(fin + 1, j, f"=SUM({col}{ini + 1}:{col}{fin})").font = f_negrita
wn.cell(fin + 1, 5).number_format = "$#,##0"
fila += 1

# 5g. Riesgos
UF = '{"Laptop","Desktop","Mobile Phone","Tablet"}'
SO = '{"Laptop","Desktop","Server","Virtual Machine","Mobile Phone","Tablet"}'
riesgos = [
    ["Security", "OS out of vendor support (Win 7/8.1/10, Server 2008/2012 R2, CentOS 7, ESXi 7)",
     f'=COUNTIFS({ACT},{rango("os_unsupported")},"Yes")', "No security patches; audit finding"],
    ["Security", "End-user device with encryption not confirmed (No / Pending / unknown)",
     f'=SUM(COUNTIFS({ACT},{rango("asset_type")},{UF},{rango("encrypted")},"<>Yes"))', "Data-loss exposure if lost; blank = unknown"],
    ["Security", "Lost/Stolen device still checking in (last 30 days)",
     f'=COUNTIFS({rango("asset_status")},"Lost/Stolen",{rango("days_since_seen")},"<=30")', "Someone is using it: remote wipe / lock"],
    ["Compliance", "Patch Non-Compliant or Pending Reboot",
     f'=COUNTIFS({ACT},{rango("patch_compliance")},"Non-Compliant")+COUNTIFS({ACT},{rango("patch_compliance")},"Pending Reboot")', "Of devices reporting patch status"],
    ["Compliance", f"In Use but not seen for more than the threshold (days, Assumptions)",
     f'=SUM(COUNTIFS({rango("asset_status")},"In Use",{rango("asset_type")},{SO},{rango("days_since_seen")},">"&{UMBRAL}))', "Possibly lost, off or unmanaged"],
    ["Compliance", "Device still assigned to a leaver (Terminated)",
     f'=COUNTIFS({ACT},{rango("flag_user_terminated")},TRUE)', "Leaver process did not recover the device"],
    ["Compliance", "Lease ended but device still In Use",
     f'=COUNTIFS({rango("contradictions")},"*lease_expired_in_use*")', "Contract penalty risk"],
    ["Support", "In Use with warranty expired",
     f'=COUNTIFS({rango("asset_status")},"In Use",{rango("reliable_for_dates")},TRUE,{rango("warranty_days_left")},"<0")', "Repairs at full cost"],
    ["Support", "Warranty expiring within the window (days, Assumptions)",
     f'=COUNTIFS({ACT},{rango("reliable_for_dates")},TRUE,{rango("warranty_days_left")},">=0",{rango("warranty_days_left")},"<="&{VENTANA})', "Decide renew vs replace"],
    ["Support", "Critical assets (criticality 4-5) with expired warranty or unsupported OS",
     f'=SUMPRODUCT(({rango("in_active_inventory")}=TRUE)*ISNUMBER({rango("business_criticality")})*({rango("business_criticality")}>=4)*((({rango("warranty_days_left")}<>"")*({rango("warranty_days_left")}<0))+({rango("os_unsupported")}="Yes")>0))', "Highest-impact exposure"],
]
ini, fin = bloque("Risk register (active inventory unless stated)", ["Area", "Risk", "Assets", "Why it matters"], riesgos,
                  formatos=[None, None, "#,##0", None], nombre="riesgo")
wn.column_dimensions["B"].width = 18

# 5h. SO sin soporte por versión
filas_o = [[o, f'=COUNTIFS({ACT},{rango("os_version")},"{o}")'] for o, _ in EOL]
bloque("Active assets on unsupported OS, by version", ["os_version", "Active assets"], filas_o, nombre="so")

# 5i. Marcas: incidentes por equipo (usuario final activo)
uf = act[act["asset_type"].isin(USUARIO_FINAL)]
marcas = uf["manufacturer"].value_counts()
marcas = marcas[marcas >= 10].index.tolist()
filas_m = []
for m in marcas:
    filas_m.append([m, None, None, None, None, None])
ini, fin = bloque("End-user devices by manufacturer (active, brands with 10+ devices): support load",
                  ["Manufacturer", "Devices", "Open incidents", "Incidents per device", "Hardware-failure replacements",
                   "Swollen batteries"], filas_m, formatos=[None, "#,##0", "#,##0", "0.00", "#,##0", "#,##0"], nombre="marca")
for k in range(ini + 1, fin + 1):
    m = f"A{k}"
    wn.cell(k, 2).value = f'=SUM(COUNTIFS({ACT},{rango("manufacturer")},{m},{rango("asset_type")},{UF}))'
    wn.cell(k, 3).value = f'=SUM(SUMIFS({rango("open_incidents")},{ACT},{rango("manufacturer")},{m},{rango("asset_type")},{UF}))'
    wn.cell(k, 4).value = f"=IFERROR(C{k}/B{k},\"\")"
    wn.cell(k, 5).value = f'=SUM(COUNTIFS({ACT},{rango("manufacturer")},{m},{rango("asset_type")},{UF},{rango("evt_replacement_failure")},TRUE))'
    wn.cell(k, 6).value = f'=SUM(COUNTIFS({ACT},{rango("manufacturer")},{m},{rango("asset_type")},{UF},{rango("evt_battery_swelling")},TRUE))'

# 5j. Calidad: completitud donde aplica (top) y confiabilidad
comp = completitud[completitud["pct_vacio_donde_aplica"].fillna(0) > 0].head(12)
filas_c = [[r.variable, r.filas_esperadas, r.vacias_donde_aplica, r.pct_vacio_donde_aplica / 100] for r in comp.itertuples()]
bloque("Missing data where the field applies (by asset type and status) — top fields",
       ["Field", "Rows where it applies", "Blank", "% blank"], filas_c, formatos=[None, "#,##0", "#,##0", "0.0%"], nombre="completitud")
filas_q = [["Reliable for inventory counts", f'=COUNTIFS({rango("reliable_for_inventory")},TRUE)'],
           ["Reliable for cost / value", f'=COUNTIFS({rango("reliable_for_cost")},TRUE)'],
           ["Reliable for age / warranty", f'=COUNTIFS({rango("reliable_for_dates")},TRUE)'],
           ["Needs review (any reason)", f'=COUNTIFS({rango("needs_review")},TRUE)']]
ini, fin = bloque("How much can be relied on (rows of the clean register)", ["Use", "Rows", "% of register"],
                  [r + [None] for r in filas_q], formatos=[None, "#,##0", "0.0%"], nombre="confianza")
for k in range(ini + 1, fin + 1):
    wn.cell(k, 3).value = f"=B{k}/{KPI_CELDA['registros']}"
wn.column_dimensions["A"].width = 34
for c in "CDEFGHIJK":
    wn.column_dimensions[c].width = 16
wn.column_dimensions["B"].width = 62

# %% 6. Dashboard (gráficos nativos sobre Analysis) ----------------------------------------------------
wd = ws_["Dashboard"]
celda(wd, "A1", "IT Asset Register — Dashboard", f_titulo)
celda(wd, "A2", f"Cutoff {FECHA_CORTE:%Y-%m-%d} · active inventory unless stated · source: Analysis tab (formulas on Clean_Data)", f_nota)
etiquetas = [("Active assets", "activos", "#,##0"), ("Purchase spend (USD)", "gasto", "$#,##0,,\"M\""),
             ("Book value (USD)", "bv", "$#,##0,,\"M\""), ("Median age (yrs)", "edad", "0.0"),
             ("Rows needing review", "revisar", "#,##0")]
for j, (txt, k, fmt) in enumerate(etiquetas):
    col = get_column_letter(1 + j * 3)
    celda(wd, f"{col}4", txt, Font(name=FUENTE, size=9, color="595959"))
    celda(wd, f"{col}5", f"={KPI_CELDA[k]}", Font(name=FUENTE, size=18, bold=True, color=AZUL), number_format=fmt)


def grafico(nombre_tabla, titulo, col_val, ancla, eje_y, horizontal=True, ancho=16, alto=8, formato=None,
            col_cat=1, filas=None):
    ini, fin = ANCLAS[nombre_tabla]
    if filas:
        fin = ini + filas
    ch = BarChart()
    ch.type = "bar" if horizontal else "col"
    ch.title = titulo
    ch.style = 10
    ch.y_axis.title = eje_y
    ch.legend = None
    datos_ref = Reference(wn, min_col=col_val, min_row=ini, max_row=fin)
    cats = Reference(wn, min_col=col_cat, min_row=ini + 1, max_row=fin)
    ch.add_data(datos_ref, titles_from_data=True)
    ch.set_categories(cats)
    ch.dataLabels = DataLabelList()
    ch.dataLabels.showVal = True
    if formato:
        ch.dataLabels.numFmt = formato
        ch.y_axis.number_format = formato
    ch.y_axis.majorGridlines = None
    ch.y_axis.delete = False
    ch.x_axis.delete = False
    if horizontal:
        ch.x_axis.scaling.orientation = "maxMin"  # el más grande arriba
    ch.series[0].graphicalProperties.solidFill = AZUL
    ch.width, ch.height = ancho, alto
    wd.add_chart(ch, ancla)


grafico("tipo_estado", "Active assets by type", 11, "A7", "Assets")
grafico("sitio", "Active assets by site", 3, "J7", "Assets")
grafico("edad", "Age of active assets", 2, "A24", "Assets", horizontal=False)
grafico("garantia", "Warranty status of active assets", 2, "J24", "Assets", horizontal=False)
grafico("riesgo", "Risk register: assets exposed", 3, "A41", "Assets", col_cat=2, ancho=33, alto=9)
grafico("sitio", "Purchase spend by site (USD, reliable cost)", 4, "A60", "USD", formato="$#,##0,\"K\"")
grafico("marca", "Open incidents per end-user device, by brand", 4, "J60", "Incidents per device", formato="0.00")
grafico("completitud", "Missing data where the field applies", 4, "A77", "% blank", formato="0%")
grafico("refresh", "Assets due for refresh, by type", 4, "J77", "Assets")

# %% 7. Executive_Summary ---------------------------------------------------------------------------
we = ws_["Executive_Summary"]
we.column_dimensions["A"].width = 4
we.column_dimensions["B"].width = 120
celda(we, "B1", "Executive Summary", f_titulo)
celda(we, "B2", f"IT asset register review ahead of the hardware refresh and internal audit · cutoff {FECHA_CORTE:%d %b %Y} · "
      f"{NOMBRE_COMPLETO}, {CIUDAD}", f_nota)
M = 1e6
resumen = [
    ("Bottom line", None),
    (f"The register can be used for the refresh and the audit, but not as is. After cleaning, {K['registros']} records remain "
     f"({K['excluidas']} duplicate rows removed, nothing else deleted), of which {K['activos']} are today's active estate. "
     f"Counts are reliable for {K['conf_inv'] / K['registros']:.0%} of records, costs for {K['conf_cost'] / K['registros']:.0%} "
     f"and dates for {K['conf_fech'] / K['registros']:.0%}. Every doubtful value is flagged, not hidden: {K['revisar']} rows "
     "carry a specific reason to check (needs_review).", "p"),
    ("Key findings", None),
    (f"1. Security exposure is the biggest issue. {K['so_sin_soporte']} active assets run an OS without vendor support "
     f"({K['win10']} on Windows 10, out of support since Oct-2025). For {K['cifrado_desconocido']} of {K['usuario_final']} "
     f"active end-user devices, encryption is not confirmed (mostly blank = unknown, which an auditor treats as a gap). "
     f"{K['perdidos_reportando']} devices marked Lost/Stolen are still checking in.", "p"),
    (f"2. Managed-device hygiene. {K['no_parcheado']} of {K['con_parche']} devices reporting patch status are Non-Compliant or "
     f"Pending Reboot; {K['sin_ver_90']} 'In Use' devices have not checked in for more than 90 days (lost, off or unmanaged); "
     f"{K['ex_empleado']} active devices are still assigned to leavers.", "p"),
    (f"3. Refresh and support. The median active asset is {K['edad_mediana']:.1f} years old; {K['refresh']} active assets are "
     f"past the refresh age for their type (about ${K['refresh_usd'] / M:.1f}M at original prices, a budget proxy). "
     f"{K['garantia_vencida_en_uso']} devices in use are out of warranty and {K['garantia_180']} more expire within 180 days.", "p"),
    (f"4. Value. Purchase spend of the active estate is about ${K['gasto_usd_conf'] / M:.1f}M (USD, reliable rows) with a book "
     f"value of ${K['bv_usd'] / M:.1f}M. Costs come in 10 currencies, so sites can only be compared after converting to USD and "
     "comparing the same asset type per unit (e.g. average laptop cost), not totals. "
     f"{K['atipicos']} amounts are implausible for their type (> 3x the type median, e.g. a US$442,796 laptop): they alone "
     f"would add ${K['atipicos_usd'] / M:.1f}M, so they were removed from the cost field and flagged until confirmed (most look like local "
     "currency loaded as USD or a lost decimal point).", "p"),
    ("5. The data shows broken processes, not just typos. Several tools feed the register without a shared key: the same "
     f"device loaded twice by a second source (tag as a bare number, no SCCM data), {K['dup_host']} live devices sharing a "
     f"hostname with another live device, {K['email_otro']} devices whose email belongs to someone else (re-assigned without "
     f"updating), {K['leasing']} leases expired but still in use, retired assets still reporting. Book value is filled in only "
     "18% of rows, with a 5-year life that contradicts the documented 4 years.", "p"),
    (f"6. Format chaos, now fixed. {K['cambios']:,} cell-level changes were logged: 55 spellings of 10 asset types, 53 of 18 "
     "brands, 9 asset-tag formats, mixed date formats (incl. dd/mm vs mm/dd and quarters), broken accents, costs as text, "
     "MB vs GB. All are standardized to catalogs and traceable in Change_Log.", "p"),
    ("Recommendations", None),
    ("1. Immediate (audit): remote-wipe / lock the Lost/Stolen devices still checking in; recover leaver devices; plan Windows 10 "
     "upgrades or Extended Security Updates; confirm encryption status from Intune / BitLocker reports for every end-user device.", "p"),
    ("2. Refresh: prioritize assets past refresh age that are also out of warranty, critical (4-5) or on unsupported OS; use the "
     "Analysis tab budget proxy by type, re-priced with current quotes.", "p"),
    ("3. One source of truth: make asset_tag the mandatory key in every feeding tool (SCCM / Intune, procurement, service desk) and "
     "reconcile daily on asset_tag + serial; block records without a valid tag (AST-######).", "p"),
    ("4. Controlled values: drop-down catalogs for type, brand, status, site, currency and department; ISO dates; numeric cost "
     "with a separate currency; email and person taken from the directory (HR / Entra ID), not typed.", "p"),
    ("5. Lifecycle triggers: when an employee leaves, a device is returned or retired, or a lease ends, the status and "
     "assignment must change in the same ticket; alert on retired devices that check in and on In Use devices silent > 90 days.", "p"),
    ("6. Governance: document naming standards for servers and network devices (today only end-user devices are covered); "
     "do not reuse a hostname until the old device is retired; compute book value centrally with one useful-life policy; "
     "replace the unnamed review column with a tracked data-quality queue; run the Validation_Rules checks monthly.", "p"),
    ("How to read this workbook", None),
    ("Dashboard and Analysis show the numbers (formulas, change Assumptions to update). Data_Description, Data_Quality and "
     "Validation_Rules explain the data, the issues and how they were treated. Clean_Data is the cleaned register: filter "
     "needs_review / reliable_for_* to choose rows; 'Not applicable' = field does not apply to that type or status, blank = "
     "missing. Change_Log traces every change back to Original_Data (untouched source).", "p"),
]
r = 4
for texto, tipo in resumen:
    if tipo is None:
        r += 1
        celda(we, f"B{r}", texto, f_subtitulo)
    else:
        celda(we, f"B{r}", texto, f_normal, alignment=envolver)
        we.row_dimensions[r].height = 15 * (len(texto) // 115 + 1)
    r += 1

# %% 8. Cover ---------------------------------------------------------------------------------------------
wp = ws_["Cover"]
wp.column_dimensions["A"].width = 4
wp.column_dimensions["B"].width = 28
wp.column_dimensions["C"].width = 110
celda(wp, "B2", "ITSAM Data Challenge — IT Asset Register Review", f_titulo)
datos_portada = [("Candidate", NOMBRE_COMPLETO), ("City of residence", CIUDAD), ("Position", "Report Analyst"),
                 ("Submitted", date.today().strftime("%Y-%m-%d")), ("Analysis cutoff", f"{FECHA_CORTE:%Y-%m-%d} (see Assumptions)"),
                 ("Source file", "DataChallenge.xlsx (1,009 rows incl. 1 documentation row and 9 empty rows; 34 columns)")]
for i, (k, v) in enumerate(datos_portada, start=4):
    celda(wp, f"B{i}", k, f_negrita)
    celda(wp, f"C{i}", v)
r = 4 + len(datos_portada) + 1
celda(wp, f"B{r}", "Contents", f_subtitulo)
contenido = [
    ("Executive_Summary", "Bottom line, key findings and recommendations for the IT Asset Manager (Presentation)."),
    ("Dashboard", "Charts of the estate, risk, spend, refresh and data quality (Visualization)."),
    ("Analysis", "Formula tables behind the dashboard: estate, sites, age, warranty, refresh, risk, brands (Data Mining)."),
    ("Data_Description", "Grain, structure and field dictionary: meaning, type, where each field applies (Data Description)."),
    ("Data_Quality", "Issue catalog: what, where, how many, why it matters and how it was treated; completeness (Cleansing)."),
    ("Validation_Rules", "Checks that prove the clean data is fit for use, with live results (Validation)."),
    ("Assumptions", "Cutoff date, FX rates, useful life, refresh ages, unsupported OS list — editable inputs."),
    ("Clean_Data", "The cleaned register (985 assets): original fields standardized, notes, review / reliability columns and calculated columns."),
    ("Flags", "One row per cleaning flag raised (source_row, flag, note): filter by flag to see every row a decision touched."),
    ("Change_Log", "Every cell changed and every row / column excluded, with the rule applied (traceability)."),
    ("Original_Data", "The source sheet exactly as received (values)."),
]
for i, (h, t) in enumerate(contenido, start=r + 1):
    c = celda(wp, f"B{i}", h, Font(name=FUENTE, size=10, color="0563C1", underline="single"))
    c.hyperlink = f"#'{h}'!A1"
    celda(wp, f"C{i}", t)
r = r + len(contenido) + 2
celda(wp, f"B{r}", "Use of AI", f_subtitulo)
ia = ("I used Claude (Anthropic), through Claude Code, as a pair-programming assistant. I set the goals and the "
      "cleaning rules and made the decisions (for example: which duplicates to drop, how to resolve ambiguous dates, "
      "rebuilding impossible purchase years from the warranty, labelling 'Not applicable' vs missing). Claude helped "
      "organize the project folders, write and run the Python code step by step, propose checks and options for each "
      "decision, and draft the English text. I reviewed every step and every decision is documented (Data_Quality, "
      "Change_Log, code comments).")
celda(wp, f"C{r + 1}", ia, alignment=envolver)
wp.row_dimensions[r + 1].height = 80
celda(wp, f"B{r + 3}", "Tools and libraries", f_subtitulo)
herramientas = [("Python 3.10", "Language for cleaning, validation and building this workbook."),
                ("pandas 2.1", "Data loading, cleaning and analysis."),
                ("openpyxl 3.1", "Reading the source and writing this workbook (formulas, native charts, formatting)."),
                ("Python standard library", "re (patterns), datetime, pathlib, unicodedata (accents / encoding)."),
                ("Microsoft Excel", "Formula recalculation and final review.")]
for i, (k, v) in enumerate(herramientas, start=r + 4):
    celda(wp, f"B{i}", k, f_negrita)
    celda(wp, f"C{i}", v)
r = r + 4 + len(herramientas) + 1
celda(wp, f"B{r}", "Reproducibility", f_subtitulo)
celda(wp, f"C{r + 1}", "Three scripts run in order: 01 explorar_activos_ti.py (description + surface cleaning), "
      "02 revisar_contradicciones.py (contradictions, reliability, completeness), 03 crear_excel_entrega.py (this file). "
      "The source file is never modified.", alignment=envolver)
wp.row_dimensions[r + 1].height = 30

# %% 9. Data_Description ----------------------------------------------------------------------------------
wdd = ws_["Data_Description"]
celda(wdd, "A1", "Data Description", f_titulo)
desc = [
    ("What it is", "Export of the IT asset register of a fictional global company: end-user devices (laptops, desktops, "
     "monitors, phones, tablets), infrastructure (servers, switches, firewalls) and virtual machines, 14 sites in 10 countries."),
    ("Grain", "One row = one IT asset (one physical device or one VM). Identified by asset_tag (internal label); serial_number "
     "and hostname are secondary identifiers."),
    ("Structure", "Source sheet: 1,009 rows x 34 columns. Row 2 is a documentation row (valid values / formulas as text), "
     "9 rows are completely empty, 3 columns (AF-AH) have no header and no data, one column has no header (reviewer marks), "
     "and two columns are called 'Status' (asset status and patch compliance)."),
    ("After cleaning", f"{K['registros']} assets x {len(columnas)} columns: 31 source fields (renamed to snake_case), "
     "flags (flag_*), events extracted from notes (evt_*), duplicate and reliability markers, and calculated columns "
     "(green header, formulas)."),
    ("Data types", "Found vs expected: dates came as Excel dates, Excel serial numbers, ISO text, dd/mm and mm/dd text, "
     "'Q4 2023' and words ('Never'); costs, RAM, storage and criticality came as text with units or words; asset tags and "
     "PO numbers sometimes as numbers (leading zeros lost). All converted to the expected type (see Data_Quality)."),
    ("Completeness", "Measured only where a field applies to the asset type AND status (e.g. a monitor has no OS; a laptop in "
     "stock has no user). Where it applies, most fields are 95-100% complete; the real gaps are encryption status (29% "
     "unknown on end-user devices), criticality and department (5.6%), PO (5.5%), warranty and IP (4.7%)."),
    ("Reliability", f"Counts reliable for {K['conf_inv']} rows, cost for {K['conf_cost']}, dates for {K['conf_fech']} "
     "(flags reliable_for_*). Assumptions are listed in Assumptions and Data_Quality."),
]
r = 3
for k, v in desc:
    celda(wdd, f"A{r}", k, f_negrita, alignment=envolver)
    celda(wdd, f"B{r}", v, alignment=envolver)
    wdd.merge_cells(f"B{r}:H{r}")
    wdd.row_dimensions[r].height = 15 * (len(v) // 140 + 1)
    r += 1
r += 1
celda(wdd, f"A{r}", "Field dictionary", f_subtitulo)
comp_idx = completitud.set_index("variable")
filas_dic = []
for c in columnas:
    base = {v: k for k, v in COLUMNAS.items()}.get(c, c)
    if base in DICC_CRUDO:
        orig, tipo, defin = DICC_CRUDO[base]
        tipos = APLICA_A.get(base)
        aplica = ("all types" if tipos and set(tipos) == set(TIPOS_ACTIVO) else ", ".join(tipos) if tipos else "optional")
        estados = ESTADOS_DONDE_APLICA.get(base)
        cuando = ", ".join(estados) if estados else ("any status" if tipos else "")
        pct = comp_idx["pct_vacio_donde_aplica"].get(base)
        pct = None if pct is None or pd.isna(pct) else pct / 100
        filas_dic.append([c, "source", orig, tipo, aplica, cuando, pct, defin])
    else:
        tipo, defin = DICC_DERIVADAS.get(c, ("", ""))
        assert defin, f"Columna sin definición en inglés: {c}"
        filas_dic.append([c, "formula" if c in FORMULAS else "added in cleaning", "", tipo, "", "", None, defin])
tabla(wdd, r + 1, 1, ["Column", "Origin", "Source header", "Type", "Applies to (asset type)", "Expected when (status)",
                      "% blank where it applies", "Definition"], filas_dic,
      anchos=[28, 16, 22, 14, 34, 30, 12, 100], formatos=[None] * 6 + ["0.0%", None])
wdd.freeze_panes = f"B{r + 2}"

# %% 10. Data_Quality ------------------------------------------------------------------------------------
wq = ws_["Data_Quality"]
celda(wq, "A1", "Data Quality — issues found and how each was treated", f_titulo)
celda(wq, "A2", "Approach: the source is never edited; every change is logged (Change_Log); values are standardized to catalogs, "
      "impossible values are blanked and flagged, deductions are flagged (cleaning_notes), only true duplicate rows are excluded.", f_nota)
issues = [
    # área, issue, campo, ejemplo, filas, por qué importa, tratamiento, dónde ver
    ("Structure", "Documentation row inside the data", "row 2", "lists of valid values, formulas as text", "1", "Would be counted as an asset", "Excluded; used only as documentation", "Change_Log"),
    ("Structure", "Completely empty rows", "all", "", "9", "Inflate counts", "Excluded", "Change_Log"),
    ("Structure", "Columns without header and data", "AF-AH", "", "3", "Noise", "Excluded", "Change_Log"),
    ("Structure", "Unnamed column with reviewer marks", "col. S", "dup?, verify, check w/ Mike", "12", "Informal, untracked review", "Kept as review_flag; compared with checks", "Clean_Data"),
    ("Structure", "Two columns named 'Status'", "H, AA", "In Use / Compliant", "all", "Ambiguous meaning", "asset_status and patch_compliance", "Data_Description"),
    ("Structure", "Inconsistent headers", "all", "'Manufacturer ' (trailing space), 'encrypted?'", "31", "Fragile references", "snake_case English names", "Data_Description"),
    ("Encoding", "Broken accents (mojibake) and stray spaces", "text fields", "MontrÃ©al, S√£o Paulo, \\xa0", "252 cells", "Same site counted as several", "Re-decoded (latin-1 / cp1252 / mac-roman), trimmed", "Change_Log"),
    ("Catalogs", "Asset type spellings", "asset_type", "laptops, LAPTOP, Lap top, Network", "55 variants", "Wrong counts by type", "Mapped to 10-type catalog; 'Network' resolved by hostname code", "Change_Log"),
    ("Catalogs", "Brand spellings; HP vs HPE", "manufacturer", "MSFT, Lenova, H.P.", "53 variants", "Wrong brand stats", "18 brands; HP on servers -> HPE", "Change_Log"),
    ("Catalogs", "Status variants", "asset_status", "Active, Deployed, EOL, RMA, e-Waste", "133", "Lifecycle counts wrong", "8-status catalog (EOL flagged)", "Assumptions"),
    ("Catalogs", "Site / country / currency / vendor / department variants", "location, country…", "NYC, NY Office; USA; $; cdw", "~1,100", "Sites and spend split or merged wrongly", "Catalogs; country and currency inferred where blank (flagged)", "Change_Log"),
    ("Catalogs", "Cost-center codes instead of department; two areas", "department", "CC-1200; 'Finance, Legal'", "34 + 16", "Mixed concepts; spend by area incomplete", "Codes moved to cost_center, department left blank to complete (flag_department_was_cc); two areas -> Shared (flagged)", "cleaning_notes"),
    ("Identity", "Asset tag in 9 formats", "asset_tag", "AST004007, 4007, ast-004007", "69", "Same asset not matched", "Rebuilt to AST-######", "Change_Log"),
    ("Identity", "Exact duplicate rows", "all", "", "6", "Double counting", "Excluded (first kept)", "Change_Log"),
    ("Identity", "Same device loaded twice (second source, tag as number, no SCCM data)", "asset_tag + serial", "AST-004072", "9 pairs", "Double counting; shows a feed without a key", "More complete row / well-formed tag + In Use with last_seen kept; 1 cost conflict to confirm", "dup_tag_type"),
    ("Identity", "Same tag on different devices", "asset_tag", "AST-004399 on AUS-DT-04399 and MTL-DT-05003", "6 tags / 12 rows", "Two assets, one identity", "Wrong row found by hostname number: 3 re-tagged from hostname (flag_tag_reassigned), 4 blanked as not deducible (flag_tag_blanked: new label needed). No tag repeats now", "cleaning_notes"),
    ("Identity", "Serial placeholders", "serial_number", "SEE STICKER, TO BE FILLED, 0, UNKNOWN", "19", "Fake repeats; device not identifiable", "Blanked; flagged (flag_serial_placeholder). Serials are unique", "Validation_Rules"),
    ("Identity", "Missing asset tags", "asset_tag", "", "12", "Asset cannot be tracked", "7 deduced from hostname (100% match rule), 5 left blank", "cleaning_notes"),
    ("Identity", "Same hostname on several live devices", "hostname", "EAG-FW-EDGE04 x3", "11 names / 23 rows", "Network conflict or stale record", "Each is a distinct device (own serial, model, MAC, IP). Flagged and out of reliable_for_inventory; per name, the most recent In Use keeps it and 10 others show evidence of a replaced device never retired (hostname_conflict_role / _evidence), 2 undetermined. 5 legitimate reuses kept", "hostname_conflict_role"),
    ("Identity", "Four hostname conventions, only one documented", "hostname", "EAGSQL33, mex-fw-core08, MANILA-LAPTOP-129", "227", "Hard to automate / reconcile", "Classified (hostname_convention); case and domain normalized", "Clean_Data"),
    ("Dates", "Mixed formats", "dates", "44497, 05/06/2024, 'May 31, 2019', '20260705'", "~230", "Ages and warranties wrong", "Parsed to dates", "Change_Log"),
    ("Dates", "Ambiguous dd/mm vs mm/dd", "purchase_date", "05/06/2024", "26", "Up to 11 months error", "mm/dd assumed (US, as in last_seen); flagged", "cleaning_notes"),
    ("Dates", "Quarter instead of date; warranty as duration", "purchase / warranty", "Q4 2023; '3 years'", "18 + 23", "Not a date", "Mid-quarter; purchase + N years; flagged", "cleaning_notes"),
    ("Dates", "Impossible purchase year", "purchase_date", "2028-09-07, 1999-10-26", "20", "Negative ages", "Year rebuilt from warranty end (same day/month, standard years by type); 5 blank; flagged", "cleaning_notes"),
    ("Dates", "Warranty end blank on physical devices with a purchase date", "warranty_end_date", "blank, 'TBD', 'Expired'", "36 + 2",
     "Warranty status unknown", "Estimated = purchase + standard years for the type (1 phone/tablet, 3 PC/monitor/printer, 5 server/network; the pattern of 716 of 843 rows); flagged (flag_warranty_estimated), excluded from reliable_for_dates. 'Lifetime' left blank. 2 'Expired' with a 2024/2026 purchase contradict the estimate: left blank and flagged (flag_warranty_expired_contradiction)", "cleaning_notes"),
    ("Dates", "Warranty ends before purchase", "warranty_end_date", "", "25", "Wrong warranty status", "Kept; excluded from reliable_for_dates", "Validation_Rules"),
    ("Dates", "Placeholder last_seen (09:00:00, before purchase)", "last_seen", "1998-10-04 09:00", "8", "False 'seen'", "Blanked; flagged", "cleaning_notes"),
    ("Cost", "Cost as text, negative or with symbols", "purchase_cost", "'$1,200', '1.234,50', -899, 'free'", "98", "Spend wrong", "Parsed per locale; negatives -> absolute (flagged)", "Change_Log"),
    ("Cost", "10 currencies, some blank or foreign", "currency", "EUR in UK; blank", "28 + 4", "Sites not comparable", "ISO codes; inferred (flagged); USD conversion in Analysis", "Assumptions"),
    ("Cost", "Book value only in 18% of rows, mixed 4/5-year life", "book_value", "", "175", "Value of estate unknown", "Recomputed for all physical assets (5 years) in book_value_calc", "Assumptions"),
    ("Cost", "Implausible amounts (outliers by type, in USD)", "purchase_cost", "laptop US$442,796 in Hungary; firewall US$2.27M", "19",
     "A few rows inflate spend by millions", "> 3x the median of the type: amount removed (kept in purchase_cost_removed), flagged, excluded from reliable_for_cost and from all statistics; probable cause noted (11 local currency loaded as USD, 8 decimal point lost). Observed: in those 11 the currency is mis-recorded (says USD, the amount only makes sense in local currency): all USD rows in Brazil, Mexico and Hungary and 3 of 4 in India. Not a site practice (most rows there use the local currency). Where the local currency is close to USD (AUD, CAD, EUR, GBP) this error cannot be detected from the amount", "cost_outlier_probable_cause"),
    ("Cost", "Amount carries a different currency than the currency column", "purchase_cost vs currency", "'3785 AUD' with currency USD", "1",
     "Unknown which currency is right", "currency kept; flagged; excluded from reliable_for_cost", "cleaning_notes"),
    ("Cost", "VMs with purchase cost", "purchase_cost", "", "20", "Spend mixes capex and subscription", "Flagged; excluded from reliable_for_cost", "Validation_Rules"),
    ("Cost", "Physical devices with cost 0 or 'free'", "purchase_cost", "Latitude 5440 = 0 with PO 4500625396", "16",
     "Pulls averages and spend down; a purchased device (PO, reseller, no lease) does not cost 0", "Blanked; flagged (flag_cost_zero_blanked), real cost to complete. VM 0 kept (no purchase)", "cleaning_notes"),
    ("Hardware", "RAM / storage in mixed units; impossible values", "ram_gb, storage_gb", "'16GB', 16384 MB, 1600 (MHz), '1TB SSD'", "~1,000", "Specs wrong", "GB numbers; disk type split; impossible -> blank (flagged)", "Change_Log"),
    ("Hardware", "Invalid MAC / IP / criticality / incidents", "mac, ip, …", "truncated MAC, 300.1.2.3, criticality 9, -1", "28+16+23+5", "Bad identifiers / scales", "Blanked and flagged", "cleaning_notes"),
    ("Hardware", "Brand does not match model", "manufacturer", "Dell 'MacBook Pro', Lenovo 'S27R650'", "7", "Wrong brand stats", "Brand taken from the model (OS confirms); flagged (flag_manufacturer_from_model)", "cleaning_notes"),
    ("Field vs type", "Values impossible for the asset type", "os, ram_gb, encrypted", "Monitor with 'Windows 10' and 16 GB RAM", "7 + 6 + 8", "Data of another device (the attached PC?)", "Blanked; flagged (flag_value_impossible_for_type)", "cleaning_notes"),
    ("Field vs type", "Values unusual for the type but possible", "last_seen, mac_address", "last_seen on monitors; MAC on monitors / VMs", "59 + 20", "Monitors inventoried via the PC, USB-C docks, virtual MACs", "Kept; flagged as contradiction (not counted as managed devices)", "Validation_Rules"),
    ("Hardware", "VM named as hypervisor (ESX)", "hostname vs asset_type", "VM 'BBLESX08' on AWS", "15", "Wrong type or copied name", "Flagged", "Validation_Rules"),
    ("People", "Names in several formats", "assigned_to", "'Lewis, Jennifer', 'Lewis Jennifer', 'jlewis', '(Terminated)'", "~130", "User not matched", "'First Last'; from email when username / reversed; leaver flagged", "cleaning_notes"),
    ("People", "Email typos and non-corporate domains", "user_email", "@@, '.con', @gmail.com", "36 + 15", "Notifications fail; personal data", "Typos fixed; domains flagged, not changed", "Clean_Data"),
    ("People", "Email of another person; device with leaver", "assigned_to vs email", "Ana White / joseph.thomas@", "16 + 14", "Accountability unclear", "Kept both, flagged", "Validation_Rules"),
    ("Lifecycle", "Retired / lost but still reporting; In Use silent > 90 days", "status vs last_seen", "", "12 + 142", "Status not trusted", "Flagged as findings", "Analysis"),
    ("Notes", "Free-text notes hide events and contradict fields", "notes", "'Battery swelling' on a server; lease ended", "~300", "Events invisible to reporting", "9 event columns extracted (evt_*); contradictions flagged", "Clean_Data"),
    ("Completeness", "Blank fields that do apply", "several", "encryption 29%, criticality 5.6%", "see table", "Audit evidence missing", "Measured where applicable; 'Not applicable' label vs blank", "below"),
]
fin = tabla(wq, 4, 1, ["Area", "Issue", "Field", "Example", "How widespread", "Why it matters", "Treatment", "Where to see it"],
            [list(i) for i in issues], anchos=[12, 40, 18, 34, 14, 30, 48, 16])
r = fin + 3
celda(wq, f"A{r}", "Completeness where the field applies (asset type + status rules; 'Not applicable' cells are not missing)", f_subtitulo)
APLICA_EN = {"todos": "all", "(no se exige)": "(not required)"}
filas_comp = [[c.variable, APLICA_EN.get(c.aplica_a_tipos, c.aplica_a_tipos), APLICA_EN.get(c.aplica_a_estados, c.aplica_a_estados),
               c.filas_esperadas, c.vacias_donde_aplica,
               None if pd.isna(c.pct_vacio_donde_aplica) else c.pct_vacio_donde_aplica / 100, c.pct_vacio_total / 100,
               c.lleno_en_tipo_que_no_aplica] for c in completitud.itertuples()]
tabla(wq, r + 1, 1, ["Field", "Required for types", "Required in status", "Rows where required", "Blank where required",
                     "% blank where required", "% blank overall", "Filled where not applicable"], filas_comp,
      formatos=[None, None, None, "#,##0", "#,##0", "0.0%", "0.0%", "#,##0"])
wq.freeze_panes = "C5"

# %% 11. Validation_Rules ----------------------------------------------------------------------------------
wv = ws_["Validation_Rules"]
celda(wv, "A1", "Validation Rules — checks that prove the clean data is fit for use", f_titulo)
celda(wv, "A2", "Part A: format and domain rules on Clean_Data (live formulas). Part B: cross-field contradiction checks run "
      "in Python (02_limpiar); 'Rows now' counts them live from the contradictions column.", f_nota)
reglas = [
    ("Records = 985 after exclusions", f'=COUNTA({rango("source_row")})', 985),
    ("asset_tag format AST-###### (non-blank)", f'=COUNTIFS({rango("asset_tag")},"<>")-COUNTIF({rango("asset_tag")},"AST-??????")', 0),
    ("asset_tag blank (5 never had one, 4 blanked: need a new label)", f'=COUNTBLANK({rango("asset_tag")})', 9),
    ("asset_tag unique (rows sharing a tag)", f'=SUMPRODUCT(({rango("asset_tag")}<>"")*(COUNTIF({rango("asset_tag")},{rango("asset_tag")})>1))', 0),
    ("serial_number unique (rows sharing a serial)", f'=SUMPRODUCT(({rango("serial_number")}<>"")*(COUNTIF({rango("serial_number")},{rango("serial_number")})>1))', 0),
    ("asset_type in catalog", f'=SUMPRODUCT(--(COUNTIF({RF_TIPO},{rango("asset_type")})=0))', 0),
    ("asset_status in catalog", f'=SUMPRODUCT(--(COUNTIF({ES_RNG},{rango("asset_status")})=0))', 0),
    ("location and country filled", f'=COUNTBLANK({rango("location")})+COUNTBLANK({rango("country")})', 0),
    ("currency is a known ISO code (non-blank)", f'=SUMPRODUCT(({rango("currency")}<>"")*(COUNTIF({FX_COD},{rango("currency")})=0))', 0),
    ("purchase_cost >= 0", f'=COUNTIF({rango("purchase_cost")},"<0")', 0),
    ("purchase_date <= cutoff", f'=COUNTIF({rango("purchase_date")},">"&{CUTOFF})', 0),
    ("purchase_date >= 2010", f'=COUNTIF({rango("purchase_date")},"<"&DATE(2010,1,1))', 0),
    ("last_seen <= cutoff", f'=COUNTIF({rango("last_seen")},">"&{CUTOFF})', 0),
    ("last_seen >= purchase_date", f'=SUMPRODUCT(ISNUMBER({rango("last_seen")})*ISNUMBER({rango("purchase_date")})*({rango("last_seen")}<{rango("purchase_date")}))', 0),
    ("warranty_end_date >= purchase_date (flagged, kept)", f'=SUMPRODUCT(ISNUMBER({rango("warranty_end_date")})*ISNUMBER({rango("purchase_date")})*({rango("warranty_end_date")}<{rango("purchase_date")}))', 25),
    ("business_criticality in 1-5", f'=COUNTIF({rango("business_criticality")},"<1")+COUNTIF({rango("business_criticality")},">5")', 0),
    ("ram_gb in 1-1024", f'=COUNTIF({rango("ram_gb")},"<1")+COUNTIF({rango("ram_gb")},">1024")', 0),
    ("open_incidents >= 0", f'=COUNTIF({rango("open_incidents")},"<0")', 0),
    ("mac_address format AA:BB:CC:DD:EE:FF", f'=COUNTIFS({rango("mac_address")},"<>",{rango("mac_address")},"<>{NO_APLICA}")-COUNTIF({rango("mac_address")},"??:??:??:??:??:??")', 0),
    ("user_email has name@domain.com", f'=COUNTIFS({rango("user_email")},"<>",{rango("user_email")},"<>{NO_APLICA}")-COUNTIF({rango("user_email")},"*?@?*.com")', 0),
    ("user_email corporate domain (flagged, kept)", f'=COUNTIF({rango("flag_email_not_corporate")},TRUE)', 15),
    # Sin el <>"" las celdas vacías cuentan como fallo: --"" no es error en SUMPRODUCT y LEN("") = 0.
    ("po_number 10 digits (non-blank)", f'=SUMPRODUCT(({rango("po_number")}<>"")*ISNUMBER(--{rango("po_number")})*(LEN({rango("po_number")})<>10))', 0),
    ("Every cleaning deduction has a note", f'=SUMPRODUCT(({rango("cleaning_to_confirm")}=TRUE)*({rango("cleaning_notes")}=""))', 0),
]
filas_v = []
for i, (txt, formula, esperado) in enumerate(reglas):
    filas_v.append([f"A{i + 1:02d}", txt, formula, esperado, None])
ini = 4
fin = tabla(wv, ini, 1, ["ID", "Rule", "Rows failing (live)", "Expected", "Result"], filas_v, anchos=[8, 60, 16, 12, 12])
for k in range(ini + 1, fin + 1):
    wv.cell(k, 5).value = f'=IF(C{k}=D{k},"PASS","CHECK")'
celda(wv, f"A{fin + 1}", "Expected > 0 means the rows are known exceptions: kept, flagged and explained in Data_Quality.", f_nota)
r = fin + 3
celda(wv, f"A{r}", "Part B — contradiction checks (cross-field)", f_subtitulo)
filas_b = []
for i, c in enumerate(chequeos.sort_values("filas", ascending=False).itertuples()):
    cid, area, regla, trat = CHEQUEOS[c.chequeo]
    filas_b.append([f"B{i + 1:02d}", area, cid, regla, c.filas, None, trat])
fin = tabla(wv, r + 1, 1, ["ID", "Area", "Check", "Rule", "Rows (Python)", "Rows now (live)", "Treatment"], filas_b,
            anchos=[8, 14, 34, 70, 12, 14, 60])
for k in range(r + 2, fin + 1):
    wv.cell(k, 6).value = f'=COUNTIF({rango("contradictions")},"*"&C{k}&"*")'
wv.column_dimensions["B"].width = 50

# %% 12b. Flags (formato largo) ----------------------------------------------------------------------------
wf = ws_["Flags"]
filas_f = [[valor_excel(v) for v in fila] for fila in banderas.itertuples(index=False)]
tabla(wf, 1, 1, list(banderas.columns), filas_f, anchos=[11, 14, 16, 36, 100])
wf.freeze_panes = "A2"
wf.auto_filter.ref = f"A1:E{len(banderas) + 1}"

# %% 12. Change_Log ------------------------------------------------------------------------------------
wl = ws_["Change_Log"]
log = registro.rename(columns={"fila_excel": "source_row", "campo": "field", "valor_original": "original_value",
                               "valor_nuevo": "new_value", "regla": "rule"})
log["rule"] = log["rule"].map(REGLAS)
log["field"] = log["field"].replace({"(fila completa)": "(whole row)"}).str.replace("columna", "column").str.replace(
    "sin nombre", "no name")
log["new_value"] = log["new_value"].replace(r"^se conserva la fila (\d+)$", r"row \1 kept", regex=True)
filas_l = [[valor_excel(v) for v in fila] for fila in log.itertuples(index=False)]
tabla(wl, 1, 1, list(log.columns), filas_l, anchos=[11, 22, 40, 40, 70])
wl.freeze_panes = "A2"
wl.auto_filter.ref = f"A1:E{len(log) + 1}"

# %% 13. Original_Data -----------------------------------------------------------------------------------
wo = ws_["Original_Data"]
crudo = openpyxl.load_workbook(ARCHIVO_CRUDO, data_only=True).worksheets[0]
for fila in crudo.iter_rows():
    for c in fila:
        if c.value is not None:
            celda_o = wo.cell(c.row, c.column, c.value)
            celda_o.font = f_normal
            # La fila de documentación trae una fórmula como TEXTO ('=IF(P3>0,…'): se escribe como texto,
            # si no openpyxl la guarda como fórmula (inválida) y Excel pide reparar el archivo.
            if isinstance(c.value, str) and c.value.startswith("="):
                celda_o.data_type = "s"
wo.freeze_panes = "A2"

# %% 14. Formato final y guardar --------------------------------------------------------------------------
for ws in wb.worksheets:
    ws.sheet_view.showGridLines = ws.title in ("Clean_Data", "Flags", "Change_Log", "Original_Data")
ws_["Cover"].sheet_properties.tabColor = AZUL
ws_["Executive_Summary"].sheet_properties.tabColor = AZUL
ws_["Dashboard"].sheet_properties.tabColor = AZUL
# Que abra en la portada (nombre y ciudad a la vista).
wb.active = 0
for ws in wb.worksheets:
    ws.sheet_view.tabSelected = ws.title == "Cover"
wb.calculation.fullCalcOnLoad = True
ARCHIVO_ENTREGA.parent.mkdir(parents=True, exist_ok=True)
wb.save(ARCHIVO_ENTREGA)
print(f"Guardado: {ARCHIVO_ENTREGA}")
