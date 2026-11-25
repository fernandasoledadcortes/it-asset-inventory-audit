# -*- coding: utf-8 -*-
"""Revisión de contradicciones en la base limpia (limpieza superficial ya hecha).

Lee elaborado/activos_ti_2026_limpio.pkl (sale de codigo/01_describir/explorar_activos_ti.py).
Cada chequeo:
  - define una regla (qué es contradictorio y por qué),
  - agrega una columna chk_<nombre> (True = la fila no cumple la regla),
  - muestra cuántas filas y ejemplos para revisar a mano.
No corrige nada todavía: lo que se decida con cada chequeo se agrega después.
Correr por celdas (# %%) o entero desde la raíz del proyecto:
    python3 codigo/02_limpiar/revisar_contradicciones.py
"""

# %% 0. Preparación ------------------------------------------------------------------------------
import re
import sys
from pathlib import Path

import pandas as pd

# Si el editor lee este .py con otra codificación (no UTF-8), los acentos del código se rompen
# y la limpieza da otros resultados: mejor frenar acá.
assert "ñ" == "\u00f1", "Este archivo se leyó con una codificación distinta de UTF-8: abrirlo como UTF-8"
import warnings

# Cada chequeo agrega una columna chk_: pandas avisa fragmentación, no afecta resultados
warnings.simplefilter("ignore", pd.errors.PerformanceWarning)

# Carpeta de este script: con __file__ si se corre entero; por celdas (VS Code / Jupyter) no
# existe __file__ y se busca la raíz del proyecto desde la carpeta actual.
try:
    AQUI = Path(__file__).resolve().parent
except NameError:
    RAIZ_PROYECTO = next(p for p in [Path.cwd(), *Path.cwd().parents] if (p / "codigo" / "00_comun").is_dir())
    AQUI = RAIZ_PROYECTO / "codigo" / "02_limpiar"
sys.path.insert(0, str(AQUI.parent / "00_comun"))
from config import BASE_LIMPIA, FECHA_CORTE  # noqa: E402
from campos import APLICA_A, CON_RESPONSABLE, GESTIONADOS, USUARIO_FINAL  # noqa: E402

pd.set_option("display.max_columns", 50)
pd.set_option("display.width", 220)
pd.set_option("display.max_colwidth", 60)

df = pd.read_pickle(BASE_LIMPIA.with_suffix(".pkl"))
CORTE = pd.Timestamp(FECHA_CORTE)
print(f"Base limpia: {len(df)} activos x {df.shape[1]} columnas | corte {CORTE.date()}")

chequeos = []  # resumen: una fila por chequeo

def revisar(nombre, mascara, columnas, regla, n_ejemplos=15):
    """Guarda el chequeo como columna chk_<nombre> y muestra los casos."""
    mascara = pd.Series(mascara, index=df.index).fillna(False).astype(bool)
    df[f"chk_{nombre}"] = mascara
    chequeos.append({"chequeo": nombre, "filas": int(mascara.sum()), "regla": regla})
    print(f"\n=== {nombre}: {mascara.sum()} filas ===\n{regla}")
    if mascara.any():
        print(df.loc[mascara, ["asset_tag", "asset_type"] + columnas].head(n_ejemplos).to_string())

BAJA = ["Retired", "Disposed", "Lost/Stolen"]
sitio_de_hostname = df["hostname"].str.extract(r"^([A-Z]{3})")[0]

# %% 1. Duplicados que quedaron marcados --------------------------------------------------------
# Ya resueltos en la limpieza (sección 21): exactos, "una fila más completa" y "mismo equipo con
# datos que chocan" (se conservó la fila con el tag bien escrito / In Use con last_seen).
revisar("dup_valor_a_confirmar", df["flag_revisar_duplicado"],
        ["hostname", "purchase_cost", "currency", "po_number", "vendor", "dup_valores_descartados"],
        "Duplicado resuelto, pero la fila excluida tenía otro valor en un campo que no es de estado: "
        "confirmar con la orden de compra (PO) / factura.")
revisar("tag_sospechoso", df["flag_tag_sospechoso"],
        ["hostname", "serial_number", "model", "location", "asset_status"],
        "asset_tag repetido en otro equipo y el número del hostname no coincide con el tag (o no se puede "
        "comparar): ese tag probablemente está mal cargado.")
revisar("dup_hostname_conflicto", df["dup_hostname_tipo"].eq("conflicto: varios vivos con el mismo nombre"),
        ["hostname", "serial_number", "model", "asset_status", "last_seen"],
        "Dos o más equipos vivos con el mismo hostname.")

# Referencia para confirmar un costo: otros equipos del mismo modelo y moneda
for _, f in df[df["flag_revisar_duplicado"]].iterrows():
    ref = df[(df["model"] == f["model"]) & (df["currency"] == f["currency"])]["purchase_cost"]
    print(f"{f['asset_tag']} {f['model']} ({f['currency']}): mismo modelo n={ref.count()}, "
          f"mediana {ref.median():.2f}, rango {ref.min():.2f}–{ref.max():.2f}")

# %% 2. Fechas imposibles o incoherentes --------------------------------------------------------
revisar("compra_posterior_al_corte", df["purchase_date"] > CORTE,
        ["purchase_date", "asset_status", "last_seen"],
        "Fecha de compra posterior al corte (2026-10-01): no puede haberse recibido todavía "
        "(salvo On Order).")
revisar("compra_muy_antigua", df["purchase_date"] < pd.Timestamp("2010-01-01"),
        ["purchase_date", "model", "asset_status"],
        "Compra antes de 2010 para modelos que salieron después (revisar modelo vs. año).")
revisar("garantia_antes_de_compra", df["warranty_end_date"] < df["purchase_date"],
        ["purchase_date", "warranty_end_date", "flag_garantia_derivada"],
        "La garantía termina antes de la compra.")
revisar("visto_antes_de_compra", df["last_seen"] < df["purchase_date"],
        ["purchase_date", "last_seen", "asset_status"],
        "Se reportó a SCCM/Intune antes de haberse comprado.")
revisar("dado_de_baja_pero_activo", df["asset_status"].isin(BAJA) & (df["last_seen"] >= CORTE - pd.Timedelta(days=30)),
        ["asset_status", "last_seen", "assigned_to"],
        "Retirado / desechado / perdido pero se reportó en los últimos 30 días: o el estado está mal, "
        "o alguien lo sigue usando.")
dias_sin_ver = (CORTE - df["last_seen"]).dt.days
revisar("en_uso_sin_reportar_90d",
        df["asset_status"].eq("In Use") & df["asset_type"].isin(GESTIONADOS) & (dias_sin_ver > 90),
        ["asset_status", "last_seen", "assigned_to", "location"],
        "In Use pero sin reportarse hace más de 90 días: ¿perdido, apagado o fuera de gestión?")

# PREGUNTAS: ¿90 días es un buen umbral? ¿On Order con fecha futura es aceptable?

# %% 3. Estado vs. persona asignada --------------------------------------------------------------
revisar("sin_uso_con_persona", df["asset_status"].isin(["In Stock", "On Order"] + BAJA) & df["assigned_to"].notna(),
        ["asset_status", "assigned_to", "last_seen"],
        "En bodega / pedido / dado de baja pero con persona asignada.")
compartido = df["notes"].str.contains("Shared device|Loaner pool", na=False)
revisar("en_uso_sin_persona",
        df["asset_status"].eq("In Use") & df["asset_type"].isin(USUARIO_FINAL) & df["assigned_to"].isna() & ~compartido,
        ["asset_status", "assigned_to", "user_email", "notes"],
        "Equipo de usuario final In Use sin persona (se excluyen los marcados como compartidos / préstamo).")
revisar("ex_empleado_con_equipo", df["flag_user_terminated"] & ~df["asset_status"].isin(BAJA),
        ["asset_status", "assigned_to", "last_seen"],
        "Equipo asignado a un ex-empleado (Terminated) que no fue dado de baja ni recuperado.")

def apellido(s):
    # sin puntuación: O'Brien = obrien
    return None if pd.isna(s) else re.sub(r"[^a-z]", "", re.split(r"[ .]", s.split("@")[0].lower())[-1])

revisar("email_no_coincide_con_persona",
        df["assigned_to"].notna() & df["user_email"].notna() &
        (df["assigned_to"].map(apellido) != df["user_email"].map(apellido)),
        ["assigned_to", "user_email"],
        "El email es de otra persona (el apellido no coincide): ¿reasignado sin actualizar?")

# department con código CC (department quedó vacío a propósito, para completar). Referencia para
# completarlo: el área que tiene la misma persona en sus otros equipos.
area_persona = df[df["department"].notna() & df["user_email"].notna()].groupby("user_email")["department"].agg(
    lambda s: ", ".join(sorted(set(s))))
cc = df[df["flag_department_cc"]]
print(f"\ndepartment vacío por código CC: {len(cc)} filas; código -> área de la misma persona en otro equipo:")
print(cc.assign(area_de_la_persona=cc["user_email"].map(area_persona)).dropna(subset=["area_de_la_persona"])
      .groupby("cost_center")["area_de_la_persona"].agg(lambda s: s.value_counts().to_dict()).to_string())

# %% 4. Datos que no corresponden al tipo de activo ---------------------------------------------
# Generaliza lo de la sección 4 de la limpieza: campo lleno en un tipo al que no aplica.
for col, aplica in APLICA_A.items():
    if aplica and col in df.columns and col != "purchase_cost":  # costo en VM: ver vm_con_costo
        mascara = df[col].notna() & ~df["asset_type"].isin(aplica)
        if mascara.any():
            revisar(f"{col}_en_tipo_que_no_aplica", mascara, [col],
                    f"{col} lleno en tipos donde no aplica: {sorted(df.loc[mascara, 'asset_type'].unique())}",
                    n_ejemplos=5)
revisar("vm_con_costo", df["asset_type"].eq("Virtual Machine") & (df["purchase_cost"] > 0),
        ["manufacturer", "purchase_cost", "currency"],
        "VM con costo de compra: las VMs se pagan por suscripción / capacidad, no se compran.")

# %% 5. Marca, modelo y sistema operativo --------------------------------------------------------
# Marca esperada de cada modelo = la marca más frecuente para ese modelo en la base.
marca_del_modelo = df.groupby("model")["manufacturer"].agg(lambda s: s.mode().iat[0])
revisar("marca_no_coincide_con_modelo",
        df["model"].notna() & (df["manufacturer"] != df["model"].map(marca_del_modelo)),
        ["manufacturer", "model"],
        "La marca no es la que corresponde al modelo (ej. HP con 'MacBook Pro').")

APPLE_OS = ["macOS", "iOS", "iPadOS"]
revisar("so_no_coincide_con_marca",
        (df["os_family"].isin(APPLE_OS) & df["manufacturer"].ne("Apple")) |
        (df["manufacturer"].eq("Apple") & df["os_family"].notna() & ~df["os_family"].isin(APPLE_OS)),
        ["manufacturer", "model", "os"],
        "Sistema de Apple en equipo no Apple, o equipo Apple con Windows / Android / Linux.")
revisar("so_no_coincide_con_tipo",
        (df["os_family"].isin(["iOS", "Android"]) & ~df["asset_type"].isin(["Mobile Phone", "Tablet"])) |
        (df["os_family"].eq("Windows Server") & df["asset_type"].isin(USUARIO_FINAL)) |
        (df["os_family"].isin(["macOS", "Windows"]) & df["asset_type"].isin(["Server", "Virtual Machine", "Mobile Phone"])),
        ["model", "os"],
        "Sistema operativo que no corresponde al tipo (ej. Android en laptop, Windows 10 en servidor).")

# %% 6. Ubicación, país y moneda ----------------------------------------------------------------
SITIO = {"AUS": "Austin", "BBL": "Blue Bell", "CHI": "Chicago", "EAG": "Eagan Data Center",
         "NYC": "New York", "BLR": "Bangalore", "BUD": "Budapest", "LON": "London",
         "MEX": "Mexico City", "MNL": "Manila", "MTL": "Montreal", "MUC": "Munich",
         "SAO": "São Paulo", "SYD": "Sydney"}
sitio_host = sitio_de_hostname.map(SITIO)
revisar("hostname_en_otro_sitio", sitio_host.notna() & (sitio_host != df["location"]),
        ["hostname", "location", "notes"],
        "El código de sitio del hostname no coincide con location (¿equipo trasladado sin renombrar?).")

TIPO_HOST = {"LT": "Laptop", "DT": "Desktop", "MB": "Mobile Phone", "TB": "Tablet", "PR": "Printer",
             "SW": "Network Switch", "FW": "Firewall"}
tipo_host = df["hostname"].str.extract(r"^[A-Z]{3}-([A-Z]{2})-")[0].map(TIPO_HOST)
revisar("hostname_de_otro_tipo", tipo_host.notna() & (tipo_host != df["asset_type"]),
        ["hostname", "model"], "El código de tipo del hostname no coincide con asset_type.")

# Hostnames legados: la ciudad y el tipo van escritos enteros (MANILA-LAPTOP-129).
legado = df["hostname_convention"].eq("legado")
CIUDAD_LEGADO = {"MANILA": "Manila", "NEWYORK": "New York", "MONTREAL": "Montreal", "BUDAPEST": "Budapest",
                 "LONDON": "London", "DATACENTER": "Eagan Data Center"}
sitio_legado = df["hostname"].str.split("-").str[0].map(CIUDAD_LEGADO)
tipo_legado_ok = [t.replace(" ", "").upper() in h for h, t in zip(df["hostname"].fillna(""), df["asset_type"])]
revisar("hostname_legado_no_coincide",
        legado & ((sitio_legado != df["location"]) | ~pd.Series(tipo_legado_ok, index=df.index)),
        ["hostname", "location"], "Hostname legado cuya ciudad o tipo escritos no coinciden con location / asset_type.")

# Cobertura: cuánto se pudo comparar y cuánto coincide "por construcción" (la limpieza dedujo
# location o asset_type desde el hostname, así que no es una confirmación independiente).
sitio_cmp = sitio_host.where(~legado, sitio_legado)
print(f"\nHostname vs. location: comparables {sitio_cmp.notna().sum()} de {df['hostname'].notna().sum()} con hostname; "
      f"{df['flag_location_desde_hostname'].sum()} coinciden por construcción (location venía HQ / DC1); "
      f"sin hostname: {df['hostname'].isna().sum()} ({df['asset_type'][df['hostname'].isna()].value_counts().to_dict()})")
print(f"Hostname vs. asset_type: con código de tipo {tipo_host.notna().sum()}; legados {legado.sum()}; "
      f"servidores/VM con código de función (APP, DC, ESX, FS, SQL, WEB), sin tipo: "
      f"{df['hostname_convention'].eq('servidor').sum()}")

# Un nombre ESX es de hipervisor: el servidor FÍSICO que aloja VMs. En una VM (más aún en la nube)
# sugiere tipo mal cargado o nombre copiado del host que la aloja.
revisar("vm_con_nombre_de_hipervisor",
        df["asset_type"].eq("Virtual Machine") & df["hostname"].str.fullmatch(r"[A-Z]{3}ESX\d+", na=False),
        ["hostname", "manufacturer", "model"],
        "VM con nombre de hipervisor (ESX = servidor físico que aloja VMs): ¿tipo mal cargado o nombre copiado?")

MONEDA_LOCAL = {"United States": "USD", "India": "INR", "United Kingdom": "GBP", "Mexico": "MXN",
                "Germany": "EUR", "Canada": "CAD", "Philippines": "PHP", "Brazil": "BRL",
                "Australia": "AUD", "Hungary": "HUF"}
moneda_local = df["country"].map(MONEDA_LOCAL)
revisar("moneda_de_otro_pais", df["currency"].notna() & ~df["currency"].isin([moneda_local, "USD"]) &
        (df["currency"] != moneda_local), ["country", "currency", "purchase_cost"],
        "Moneda que no es la local ni USD (ej. EUR en Reino Unido).")
print("\nCompras en USD fuera de EE.UU. (no es error, pero afecta la comparación):",
      (df["currency"].eq("USD") & df["country"].ne("United States")).sum())

# %% 7. Notas que contradicen los campos ---------------------------------------------------------
# Usa los eventos extraídos de notes en la limpieza (columnas evt_* y sus datos).
revisar("nota_bateria_en_equipo_sin_bateria",
        df["evt_battery_swelling"] & ~df["asset_type"].isin(["Laptop", "Mobile Phone", "Tablet"]),
        ["model", "battery_incident"], "Nota de batería hinchada en un equipo que no tiene batería.")
revisar("nota_servidor_en_equipo_de_usuario",
        df["evt_hosts_service"] & ~df["asset_type"].isin(["Server", "Virtual Machine"]),
        ["model", "hosted_service", "service_team"],
        "Nota de servicio de servidor ('Hosts …') en un equipo que no es servidor.")
revisar("leasing_vencido_en_uso", (df["lease_end_date"] < CORTE) & df["asset_status"].eq("In Use"),
        ["asset_status", "lease_end_date"], "El leasing terminó (había que devolverlo) y sigue In Use.")
revisar("nota_traslado_no_coincide", df["moved_to"].notna() & (df["moved_to"] != df["location"]),
        ["location", "moved_from", "moved_to"], "La nota dice que se trasladó a otro sitio distinto de location.")
revisar("auditado_pero_perdido", df["evt_physical_audit"] & df["asset_status"].isin(["Missing", "Lost/Stolen"]),
        ["asset_status", "audit_month"], "Verificado en auditoría física pero marcado como perdido.")
revisar("auditado_antes_de_compra", df["audit_month"] < df["purchase_date"].dt.to_period("M").dt.to_timestamp(),
        ["purchase_date", "audit_month"], "Verificado en auditoría antes de haberse comprado.")
revisar("dueño_anterior_igual_al_actual", df["previous_owner"].notna() & (df["previous_owner"] == df["assigned_to"]),
        ["assigned_to", "previous_owner"], "La nota dice que la persona asignada es el dueño anterior.")
revisar("compartido_o_prestamo_con_persona", (df["evt_shared"] | df["evt_loaner"]) & df["assigned_to"].notna(),
        ["assigned_to", "notes"], "Equipo compartido / de préstamo con una persona asignada.")
revisar("refresh_antes_de_compra", df["refresh_fy"] < df["purchase_date"].dt.year - 1,
        ["purchase_date", "refresh_fy"], "Entregado en un refresh de un año fiscal anterior a su compra.")

# %% 8. Valor en libros --------------------------------------------------------------------------
# La fórmula del crudo: costo × (1 − días desde compra / vida útil), mínimo 0, al 2026-10-01.
# La fila 2 documenta 4 años (1460 días) y las fórmulas usan 5 años (1825 días).
dias = (CORTE - df["purchase_date"]).dt.days
for anios in (4, 5):
    df[f"book_value_{anios}a"] = (df["purchase_cost"] * (1 - dias / (365 * anios))).clip(lower=0).round(2)
con_valor = df["book_value"].notna()
print("\nbook_value informado en", con_valor.sum(), "filas. Coincide (±1) con vida útil de:")
for anios in (4, 5):
    print(f"  {anios} años:", ((df["book_value"] - df[f"book_value_{anios}a"]).abs() <= 1)[con_valor].sum())
revisar("book_value_no_coincide",
        con_valor & ((df["book_value"] - df["book_value_5a"]).abs() > 1) & ((df["book_value"] - df["book_value_4a"]).abs() > 1),
        ["purchase_cost", "purchase_date", "book_value", "book_value_4a", "book_value_5a"],
        "book_value informado no coincide con la fórmula ni a 4 ni a 5 años.")

# Decidido: se recalcula para todos los físicos con costo y fecha a 5 años (lo que usan 149 de 175
# fórmulas del crudo; la fila 2 decía 4). book_value original se conserva para comparar.
df["book_value_calc"] = df["book_value_5a"].where(df["asset_type"].isin(APLICA_A["book_value"]))
df = df.drop(columns=["book_value_4a", "book_value_5a"])
print(f"book_value_calc: {df['book_value_calc'].notna().sum()} físicos con valor (antes {con_valor.sum()})")

# %% 9. Costos -----------------------------------------------------------------------------------
revisar("fisico_con_costo_cero", df["asset_type"].ne("Virtual Machine") & df["purchase_cost"].eq(0),
        ["model", "purchase_cost", "vendor", "notes"],
        "Equipo físico con costo 0 (¿leasing, donación, dato faltante?).")
# Debería dar 0: en 01_describir los 0 de físicos pasaron a vacío (flag_costo_cero).
# Para comparar costos entre sitios hace falta pasar a USD: se hace en el análisis.

# %% 10. Filas que alguien marcó en review_flag --------------------------------------------------
chk = [c for c in df.columns if c.startswith("chk_")]
marcadas = df[df["review_flag"].notna()]
print("\nreview_flag vs. chequeos que fallan en esa fila:")
print(marcadas[["review_flag", "asset_tag", "asset_type"]].assign(
    chequeos=marcadas[chk].apply(lambda f: [c[4:] for c in chk if f[c]], axis=1)).to_string())

# %% 11. Resumen ---------------------------------------------------------------------------------
resumen = pd.DataFrame(chequeos).sort_values("filas", ascending=False)
print("\n", resumen.to_string(index=False))
print(f"\nActivos con al menos una contradicción: {df[chk].any(axis=1).sum()} de {len(df)}")

# %% 11b. Completitud real: vacíos solo donde el dato corresponde ------------------------------
# Un vacío es faltante solo si el campo aplica al TIPO (campos.APLICA_A) y al ESTADO
# (campos.ESTADOS_DONDE_APLICA). Ej.: monitor sin SO o laptop en bodega sin persona = esperable.
# También se cuenta lo inverso: dato lleno en un tipo donde no aplica.
from campos import ESTADOS_DONDE_APLICA, NO_EXIGIDO_EN, TIPOS_ACTIVO  # noqa: E402

ESTADOS = sorted(df["asset_status"].dropna().unique())
filas = []
for col, tipos in APLICA_A.items():
    if not tipos or col not in df:
        continue  # review_flag y notes: opcionales
    estados = ESTADOS_DONDE_APLICA.get(col, ESTADOS)
    exigidos = [t for t in tipos if t not in NO_EXIGIDO_EN.get(col, [])]
    esperado = df["asset_type"].isin(exigidos) & df["asset_status"].isin(estados)
    vacio = df[col].isna()
    filas.append({
        "variable": col,
        "aplica_a_tipos": "todos" if set(exigidos) == set(TIPOS_ACTIVO) else ", ".join(exigidos) or "(no se exige)",
        "aplica_a_estados": "todos" if set(estados) >= set(ESTADOS) else ", ".join(estados),
        "filas_esperadas": int(esperado.sum()),
        "vacias_donde_aplica": int((esperado & vacio).sum()),
        "pct_vacio_donde_aplica": round((esperado & vacio).sum() / esperado.sum() * 100, 1) if esperado.any() else None,
        "pct_vacio_total": round(vacio.mean() * 100, 1),
        "vacias_esperables": int((~esperado & vacio).sum()),
        "lleno_en_tipo_que_no_aplica": int((~df["asset_type"].isin(tipos) & ~vacio).sum()),
    })
completitud = pd.DataFrame(filas).sort_values("pct_vacio_donde_aplica", ascending=False)
print("\n=== Completitud donde el dato corresponde (tipo + estado) ===")
print(completitud.drop(columns=["aplica_a_tipos", "aplica_a_estados"]).to_string(index=False))

# Detalle por tipo de activo: % vacío donde aplica (para ver DÓNDE faltan los datos)
detalle = {}
for col in completitud["variable"]:
    exigidos = [t for t in APLICA_A[col] if t not in NO_EXIGIDO_EN.get(col, [])]
    esperado = df["asset_type"].isin(exigidos) & df["asset_status"].isin(ESTADOS_DONDE_APLICA.get(col, ESTADOS))
    detalle[col] = df[esperado].groupby("asset_type")[col].apply(lambda s: round(s.isna().mean() * 100))
completitud_por_tipo = pd.DataFrame(detalle).T
print("\n% vacío donde aplica, por tipo de activo (vacío = no aplica):")
print(completitud_por_tipo.to_string())

# %% 12. Contradicciones por fila y guardar ------------------------------------------------------
# Igual que notas_limpieza: una columna de texto con los chequeos que fallan en la fila, para
# filtrar en Excel. revisar = la fila tiene una decisión de limpieza para corroborar o una contradicción.
df = df.copy()  # desfragmenta (se agregaron muchas columnas chk_ una a una)
df["contradicciones"] = df[chk].apply(lambda f: "; ".join(c[4:] for c in chk if f[c]) or None, axis=1)
df["n_contradicciones"] = df[chk].sum(axis=1)
df["revisar"] = df["revisar_limpieza"] | df["n_contradicciones"].gt(0)
print(f"\nFilas para revisar (decisión de limpieza o contradicción): {df['revisar'].sum()} de {len(df)}")

# Confiabilidad por tema: para estadísticas se saca solo lo que ensucia ESE cálculo (sacar todo
# 'revisar' dejaría 40% de la base y sesgada). Hallazgos como en_uso_sin_reportar_90d no restan.
# En conflicto de hostname se usa la bandera de limpieza (solo los vivos), no el chk (incluye al Retired).
NO_CONFIABLE = {
    # Conservador (Fernanda): también fuera los tags que salen de una deducción (a confirmar con el sticker)
    "confiable_inventario": ["flag_tag_sospechoso", "flag_hostname_conflicto", "chk_dado_de_baja_pero_activo",
                             "chk_compra_posterior_al_corte", "flag_revisar_duplicado",
                             "flag_tag_desde_hostname", "flag_tag_reasignado", "flag_tag_vaciado"],
    "confiable_costos": ["flag_costo_negativo", "chk_fisico_con_costo_cero", "chk_vm_con_costo",
                         "chk_moneda_de_otro_pais", "flag_moneda_inferida", "flag_revisar_duplicado",
                         "chk_book_value_no_coincide", "flag_moneda_monto_distinta", "flag_costo_atipico", "flag_costo_cero"],
    "confiable_fechas": ["flag_fecha_ambigua", "flag_fecha_aproximada", "chk_garantia_antes_de_compra",
                         "chk_visto_antes_de_compra", "chk_auditado_antes_de_compra", "chk_compra_muy_antigua",
                         "chk_compra_posterior_al_corte", "flag_compra_reconstruida", "flag_last_seen_invalido",
                         "flag_garantia_estimada", "flag_garantia_contradictoria"],
}
for col, motivos in NO_CONFIABLE.items():
    df[col] = ~df[motivos].any(axis=1)
    print(f"{col}: {df[col].sum()} confiables, {(~df[col]).sum()} fuera")

from config import ELABORADO  # noqa: E402
salida = ELABORADO / "activos_ti_2026_revision"
df.drop(columns=chk).to_pickle(salida.with_suffix(".pkl"))
df.drop(columns=chk).to_csv(salida.with_suffix(".csv"), encoding="utf-8-sig")
resumen.to_csv(ELABORADO / "activos_ti_2026_chequeos.csv", index=False, encoding="utf-8-sig")
completitud.to_csv(ELABORADO / "activos_ti_2026_completitud.csv", index=False, encoding="utf-8-sig")
completitud_por_tipo.to_csv(ELABORADO / "activos_ti_2026_completitud_por_tipo.csv", encoding="utf-8-sig")
print(f"Guardado: {salida.name}.pkl/.csv y activos_ti_2026_chequeos.csv (regla de cada chequeo)")
