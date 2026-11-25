# -*- coding: utf-8 -*-
"""Exploración (y de a poco, limpieza) de la base cruda DataChallenge.xlsx.

Script que vamos armando juntas. Está dividido en celdas `# %%`: en VS Code / Positron se
pueden correr una por una (Shift+Enter) o el archivo entero desde la raíz del proyecto:
    python3 codigo/01_describir/explorar_activos_ti.py

El crudo nunca se modifica. Todo lo que se excluye o cambia queda en `registro` (log de cambios),
que después se vuelve la pestaña Change_Log de la entrega.
Definiciones de cada campo: codigo/00_comun/campos.py.
Supuestos y hallazgos: docs/notas_supuestos.md.
"""

# %% 0. Preparación ------------------------------------------------------------------------------
import re
import sys
import unicodedata
from collections import Counter
from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd

# Si el editor lee este .py con otra codificación (no UTF-8), los acentos del código se rompen
# y la limpieza da otros resultados: mejor frenar acá.
assert "ñ" == "\u00f1", "Este archivo se leyó con una codificación distinta de UTF-8: abrirlo como UTF-8"

# Carpeta de este script: con __file__ si se corre entero; por celdas (VS Code / Jupyter) no
# existe __file__ y se busca la raíz del proyecto desde la carpeta actual.
try:
    AQUI = Path(__file__).resolve().parent
except NameError:
    RAIZ_PROYECTO = next(p for p in [Path.cwd(), *Path.cwd().parents] if (p / "codigo" / "00_comun").is_dir())
    AQUI = RAIZ_PROYECTO / "codigo" / "01_describir"
sys.path.insert(0, str(AQUI.parent / "00_comun"))
from config import ARCHIVO_CRUDO, FECHA_CORTE, FX  # noqa: E402
from campos import APLICA_A, DEFINICION, RENOMBRAR, TIPO_ESPERADO, TIPOS_ACTIVO  # noqa: E402

pd.set_option("display.max_columns", 50)
pd.set_option("display.width", 200)
pd.set_option("display.max_colwidth", 60)

# Log de cambios: una fila por cada exclusión / cambio, con la regla que lo justifica
registro = []

def registrar(fila_excel, campo, valor_original, valor_nuevo, regla):
    registro.append({"fila_excel": fila_excel, "campo": campo, "valor_original": valor_original,
                     "valor_nuevo": valor_nuevo, "regla": regla})

def es_vacio(v):
    """Vacío real o 'falso lleno' (texto que es solo espacios)."""
    return pd.isna(v) or (isinstance(v, str) and v.strip() == "")

# %% 1. Carga y grano ----------------------------------------------------------------------------
# dtype=object: que pandas NO convierta tipos, para ver los datos tal como vienen.
# Las fórmulas se leen con su último valor calculado (o vacío si nunca se calculó).
crudo = pd.read_excel(ARCHIVO_CRUDO, sheet_name=0, header=0, dtype=object)
crudo.index = crudo.index + 2  # índice = número de fila en Excel, para poder ir a revisar a mano

# GRANO: una fila = un activo de TI registrado. No son solo equipos entregados a empleados
# (laptops, celulares): también hay infraestructura (servidores, switches, firewalls) y máquinas
# virtuales, que no tienen persona asignada.

# Filas que no son activos. No se asume que sea solo la fila 2: se busca en toda la base.
# Una fila de documentación tiene varias celdas con listas (saltos de línea) o fórmulas como texto.
es_doc = crudo.apply(
    lambda fila: sum(isinstance(v, str) and ("\n" in v or v.startswith("=")) for v in fila) >= 2,
    axis=1,
)
filas_doc = crudo.index[es_doc].tolist()
print("Filas de documentación (no son activos):", filas_doc)

metadatos = crudo.loc[filas_doc[0]] if filas_doc else None  # se guarda solo como referencia
for f in filas_doc:
    registrar(f, "(fila completa)", None, None, "Excluida: fila de documentación, no es un activo")
df = crudo.drop(index=filas_doc).copy()

print(f"Filas leídas (sin encabezado): {len(crudo)}  ->  candidatas a activo: {len(df)}")

# %% 2. Estructura: columnas ---------------------------------------------------------------------
# La fila 2 no está completa ni es confiable: no se usa como fuente de verdad, solo como pista.
# Reglas:
#   - columna vacía (sin datos)            -> se elimina
#   - columna idéntica a otra              -> se elimina (es repetida)
#   - nombre repetido con contenido distinto -> se renombra (Status vs Status.1)
#   - columna con datos y sin nombre       -> se le pone nombre (Unnamed: 18 -> review_flag)
#   - todos los nombres a snake_case en inglés (ver campos.py)

# Encabezados reales del Excel (pandas renombra duplicados y omite columnas sin nombre ni datos)
ws = openpyxl.load_workbook(ARCHIVO_CRUDO, read_only=True).worksheets[0]
encabezados = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
print(f"Columnas en la hoja: {len(encabezados)} | leídas por pandas: {crudo.shape[1]}")
repetidos = [n for n, k in Counter(encabezados).items() if k > 1 and n is not None]
print("Nombres repetidos en el encabezado:", repetidos)
for i in range(crudo.shape[1], len(encabezados)):
    registrar(None, f"columna {i + 1} (sin nombre)", None, None,
              "Excluida: columna sin encabezado y sin datos")

vacias = [c for c in df.columns if df[c].map(es_vacio).all()]
for c in vacias:
    registrar(None, c, None, None, "Excluida: columna sin datos")

identicas = []
cols = [c for c in df.columns if c not in vacias]
for i, a in enumerate(cols):
    for b in cols[i + 1:]:
        if b not in identicas and df[a].astype(str).equals(df[b].astype(str)):
            identicas.append(b)
            registrar(None, b, None, a, f"Excluida: columna idéntica a '{a}'")
print("Columnas vacías:", vacias, "| idénticas a otra:", identicas)
df = df.drop(columns=vacias + identicas)

sin_definir = [c for c in df.columns if c not in RENOMBRAR]
assert not sin_definir, f"Columnas nuevas sin definir en campos.py: {sin_definir}"

for orig in df.columns:
    if orig != RENOMBRAR[orig]:
        registrar(None, orig, orig, RENOMBRAR[orig], "Renombrada a snake_case")
df = df.rename(columns=RENOMBRAR)

print("\nVariables:")
for orig, nuevo in RENOMBRAR.items():
    if nuevo in df.columns:
        print(f"  {orig!r:24} -> {nuevo}")

# %% 3. Tipos de dato por columna ----------------------------------------------------------------
# Compara el tipo esperado (campos.py) contra los tipos que aparecen de verdad.
# Una columna "limpia" tiene un solo tipo y es el esperado.
def cumple_tipo(v, tipo):
    if tipo in ("texto", "categoria"):
        return isinstance(v, str)
    if tipo == "entero":
        return isinstance(v, int) or (isinstance(v, float) and v.is_integer())
    if tipo == "numero":
        return isinstance(v, (int, float))
    if tipo in ("fecha", "fecha_hora"):
        return isinstance(v, datetime)
    return True

filas = []
for col in df.columns:
    llenos = df[col][~df[col].map(es_vacio)]
    no_cumple = llenos[~llenos.map(lambda v: cumple_tipo(v, TIPO_ESPERADO[col]))]
    filas.append({
        "variable": col,
        "esperado": TIPO_ESPERADO[col],
        "tipos_encontrados": dict(Counter(type(v).__name__ for v in llenos)),
        "no_cumple": len(no_cumple),
        "ejemplos": list(dict.fromkeys(no_cumple.map(repr)))[:5],
    })
tipos = pd.DataFrame(filas).set_index("variable")
print(tipos.to_string())

# review_flag (antes Unnamed: 18): marcas de revisión, se usan después para cruzar
print("\nreview_flag:\n", df["review_flag"].dropna().to_string())

# Lo que ya sabemos por variable (se trata en las secciones siguientes):
#   - asset_tag, hostname: texto. serial_number: texto aunque venga como número (IMEI, ceros).
#   - purchase_date, warranty_end_date, last_seen: fechas mezcladas con números seriales de Excel
#     y textos en distintos formatos -> homogeneizar a fecha.
#   - purchase_cost: textos con separadores y símbolos -> número. book_value depende de cost,
#     purchase_date y la vida útil: se recalcula después con una sola regla.
#   - business_criticality: entero 1–5; los textos ('Low', 'High'…) se mapean con la escala de la fila 2.
#   - patch_compliance (Status.1) es distinto de asset_status (Status).

# %% 4. Completitud ------------------------------------------------------------------------------
# 4a. Primero las filas: vacías completas se excluyen; las que tienen muy pocos datos se revisan.
llenos_por_fila = (~df.map(es_vacio)).sum(axis=1)
print("Campos llenos por fila (distribución):")
print(llenos_por_fila.value_counts().sort_index().to_string())

filas_vacias = llenos_por_fila.index[llenos_por_fila == 0].tolist()
for f in filas_vacias:
    registrar(f, "(fila completa)", None, None, "Excluida: fila completamente vacía")
df = df.drop(index=filas_vacias)
llenos_por_fila = llenos_por_fila.drop(index=filas_vacias)
print(f"\nFilas vacías excluidas: {len(filas_vacias)} -> {filas_vacias}")

UMBRAL_POCOS = df.shape[1] // 2  # menos de la mitad de los campos llenos
pocos = df[llenos_por_fila < UMBRAL_POCOS]
print(f"Filas con menos de {UMBRAL_POCOS} campos llenos (revisar a mano): {len(pocos)}")
if len(pocos):
    print(pocos.to_string())

print(f"\nActivos: {len(df)}")

# 4b. Después las columnas. Un vacío solo es problema donde el campo aplica (campos.APLICA_A):
# una VM sin MAC o un monitor sin SO es esperable; una laptop sin SO, no.
# Hasta estandarizar asset_type (sección 5), "donde aplica" se calcula solo sobre las filas cuyo
# tipo ya viene bien escrito. También se cuenta lo contrario: dato lleno donde NO debería haber.
def tabla_completitud(df):
    """% vacío total y donde aplica, usando solo filas con asset_type del catálogo."""
    bien_tipadas = df[df["asset_type"].isin(TIPOS_ACTIVO)]
    vacio = df.map(es_vacio)
    filas = []
    for col in [c for c in df.columns if c in APLICA_A]:
        aplica = APLICA_A[col]
        fila = {"variable": col, "pct_vacio_total": round(vacio[col].mean() * 100, 1)}
        if aplica:
            dentro = bien_tipadas["asset_type"].isin(aplica)
            fila["pct_vacio_donde_aplica"] = round(vacio.loc[bien_tipadas.index[dentro], col].mean() * 100, 1)
            fila["lleno_donde_no_aplica"] = int((~vacio.loc[bien_tipadas.index[~dentro], col]).sum())
        filas.append(fila)
    print(f"(donde aplica: calculado sobre {len(bien_tipadas)} de {len(df)} filas con asset_type del catálogo)")
    return pd.DataFrame(filas).set_index("variable").sort_values("pct_vacio_donde_aplica", ascending=False)

print(tabla_completitud(df).to_string())

# Supuestos para leer esta tabla:
#   - encrypted vacío = desconocido (no 'no'). Solo se espera en laptops y desktops.
#   - patch_compliance y last_seen solo aplican a equipos gestionados por SCCM/Intune (con SO).
#   - assigned_to / user_email solo en equipos de usuario final; en bodega (In Stock) es normal que falte.
#   - ip_address: se espera en equipos de red fija; en laptops y celulares puede faltar (DHCP).
#   - country vacío se puede deducir de location (sección 5).
#   - po_number debería estar en todo lo comprado (no en VMs).

# Para ver un campo por tipo de activo:
# df.groupby("asset_type")["user_email"].apply(lambda s: s.map(es_vacio).mean()).round(2)

print(f"\nRegistro de cambios hasta aquí: {len(registro)} entradas")
print(pd.DataFrame(registro)["regla"].value_counts().to_string())

# %% 5. Homogeneización general de texto ---------------------------------------------------------
# Criterio: NO se pasa todo a minúsculas ni se quitan acentos en los datos. Eso se perdería
# ("Ciudad de México" es el nombre correcto) y en nombres, notas o modelos no hace falta.
# En cambio:
#   1. Limpieza que sí se aplica a todo el texto: espacios (incluido el invisible \xa0) y
#      caracteres mal codificados (mojibake: "MÃ©xico" -> "México").
#   2. Para comparar variantes se usa una CLAVE auxiliar (minúsculas, sin acentos ni signos) que
#      no se guarda: solo sirve para buscar en el diccionario de cada variable, y el valor final
#      sale del catálogo, bien escrito ("laptops", "LAPTOP " -> clave "laptops"/"laptop" -> "Laptop").
def arreglar_mojibake(s):
    """UTF-8 leído con otra codificación y vuelto a guardar. Si no se puede, se deja igual.
    Latin-1 / Windows: 'MÃ©xico' -> 'México'.  Mac: 'S√£o Paulo' -> 'São Paulo'."""
    if not re.search("[\u00c3\u00c2\u221a]|\u00e2\u20ac", s):  # Ã Â √ â€ (escapados: no dependen de cómo se lea el .py)
        return s
    for codificacion in ("latin-1", "cp1252", "mac_roman"):
        try:
            return s.encode(codificacion).decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            continue
    return s

def limpiar_texto(v):
    if not isinstance(v, str):
        return v
    s = arreglar_mojibake(v.replace("\xa0", " "))
    s = re.sub(r"\s+", " ", s).strip()
    return s if s else None  # texto solo con espacios -> vacío

def clave(v):
    """Minúsculas, sin acentos, sin signos ni espacios dobles. Solo para comparar."""
    s = unicodedata.normalize("NFKD", str(v)).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", " ", s.lower()).strip()

def aplicar(df, col, nuevos, regla):
    """Reemplaza la columna y registra cada celda que cambió."""
    antes = df[col]
    cambio = ~((antes == nuevos) | (antes.isna() & nuevos.isna()))
    for fila in antes.index[cambio]:
        registrar(fila, col, antes[fila], nuevos[fila], regla)
    df[col] = nuevos
    return int(cambio.sum())

for col in df.columns:
    n = aplicar(df, col, df[col].map(limpiar_texto), "Texto: espacios / codificación")
    if n:
        print(f"{col:22} {n:4} celdas")

def ver_valores(col, n=60):
    vc = df[col].map(repr).value_counts(dropna=False)  # repr para ver mayúsculas y vacíos
    print(f"\n=== {col}  ({vc.size} valores distintos) ===")
    print(vc.head(n).to_string())

ORIGINAL = {limpio: orig for orig, limpio in RENOMBRAR.items()}

def homogeneizar(col, mapa, resolver=None, via_resolver="otra columna"):
    """Lleva cada variante al valor del catálogo según su clave.

    mapa: clave -> valor final. Si el valor es None, la variante es ambigua y se toma de
    `resolver` (una Serie con el valor correcto para esas filas). Si aparece una clave sin regla,
    el script se detiene para revisarla. Imprime qué valor crudo quedó asignado a qué.
    """
    llenos = df[col].notna()
    claves = df.loc[llenos, col].map(clave)
    sin_mapa = sorted(set(claves) - set(mapa))
    assert not sin_mapa, f"{col} sin regla: {sin_mapa}"

    nuevo = claves.map(mapa).reindex(df.index)
    ambiguos = llenos & nuevo.isna()
    if ambiguos.any():
        assert resolver is not None and resolver[ambiguos].notna().all(), \
            f"{col}: hay valores ambiguos sin resolver"
        nuevo = nuevo.where(~ambiguos, resolver)
    aplicar(df, col, nuevo, f"{col}: variante -> catálogo")

    asignacion = (pd.DataFrame({"valor_crudo": crudo.loc[df.index, ORIGINAL[col]].map(repr),
                                "asignado": df[col],
                                "via": ambiguos.map({True: via_resolver, False: "diccionario"})})
                    [llenos].value_counts().rename("n").reset_index()
                    .sort_values(["asignado", "n"], ascending=[True, False]))
    print(asignacion.to_string(index=False))
    ver_valores(col)

# %% 6. asset_type -------------------------------------------------------------------------------
# Catálogo: los 10 tipos de campos.TIPOS_ACTIVO. Cada variante (por su clave) -> valor del catálogo.
# Decisiones:
#   - Notebook, LT, Laptop/Tablet -> Laptop (todos tienen hostname -LT- y modelos de laptop; el
#     Surface Pro 9 es un 2-en-1 registrado como laptop).
#   - Workstation, Desktop PC -> Desktop (hostname -DT-).
#   - Screen, Display -> Monitor. MFP, Printer/Scanner -> Printer. Security Appliance, FW -> Firewall.
#   - Physical Server, SRV -> Server. Virtual Server, VM -> Virtual Machine.
#   - Network / Network Device son ambiguos (switch o firewall) -> se resuelven con el código de
#     tipo del hostname (SITE-SW-… / SITE-FW-…).
ASSET_TYPE = {
    "laptop": "Laptop", "laptops": "Laptop", "notebook": "Laptop", "lt": "Laptop",
    "laptop tablet": "Laptop",
    "desktop": "Desktop", "desktop pc": "Desktop", "workstation": "Desktop",
    "monitor": "Monitor", "monitors": "Monitor", "screen": "Monitor", "display": "Monitor",
    "server": "Server", "physical server": "Server", "server physical": "Server", "srv": "Server",
    "virtual machine": "Virtual Machine", "virtual server": "Virtual Machine", "vm": "Virtual Machine",
    "mobile phone": "Mobile Phone", "mobile": "Mobile Phone", "phone": "Mobile Phone",
    "cell phone": "Mobile Phone", "smartphone": "Mobile Phone", "iphone": "Mobile Phone",
    "tablet": "Tablet", "ipad": "Tablet",
    "printer": "Printer", "printer scanner": "Printer", "mfp": "Printer",
    "network switch": "Network Switch", "switch": "Network Switch",
    "firewall": "Firewall", "fw": "Firewall", "security appliance": "Firewall",
    "network": None, "network device": None,  # ambiguos: se resuelven por hostname
}
CODIGO_HOSTNAME = {"LT": "Laptop", "DT": "Desktop", "MB": "Mobile Phone", "TB": "Tablet",
                   "PR": "Printer", "SW": "Network Switch", "FW": "Firewall"}

ver_valores("asset_type")
codigo_host = df["hostname"].str.upper().str.extract(r"^[A-Z]{3}-([A-Z]{2})-")[0].map(CODIGO_HOSTNAME)
homogeneizar("asset_type", ASSET_TYPE, resolver=codigo_host, via_resolver="hostname")

# Ahora la completitud "donde aplica" ya usa todas las filas
print(tabla_completitud(df).to_string())

# %% 7. manufacturer -----------------------------------------------------------------------------
# Catálogo: nombre comercial corto de la marca (sin "Inc.", "Corp.", "Group", "Electronics").
# Decisiones:
#   - HP y HPE son empresas distintas desde 2015: HP Inc. hace PCs, monitores e impresoras;
#     Hewlett Packard Enterprise (HPE) hace servidores (ProLiant). Las variantes genéricas
#     ("HP", "Hewlett-Packard", "H.P.") se resuelven por tipo: Server -> HPE, el resto -> HP.
#   - Aruba (incluye "HPE Aruba") se deja como Aruba: es la marca de los switches.
#   - Microsoft (Surface, Hyper-V) se separa de Microsoft Azure (nube): son cosas distintas para el
#     análisis (equipo comprado vs. suscripción). MSFT -> Microsoft.
#   - Amazon -> Amazon Web Services. Broadcom/VMware -> VMware. Lenova (error de tipeo) -> Lenovo.
#   - Marca que no corresponde al modelo (Dell 'MacBook Pro', Lenovo 'S27R650'): se corrige con la
#     marca más frecuente de ese modelo en la base (el modelo es más específico y el SO lo confirma)
#     + flag_manufacturer_from_model (decidido en la revisión de contradicciones).
MANUFACTURER = {
    "apple": "Apple", "apple inc": "Apple",
    "lenovo": "Lenovo", "lenova": "Lenovo", "lenovo group": "Lenovo",
    "dell": "Dell", "dell inc": "Dell", "dell technologies": "Dell",
    "hp": None, "h p": None, "hewlett packard": None,           # ambiguos: HP o HPE según tipo
    "hp inc": "HP",
    "hpe": "HPE", "hp enterprise": "HPE", "hewlett packard enterprise": "HPE",
    "microsoft": "Microsoft", "microsoft corp": "Microsoft", "msft": "Microsoft",
    "microsoft azure": "Microsoft Azure",
    "amazon": "Amazon Web Services",
    "vmware": "VMware", "vmware inc": "VMware", "broadcom vmware": "VMware",
    "samsung": "Samsung", "samsung electronics": "Samsung",
    "lg": "LG", "lg electronics": "LG",
    "cisco": "Cisco", "cisco systems": "Cisco", "cisco systems inc": "Cisco",
    "aruba": "Aruba", "aruba networks": "Aruba", "hpe aruba": "Aruba",
    "juniper": "Juniper", "fortinet": "Fortinet",
    "palo alto": "Palo Alto Networks", "palo alto networks": "Palo Alto Networks",
    "canon": "Canon", "xerox": "Xerox",
}

ver_valores("manufacturer")
hp_por_tipo = df["asset_type"].map(lambda t: "HPE" if t == "Server" else "HP")
homogeneizar("manufacturer", MANUFACTURER, resolver=hp_por_tipo, via_resolver="tipo de activo")

modelo = df["model"].map(clave)
marcas_modelo = df.groupby(modelo)["manufacturer"].agg(lambda s: s.value_counts(normalize=True))
marca_modelo = df.groupby(modelo)["manufacturer"].agg(lambda s: s.mode().iat[0] if s.notna().any() else None)
peso = df.groupby(modelo)["manufacturer"].agg(lambda s: s.value_counts(normalize=True).max() if s.notna().any() else 0)
esperada = modelo.map(marca_modelo)
df["flag_manufacturer_from_model"] = (esperada.notna() & df["manufacturer"].ne(esperada)
                                      & modelo.map(peso).ge(0.8) & modelo.map(df.groupby(modelo).size()).ge(3))
print(df.loc[df["flag_manufacturer_from_model"], ["model", "manufacturer"]].assign(nueva=esperada).to_string())
aplicar(df, "manufacturer", df["manufacturer"].where(~df["flag_manufacturer_from_model"], esperada),
        "manufacturer: no corresponde al modelo -> marca del modelo")
print(pd.crosstab(df["manufacturer"], df["asset_type"]).to_string())

# %% 8. asset_tag --------------------------------------------------------------------------------
# Formato correcto: AST-###### (6 dígitos). Variantes encontradas: 4059 (número), 'ast-004143',
# 'AST 004246', 'AST004007', 'AST-4081', '004438'. Todas tienen el número completo, así que se
# reconstruye: prefijo AST- + número con ceros a la izquierda hasta 6 dígitos.
# Supuesto: los números cortos ('AST-4081', 4059) son el mismo número sin ceros (los tags van de
# 004003 a 005947).
# Los vacíos NO se inventan: aunque el número del hostname suele coincidir con el del tag, no
# coincide siempre (ver abajo), así que deducirlo sería adivinar. Quedan vacíos y se marcan.
def patron(v):
    """Forma del valor: letras -> A/a, dígitos -> 9. 'AST-004003' -> 'AAA-999999'."""
    if pd.isna(v):
        return "(vacío)"
    s = re.sub(r"\d", "9", str(v))
    return re.sub("[a-z]", "a", re.sub("[A-Z]", "A", s)) + f"  [{type(v).__name__}]"

print(crudo.loc[df.index, "Asset Tag"].map(patron).value_counts().to_string())

def normalizar_tag(v):
    if pd.isna(v):
        return None
    m = re.fullmatch(r"(?i)\s*(?:AST)?[\s-]*(\d{1,6})\s*", str(v))
    return f"AST-{int(m.group(1)):06d}" if m else v

aplicar(df, "asset_tag", df["asset_tag"].map(normalizar_tag), "asset_tag: formato AST-######")
llenos = df["asset_tag"].dropna()
assert llenos.str.fullmatch(r"AST-\d{6}").all(), llenos[~llenos.str.fullmatch(r"AST-\d{6}")]

print(f"\nasset_tag vacíos: {df['asset_tag'].isna().sum()}")
print(f"asset_tag repetidos (mismo tag en más de una fila): {llenos.duplicated(keep=False).sum()} filas"
      " -> se revisan en la sección de identificadores")

# ¿El número del hostname coincide con el del tag? (solo hostnames estándar SITE-TYPE-#####)
num_host = df["hostname"].str.extract(r"-(\d{5})(?:\b|$)")[0]
num_tag = df["asset_tag"].str[-5:]
comparables = num_host.notna() & num_tag.notna()
print(f"Número de hostname = número de tag: {(num_host == num_tag)[comparables].sum()} de {comparables.sum()}")
print(df.loc[comparables & (num_host != num_tag), ["asset_tag", "hostname"]].head(10).to_string())

# %% 9. hostname ---------------------------------------------------------------------------------
# Estándar documentado: SITE-TYPE-##### (ej. NYC-LT-04020). En la base conviven varias convenciones:
#   - estandar:  SITE-TIPO-#####           laptops, desktops, móviles, tablets, impresoras
#   - red:       SITE-SW|FW-ZONA##         switches y firewalls (zona CORE / EDGE / ACC / DMZ)
#   - servidor:  SITE + ROL + ##           servidores y VMs (EAGSQL33: Eagan, SQL, 33; roles
#                                          DC, FS, SQL, ESX, WEB, APP)
#   - legado:    CIUDAD-TIPO-###           nombres viejos (MANILA-LAPTOP-129)
# Las convenciones de red y servidor son razonables para ese tipo de equipo: no son errores, pero
# el estándar documentado no las contempla -> se reporta como hallazgo de gobierno de datos.
# Limpieza (no cambia el nombre, solo la forma):
#   - mayúsculas ('mex-fw-core08' -> 'MEX-FW-CORE08')
#   - se quita el dominio ('CHI-LT-04190.corp.contoso-corp.com' -> 'CHI-LT-04190')
#   - marcadores de vacío ('0', '-', 'none') -> vacío (son monitores, que no tienen hostname)
# El sufijo -OLD se deja tal cual: indica un equipo renombrado / reemplazado y se marca aparte,
# porque puede ser un registro viejo duplicado.
print(crudo.loc[df.index, "hostname"].map(patron).value_counts().to_string())

VACIO_DISFRAZADO = {"0", "", "none", "n a", "na", "null", "tbd", "unknown"}

def normalizar_host(v):
    if pd.isna(v) or clave(v) in VACIO_DISFRAZADO:
        return None
    return str(v).strip().upper().split(".")[0]

aplicar(df, "hostname", df["hostname"].map(normalizar_host), "hostname: mayúsculas / sin dominio / vacío")

CONVENCIONES = [
    ("estandar", r"[A-Z]{3}-(LT|DT|MB|TB|PR)-\d{5}"),
    ("red", r"[A-Z]{3}-(SW|FW)-[A-Z]{3,4}\d{2}"),
    ("servidor", r"[A-Z]{3}(DC|FS|SQL|ESX|WEB|APP)\d{2}"),
    ("legado", r"[A-Z]+(-[A-Z]+)+-\d{3}"),
]

def convencion(h):
    if pd.isna(h):
        return "(sin hostname)"
    base = h.removesuffix("-OLD")
    return next((nombre for nombre, rx in CONVENCIONES if re.fullmatch(rx, base)), "otro")

df["hostname_convention"] = df["hostname"].map(convencion)
df["flag_hostname_old"] = df["hostname"].str.endswith("-OLD", na=False)
print(pd.crosstab(df["asset_type"], df["hostname_convention"], margins=True).to_string())
print(f"\nHostnames con -OLD: {df['flag_hostname_old'].sum()}")
print("Hostnames repetidos:", df["hostname"].dropna().duplicated(keep=False).sum(), "filas")

# %% 10. asset_status ---------------------------------------------------------------------------
# Catálogo de 8 estados (fila 2). Decisiones (acordadas):
#   Active / Assigned / Deployed -> In Use · Available / Stock / Spare -> In Stock
#   Decommissioned / EOL -> Retired (EOL = sin soporte del fabricante; podría seguir en uso -> se marca)
#   e-Waste / Recycled -> Disposed · Not Found -> Missing · RMA (devuelto al fabricante) -> In Repair
ASSET_STATUS = {
    "in use": "In Use", "inuse": "In Use", "active": "In Use", "assigned": "In Use", "deployed": "In Use",
    "in stock": "In Stock", "stock": "In Stock", "available": "In Stock", "spare": "In Stock",
    "in repair": "In Repair", "rma": "In Repair",
    "retired": "Retired", "decommissioned": "Retired", "eol": "Retired",
    "disposed": "Disposed", "e waste": "Disposed", "recycled": "Disposed",
    "missing": "Missing", "not found": "Missing",
    "lost stolen": "Lost/Stolen",
    "on order": "On Order",
}
df["flag_status_eol"] = df["asset_status"].map(clave).eq("eol")
homogeneizar("asset_status", ASSET_STATUS)

# %% 11. os --------------------------------------------------------------------------------------
# Se conserva el máximo detalle que trae cada valor (edición, build) y se agregan dos columnas
# derivadas para el análisis: os_family (Windows, macOS, …) y os_version (Windows 10, macOS 13, …).
# Cuando el valor no dice la versión ('iOS', 'Linux', 'macOS') queda solo la familia.
OS = {
    "windows 11 enterprise 23h2": "Windows 11 Enterprise 23H2", "windows 11 pro": "Windows 11 Pro",
    "windows 11 ent": "Windows 11 Enterprise",
    "windows 11": "Windows 11", "win 11": "Windows 11", "w11": "Windows 11",
    "windows 10 enterprise 22h2": "Windows 10 Enterprise 22H2", "win10 ent": "Windows 10 Enterprise",
    "windows 10": "Windows 10", "win 10": "Windows 10", "w10": "Windows 10", "windows10": "Windows 10",
    "windows 8 1": "Windows 8.1", "windows 7 professional": "Windows 7 Professional",
    "windows server 2022": "Windows Server 2022", "win2022": "Windows Server 2022",
    "w2k22": "Windows Server 2022", "server 2022": "Windows Server 2022",
    "windows server 2019": "Windows Server 2019", "win2019": "Windows Server 2019",
    "ws2019": "Windows Server 2019", "server 2019": "Windows Server 2019",
    "windows server 2016": "Windows Server 2016", "windows server 2012 r2": "Windows Server 2012 R2",
    "windows server 2008 r2": "Windows Server 2008 R2",
    "macos sonoma 14": "macOS Sonoma 14", "macos ventura 13": "macOS Ventura 13",
    "macos": "macOS", "mac os": "macOS", "osx": "macOS",
    "ios 16": "iOS 16", "ios 17": "iOS 17", "ios": "iOS", "ipados 17": "iPadOS 17",
    "android 12": "Android 12", "android 13": "Android 13", "android 14": "Android 14", "android": "Android",
    "ubuntu 22 04 lts": "Ubuntu 22.04 LTS", "ubuntu 22": "Ubuntu 22.04 LTS",
    "red hat enterprise linux 8 9": "Red Hat Enterprise Linux 8.9", "redhat 8": "Red Hat Enterprise Linux 8",
    "rhel 8": "Red Hat Enterprise Linux 8", "rhel 9 3": "Red Hat Enterprise Linux 9.3",
    "centos 7": "CentOS 7", "linux": "Linux", "vmware esxi 7 0": "VMware ESXi 7.0",
}
homogeneizar("os", OS)

def familia_version(os):
    if pd.isna(os):
        return None, None
    reglas = [
        (r"Windows Server (\d{4}( R2)?)", "Windows Server", "Windows Server {}"),
        (r"Windows (\d+(\.\d)?)", "Windows", "Windows {}"),
        (r"macOS(?: \w+)? (\d+)", "macOS", "macOS {}"),
        (r"iPadOS (\d+)", "iPadOS", "iPadOS {}"),
        (r"iOS (\d+)", "iOS", "iOS {}"),
        (r"Android (\d+)", "Android", "Android {}"),
        (r"Red Hat Enterprise Linux (\d+)", "Linux", "RHEL {}"),
        (r"Ubuntu (\d+\.\d+)", "Linux", "Ubuntu {}"),
        (r"CentOS (\d+)", "Linux", "CentOS {}"),
        (r"VMware ESXi (\d+)", "VMware ESXi", "ESXi {}"),
    ]
    for rx, familia, version in reglas:
        m = re.match(rx, os)
        if m:
            return familia, version.format(m.group(1))
    familia = {"macOS": "macOS", "iOS": "iOS", "Android": "Android", "Linux": "Linux"}[os]
    return familia, None  # sin versión

df["os_family"], df["os_version"] = zip(*df["os"].map(familia_version))
print(df.groupby(["os_family", "os_version"], dropna=False).size().to_string())

# %% 12. department y cost_center ----------------------------------------------------------------
# Decisiones (acordadas):
#   - Códigos de centro de costos (CC-1200…) -> columna nueva cost_center; department queda vacío
#     (no hay tabla que diga a qué área corresponde cada código).
#   - Dos áreas distintas ('Finance, Operations') -> 'Shared' + bandera. Repetida ('Finance, Finance') -> Finance.
#   - Service Desk -> Information Technology (es la mesa de ayuda interna de TI).
#   - R&D se deja como área propia (Research & Development), no se junta con Engineering.
es_cc = df["department"].str.fullmatch(r"CC-\d{4}", na=False)
df["cost_center"] = df["department"].where(es_cc)
df["flag_department_cc"] = es_cc  # department a completar: venía el código, no el área
aplicar(df, "department", df["department"].mask(es_cc), "department: código CC -> cost_center")

DEPARTAMENTO = {
    "information technology": "Information Technology", "it dept": "Information Technology",
    "it": "Information Technology", "i t": "Information Technology", "info tech": "Information Technology",
    "service desk": "Information Technology",
    "engineering": "Engineering", "eng": "Engineering",
    "r d": "Research & Development",
    "customer support": "Customer Support", "support": "Customer Support", "cust support": "Customer Support",
    "sales": "Sales", "sales americas": "Sales",
    "operations": "Operations", "ops": "Operations",
    "finance": "Finance", "fin": "Finance", "finance accounting": "Finance",
    "marketing": "Marketing", "mktg": "Marketing",
    "human resources": "Human Resources", "hr": "Human Resources", "h r": "Human Resources",
    "legal": "Legal", "legal compliance": "Legal",
    "executive": "Executive", "exec": "Executive", "c suite": "Executive",
}

def departamento(v):
    areas = {DEPARTAMENTO[clave(p)] for p in str(v).split(",")}
    return areas.pop() if len(areas) == 1 else "Shared"

# Se arma el diccionario con cada valor observado (incluidas las combinaciones con coma)
MAPA_DEP = {clave(v): departamento(v) for v in df["department"].dropna().unique()}
df["flag_department_shared"] = df["department"].map(lambda v: isinstance(v, str) and departamento(v) == "Shared")
homogeneizar("department", MAPA_DEP)
print(df["cost_center"].value_counts().to_string())

# %% 13. location y country ----------------------------------------------------------------------
# Un nombre por sitio (en inglés). Ambiguos: 'HQ' y 'DC1' se resuelven con el código de sitio del
# hostname (las 3 primeras letras: BBL = Blue Bell, EAG = Eagan…).
SITIOS = {  # código de sitio -> (nombre, país)
    "AUS": ("Austin", "United States"), "BBL": ("Blue Bell", "United States"),
    "CHI": ("Chicago", "United States"), "EAG": ("Eagan Data Center", "United States"),
    "NYC": ("New York", "United States"), "BLR": ("Bangalore", "India"),
    "BUD": ("Budapest", "Hungary"), "LON": ("London", "United Kingdom"),
    "MEX": ("Mexico City", "Mexico"), "MNL": ("Manila", "Philippines"),
    "MTL": ("Montreal", "Canada"), "MUC": ("Munich", "Germany"),
    "SAO": ("São Paulo", "Brazil"), "SYD": ("Sydney", "Australia"),
}
LOCATION = {
    "new york": "New York", "nyc": "New York", "ny office": "New York", "new york ny": "New York",
    "bangalore": "Bangalore", "bengaluru": "Bangalore", "blr": "Bangalore",
    "data center eagan": "Eagan Data Center", "eagan dc": "Eagan Data Center",
    "blue bell": "Blue Bell", "blue bell pa": "Blue Bell", "bluebell": "Blue Bell",
    "london": "London", "london uk": "London", "ldn": "London",
    "chicago": "Chicago", "chicago il": "Chicago", "chi": "Chicago",
    "austin": "Austin", "austin tx": "Austin", "aus": "Austin",
    "manila": "Manila", "manila ph": "Manila", "mnl": "Manila",
    "sydney": "Sydney", "sydney nsw": "Sydney", "syd": "Sydney",
    "munich": "Munich", "munchen": "Munich", "muenchen": "Munich",
    "budapest": "Budapest", "budapest hu": "Budapest", "bud": "Budapest",
    "montreal": "Montreal", "mtl": "Montreal",
    "mexico city": "Mexico City", "ciudad de mexico": "Mexico City", "cdmx": "Mexico City",
    "sao paulo": "São Paulo", "sp office": "São Paulo",
    "hq": None, "dc1": None,  # ambiguos: se resuelven por hostname
}
sitio_host = df["hostname"].str[:3].map(lambda c: SITIOS.get(c, (None,))[0] if isinstance(c, str) else None)
df["flag_location_desde_hostname"] = df["location"].map(lambda v: pd.notna(v) and LOCATION.get(clave(v), "") is None)
homogeneizar("location", LOCATION, resolver=sitio_host, via_resolver="hostname")

COUNTRY = {
    "united states": "United States", "us": "United States", "usa": "United States",
    "u s": "United States", "u s a": "United States",
    "india": "India", "in": "India", "ind": "India",
    "united kingdom": "United Kingdom", "uk": "United Kingdom", "gb": "United Kingdom",
    "great britain": "United Kingdom", "england": "United Kingdom",
    "mexico": "Mexico", "mx": "Mexico", "germany": "Germany", "de": "Germany", "deutschland": "Germany",
    "canada": "Canada", "ca": "Canada", "can": "Canada",
    "philippines": "Philippines", "phillipines": "Philippines", "ph": "Philippines",
    "brazil": "Brazil", "br": "Brazil", "brasil": "Brazil",
    "australia": "Australia", "au": "Australia", "aus": "Australia",
    "hungary": "Hungary", "hu": "Hungary", "magyarorszag": "Hungary",
}
homogeneizar("country", COUNTRY)

# País vacío -> se deduce del sitio (cada sitio está en un solo país)
PAIS_DE_SITIO = {nombre: pais for nombre, pais in SITIOS.values()}
df["flag_country_desde_location"] = df["country"].isna()
aplicar(df, "country", df["country"].fillna(df["location"].map(PAIS_DE_SITIO)),
        "country: vacío -> deducido de location")
print(pd.crosstab(df["location"], df["country"]).to_string())

# %% 14. business_criticality, encrypted, patch_compliance, vendor ------------------------------
# business_criticality: entero 1–5 (escala de la fila 2: 1 Low, 2 Minor, 3 Moderate, 4 High,
# 5 Critical). Textos -> número. Fuera de rango (0, 7, 9) -> vacío + bandera (acordado).
CRITICIDAD_TEXTO = {"low": 1, "minor": 2, "med": 3, "medium": 3, "moderate": 3, "high": 4, "critical": 5}

def criticidad(v):
    if pd.isna(v):
        return None
    if isinstance(v, str):
        m = re.match(r"\s*(\d)", v)
        v = int(m.group(1)) if m else CRITICIDAD_TEXTO[clave(v)]
    return int(v) if 1 <= v <= 5 else None

df["flag_criticality_invalida"] = df["business_criticality"].notna() & df["business_criticality"].map(criticidad).isna()
aplicar(df, "business_criticality", df["business_criticality"].map(criticidad).astype("Int64"),
        "business_criticality: texto -> 1-5 / fuera de rango -> vacío")
print(df["business_criticality"].value_counts(dropna=False).sort_index().to_string())

# encrypted: Yes / No / Pending. Vacío = desconocido (no se rellena).
ENCRYPTED = {"y": "Yes", "yes": "Yes", "true": "Yes", "1": "Yes", "no": "No", "n": "No", "pending": "Pending"}
homogeneizar("encrypted", ENCRYPTED)

# patch_compliance: Compliant / Non-Compliant / Pending Reboot. Supuesto: OK y Yes = Compliant;
# FAIL y No = Non-Compliant.
PATCH = {
    "compliant": "Compliant", "ok": "Compliant", "yes": "Compliant",
    "non compliant": "Non-Compliant", "noncompliant": "Non-Compliant", "fail": "Non-Compliant", "no": "Non-Compliant",
    "pending reboot": "Pending Reboot", "pending": "Pending Reboot", "reboot req": "Pending Reboot",
}
homogeneizar("patch_compliance", PATCH)

# vendor: revendedor. CDW-G es la división de gobierno de CDW -> CDW. 'Misc vendor' -> Other.
VENDOR = {
    "cdw": "CDW", "c d w": "CDW", "cdw g": "CDW", "cdw corp": "CDW",
    "shi international": "SHI International", "shi intl": "SHI International", "s h i": "SHI International",
    "shi": "SHI International",
    "insight": "Insight", "insight enterprises": "Insight",
    "dell direct": "Dell Direct", "dell com": "Dell Direct", "dell": "Dell Direct",
    "softchoice": "Softchoice", "soft choice": "Softchoice",
    "amazon business": "Amazon Business", "amzn": "Amazon Business", "amazon": "Amazon Business",
    "local reseller": "Local Reseller", "local": "Local Reseller", "misc vendor": "Other",
}
homogeneizar("vendor", VENDOR)

# %% 15. Fechas ----------------------------------------------------------------------------------
# Formatos encontrados: fecha de Excel, número serial de Excel (44497), ISO ('2021-01-28',
# '2026-08-17T02:11:04Z' en UTC), '20260705', 'dd/mm/aaaa' y 'mm/dd/aaaa', 'mm/dd/aa',
# 'May 31, 2019', '25-Aug-2026 16:44', trimestres ('Q4 2023') y textos ('Never', 'TBD').
# Decisiones (acordadas):
#   - dd/mm vs mm/dd: si un número es > 12 se sabe cuál es el día; si no, se asume mm/dd (formato
#     de EE.UU., el mismo de last_seen y de las notas) y se marca flag_fecha_ambigua.
#   - Trimestre -> día 15 del mes del medio del trimestre (Q4 2023 -> 2023-11-15) + bandera.
#   - 1900-01-01 es el "cero" de Excel, no una fecha real -> vacío.
#   - Garantía en texto ('3 years', '5Y', '1yr ProSupport') -> compra + N años, marcado como derivado.
#     'Lifetime', 'Expired', 'TBD' -> vacío.
#   - last_seen 'Never' / 'never seen' -> vacío + flag_never_seen (dato útil: nunca se conectó).
SIN_FECHA = {"never", "never seen", "0", "tbd", "expired", "lifetime", "n a", ""}

def leer_fecha(v):
    """Devuelve (fecha, nota). nota: None | 'ambigua' | 'trimestre'."""
    if pd.isna(v) or clave(v) in SIN_FECHA:
        return None, None
    if isinstance(v, datetime):
        return (None if v.year == 1900 else pd.Timestamp(v)), None
    if isinstance(v, (int, float)):
        return pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(v)), None
    s = v.strip()
    if m := re.fullmatch(r"Q([1-4]) (\d{4})", s, flags=re.I):
        return pd.Timestamp(int(m.group(2)), 3 * int(m.group(1)) - 1, 15), "trimestre"
    if re.match(r"(?i)\d+\s*(y|yr|yrs|year|years)\b", s):
        return None, "duracion"  # garantía en años: se calcula aparte
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}(T[\d:]+Z?)?", s):
        t = pd.Timestamp(s)
        return (t.tz_convert(None) if t.tzinfo else t), None
    if re.fullmatch(r"\d{8}", s):
        return pd.to_datetime(s, format="%Y%m%d"), None
    if m := re.fullmatch(r"(\d{2})/(\d{2})/(\d{2}|\d{4})", s):
        a, b, anio = int(m.group(1)), int(m.group(2)), int(m.group(3))
        anio = anio + 2000 if anio < 100 else anio
        if a > 12:
            return pd.Timestamp(anio, b, a), None
        return pd.Timestamp(anio, a, b), (None if b > 12 else "ambigua")
    for formato in ("%B %d, %Y", "%d-%b-%Y %H:%M"):
        try:
            return pd.to_datetime(s, format=formato), None
        except ValueError:
            pass
    raise ValueError(f"Formato de fecha sin regla: {v!r}")

for col in ["purchase_date", "warranty_end_date", "last_seen"]:
    leidas = df[col].map(leer_fecha)
    fechas = leidas.map(lambda t: t[0])
    notas = leidas.map(lambda t: t[1])
    if col == "purchase_date":
        df["flag_fecha_ambigua"] = notas.eq("ambigua")
        df["flag_fecha_aproximada"] = notas.eq("trimestre")
    if col == "warranty_end_date":
        # texto con duración -> se calcula desde la compra
        anios = df[col].map(lambda v: re.match(r"(?i)\s*(\d+)\s*(y|yr|yrs|year|years)\b", v)
                            if isinstance(v, str) else None).map(lambda m: int(m.group(1)) if m else None)
        derivada = anios.notna() & df["purchase_date"].notna()
        df["flag_garantia_derivada"] = derivada
        calculada = pd.Series([c + pd.DateOffset(years=int(n)) if d else pd.NaT
                               for c, n, d in zip(df["purchase_date"], anios, derivada)], index=df.index)
        fechas = fechas.where(~derivada, calculada)
    if col == "last_seen":
        df["flag_never_seen"] = df[col].map(lambda v: isinstance(v, str) and clave(v) in {"never", "never seen"})
    aplicar(df, col, pd.to_datetime(fechas), f"{col}: formato -> fecha")
    print(f"{col}: {df[col].notna().sum()} fechas | rango {df[col].min()} a {df[col].max()}")
print("Fechas de compra ambiguas (mm/dd asumido):", df["flag_fecha_ambigua"].sum(),
      "| trimestre:", df["flag_fecha_aproximada"].sum(), "| garantías derivadas:", df["flag_garantia_derivada"].sum())

# 15b. Año de compra imposible (decidido en la revisión de contradicciones)
#   La garantía es casi siempre compra + N años con el mismo día y mes (727 de 830), con N fijo por
#   tipo (1 teléfono/tablet, 3 PC/monitor/impresora, 5 servidor/red). Las compras posteriores al
#   corte (2027-2029) o anteriores a 2010 (todas 1999) conservan día y mes: el error es el año.
#   -> purchase_date = fin de garantía - N años del tipo + flag_compra_reconstruida.
#   Sin garantía propia (VM, o garantía derivada de esa misma compra) no se puede: vacío + misma bandera.
compra, garantia = df["purchase_date"], df["warranty_end_date"]
mismo_dia = compra.dt.strftime("%m-%d").eq(garantia.dt.strftime("%m-%d"))
anios_garantia = (garantia.dt.year - compra.dt.year)[mismo_dia & ~df["flag_garantia_derivada"]]
print(f"Garantía = compra + N años con el mismo día y mes: {(mismo_dia & ~df['flag_garantia_derivada']).sum()} "
      f"de {(compra.notna() & garantia.notna()).sum()} con ambas fechas")
ANIOS_GARANTIA = anios_garantia[anios_garantia.between(1, 5)].groupby(df["asset_type"]).agg(lambda s: s.mode()[0])
print("Años de garantía estándar por tipo:", ANIOS_GARANTIA.to_dict())

anio_imposible = (compra > pd.Timestamp(FECHA_CORTE)) | (compra < pd.Timestamp("2010-01-01"))
df["flag_compra_reconstruida"] = anio_imposible
n = df["asset_type"].map(ANIOS_GARANTIA)
se_puede = anio_imposible & garantia.notna() & ~df["flag_garantia_derivada"] & n.notna()
reconstruida = pd.Series([g - pd.DateOffset(years=int(k)) if ok else pd.NaT
                          for g, k, ok in zip(garantia, n, se_puede)], index=df.index)
aplicar(df, "purchase_date", compra.where(~anio_imposible, reconstruida),
        "purchase_date: año imposible -> fin de garantía - años estándar del tipo (vacío si no hay garantía)")
# la garantía que se había calculado desde esa compra errónea tampoco vale
aplicar(df, "warranty_end_date", garantia.where(~(anio_imposible & df["flag_garantia_derivada"])),
        "warranty_end_date: derivada de una compra con año imposible -> vacío")
print(f"Compras con año imposible: {anio_imposible.sum()} | reconstruidas: {se_puede.sum()} | "
      f"quedan vacías: {(anio_imposible & ~se_puede).sum()}")
print(df.loc[anio_imposible, ["asset_type", "purchase_date", "warranty_end_date"]].to_string())

# 15b'. Garantía vacía con compra (equipos físicos): en el crudo venía vacía, 'TBD' o 'Expired'.
#   Decidido (Fernanda): se estima = compra + años estándar del tipo (el patrón que cumplen 716 de 843)
#   + flag_garantia_estimada, y sale de confiable_fechas (es un supuesto).
#   'Lifetime' no tiene fin -> queda vacía. VM no llevan garantía (no hay años estándar).
garantia_cruda = crudo.loc[df.index, "warranty_end"].map(lambda v: clave(v) if isinstance(v, str) else None)
df["flag_garantia_estimada"] = (df["warranty_end_date"].isna() & df["purchase_date"].notna()
                                & n.notna() & garantia_cruda.ne("lifetime"))
estimada = pd.Series([c + pd.DateOffset(years=int(k)) if ok else pd.NaT
                      for c, k, ok in zip(df["purchase_date"], n, df["flag_garantia_estimada"])], index=df.index)
# 'Expired' cuya estimación queda vigente al corte (compra reciente): el crudo contradice al supuesto
# -> no se estima, queda vacía + flag_garantia_contradictoria (2 casos: compras 2024 y 2026).
df["flag_garantia_contradictoria"] = (df["flag_garantia_estimada"] & garantia_cruda.eq("expired")
                                      & (estimada >= pd.Timestamp(FECHA_CORTE)))
df["flag_garantia_estimada"] &= ~df["flag_garantia_contradictoria"]
estimada = estimada.where(df["flag_garantia_estimada"])
aplicar(df, "warranty_end_date", df["warranty_end_date"].where(~df["flag_garantia_estimada"], estimada),
        "warranty_end_date: vacía -> compra + años estándar del tipo (estimada)")
print("Garantías estimadas:", df["flag_garantia_estimada"].sum(), "| por valor crudo:",
      garantia_cruda[df["flag_garantia_estimada"]].fillna("(vacío)").value_counts().to_dict())
print("'Expired' con compra reciente (no se estima, vacía):", df["flag_garantia_contradictoria"].sum())

# 15c. last_seen a las 09:00:00 exactas y anterior a la compra: valor por defecto de otra fuente,
#   no un reporte real (los 8 casos cumplen ambas cosas) -> vacío + flag_last_seen_invalido.
df["flag_last_seen_invalido"] = (df["last_seen"].dt.strftime("%H:%M:%S").eq("09:00:00")
                                 & (df["last_seen"] < df["purchase_date"]))
aplicar(df, "last_seen", df["last_seen"].where(~df["flag_last_seen_invalido"]),
        "last_seen: 09:00:00 exacto y anterior a la compra -> vacío")
print("last_seen inválidos (09:00 antes de la compra):", df["flag_last_seen_invalido"].sum())

# %% 16. purchase_cost y currency ----------------------------------------------------------------
# Montos en texto con separadores de distintos países ('2.013,90' = 2013.90; '442,796.40' =
# 442796.40; '569,99' = 569.99), símbolos ('$1,110', '€690.00') y moneda pegada ('3785 AUD').
# Regla de separadores: si hay punto y coma, el último es el decimal; si solo hay coma, es decimal
# cuando le siguen 1–2 dígitos y de miles cuando le siguen 3.
# Decisiones (acordadas): negativos -> valor absoluto + bandera. 'free' -> 0. 'see PO', '-', 'TBD' -> vacío.
SIMBOLOS = {"us$": "USD", "c$": "CAD", "a$": "AUD", "r$": "BRL", "€": "EUR", "£": "GBP", "$": "$"}

def leer_monto(v):
    """Devuelve (monto, moneda que venía pegada al monto o None)."""
    if pd.isna(v):
        return None, None
    if isinstance(v, (int, float)):
        return float(v), None
    s = v.strip().lower()
    if s in {"see po", "-", "tbd", ""}:
        return None, None
    if s == "free":
        return 0.0, None
    moneda = None
    if m := re.search(r"\s*([a-z]{3})$", s):
        moneda, s = m.group(1).upper(), s[:m.start()]
    for simbolo, codigo in SIMBOLOS.items():
        if s.startswith(simbolo):
            moneda, s = codigo, s[len(simbolo):]
            break
    s = s.replace(" ", "")
    if "," in s and "." in s:
        decimal = "," if s.rfind(",") > s.rfind(".") else "."
        s = s.replace("." if decimal == "," else ",", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", "") if re.search(r",\d{3}$", s) else s.replace(",", ".")
    return float(s), moneda

leidos = df["purchase_cost"].map(leer_monto)
monto = leidos.map(lambda t: t[0])
moneda_en_monto = leidos.map(lambda t: t[1])
df["flag_costo_negativo"] = monto < 0
aplicar(df, "purchase_cost", monto.abs(), "purchase_cost: texto -> número / negativo -> absoluto")

# currency: código ISO. '$' y 'Peso' solo aparecen en filas de EE.UU. y México respectivamente.
def codigo_moneda(v):
    if pd.isna(v):
        return None
    s = str(v).strip().lower()
    return {"usd": "USD", "us dollar": "USD", "us$": "USD", "$": "USD", "inr": "INR", "rs": "INR",
            "gbp": "GBP", "£": "GBP", "pound": "GBP", "eur": "EUR", "€": "EUR", "euro": "EUR",
            "brl": "BRL", "r$": "BRL", "cad": "CAD", "c$": "CAD", "aud": "AUD", "a$": "AUD",
            "huf": "HUF", "ft": "HUF", "mxn": "MXN", "peso": "MXN", "php": "PHP"}[s]

aplicar(df, "currency", df["currency"].map(codigo_moneda), "currency: variante -> código ISO")

# Moneda vacía (acordado): 1) la que venía pegada al monto; 2) la moneda local del país. Se marca.
MONEDA_LOCAL = {"United States": "USD", "India": "INR", "United Kingdom": "GBP", "Mexico": "MXN",
                "Germany": "EUR", "Canada": "CAD", "Philippines": "PHP", "Brazil": "BRL",
                "Australia": "AUD", "Hungary": "HUF"}
dolar_por_pais = df["country"].map({"United States": "USD", "Canada": "CAD", "Australia": "AUD", "Mexico": "MXN"})
moneda_en_monto = moneda_en_monto.where(moneda_en_monto != "$", dolar_por_pais)
falta = df["currency"].isna() & df["purchase_cost"].notna()
inferida = moneda_en_monto.where(moneda_en_monto.notna(), df["country"].map(MONEDA_LOCAL))
df["flag_moneda_inferida"] = falta
aplicar(df, "currency", df["currency"].where(~falta, inferida), "currency: vacía -> inferida (monto / país)")
print("Monedas inferidas:", falta.sum(), "| costos negativos:", df["flag_costo_negativo"].sum())
# Moneda pegada al monto distinta de currency (acordado): se deja currency como viene, se marca
# y sale de confiable_costos (no se sabe cuál de las dos es la correcta).
choque_moneda = moneda_en_monto.notna() & df["currency"].notna() & (moneda_en_monto != df["currency"])
print("Moneda del monto ≠ currency:", choque_moneda.sum())
print(pd.DataFrame({"monto_crudo": crudo.loc[choque_moneda[choque_moneda].index, "Purchase Cost"],
                    "curr_crudo": crudo.loc[choque_moneda[choque_moneda].index, "Curr."],
                    "moneda_en_monto": moneda_en_monto[choque_moneda],
                    "currency": df.loc[choque_moneda, "currency"],
                    "country": df.loc[choque_moneda, "country"],
                    "location": df.loc[choque_moneda, "location"]}).to_string())
df["flag_moneda_monto_distinta"] = choque_moneda

# Costo 0 en equipo físico (15 + 'free'): todos tienen PO de un revendedor y ninguno es leasing ->
# no es creíble que costaran 0 (alguno podría ser reemplazo por garantía). Decidido (Fernanda,
# conservador): vacío + bandera, a revisar; así no bajan los promedios de dinero.
# En VM el 0 se deja: la VM no se compra (suscripción), 0 = sin costo de compra.
df["flag_costo_cero"] = df["purchase_cost"].eq(0) & df["asset_type"].ne("Virtual Machine")
print("Físicos con costo 0 -> vacío:", df["flag_costo_cero"].sum())
aplicar(df, "purchase_cost", df["purchase_cost"].mask(df["flag_costo_cero"]), "purchase_cost: 0 en equipo físico -> vacío")

# purchase_cost_usd: el mismo monto homogeneizado a USD (FX de config) para comparar entre países.
# purchase_cost y currency quedan como vienen (moneda original). En el Excel se recalcula con fórmula.
df["purchase_cost_usd"] = (df["purchase_cost"] * df["currency"].map(FX)).round(2)
costo_usd = df["purchase_cost_usd"]
# Estadística descriptiva del costo por tipo de activo, en USD. Sirve para ver si hay montos no
# razonables (muy chicos o muy grandes para el tipo).
print("\nCosto de compra en USD por asset_type:")
print(costo_usd.groupby(df["asset_type"]).describe(percentiles=[.05, .5, .95]).round(0).to_string())

# Revisión del parseo (texto con coma/punto): los 89 textos se leyeron bien ('2.013,90' -> 2013.90,
# '442,796.40' -> 442796.40, '569,99' -> 569.99); los montos enormes vienen como NÚMERO en el crudo,
# no son un error de lectura de separadores.
# Montos chicos: sin los 0 no hay montos chicos raros; el menor > 0 es ~0.2 veces la mediana del tipo (impresoras, monitores baratos).
# Montos grandes: hay equipos de 10 a 220 veces la mediana (laptop de US$442,796).
# Decisión (Fernanda, conservadora): > 3 veces la mediana del tipo en USD -> se QUITA el monto (vacío),
# se marca y sale de confiable_costos: no se usa para estadísticas. El monto original queda en
# purchase_cost_descartado (moneda original) para revisarlo. No se corrige porque la causa es una
# suposición; se anota la causa probable:
#   - moneda: currency USD fuera de EE.UU. y, leído en moneda local, el monto es normal para el tipo;
#   - decimal: dividido por 100 el monto es normal para el tipo (se perdió el punto decimal);
# 'normal' = entre 0.3 y 3 veces la mediana. Las VM no entran: su costo no se exige (mediana 0).
# Observación (Fernanda): en los 11 de 'moneda' la currency está mal anotada: dice USD pero el monto
# solo tiene sentido en moneda local. Chequeado: NO es la práctica del sitio (Brasil carga 40 en BRL
# y 5 en USD). De las 29 compras en USD fuera de EE.UU. (sin VM), 11 son este error (Brasil 5 de 5,
# México 2 de 2, Hungría 1 de 1, India 3 de 4) y 18 son plausibles en USD (Australia, Canadá,
# Alemania, Reino Unido). Ojo: donde la moneda local vale parecido al USD (AUD, CAD, EUR, GBP) el
# error no se puede detectar por el monto, así que podría haber más casos no visibles.
mediana_usd = costo_usd.groupby(df["asset_type"]).transform("median")
razon = costo_usd / mediana_usd
df["flag_costo_atipico"] = (razon > 3) & df["asset_type"].ne("Virtual Machine")
razon_local = df["purchase_cost"] * df["country"].map(MONEDA_LOCAL).map(FX) / mediana_usd
normal = lambda r: r.between(0.3, 3)
df["causa_costo_atipico"] = None
df.loc[df["flag_costo_atipico"] & normal(razon / 100), "causa_costo_atipico"] = "decimal perdido (monto x100)"
df.loc[df["flag_costo_atipico"] & df["currency"].eq("USD") & df["country"].ne("United States")
       & normal(razon_local), "causa_costo_atipico"] = "moneda local cargada como USD"
df["causa_costo_atipico"] = df["causa_costo_atipico"].where(~df["flag_costo_atipico"] | df["causa_costo_atipico"].notna(),
                                                            "sin causa clara")
print("\nCostos atípicos (> 3 veces la mediana del tipo):", df["flag_costo_atipico"].sum())
print(df.loc[df["flag_costo_atipico"], ["asset_type", "model", "country", "currency", "purchase_cost",
                                        "causa_costo_atipico"]].assign(usd=costo_usd.round(0),
                                        veces_mediana=razon.round(1)).sort_values("veces_mediana").to_string())
print("Más chicos (> 0) por tipo, en veces la mediana:")
print(razon[costo_usd > 0].groupby(df["asset_type"]).min().round(2).to_string())

df["purchase_cost_descartado"] = df["purchase_cost"].where(df["flag_costo_atipico"])
aplicar(df, "purchase_cost", df["purchase_cost"].mask(df["flag_costo_atipico"]), "purchase_cost: atípico (> 3x mediana del tipo) -> vacío")
df["purchase_cost_usd"] = df["purchase_cost_usd"].mask(df["flag_costo_atipico"])
print("\nCosto de compra en USD por asset_type, SIN atípicos (lo que se usa para estadísticas):")
print(df["purchase_cost_usd"].groupby(df["asset_type"]).describe(percentiles=[.05, .5, .95]).round(0).to_string())

# %% 17. ram_gb, storage, open_incidents ---------------------------------------------------------
# ram_gb: número en GB. '16384 MB' -> 16. '16G', '16gb', '16 GB' -> 16.
# 1600 y 3200 parecen la velocidad de la memoria (MHz), no GB -> vacío + bandera (acordado).
def leer_ram(v):
    if pd.isna(v):
        return None
    if isinstance(v, (int, float)):
        return float(v) if v <= 1024 else None
    m = re.fullmatch(r"(?i)\s*(\d+(?:\.\d+)?)\s*(mb|gb|g)\s*", v)
    n, unidad = float(m.group(1)), m.group(2).lower()
    return n / 1024 if unidad == "mb" else n

df["flag_ram_invalida"] = df["ram_gb"].notna() & df["ram_gb"].map(leer_ram).isna()
aplicar(df, "ram_gb", df["ram_gb"].map(leer_ram), "ram_gb: texto/MB -> GB / no plausible -> vacío")

# storage_gb: número en GB; el tipo de disco (SSD/HDD) se separa en storage_type.
# La unidad va en el nombre de la variable (como ram_gb): todo queda en GB.
# 1 TB = 1024 GB (así viene en la base: '0.5 TB' = 512 GB). '2 x 960 GB SSD' = 1920 GB.
# Supuesto: número sin unidad < 20 está en TB ('1', '4', '12'); si no, en GB ('256', '512').
def leer_storage(v):
    if pd.isna(v):
        return None, None
    s = str(v).strip().lower().replace(" ", "")
    m = re.fullmatch(r"(?:(\d+)x)?(\d+(?:\.\d+)?)(tb|gb)?(ssd|hdd)?", s)
    cantidad, n, unidad, tipo = int(m.group(1) or 1), float(m.group(2)), m.group(3), m.group(4)
    unidad = unidad or ("tb" if n < 20 else "gb")
    gb = round(cantidad * n * (1024 if unidad == "tb" else 1))
    return gb, (tipo.upper() if tipo else None)

leidos = df["storage_gb"].map(leer_storage)
df["storage_type"] = leidos.map(lambda t: t[1])
aplicar(df, "storage_gb", leidos.map(lambda t: t[0]).astype("Int64"), "storage_gb: texto/TB -> GB")
print(df.groupby("storage_type", dropna=False)["storage_gb"].describe().to_string())

# open_incidents: entero >= 0. '1 open' -> 1, 'none' -> 0, 'TBD' -> vacío, -1 -> vacío + bandera.
def leer_incidentes(v):
    if pd.isna(v) or clave(v) == "tbd":
        return None
    if clave(v) == "none":
        return 0
    n = int(re.match(r"\s*(-?\d+)", str(v)).group(1))
    return n if n >= 0 else None

df["flag_incidentes_invalido"] = df["open_incidents"].notna() & df["open_incidents"].map(
    lambda v: clave(v) != "tbd" and leer_incidentes(v) is None)
aplicar(df, "open_incidents", df["open_incidents"].map(leer_incidentes).astype("Int64"),
        "open_incidents: texto -> entero / negativo -> vacío")

# %% 18. Identificadores secundarios: serial, MAC, IP, PO ---------------------------------------
# serial_number: texto en mayúsculas; '0', 'UNKNOWN', 'SEE STICKER' y 'TO BE FILLED' son vacíos
# disfrazados -> vacío + flag_serial_placeholder (un serial de relleno se "repite" entre equipos).
SERIAL_RELLENO = {"0", "unknown", "see sticker", "to be filled"}
df["flag_serial_placeholder"] = df["serial_number"].map(lambda v: isinstance(v, (str, int)) and clave(v) in SERIAL_RELLENO)
aplicar(df, "serial_number",
        df["serial_number"].map(lambda v: None if pd.isna(v) or clave(v) in SERIAL_RELLENO else str(v).strip().upper()),
        "serial_number: texto / vacío disfrazado")
print("Seriales de relleno -> vacío:", df["flag_serial_placeholder"].sum())

# mac_address: formato AA:BB:CC:DD:EE:FF. Se aceptan ':', '-', '.' o sin separador. Las que no
# tienen 12 dígitos hex (truncadas, ej. '6B:CF:88:F8:B7') -> vacío + bandera.
def leer_mac(v):
    if pd.isna(v):
        return None
    h = re.sub(r"[^0-9a-f]", "", str(v).lower())
    return ":".join(h[i:i + 2] for i in range(0, 12, 2)).upper() if len(h) == 12 else None

df["flag_mac_invalida"] = df["mac_address"].notna() & df["mac_address"].map(leer_mac).isna()
aplicar(df, "mac_address", df["mac_address"].map(leer_mac), "mac_address: formato / inválida -> vacío")

# ip_address: 'DHCP', 'dynamic', '0.0.0.0' -> vacío (IP dinámica / sin IP). Octeto > 255 -> inválida.
# Se quitan ceros a la izquierda de cada octeto ('010.033.033.115' -> '10.33.33.115').
def leer_ip(v):
    if pd.isna(v) or clave(v) in {"dhcp", "dynamic"} or v == "0.0.0.0":
        return None
    octetos = str(v).split(".")
    if len(octetos) == 4 and all(o.isdigit() and int(o) <= 255 for o in octetos):
        return ".".join(str(int(o)) for o in octetos)  # sin ceros a la izquierda: '010.033' -> '10.33'
    return None

df["flag_ip_invalida"] = df["ip_address"].notna() & df["ip_address"].map(leer_ip).isna() & \
    ~df["ip_address"].map(lambda v: clave(v) in {"dhcp", "dynamic"} or v == "0.0.0.0")
aplicar(df, "ip_address", df["ip_address"].map(leer_ip), "ip_address: placeholder / inválida -> vacío")

# po_number: 10 dígitos. Se quita el prefijo 'PO' / 'PO-' y el sufijo de línea '/1'.
aplicar(df, "po_number",
        df["po_number"].map(lambda v: None if pd.isna(v) else re.search(r"\d{10}", str(v)).group()),
        "po_number: formato 10 dígitos")

# %% 19. Personas: assigned_to y user_email ------------------------------------------------------
# user_email: minúsculas; se arreglan errores de tipeo evidentes: '@@' -> '@', ';' final, espacio en
# lugar de punto, falta de '@' antes del dominio y '.con' -> '.com'.
# Dominios distintos de contoso-corp.com NO se cambian (no se sabe si @contoso.com es un alias) y se
# marcan; @gmail.com es un correo personal en un registro corporativo (hallazgo de cumplimiento).
def leer_email(v):
    if pd.isna(v):
        return None
    s = str(v).strip().lower().rstrip(";").replace("@@", "@").replace(" ", ".")
    if "@" not in s:
        s = s.replace("contoso-corp.com", "@contoso-corp.com")
    return re.sub(r"\.con$", ".com", s)

aplicar(df, "user_email", df["user_email"].map(leer_email), "user_email: formato")
emails = df["user_email"].dropna()
assert emails.str.fullmatch(r"[a-z0-9.'_-]+@[a-z0-9-]+\.com").all(), emails[~emails.str.fullmatch(r"[a-z0-9.'_-]+@[a-z0-9-]+\.com")].tolist()
df["flag_email_no_corporativo"] = df["user_email"].notna() & ~df["user_email"].str.endswith("@contoso-corp.com", na=False)
print(df.loc[df["flag_email_no_corporativo"], "user_email"].str.split("@").str[1].value_counts().to_string())

# assigned_to: 'Nombre Apellido'.
#   - '(Terminated)' se quita del nombre y se marca: equipo asignado a un ex-empleado (hallazgo).
#   - 'Apellido, Nombre' -> 'Nombre Apellido'. MAYÚSCULAS -> Nombre Propio.
#   - Usuario ('nreyes') o inicial ('A. Miller'): si el email de la fila tiene el mismo apellido, se
#     toma el nombre del email (nombre.apellido@) y se marca. Si no coincide, queda como está.
df["flag_user_terminated"] = df["assigned_to"].str.contains(r"\(Terminated\)", na=False)

def nombre_del_email(email):
    if pd.isna(email):
        return None
    partes = email.split("@")[0].split(".")
    return " ".join(p.capitalize() for p in partes) if len(partes) == 2 else None

def leer_nombre(v, email):
    if pd.isna(v):
        return None, False
    s = re.sub(r"\s*\(Terminated\)", "", v).strip()
    if "," in s:
        apellido, nombre = [p.strip() for p in s.split(",", 1)]
        s = f"{nombre} {apellido}"
    if s.isupper():
        s = s.title()
    del_email = nombre_del_email(email)
    es_usuario = re.fullmatch(r"[a-z]+", s) is not None
    es_inicial = re.fullmatch(r"[A-Z]\. \w+", s) is not None
    if (es_usuario or es_inicial) and del_email:
        apellido = s.split()[-1].lower() if es_inicial else s[1:]
        if del_email.split()[-1].lower() == apellido:
            return del_email, True
    return s, False

leidos = [leer_nombre(v, e) for v, e in zip(df["assigned_to"], df["user_email"])]
df["flag_nombre_desde_email"] = [t[1] for t in leidos]
aplicar(df, "assigned_to", pd.Series([t[0] for t in leidos], index=df.index), "assigned_to: formato 'Nombre Apellido'")
# Decidido en la revisión de contradicciones (estado vs. persona):
#   - 'Apellido Nombre' sin coma: si invertido coincide exacto con el nombre del email -> se invierte.
#   - In Use sin persona pero con email -> nombre del email (misma regla que usuario/inicial).
del_email = df["user_email"].map(nombre_del_email)
invertido = df["assigned_to"].map(lambda s: " ".join(s.split()[::-1]) if isinstance(s, str) else None)
df["flag_nombre_invertido"] = invertido.notna() & invertido.eq(del_email) & df["assigned_to"].ne(del_email)
aplicar(df, "assigned_to", df["assigned_to"].where(~df["flag_nombre_invertido"], invertido),
        "assigned_to: 'Apellido Nombre' -> 'Nombre Apellido' (coincide con el email)")
sin_persona = df["assigned_to"].isna() & del_email.notna() & df["asset_status"].eq("In Use")
df["flag_nombre_desde_email"] |= sin_persona
aplicar(df, "assigned_to", df["assigned_to"].where(~sin_persona, del_email),
        "assigned_to: vacío en equipo In Use -> nombre del email")
print("Nombres invertidos:", df["flag_nombre_invertido"].sum(), "| In Use sin persona completados:", sin_persona.sum())
raros = df["assigned_to"].dropna()
print("Nombres que siguen sin formato 'Nombre Apellido':",
      raros[~raros.str.fullmatch(r"[A-Z][a-zA-Z']+ [A-Z][a-zA-Z']+")].tolist())
print("Asignados a ex-empleados (Terminated):", df["flag_user_terminated"].sum(),
      "| nombre tomado del email:", df["flag_nombre_desde_email"].sum())

# %% 20. Eventos en notes -------------------------------------------------------------------------
# notes es texto libre, pero sale de ~9 plantillas fijas. Cada plantilla se convierte en un
# indicador de evento (evt_*) y, cuando la nota trae datos (fechas, incidentes, sitios), en columnas.
# notes se conserva tal cual. Las fechas de las notas están en mm/dd/aaaa (formato EE.UU.).
notas = df["notes"].fillna("").str.replace("–", "-")

def extraer(patron):
    return notas.str.extract(patron)

m = extraer(r"Device re-imaged (\d{2}/\d{2}/\d{4}) after malware alert \((INC\d+)\)\. "
            r"Previous owner: (.+?), ext\. \d+\. Asset moved from (.+?) to (.+?)\.")
df["evt_malware"] = m[0].notna()
df["malware_reimage_date"] = pd.to_datetime(m[0], format="%m/%d/%Y")
df["malware_incident"] = m[1]
df["previous_owner"] = m[2]
df["moved_from"], df["moved_to"] = m[3], m[4]

m = extraer(r"Lease ends (\d{2}/\d{2}/\d{4})")
df["evt_lease"] = m[0].notna()
df["lease_end_date"] = pd.to_datetime(m[0], format="%m/%d/%Y")

m = extraer(r"battery swelling; (\w+) raised")
df["evt_battery_swelling"] = m[0].notna()
df["battery_incident"] = m[0]

m = extraer(r"Replacement for asset that failed \((.+?)\)\. See (\w+)")
df["evt_replacement_failure"] = m[0].notna()
df["failed_component"], df["replacement_incident"] = m[0], m[1]

df["evt_loaner"] = notas.str.contains("Loaner pool device")
df["evt_shared"] = notas.str.contains("Shared device")

m = extraer(r"verified during physical audit (\w{3} \d{4})")
df["evt_physical_audit"] = m[0].notna()
df["audit_month"] = pd.to_datetime(m[0], format="%b %Y")

m = extraer(r"FY(\d{2}) hardware refresh")
df["evt_refresh"] = m[0].notna()
df["refresh_fy"] = (2000 + pd.to_numeric(m[0])).astype("Int64")

m = extraer(r"Hosts (.+?) for the (.+?) team")
df["evt_hosts_service"] = m[0].notna()
df["hosted_service"], df["service_team"] = m[0], m[1]

eventos = [c for c in df.columns if c.startswith("evt_")]
sin_plantilla = df["notes"].notna() & ~df[eventos].any(axis=1)
print(df[eventos].sum().sort_values(ascending=False).to_string())
print("Notas que no calzan con ninguna plantilla:", df.loc[sin_plantilla, "notes"].tolist())
# ¿Hay notas largas (malware) idénticas en activos distintos? Sería señal de copia y pega.
malware = df[df["evt_malware"]]
repetidas = malware[malware["notes"].duplicated(keep=False)]
print("Notas de malware repetidas en activos con distinto asset_tag:",
      repetidas.groupby("notes")["asset_tag"].nunique().gt(1).sum())

# %% 21. Identificadores: duplicados ------------------------------------------------------------
# ¿Puede repetirse un asset_tag? NO: es la etiqueta física única de cada activo (la llave del
# registro). Si se repite es porque (a) el mismo equipo está cargado dos veces, o (b) dos equipos
# distintos recibieron el mismo tag por error.
# ¿Puede repetirse un hostname? Al mismo tiempo NO (dos equipos vivos con el mismo nombre chocan
# en la red), pero a lo largo del tiempo SÍ: cuando se reemplaza un equipo, el nuevo suele heredar
# el nombre del viejo, que queda dado de baja. Por eso el hostname no sirve como llave.
# Para distinguir "mismo equipo" de "equipo distinto" se usa el serial del fabricante.
# Esta sección va después de homogeneizar: así 'INUSE' e 'In Use' ya cuentan como el mismo valor.
COLS_DATOS = [c for c in df.columns if c in APLICA_A]  # solo campos originales, no derivados

def campos_distintos(grupo):
    return [c for c in COLS_DATOS if grupo[c].astype(str).nunique() > 1]

def solo_difiere_en_vacios(grupo):
    """True si, en cada campo, las filas que tienen dato tienen el mismo valor (una solo agrega info)."""
    return all(grupo[c].dropna().astype(str).nunique() <= 1 for c in COLS_DATOS)

def mas_completa(filas):
    """Fila con más campos llenos (si empatan, la primera)."""
    return df.loc[filas, COLS_DATOS].notna().sum(axis=1).idxmax()

# 21a. asset_tag repetido
df["dup_tag_tipo"] = None
rep = df[df["asset_tag"].notna() & df["asset_tag"].duplicated(keep=False)]
filas = []
for tag, g in rep.groupby("asset_tag"):
    distintos = campos_distintos(g)
    if not distintos:
        tipo = "duplicado exacto"
    elif solo_difiere_en_vacios(g):
        tipo = "mismo equipo, una fila más completa"
    elif g["serial_number"].astype(str).nunique() == 1:
        tipo = "mismo equipo, datos distintos"
    else:
        tipo = "tag repetido en equipos distintos"
    df.loc[g.index, "dup_tag_tipo"] = tipo
    filas.append({"asset_tag": tag, "filas_excel": list(g.index), "tipo": tipo,
                  "hostnames": list(g["hostname"]), "seriales": list(g["serial_number"]),
                  "campos_distintos": distintos})
dup_tags = pd.DataFrame(filas).sort_values("tipo")
print(dup_tags["tipo"].value_counts().to_string())
print(dup_tags.to_string(index=False))

# 21b. hostname repetido (sin contar los grupos que ya son el mismo tag repetido)
DE_BAJA = {"Retired", "Disposed", "Lost/Stolen", "On Order"}
df["dup_hostname_tipo"] = None
rep = df[df["hostname"].notna() & df["hostname"].duplicated(keep=False)]
filas = []
for host, g in rep.groupby("hostname"):
    if g["asset_tag"].nunique() == 1 and g["asset_tag"].notna().all():
        continue  # mismo activo cargado dos veces: ya está en 11a
    vivos = (~g["asset_status"].isin(DE_BAJA)).sum()
    tipo = "nombre reutilizado (uno vivo)" if vivos <= 1 else "conflicto: varios vivos con el mismo nombre"
    df.loc[g.index, "dup_hostname_tipo"] = tipo
    filas.append({"hostname": host, "filas_excel": list(g.index), "tipo": tipo,
                  "tipos_activo": sorted(set(g["asset_type"])), "estados": list(g["asset_status"]),
                  "tags": list(g["asset_tag"])})
dup_hosts = pd.DataFrame(filas).sort_values("tipo")
# En los conflictos se marcan solo los equipos vivos: son los que chocan en la red y hay que revisar
# cuál conserva el nombre (el dado de baja del mismo grupo no molesta).
df["flag_hostname_conflicto"] = (df["dup_hostname_tipo"].eq("conflicto: varios vivos con el mismo nombre")
                                 & ~df["asset_status"].isin(DE_BAJA))

# Decidido (revisión): en cada conflicto, marcar el equipo que probablemente es el registro viejo
# sin dar de baja. Todos son equipos distintos (serial, modelo, MAC e IP propios), así que no se
# borra nada: se indica cuál debería conservar el nombre y por qué.
#   Se queda con el nombre: el In Use comprado más recientemente (empate: el último en reportar).
#   Los demás vivos del grupo: indicios de equipo viejo (cuantos más, más probable):
#     está In Stock pero el nombre lo usa uno In Use · comprado antes que el que conserva ·
#     garantía vencida al corte · sin reportar > 90 días · la nota del otro dice que lo reemplazó ·
#     VM en otra plataforma de virtualización (migración). "Comprado antes" solo no alcanza para
#     marcarlo: queda "undetermined".
CORTE_TS = pd.Timestamp(FECHA_CORTE)
df["hostname_conflict_role"] = None
df["hostname_conflict_evidence"] = None
for host, g in df[df["flag_hostname_conflicto"]].groupby("hostname"):
    candidatos = g[g["asset_status"].eq("In Use")]
    candidatos = candidatos if len(candidatos) else g
    queda = candidatos.sort_values(["purchase_date", "last_seen"], ascending=False, na_position="last").index[0]
    df.at[queda, "hostname_conflict_role"] = "keeps the name"
    q = df.loc[queda]
    for f in g.index.drop(queda):
        r, ev = df.loc[f], []
        if r["asset_status"] == "In Stock" and q["asset_status"] == "In Use":
            ev.append("In Stock while the name is in use")
        if pd.notna(r["purchase_date"]) and pd.notna(q["purchase_date"]) and r["purchase_date"] < q["purchase_date"]:
            ev.append(f"older purchase ({r['purchase_date']:%Y} vs {q['purchase_date']:%Y})")
        if pd.notna(r["warranty_end_date"]) and r["warranty_end_date"] < CORTE_TS:
            ev.append("warranty expired")
        if pd.isna(r["last_seen"]) or (CORTE_TS - r["last_seen"]).days > 90:
            ev.append("not seen > 90 days")
        if isinstance(q["notes"], str) and "Replacement for asset" in q["notes"]:
            ev.append("the other device's note says it is a replacement")
        if r["asset_type"] == q["asset_type"] == "Virtual Machine" and r["manufacturer"] != q["manufacturer"]:
            ev.append(f"VM on another platform ({r['manufacturer']} vs {q['manufacturer']}): migration")
        # "más viejo" solo no alcanza (el que conserva es siempre el más nuevo): hace falta otro indicio
        fuertes = [e for e in ev if not e.startswith("older purchase")]
        df.at[f, "hostname_conflict_role"] = "probably replaced" if fuertes else "undetermined"
        df.at[f, "hostname_conflict_evidence"] = "; ".join(ev) or None
df["flag_hostname_probable_reemplazado"] = df["hostname_conflict_role"].eq("probably replaced")
print(df.loc[df["flag_hostname_conflicto"], ["hostname", "asset_tag", "asset_status", "purchase_date",
      "hostname_conflict_role", "hostname_conflict_evidence"]].sort_values(["hostname", "hostname_conflict_role"]).to_string())
# Inventario activo: lo que la empresa tiene hoy. Se excluyen los dados de baja/perdidos y lo que
# todavía no llegó (On Order). Missing sigue contando: es propio, solo falta ubicarlo.
df["en_inventario_activo"] = ~df["asset_status"].isin(DE_BAJA)
print(f"Equipos vivos con hostname en conflicto: {df['flag_hostname_conflicto'].sum()}")
print(f"En inventario activo: {df['en_inventario_activo'].sum()} de {len(df)}")
print("\n", dup_hosts["tipo"].value_counts().to_string())
print(dup_hosts.to_string(index=False))

# Tratamiento (acordado):
#   - duplicado exacto -> se deja la primera fila y se excluye la otra.
#   - mismo equipo, una fila más completa (donde las dos tienen dato, coinciden) -> se deja la que
#     tiene más campos llenos.
#   - mismo equipo con datos que chocan -> se deja la fila cuyo tag ya venía bien escrito en el crudo
#     (AST-######); si eso no decide, la que está In Use con last_seen (viene de SCCM). Los valores
#     distintos de la fila excluida se guardan en dup_valores_descartados. Si chocan campos que no
#     son de estado (ej. purchase_cost), se marca flag_revisar_duplicado para confirmar a mano.
#   - tag repetido en equipos distintos -> no se excluye; se marca la fila sospechosa (sección 21c).
# Toda exclusión queda en el log con la fila que se conservó.
CAMPOS_DE_ESTADO = {"asset_status", "last_seen", "book_value"}  # se resuelven eligiendo la fila

def tag_bien_en_crudo(f):
    return bool(re.fullmatch(r"AST-\d{6}", str(crudo.at[f, "Asset Tag"])))

def fila_a_conservar(filas):
    puntaje = {f: (tag_bien_en_crudo(f), df.at[f, "asset_status"] == "In Use", pd.notna(df.at[f, "last_seen"]))
               for f in filas}
    return max(filas, key=lambda f: puntaje[f])

df["dup_valores_descartados"] = None
df["flag_revisar_duplicado"] = False
excluir, conservar = [], {}
for _, g in dup_tags.iterrows():
    filas_g = g["filas_excel"]
    if g["tipo"] == "duplicado exacto":
        queda = min(filas_g)
    elif g["tipo"] == "mismo equipo, una fila más completa":
        queda = mas_completa(filas_g)
    elif g["tipo"] == "mismo equipo, datos distintos":
        queda = fila_a_conservar(filas_g)
        distintos = campos_distintos(df.loc[filas_g])
        otras = [f for f in filas_g if f != queda]
        df.at[queda, "dup_valores_descartados"] = "; ".join(
            f"{c}={df.at[f, c]} (fila {f})" for f in otras for c in distintos
            if pd.notna(df.at[f, c]) and str(df.at[f, c]) != str(df.at[queda, c]))
        df.at[queda, "flag_revisar_duplicado"] = bool(set(distintos) - CAMPOS_DE_ESTADO)
    else:
        continue
    for f in filas_g:
        if f != queda:
            excluir.append(f)
            conservar[f] = queda
            registrar(f, "(fila completa)", None, f"se conserva la fila {queda}", f"Excluida: {g['tipo']}")
print(df.loc[df["dup_valores_descartados"].notna(), ["asset_tag", "flag_revisar_duplicado", "dup_valores_descartados"]].to_string())

# 21c. Tag vs. número del hostname. Fuera de los duplicados, en TODOS los hostnames estándar
# (SITE-TIPO-#####) el número coincide con el del tag. Por eso, en los tags repetidos en equipos
# distintos, la fila cuyo hostname no coincide (o no se puede comparar) es la sospechosa.
num_host = df["hostname"].str.extract(r"^[A-Z]{3}-[A-Z]{2}-(\d{5})")[0]
coincide = num_host == df["asset_tag"].str[-5:]
comparable = num_host.notna() & df["asset_tag"].notna()
sin_dup = comparable & df["dup_tag_tipo"].isna()
print(f"Sin duplicados: el número del hostname coincide con el tag en {coincide[sin_dup].sum()} de {sin_dup.sum()}")
df["flag_tag_sospechoso"] = False
for _, g in dup_tags[dup_tags["tipo"] == "tag repetido en equipos distintos"].iterrows():
    ok = [f for f in g["filas_excel"] if coincide[f]]
    sospechosas = [f for f in g["filas_excel"] if f not in ok] if ok else g["filas_excel"]
    df.loc[sospechosas, "flag_tag_sospechoso"] = True
print(df.loc[df["dup_tag_tipo"].eq("tag repetido en equipos distintos"),
             ["asset_tag", "hostname", "flag_tag_sospechoso"]].sort_values("asset_tag").to_string())

# Decidido (revisión): ningún tag queda repetido en equipos distintos.
#   - La fila sospechosa con hostname estándar toma el tag de su hostname si está libre
#     (flag_tag_reasignado).
#   - Si sigue repetido y no hay de dónde deducir el propio (hostname legado / VM / monitor sin
#     hostname) -> vacío + flag_tag_vaciado: hay que ponerle etiqueta nueva. El tag original queda
#     en el registro de cambios.
distintos = df["dup_tag_tipo"].eq("tag repetido en equipos distintos")
candidato = "AST-0" + num_host
reasignar = df["flag_tag_sospechoso"] & num_host.notna() & ~candidato.isin(set(df["asset_tag"].dropna()))
df["flag_tag_reasignado"] = reasignar
aplicar(df, "asset_tag", df["asset_tag"].where(~reasignar, candidato),
        "asset_tag: repetido en otro equipo -> número del hostname")
sigue = distintos & df["asset_tag"].notna() & df["asset_tag"].duplicated(keep=False) & ~df.index.isin(excluir)
vaciar = sigue & df["flag_tag_sospechoso"]
df["flag_tag_vaciado"] = vaciar
aplicar(df, "asset_tag", df["asset_tag"].where(~vaciar),
        "asset_tag: repetido en otro equipo sin forma de deducir el propio -> vacío (etiqueta nueva)")
df["flag_tag_sospechoso"] = False  # ya resueltos: reasignados o vaciados
quedan = df.loc[~df.index.isin(excluir), "asset_tag"].dropna()
assert not quedan.duplicated().any(), quedan[quedan.duplicated(keep=False)].tolist()
print(f"Tags repetidos resueltos: {reasignar.sum()} reasignados desde el hostname, {vaciar.sum()} vaciados")

# Tag vacío con hostname estándar -> se deduce del número del hostname (acordado, dada la
# coincidencia del 100%) y se marca. No se usa si ese tag ya existe en otra fila.
deducido = ("AST-0" + num_host).where(df["asset_tag"].isna() & num_host.notna())
choca = deducido.isin(df["asset_tag"].dropna())
assert not choca.any(), f"Tags deducidos que ya existen: {deducido[choca].tolist()}"
df["flag_tag_desde_hostname"] = deducido.notna()
aplicar(df, "asset_tag", df["asset_tag"].fillna(deducido), "asset_tag: vacío -> deducido del hostname")
print("Tags deducidos del hostname:", df.loc[df["flag_tag_desde_hostname"], ["asset_tag", "hostname"]].to_string(), sep="\n")
print("asset_tag que siguen vacíos:", df["asset_tag"].isna().sum())

# CSV para revisar a mano: todas las filas con asset_tag u hostname repetido (antes de excluir),
# con el valor crudo al lado del limpio para ver si la repetición ya venía en el crudo o apareció
# al homogeneizar (ej. 'AST-4081' y 'AST-004081' recién son iguales después de limpiar).
COLS_REVISION = ["asset_tag", "hostname", "serial_number", "asset_type", "manufacturer", "model",
                 "asset_status", "assigned_to", "location", "purchase_date", "last_seen",
                 "purchase_cost", "currency", "book_value", "mac_address", "ip_address", "review_flag", "notes"]
partes = []
for grupo, tabla, col_clave in [("asset_tag", dup_tags, "asset_tag"), ("hostname", dup_hosts, "hostname")]:
    for _, g in tabla.iterrows():
        bloque = df.loc[g["filas_excel"], COLS_REVISION].copy()
        bloque.insert(0, "grupo", grupo)
        bloque.insert(1, "valor_repetido", g[col_clave])
        bloque.insert(2, "tipo_repeticion", g["tipo"])
        bloque.insert(3, "excluida", bloque.index.isin(excluir))
        bloque.insert(4, "fila_conservada", bloque.index.map(lambda f: conservar.get(f)).astype("Int64"))
        bloque.insert(5, "campos_distintos", ", ".join(campos_distintos(df.loc[g["filas_excel"]])))
        bloque["flag_tag_sospechoso"] = df.loc[bloque.index, "flag_tag_sospechoso"]
        bloque["flag_revisar_duplicado"] = df.loc[bloque.index, "flag_revisar_duplicado"]
        bloque["dup_valores_descartados"] = df.loc[bloque.index, "dup_valores_descartados"]
        bloque["asset_tag_crudo"] = crudo.loc[bloque.index, "Asset Tag"].map(repr)
        bloque["hostname_crudo"] = crudo.loc[bloque.index, "hostname"].map(repr)
        bloque["serial_crudo"] = crudo.loc[bloque.index, "Serial #"].map(repr)
        partes.append(bloque)
revision_dup = pd.concat(partes).rename_axis("fila_excel").reset_index()
crudo_igual = revision_dup.groupby(["grupo", "valor_repetido"])[["asset_tag_crudo", "hostname_crudo"]].nunique()
revision_dup["repetido_ya_en_crudo"] = revision_dup.set_index(["grupo", "valor_repetido"]).index.map(
    lambda k: crudo_igual.loc[k, "asset_tag_crudo" if k[0] == "asset_tag" else "hostname_crudo"] == 1)
from config import ELABORADO  # noqa: E402
revision_dup.to_csv(ELABORADO / "activos_ti_2026_repetidos_revision.csv", index=False, encoding="utf-8-sig")
print(revision_dup.groupby(["grupo", "tipo_repeticion"])["valor_repetido"].nunique().to_string())
print("Grupos cuya repetición ya venía en el crudo (mismo texto):",
      revision_dup.drop_duplicates(["grupo", "valor_repetido"]).groupby("grupo")["repetido_ya_en_crudo"].agg(["sum", "size"]).to_string())

df = df.drop(index=excluir)
print(f"\nFilas excluidas por duplicado: {len(excluir)} -> activos: {len(df)}")

# %% 21d. Datos imposibles para el tipo -----------------------------------------------------------
# Decidido (revisión): SO / RAM en monitores e impresoras y encrypted en monitores, impresoras o
# firewalls no pueden existir: parecen datos de la PC conectada u otro equipo -> vacío +
# flag_dato_tipo_no_aplica. Se conservan (y se marcan en la revisión) last_seen y MAC en monitores
# y MAC en VMs: pueden ser reales (inventario vía la PC, dock USB-C, MAC virtual).
IMPOSIBLE = {"os": ["os", "os_family", "os_version"], "ram_gb": ["ram_gb"], "encrypted": ["encrypted"]}
df["flag_dato_tipo_no_aplica"] = False
for campo, columnas_campo in IMPOSIBLE.items():
    mal = df[campo].notna() & ~df["asset_type"].isin(APLICA_A[campo])
    print(f"{campo} en tipo donde no aplica -> vacío: {mal.sum()}", df.loc[mal, "asset_type"].value_counts().to_dict())
    df["flag_dato_tipo_no_aplica"] |= mal
    for c in columnas_campo:
        aplicar(df, c, df[c].where(~mal), f"{campo}: dato en tipo donde no aplica -> vacío")

# %% 22. Verificación de la limpieza superficial ---------------------------------------------------
# Tipos esperados después de limpiar (las fechas ya son Timestamp, los montos float, etc.)
TIPO_LIMPIO = {"fecha": pd.Timestamp, "fecha_hora": pd.Timestamp, "numero": (int, float),
               "entero": (int, float), "texto": str, "categoria": str}
problemas = {}
for col in [c for c in df.columns if c in TIPO_ESPERADO]:
    llenos = df[col].dropna()
    malos = llenos[~llenos.map(lambda v: isinstance(v, TIPO_LIMPIO[TIPO_ESPERADO[col]]))]
    if len(malos):
        problemas[col] = malos.map(repr).unique()[:5].tolist()
print("Columnas con tipos fuera de lo esperado:", problemas or "ninguna")
print(tabla_completitud(df).to_string())

# Identificadores únicos (sin contar vacíos)
for col in ["asset_tag", "serial_number"]:
    repetidos = df[col].dropna()[df[col].dropna().duplicated(keep=False)]
    print(f"{col}: {repetidos.nunique()} valores repetidos {sorted(repetidos.unique())[:10]} | vacíos: {df[col].isna().sum()}")
print("hostname: repetidos en equipos distintos ->", df["dup_hostname_tipo"].value_counts().to_dict())
print(pd.crosstab(df["manufacturer"], df["asset_type"]).to_string())  # tipo vs marca: ¿coherente?

banderas = [c for c in df.columns if c.startswith("flag_")]
print("\nBanderas:\n", df[banderas].sum().sort_values(ascending=False).to_string())
print(f"\nRegistro de cambios: {len(registro)} entradas")
print(pd.DataFrame(registro)["regla"].value_counts().to_string())

# %% 23. Texto libre ------------------------------------------------------------------------------
print(df["notes"].dropna().value_counts().head(25).to_string())

# Verificación: después de la sección 5 no debería quedar mojibake
textos = df.select_dtypes(include="object").astype(str)
mojibake = textos.apply(lambda s: s.str.contains("Ã|Â|√|â€", regex=True).sum())
print("\nCeldas con mojibake por columna:\n", mojibake[mojibake > 0])

# PREGUNTAS: ¿las notas contradicen otros campos? (ej. "battery" en una VM, "lease ends" vencido)

# %% 24. Tu espacio: cruces entre campos ---------------------------------------------------------
# Ideas para probar:
#  - asset_status "In Stock" con persona asignada
#  - email sin "@" o que no coincide con el nombre en assigned_to
#  - country vs location vs currency
#  - Sistema operativo incompatible con el modelo/fabricante
#  - Activos "In Use" que no se ven en SCCM hace > 90 días
#  - Filas con review_flag ("dup?", "verify"): ¿qué problema tienen?
# pd.crosstab(df["asset_type"], df["asset_status"])

# %% 25. Notas de limpieza por fila -------------------------------------------------------------
# Resume en cada fila las decisiones de la limpieza que se basaron en un supuesto o una deducción,
# para poder filtrar en Excel lo que conviene corroborar. Los cambios de forma (mayúsculas,
# formatos, catálogos) no entran: están en el registro de cambios.
#   notas_limpieza:   texto con cada decisión aplicada a la fila
#   revisar_limpieza: True si la fila tiene al menos una decisión para corroborar
DECISIONES = {  # bandera -> nota
    "flag_tag_desde_hostname": "asset_tag deducido del hostname",
    "flag_tag_reasignado": "asset_tag estaba repetido en otro equipo; se tomó el número del hostname",
    "flag_tag_vaciado": "asset_tag estaba repetido en otro equipo y no se puede deducir -> vacío; poner etiqueta nueva",
    "flag_dato_tipo_no_aplica": "os / ram / encrypted en un tipo donde es imposible -> vacío (dato de otro equipo)",
    "flag_serial_placeholder": "serial_number de relleno (0 / UNKNOWN / SEE STICKER / TO BE FILLED) -> vacío",
    "flag_manufacturer_from_model": "manufacturer no correspondía al modelo; se usó la marca del modelo",
    "flag_department_cc": "department venía como código de centro de costos (pasó a cost_center); completar el área",
    "flag_revisar_duplicado": "duplicado resuelto; confirmar valor descartado (ver dup_valores_descartados)",
    "flag_hostname_conflicto": "hostname compartido con otro equipo vivo; revisar cuál conserva el nombre",
    "flag_hostname_probable_reemplazado": "hostname en conflicto: probablemente reemplazado sin dar de baja (ver evidencia)",
    "flag_location_desde_hostname": "location deducida del código de sitio del hostname (venía HQ/DC1)",
    "flag_country_desde_location": "country deducido de location",
    "flag_status_eol": "asset_status EOL llevado a Retired",
    "flag_department_shared": "department con dos áreas -> Shared",
    "flag_fecha_ambigua": "purchase_date ambigua dd/mm-mm/dd; se asumió mm/dd",
    "flag_compra_reconstruida": "purchase_date con año imposible; reconstruida como garantía - años del tipo (o vacía)",
    "flag_last_seen_invalido": "last_seen 09:00 exacto y anterior a la compra -> vacío",
    "flag_fecha_aproximada": "purchase_date venía como trimestre; se usó el día 15 del mes del medio",
    "flag_garantia_derivada": "warranty_end_date calculada como compra + N años",
    "flag_garantia_estimada": "warranty_end_date vacía; estimada como compra + años estándar del tipo (supuesto)",
    "flag_garantia_contradictoria": "garantía 'Expired' en el crudo pero compra reciente; queda vacía, confirmar",
    "flag_costo_negativo": "purchase_cost negativo; se usó el valor absoluto",
    "flag_moneda_inferida": "currency inferida (del monto o del país)",
    "flag_moneda_monto_distinta": "el monto traía otra moneda que currency; revisar cuál es la correcta",
    "flag_costo_cero": "purchase_cost 0 en equipo físico (con PO); se dejó vacío, completar el costo real",
    "flag_costo_atipico": "purchase_cost > 3 veces la mediana del tipo (USD); quitado (ver purchase_cost_descartado y causa_costo_atipico)",
    "flag_nombre_desde_email": "assigned_to tomado del email (venía usuario, inicial o vacío en In Use)",
    "flag_nombre_invertido": "assigned_to venía 'Apellido Nombre'; se invirtió porque coincide con el email",
    "flag_criticality_invalida": "business_criticality fuera de 1-5 -> vacío",
    "flag_ram_invalida": "ram_gb no plausible -> vacío",
    "flag_incidentes_invalido": "open_incidents negativo -> vacío",
    "flag_mac_invalida": "mac_address truncada -> vacío",
    "flag_ip_invalida": "ip_address inválida -> vacío",
}
df["notas_limpieza"] = df[list(DECISIONES)].apply(
    lambda f: "; ".join(nota for bandera, nota in DECISIONES.items() if f[bandera]) or None, axis=1)
df["revisar_limpieza"] = df["notas_limpieza"].notna()
print(f"Filas con alguna decisión para corroborar: {df['revisar_limpieza'].sum()} de {len(df)}")
print(df[list(DECISIONES)].sum().sort_values(ascending=False).to_string())

# %% 26. Guardar base intermedia -----------------------------------------------------------------
# Resultado de la limpieza superficial: base + diccionario + log de cambios (elaborado/).
# Las contradicciones se revisan en codigo/02_limpiar/revisar_contradicciones.py a partir de acá.
from config import BASE_LIMPIA, REGISTRO_CAMBIOS  # noqa: E402
from campos import CAMPOS, DERIVADAS  # noqa: E402

BASE_LIMPIA.parent.mkdir(parents=True, exist_ok=True)
df.index.name = "fila_excel"
df.to_pickle(BASE_LIMPIA.with_suffix(".pkl"))
df.to_csv(BASE_LIMPIA.with_suffix(".csv"), encoding="utf-8-sig")
pd.DataFrame(registro).to_csv(REGISTRO_CAMBIOS, index=False, encoding="utf-8-sig")

dicc = [{"variable": "fila_excel", "descripcion": "Número de fila en el Excel crudo (trazabilidad).",
         "tipo": "entero", "origen": "índice", "nivel": "activo", "años": "2026 (corte 2026-10-01)"}]
for orig, limpio, tipo, _, definicion in CAMPOS:
    dicc.append({"variable": limpio, "descripcion": definicion, "tipo": tipo,
                 "origen": f"crudo: '{orig}'", "nivel": "activo", "años": "2026 (corte 2026-10-01)"})
for col, (tipo, definicion) in DERIVADAS.items():
    dicc.append({"variable": col, "descripcion": definicion, "tipo": tipo,
                 "origen": "derivada en la limpieza", "nivel": "activo", "años": "2026 (corte 2026-10-01)"})
dicc = pd.DataFrame(dicc)
faltan = set(df.columns) - set(dicc["variable"])
assert not faltan, f"Columnas sin definir en el diccionario: {faltan}"
dicc.to_csv(BASE_LIMPIA.parent / f"{BASE_LIMPIA.name}_dicc.csv", index=False, encoding="utf-8-sig")
print(f"Guardado: {BASE_LIMPIA.name}.pkl/.csv ({len(df)} filas x {df.shape[1]} col.), _dicc.csv, registro de cambios")
