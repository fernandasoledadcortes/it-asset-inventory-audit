# -*- coding: utf-8 -*-
"""Diccionario de campos del registro de activos.

Una entrada por columna del crudo: nombre original -> nombre limpio, qué significa, qué tipo
se espera y a qué tipos de activo aplica. Sirve para renombrar, para revisar tipos y completitud,
y después es la base de la pestaña Description y del _dicc del dataset limpio.

Nombres limpios en inglés snake_case porque la entrega es en inglés.
"""

# Catálogo de tipos de activo (el de la fila 2 del crudo, que sí coincide con el instructivo)
LAPTOP, DESKTOP, MONITOR, SERVER, VM = "Laptop", "Desktop", "Monitor", "Server", "Virtual Machine"
PHONE, TABLET, PRINTER, SWITCH, FIREWALL = "Mobile Phone", "Tablet", "Printer", "Network Switch", "Firewall"
TIPOS_ACTIVO = [LAPTOP, DESKTOP, MONITOR, SERVER, VM, PHONE, TABLET, PRINTER, SWITCH, FIREWALL]

# Grupos útiles para decir a qué aplica cada campo
TODOS = TIPOS_ACTIVO
FISICOS = [t for t in TIPOS_ACTIVO if t != VM]                      # se compran, tienen garantía
USUARIO_FINAL = [LAPTOP, DESKTOP, PHONE, TABLET]                     # se le entregan a una persona
CON_RESPONSABLE = USUARIO_FINAL + [MONITOR, PRINTER]                 # pueden tener persona asignada
CON_SO = [LAPTOP, DESKTOP, SERVER, VM, PHONE, TABLET]                # tienen sistema operativo
GESTIONADOS = CON_SO                                                 # reportan a SCCM / Intune
EN_RED = [t for t in TIPOS_ACTIVO if t != MONITOR]                   # se conectan a la red
CON_TARJETA_RED = [t for t in FISICOS if t != MONITOR]               # tienen MAC física
CIFRABLES = USUARIO_FINAL + [SERVER, VM]                             # BitLocker / FileVault / móvil / disco de servidor

# tipo: texto | entero | numero | fecha | fecha_hora | categoria
CAMPOS = [
    # original, limpio, tipo, aplica_a, definición
    ("Asset Tag", "asset_tag", "texto", TODOS,
     "Etiqueta de inventario que pone TI (sticker). Es el ID interno del activo en este registro; "
     "debería ser único y con formato AST-######."),
    ("hostname", "hostname", "texto", [t for t in TODOS if t != MONITOR],
     "Nombre del equipo en la red. Estándar documentado: SITE-TYPE-##### (ej. MNL-LT-05932)."),
    ("Serial #", "serial_number", "texto", TODOS,
     "Número de serie del fabricante, grabado en el equipo. Puede ser solo dígitos (en celulares "
     "suele ser el IMEI, 15 dígitos) o el ID de instancia en la nube para VMs (i-…). Se guarda como "
     "texto: si se guarda como número se pierden ceros a la izquierda."),
    ("ASSET TYPE", "asset_type", "categoria", TODOS,
     "Categoría del activo (catálogo de 10 tipos)."),
    ("Manufacturer ", "manufacturer", "categoria", TODOS,
     "Marca que fabrica el equipo (Dell, Apple…) o proveedor de nube para VMs (Azure, Amazon)."),
    ("model", "model", "texto", TODOS, "Modelo comercial del equipo o tamaño de la VM."),
    ("OS", "os", "categoria", CON_SO,
     "Sistema operativo y versión. Monitores, impresoras y equipos de red no tienen SO (tienen firmware)."),
    ("Status", "asset_status", "categoria", TODOS,
     "Estado del activo en su ciclo de vida: In Use, In Stock, In Repair, Retired, Disposed, Missing, "
     "Lost/Stolen, On Order."),
    ("Assigned To", "assigned_to", "texto", CON_RESPONSABLE,
     "Persona que tiene el equipo. Aplica a equipos de usuario final y periféricos (monitores, "
     "impresoras); servidores, VMs y equipos de red no tienen persona. En bodega (In Stock) es normal que falte."),
    ("user email", "user_email", "texto", CON_RESPONSABLE, "Correo de la persona asignada."),
    ("Dept / Cost Center", "department", "categoria", TODOS,
     "Área / centro de costos que usa y paga el activo. Viene mezclado: nombres (Finance) y códigos (CC-1200)."),
    ("LOCATION", "location", "categoria", TODOS, "Sitio u oficina donde está el activo."),
    ("country", "country", "categoria", TODOS, "País del sitio. Si falta, se puede deducir de location."),
    ("Purchase Date", "purchase_date", "fecha", TODOS,
     "Fecha de recepción del equipo (fila 2: 'date received'). En VMs, fecha de alta."),
    ("warranty_end", "warranty_end_date", "fecha", FISICOS,
     "Fin de la garantía del fabricante. Las VMs no tienen garantía."),
    ("Purchase Cost", "purchase_cost", "numero", FISICOS,
     "Costo de compra en moneda local (ver currency). Solo físicos: las VMs se pagan por suscripción "
     "(una VM con costo se marca en la revisión, chequeo vm_con_costo)."),
    ("Curr.", "currency", "categoria", TODOS, "Moneda del costo, código ISO de 3 letras (USD, MXN…)."),
    ("Book Value (calc)", "book_value", "numero", FISICOS,
     "Valor contable hoy: costo menos depreciación lineal = costo × (1 − días desde compra / vida útil), "
     "mínimo 0. Es una fórmula; solo está en ~175 filas y la vida útil no es consistente (4 vs 5 años)."),
    ("Unnamed: 18", "review_flag", "texto", [],
     "Columna sin encabezado con marcas de alguien que revisó el registro (dup?, verify, check w/ Mike). "
     "No es un atributo del activo: es una pista de problemas conocidos."),
    ("Last Seen (SCCM)", "last_seen", "fecha_hora", EN_RED,
     "Última vez que el equipo se reportó a las herramientas de gestión (SCCM / Intune). Es la prueba "
     "de que el equipo existe, está prendido y está gestionado (los equipos de red también reportan a "
     "monitoreo). Un monitor no puede reportarse solo. Mucho tiempo sin reportar = posible equipo "
     "perdido, apagado o fuera de control."),
    ("RAM", "ram_gb", "numero", CON_SO, "Memoria RAM en GB."),
    ("Storage", "storage_gb", "numero", CON_SO,
     "Capacidad de disco en GB (1 TB = 1024 GB). El tipo de disco (SSD/HDD) va en storage_type."),
    ("IP", "ip_address", "texto", EN_RED,
     "Dirección IP (la última conocida). En laptops y celulares cambia (DHCP), así que es poco confiable "
     "como identificador. Los monitores no tienen."),
    ("MAC Address", "mac_address", "texto", CON_TARJETA_RED,
     "Identificador físico de la tarjeta de red (6 pares hex). No aplica a monitores; en VMs es virtual "
     "y no se suele registrar."),
    ("Business Criticality", "business_criticality", "entero", TODOS,
     "Impacto si el activo falla, escala 1–5 (fila 2): 1 Low … 5 Critical/Business-wide."),
    ("encrypted?", "encrypted", "categoria", CIFRABLES,
     "Disco cifrado (Y = BitLocker en Windows / FileVault en macOS; en móviles, cifrado del dispositivo). "
     "Vacío = DESCONOCIDO, no 'no': "
     "para la auditoría no se puede suponer ni que sí ni que no."),
    ("Status.1", "patch_compliance", "categoria", GESTIONADOS,
     "Segunda columna 'Status' del crudo: cumplimiento de parches/actualizaciones del SO según SCCM/Intune "
     "(Compliant, Non-Compliant, Pending Reboot). Distinto de asset_status."),
    ("Vendor", "vendor", "categoria", FISICOS,
     "Revendedor que vendió el equipo (CDW, Insight…). No es lo mismo que el fabricante."),
    ("PO #", "po_number", "texto", FISICOS,
     "Número de orden de compra. Todo lo comprado debería tenerlo; las VMs se facturan por suscripción."),
    ("Open Incidents", "open_incidents", "entero", TODOS, "Tickets de soporte abiertos del activo."),
    ("Notes", "notes", "texto", [], "Texto libre. Opcional; puede contradecir otros campos."),
]

RENOMBRAR = {orig: limpio for orig, limpio, *_ in CAMPOS}
TIPO_ESPERADO = {limpio: tipo for _, limpio, tipo, *_ in CAMPOS}
APLICA_A = {limpio: aplica for _, limpio, _, aplica, _ in CAMPOS}
DEFINICION = {limpio: d for _, limpio, *_, d in CAMPOS}

# Estados en los que se espera el dato (además del tipo de APLICA_A). Si un campo no está acá,
# se espera en cualquier estado. Sirve para medir vacíos reales: un equipo en bodega sin persona o
# un pedido sin MAC no es un dato faltante.
EN_USO = ["In Use", "In Repair"]                                     # está en manos de alguien
DESPLEGADOS = ["In Use", "In Repair", "Missing", "Lost/Stolen", "Retired", "Disposed"]  # llegó a usarse
ESTADOS_DONDE_APLICA = {
    "assigned_to": EN_USO,          # en bodega, pedido o dado de baja es normal que no tenga persona
    "user_email": EN_USO,
    "patch_compliance": ["In Use"],  # SCCM/Intune solo lo informa para equipos activos
    "ip_address": ["In Use"],        # IP actual: solo tiene sentido si está conectado
    "last_seen": DESPLEGADOS,        # un pedido o un equipo nuevo en bodega aún no se conectó
    "hostname": DESPLEGADOS + ["In Stock"],   # el pedido todavía no tiene nombre
    "mac_address": DESPLEGADOS + ["In Stock"],
    "encrypted": DESPLEGADOS + ["In Stock"],
}

# Tipos donde el campo PUEDE venir (no es un error) pero no se exige: su vacío no cuenta como faltante.
NO_EXIGIDO_EN = {
    "last_seen": [PRINTER],   # sin agente SCCM/Intune; solo reporta si hay monitoreo de red
    "book_value": TODOS,      # es una fórmula, no un dato: se recalcula en book_value_calc
    "encrypted": [SERVER, VM],  # puede venir (discos de servidor cifrados) pero el control es en equipos de usuario
}

# Columnas que agrega la limpieza (no están en el crudo): definición y regla de cálculo
DERIVADAS = {
    "hostname_convention": ("categoria", "Convención de nombre del hostname: estandar (SITE-TIPO-#####), "
                            "servidor (SITE+ROL+##), red (SITE-SW|FW-ZONA##), legado (CIUDAD-TIPO-###)."),
    "cost_center": ("texto", "Código de centro de costos (CC-####) que venía en department."),
    "os_family": ("categoria", "Familia del sistema operativo, derivada de os."),
    "os_version": ("categoria", "Familia + versión principal (ej. Windows 10), derivada de os."),
    "storage_type": ("categoria", "Tipo de disco SSD / HDD, separado del texto original de Storage."),
    "dup_tag_tipo": ("categoria", "Si el asset_tag se repite: mismo equipo con datos distintos / "
                     "tag repetido en equipos distintos."),
    "dup_hostname_tipo": ("categoria", "Si el hostname se repite: nombre reutilizado (uno vivo) / "
                          "conflicto (varios vivos)."),
    "flag_hostname_old": ("bandera", "Hostname con sufijo -OLD (equipo renombrado / reemplazado)."),
    "flag_status_eol": ("bandera", "asset_status venía como 'EOL' y se llevó a Retired."),
    "flag_department_shared": ("bandera", "department tenía dos áreas distintas -> 'Shared'."),
    "flag_criticality_invalida": ("bandera", "business_criticality fuera de la escala 1-5 -> vacío."),
    "flag_fecha_ambigua": ("bandera", "purchase_date dd/mm vs mm/dd ambigua; se asumió mm/dd."),
    "flag_fecha_aproximada": ("bandera", "purchase_date venía como trimestre; se usó el día 15 del mes del medio."),
    "flag_compra_reconstruida": ("bandera", "purchase_date con año imposible (posterior al corte o anterior a 2010): "
                                 "se reconstruyó como fin de garantía - años estándar del tipo; vacía si no había garantía propia."),
    "flag_last_seen_invalido": ("bandera", "last_seen a las 09:00:00 exactas y anterior a la compra: valor por "
                                "defecto de otra fuente, no un reporte -> vacío."),
    "flag_garantia_derivada": ("bandera", "warranty_end_date calculada como compra + N años (venía en texto)."),
    "flag_never_seen": ("bandera", "last_seen = 'Never': el equipo nunca se reportó."),
    "flag_costo_negativo": ("bandera", "purchase_cost era negativo; se usó el valor absoluto."),
    "flag_moneda_inferida": ("bandera", "currency vacía; se tomó del monto o de la moneda local del país."),
    "flag_garantia_contradictoria": ("bandera", "warranty_end venía 'Expired' pero compra + años estándar sigue vigente; queda vacía."),
    "flag_garantia_estimada": ("bandera", "warranty_end_date venía vacía/TBD/Expired; estimada = compra + años estándar del tipo."),
    "flag_costo_cero": ("bandera", "purchase_cost era 0 en un equipo físico (tiene PO); se dejó vacío."),
    "purchase_cost_usd": ("número", "purchase_cost × tipo de cambio a USD (config.FX). Derivada de purchase_cost + currency."),
    "flag_costo_atipico": ("bandera", "purchase_cost en USD > 3 veces la mediana de su asset_type (sin VM); se quitó el monto."),
    "purchase_cost_descartado": ("número", "Monto original (moneda original) de los costos atípicos que se quitaron de purchase_cost."),
    "causa_costo_atipico": ("texto", "Causa probable del costo atípico: moneda local cargada como USD / decimal perdido (x100) / sin causa clara."),
    "flag_moneda_monto_distinta": ("bandera", "el monto traía pegada otra moneda distinta de currency; se dejó currency."),
    "flag_ram_invalida": ("bandera", "ram_gb no plausible (1600/3200 = MHz) -> vacío."),
    "flag_incidentes_invalido": ("bandera", "open_incidents negativo -> vacío."),
    "flag_mac_invalida": ("bandera", "mac_address sin 12 dígitos hex (truncada) -> vacío."),
    "flag_ip_invalida": ("bandera", "ip_address con octeto > 255 -> vacío."),
    "flag_email_no_corporativo": ("bandera", "user_email con dominio distinto de contoso-corp.com."),
    "flag_user_terminated": ("bandera", "assigned_to venía marcado '(Terminated)': ex-empleado."),
    "flag_nombre_desde_email": ("bandera", "assigned_to era usuario/inicial (o vacío en un equipo In Use); "
                                "se tomó el nombre del email."),
    "flag_nombre_invertido": ("bandera", "assigned_to venía 'Apellido Nombre' sin coma; se invirtió porque así "
                              "coincide exacto con el nombre del email."),
    "dup_valores_descartados": ("texto", "Valores de la fila duplicada excluida que no coincidían con la conservada."),
    "flag_revisar_duplicado": ("bandera", "Duplicado resuelto, pero chocaba un campo que no es de estado "
                               "(ej. purchase_cost): confirmar el valor a mano."),
    "flag_tag_sospechoso": ("bandera", "asset_tag repetido en otro equipo y el número del hostname no "
                            "coincide con el tag (o no se puede comparar)."),
    "flag_tag_reasignado": ("bandera", "asset_tag repetido en otro equipo: se reemplazó por el número del hostname "
                            "estándar (el tag nuevo estaba libre)."),
    "flag_tag_vaciado": ("bandera", "asset_tag repetido en otro equipo y sin forma de deducir el propio -> vacío: "
                         "requiere etiqueta nueva. El valor original está en el registro de cambios."),
    "flag_serial_placeholder": ("bandera", "serial_number de relleno (0, UNKNOWN, SEE STICKER, TO BE FILLED) -> vacío."),
    "flag_manufacturer_from_model": ("bandera", "manufacturer no correspondía al modelo; se usó la marca más frecuente "
                                     "del modelo en la base (el SO lo confirma)."),
    "flag_department_cc": ("bandera", "department venía como código de centro de costos (CC-####): se movió a "
                           "cost_center y department queda vacío para completar."),
    "flag_dato_tipo_no_aplica": ("bandera", "Tenía os / ram_gb / encrypted en un tipo donde es imposible (monitor, "
                                 "impresora, firewall): parecen datos de otro equipo -> vacío. Original en el registro de cambios."),
    "flag_tag_desde_hostname": ("bandera", "asset_tag vacío deducido del número del hostname estándar "
                                "(AST-0 + #####; coincide en el 100% de los casos comparables)."),
    "flag_hostname_conflicto": ("bandera", "Equipo vivo (no Retired/Disposed/Lost-Stolen/On Order) cuyo hostname "
                                "lo usa también otro equipo vivo: revisar cuál conserva el nombre."),
    "hostname_conflict_role": ("categoria", "En un conflicto de hostname (varios vivos con el mismo nombre): "
                               "keeps the name (In Use más reciente) / probably replaced (hay indicios de registro viejo) / undetermined."),
    "hostname_conflict_evidence": ("texto", "Indicios de que el equipo es el registro viejo: In Stock con el nombre en uso, "
                                   "compra anterior, garantía vencida, sin reportar > 90 días, nota de reemplazo, migración de VM."),
    "flag_hostname_probable_reemplazado": ("bandera", "Conflicto de hostname: este equipo probablemente fue reemplazado y "
                                           "no se dio de baja (ver hostname_conflict_evidence)."),
    "en_inventario_activo": ("bandera", "El equipo es parte del inventario actual: todo salvo Retired, Disposed, "
                             "Lost/Stolen y On Order (Missing sí cuenta). Filtro para conteos y costos."),
    "flag_location_desde_hostname": ("bandera", "location venía ambigua (HQ / DC1) y se dedujo del código de sitio del hostname."),
    "flag_country_desde_location": ("bandera", "country vacío deducido del sitio (location)."),
    "notas_limpieza": ("texto", "Decisiones de limpieza basadas en supuestos o deducciones aplicadas a la fila "
                       "(una por bandera flag_*)."),
    "revisar_limpieza": ("bandera", "La fila tiene al menos una decisión de limpieza para corroborar."),
    # Eventos extraídos de notes (plantillas fijas)
    "evt_malware": ("bandera", "Nota: equipo re-imaginado tras una alerta de malware."),
    "malware_reimage_date": ("fecha", "Fecha de la re-imagen por malware (de notes, mm/dd/aaaa)."),
    "malware_incident": ("texto", "Ticket del incidente de malware (INC…)."),
    "previous_owner": ("texto", "Dueño anterior del equipo según la nota de malware."),
    "moved_from": ("categoria", "Sitio de origen del traslado según la nota de malware."),
    "moved_to": ("categoria", "Sitio de destino del traslado según la nota de malware."),
    "evt_lease": ("bandera", "Nota: equipo en leasing (debe devolverse al arrendador)."),
    "lease_end_date": ("fecha", "Fin del leasing (de notes, mm/dd/aaaa)."),
    "evt_battery_swelling": ("bandera", "Nota: el usuario reportó batería hinchada."),
    "battery_incident": ("texto", "Ticket del reporte de batería (INC… / RITM…)."),
    "evt_replacement_failure": ("bandera", "Nota: el equipo reemplaza a otro que falló."),
    "failed_component": ("categoria", "Componente que falló en el equipo reemplazado."),
    "replacement_incident": ("texto", "Ticket del reemplazo (INC…)."),
    "evt_loaner": ("bandera", "Nota: equipo del pool de préstamo (Service Desk)."),
    "evt_shared": ("bandera", "Nota: equipo compartido (sala de reuniones)."),
    "evt_physical_audit": ("bandera", "Nota: verificado en una auditoría física."),
    "audit_month": ("fecha", "Mes de la auditoría física (primer día del mes)."),
    "evt_refresh": ("bandera", "Nota: entregado en un programa de renovación de hardware."),
    "refresh_fy": ("entero", "Año fiscal del programa de renovación (FY21 -> 2021)."),
    "evt_hosts_service": ("bandera", "Nota: el equipo aloja un servicio (print, file shares, SQL, ERP)."),
    "hosted_service": ("categoria", "Servicio que aloja según la nota."),
    "service_team": ("categoria", "Área a la que da servicio según la nota."),
}

# Para la pestaña de datos del Excel: celda vacía donde el campo no corresponde (por tipo o por
# estado) -> "Not applicable". Así un vacío queda solo para el dato que falta de verdad.
# Solo para mostrar: las bases de trabajo (.pkl) conservan el vacío para no romper tipos.
NO_APLICA = "Not applicable"
DERIVADA_DE = {"os_family": "os", "os_version": "os", "storage_type": "storage_gb",
               "book_value_calc": "purchase_cost", "hostname_convention": "hostname"}

def marcar_no_aplica(df):
    """Copia de df (como texto donde haga falta) con NO_APLICA en las celdas vacías que no corresponden."""
    salida = df.copy()
    for col in salida.columns:
        base = DERIVADA_DE.get(col, col)
        tipos = APLICA_A.get(base)
        if not tipos:
            continue  # sin regla (notes, review_flag, banderas): vacío = vacío
        exigidos = [t for t in tipos if t not in NO_EXIGIDO_EN.get(base, [])]
        if not exigidos:
            continue  # book_value del crudo: no se exige en ningún tipo, su vacío es "no informado"
        estados = ESTADOS_DONDE_APLICA.get(base)
        corresponde = salida["asset_type"].isin(exigidos)
        if estados:
            corresponde &= salida["asset_status"].isin(estados)
        no_aplica = salida[col].isna() & ~corresponde
        if no_aplica.any():
            salida[col] = salida[col].astype(object).where(~no_aplica, NO_APLICA)
    return salida
