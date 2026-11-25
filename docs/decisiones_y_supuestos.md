# Supuestos y decisiones

Registro de todo supuesto/decisión. Esto alimenta las pestañas Description, Data_Quality y la nota
de uso de IA. Formato: fecha · decisión · por qué.

## Generales
- 2026-10-31 · Entrega en un único .xlsx (manda el correo sobre el instructivo).
- 2026-10-31 · Fecha de corte del análisis ("hoy" del dataset): **2026-10-01**. Evidencia:
  el último `Last Seen (SCCM)` es 2026-09-30 22:34 (el export se sacó ~1-oct), y las 150 fórmulas de
  Book Value usan `DATE(2026,10,1)`. `TODAY()` solo aparece en la fila de metadatos (fila 2); usarlo
  haría que los resultados cambien según el día en que se abra el archivo.
- Hallazgos ligados: Book Value solo en 175/1.009 filas (150 fórmula + 25 valor fijo); la fórmula
  usa vida útil de 5 años (1825 d) pero la fila 2 documenta 4 años (1460 d) → inconsistencia.
  13 activos con fecha de compra posterior al corte (hasta 2029-12-01).
- 2026-10-31 · El original queda intacto (pestaña Original_Data); toda corrección se registra en el
  Change_Log.

## Para qué sirve la base (preguntas de negocio, ideas de Fernanda)
- Saber qué dispositivos hay, quién tiene cada uno y si sigue funcionando (last_seen).
- Detectar problemas para reportar y resolver (incidentes abiertos, notas).
- Estadística de qué marcas/modelos dan más problemas y de qué tipo → orientar la próxima compra
  hacia marcas con menos incidentes (útil para el refresh de hardware).

## Secciones 0–4 del script (2026-10-31)
- Grano: 1 fila = 1 activo de TI (usuario final + infraestructura + VMs).
- Filas de documentación se detectan con una regla (≥2 celdas con listas o fórmulas como texto),
  no se asume que sea solo la fila 2 → resultado: solo la fila 2.
- La fila 2 no se usa como fuente de verdad (está incompleta); solo como pista (p. ej. escala de
  business_criticality, catálogo de asset_type).
- Columnas: 34 en la hoja; 3 (AF–AH) sin encabezado ni datos → excluidas. `Status` repetido →
  `asset_status` y `patch_compliance`. `Unnamed: 18` → `review_flag` (marcas de revisión).
  Ninguna columna idéntica a otra. Nombres a snake_case en inglés (entrega en inglés).
- 9 filas completamente vacías excluidas → 1.000 activos. Ninguna fila "casi vacía"
  (mínimo 16 de 31 campos llenos).
- Completitud medida "donde aplica" (campos.APLICA_A), calculada provisionalmente sobre las 864
  filas con asset_type bien escrito; se recalcula tras estandarizar asset_type.
- Supuesto: encrypted vacío = desconocido, no "no".
- Señales para revisar: 62 monitores/otros con last_seen (un monitor no se reporta solo);
  monitores con hostname/MAC/OS/RAM; book_value vacío en 82%. (a confirmar en 01_describir)
- Fila 1 = encabezados inconsistentes (mayúsculas, espacios finales: `Manufacturer `).
- Fila 2 = fila de metadatos/listas válidas, no es un activo → se excluye como dato, se usa como
  documentación de los valores esperados.
- Dos columnas llamadas `Status` (H = estado del activo; AA = cumplimiento de parches).
- Columna S sin encabezado con notas sueltas (`dup?`, `check w/ Mike`, `verify`).
- Columnas AF–AH vacías.
- Asset Tag con formatos mixtos (`AST-004003` vs `AST004007`) y algunos vacíos.
- Fechas mezcladas: datetime y números de serie de Excel (44497).
- Texto con mojibake (`Ciudad de MÃ©xico`, `MontrÃ©al`), espacios y `\xa0`.
- Catálogos sin estandarizar: país (`USA`/`United States`), fabricante (`MSFT`/`Microsoft`),
  proveedor (`cdw`, `Dell.com`/`Dell Direct`), moneda en minúsculas, RAM `32gb` vs 32,
  criticidad numérica vs texto (`Low`).
- Contradicciones: garantía anterior a la compra; laptop con Ubuntu en Surface; VM con nota de
  batería hinchada; activo "In Stock" con usuario asignado; email sin `@`.
- Moneda local → para comparar valor hay que convertir a USD (tipo de cambio fijo y documentado).

## Uso de IA
- Claude (Claude Code) como asistente de programación: armar la estructura del proyecto, escribir y
  correr el código paso a paso, proponer chequeos y opciones en cada decisión, redactar textos en inglés.
  Las reglas y decisiones las tomó Fernanda (ver cada sección). Texto final en la portada del Excel.

## Homogeneización (secciones 5–6 del script)
- Criterio general: no se pasa todo a minúsculas ni se quitan acentos en los datos. Se limpian
  espacios (incl. \xa0) y mojibake en todo el texto (216 celdas; location 71 = MontrÃ©al,
  MÃ¼nchen, Ciudad de MÃ©xico). Para comparar variantes se usa una clave auxiliar (minúsculas sin
  acentos ni signos) y el valor final sale del catálogo.
- asset_type: 55 variantes → 10 tipos del catálogo. Ambiguos (Network / Network Device, 3 filas)
  resueltos con el código de tipo del hostname (SW/FW). Laptop/Tablet → Laptop (hostname -LT-).
- manufacturer: 53 variantes → 18 marcas. HP genérico (HP, Hewlett-Packard, H.P.) se resuelve
  por tipo: Server → HPE (Hewlett Packard Enterprise, ProLiant), resto → HP (HP Inc.); son
  empresas distintas desde 2015. Microsoft (Surface/Hyper-V) ≠ Microsoft Azure (nube).
  Lenova → Lenovo. Pendiente cruce marca vs modelo (HP/Dell con "MacBook Pro").
- Mojibake: también había codificación Mac ("S√£o Paulo", 21 celdas) → se agregó al arreglo
  general (location: 92 celdas corregidas en total).
- asset_tag: 9 formatos distintos (número, minúsculas, espacio, sin guion, sin ceros) → todos a
  AST-###### reconstruyendo el número. Supuesto: números cortos = mismo número sin ceros a la
  izquierda. 12 vacíos NO se imputan (hostname coincide con el tag en 642/645, no en todos).
  42 filas con tag repetido → revisar en identificadores.
- hostname: 4 convenciones conviven (estándar 652, servidor SITE+ROL+## 152, red SITE-SW|FW-ZONA##
  67, legado CIUDAD-TIPO-### 10). Solo la estándar está documentada → hallazgo de gobierno.
  Limpieza de forma: mayúsculas, sin dominio .corp.contoso-corp.com, '0'/'-'/'none' → vacío.
  -OLD (23) se conserva y se marca (flag_hostname_old). 64 filas con hostname repetido.
  Nuevas columnas: hostname_convention, flag_hostname_old.
- Duplicados (sección 11). asset_tag: 21 tags repetidos (42 filas), todos en pares:
  6 duplicado exacto · 9 mismo equipo (mismo serial) con datos distintos (3 con asset_status en
  conflicto: In Use vs In Stock/Active) · 6 tag repetido en equipos distintos (serial distinto;
  en 3 el número del hostname delata el error: MTL-DT-05003, MNL-LT-04856, CHI-DT-05523).
  hostname (excluyendo los mismos activos): 16 nombres repetidos, todos infra (servidores, VMs,
  red): 5 reutilizados legítimamente (uno vivo + otro de baja) · 11 en conflicto (≥2 vivos).
  Las 4 filas con review_flag "dup?" NO están en ningún grupo de duplicados.
  Tratamiento pendiente de decidir. Nuevas columnas: dup_tag_tipo, dup_hostname_tipo.

## Limpieza superficial completa (secciones 10–21, 2026-11-01)
Orden nuevo: primero se homogeneizan todas las variables y recién después se buscan duplicados
(sección 20), para que 'INUSE' e 'In Use' cuenten como el mismo valor.
Decisiones acordadas con Fernanda:
- Fechas dd/mm vs mm/dd ambiguas (26) → se asume mm/dd (formato EE.UU.) + flag_fecha_ambigua.
- Trimestres 'Q4 2023' (18) → día 15 del mes del medio + flag_fecha_aproximada.
- Garantía en texto ('3 years', '5Y', '1yr ProSupport', 23) → compra + N años + flag_garantia_derivada.
  'Lifetime', 'Expired', 'TBD' → vacío. 1900-01-01 (cero de Excel) → vacío.
- Costos negativos (10) → valor absoluto + flag_costo_negativo. 'free' → 0; 'see PO', '-', 'TBD' → vacío.
- Moneda vacía (28) → la pegada al monto; si no, la local del país + flag_moneda_inferida.
- department: códigos CC-#### (35 filas, 9 códigos) → columna cost_center; dos áreas → 'Shared' (16).
  Service Desk → IT; R&D queda como área propia.
- Fuera de rango → vacío + bandera: criticidad 0/7/9 (23), RAM 1600/3200 (= MHz, 2), incidentes -1 (5).
- asset_status: Active/Assigned/Deployed → In Use; Available/Stock/Spare → In Stock;
  Decommissioned/EOL → Retired (EOL marcado, 4); e-Waste/Recycled → Disposed; Not Found → Missing;
  RMA → In Repair.
- Solo se excluyen duplicados exactos (6, se conserva la primera fila) → 994 activos.
  Los demás duplicados se marcan y se revisan en la etapa de contradicciones.
Decisiones menores tomadas sin consultar (revisar si no convencen):
- os: se conserva el detalle y se agregan os_family / os_version.
- storage → texto '<n> GB <tipo>' + storage_gb + storage_type; 1 TB = 1024 GB (así viene en la
  base); número sin unidad < 20 = TB.
- MAC truncadas (29) e IP con octeto > 255 (16) → vacío + bandera. 'DHCP'/'dynamic'/'0.0.0.0' → vacío.
- PO: se quita prefijo 'PO-' y sufijo '/1' → 10 dígitos.
- Email: se corrigen '@@', ';', espacios, falta de '@', '.con'. Dominios @contoso.com y @gmail.com
  NO se cambian y se marcan (15, flag_email_no_corporativo; gmail = correo personal).
- assigned_to: se quita '(Terminated)' y se marca (15 equipos asignados a ex-empleados);
  'Apellido, Nombre' → 'Nombre Apellido'; usuario/inicial → nombre del email si coincide el apellido (27).
- Ubicación: 'HQ' y 'DC1' resueltos por código de sitio del hostname. País vacío (9) → del sitio.
- Moneda '$' y 'Peso' solo aparecen en filas de EE.UU. y México → USD y MXN.
- Mojibake cp1252 ('â€“' → '–') agregado al arreglo general.
Pendiente para contradicciones: fechas imposibles (compra 2029, last_seen 1998), 1 moneda pegada
al monto distinta de currency, marca vs modelo, last_seen en monitores, book_value (recalcular).

## Contradicciones (codigo/02_limpiar/revisar_contradicciones.py, primera corrida)
33 chequeos (columnas chk_*), 459 de 994 activos con al menos uno. Los más grandes:
In Use sin reportar > 90 días (150), monitores con last_seen (59), email ≠ persona (35),
garantía antes de compra (32), sin uso con persona (29), book_value que no cuadra (26),
nota de servidor en equipo de usuario (25), hostname en conflicto (24), leasing vencido en uso (21),
visto antes de compra (21), VM con costo (20).
- book_value informado (175): 149 cuadran con vida útil de 5 años y solo 57 con 4 años (la
  documentada en la fila 2) → la documentación no coincide con la fórmula.
- 37 compras en USD fuera de EE.UU.: no es error, pero hay que convertir para comparar sitios.
  (Actualizado 2026-11-01: en 11 de ellas la currency está mal anotada —monto local con USD—, ver
  'Costos: moneda en el monto y montos atípicos'.)
- review_flag: de las 12 filas marcadas, 6 no tienen ninguna contradicción detectada.
Tratamiento: pendiente, se decide chequeo por chequeo.

## storage y eventos de notes (2026-11-01)
- storage → storage_gb (número en GB) + storage_type (SSD/HDD; vacío en 295 sin tipo en el dato).
- review_flag: 12 valores (dup? 4, check w/ Mike 3, verify 3, x 2) = revisión informal sin trazabilidad.
- notes → 9 indicadores evt_* + datos: malware (53: fecha, INC, dueño anterior, traslado), leasing
  (41: lease_end_date), batería (35), préstamo (35), compartido (33), auditoría física (32: mes),
  servicio alojado (30), refresh (26: FY), reemplazo por falla (23: componente). Todas las notas
  calzan con alguna plantilla. notes se conserva. Fechas de notas en mm/dd/aaaa.
- Corrección: NO hay notas de malware copiadas entre activos distintos (las repetidas eran pares
  duplicados del mismo equipo).
- Nuevos chequeos: dueño anterior = persona actual (29 de 53 notas de malware), compartido/préstamo
  con persona (38), auditado antes de compra (8). Total: 495 de 994 activos con alguna contradicción.

## Regla de duplicados: quedarse con la fila más completa (Fernanda, 2026-11-01)
- Si dos filas del mismo equipo (mismo tag y serial) solo difieren en que una tiene más campos
  llenos, y donde ambas tienen dato coinciden → se deja la de más campos llenos y se excluye la otra.
- Resultado: 5 pares (AST-004072, 004366, 004581, 004993, 005725) → 5 filas excluidas más.
  Total excluidas por duplicado: 11 (6 exactos + 5) → 989 activos.
- Quedan 4 pares de "mismo equipo, datos distintos" que chocan en algún valor:
  AST-004142, 004340, 004438 (asset_status In Use vs In Stock) y 005543 (purchase_cost distinto).
- Patrón: en los 9 pares "mismo equipo", una fila trae el tag como número solo ('004072'), sin
  last_seen → parece una segunda fuente (¿inventario/compras?) cargando el mismo equipo sin
  cruzarlo con SCCM. 7 de los 21 tags repetidos recién coinciden tras normalizar el tag.
- 1 de los 16 hostnames repetidos aparece al pasar a mayúsculas (mex-fw-core08 / MEX-FW-CORE08).
- Corrección de limpieza: IP con ceros a la izquierda ('010.033.033.115') → '10.33.33.115'.
- CSV de revisión: elaborado/activos_ti_2026_repetidos_revision.csv (columnas excluida,
  fila_conservada, campos_distintos, valores crudos y repetido_ya_en_crudo).

## Reglas de duplicados con conflicto (Fernanda, 2026-11-01)
- Mismo equipo con datos que chocan (4 pares) → se conserva la fila cuyo tag ya venía bien escrito
  en el crudo (AST-######); si eso no decide, la In Use con last_seen. En los 4 casos coinciden:
  la fila con tag bien escrito es la In Use con last_seen (SCCM).
  AST-004142, 004340, 004438: chocaba solo asset_status (In Stock en la otra) → resuelto.
  AST-005543: chocaba purchase_cost (2419.59 conservado vs 2613.16 descartado, mismo PO
  4500473718) → flag_revisar_duplicado. Referencia: Surface Laptop 5 en EUR, n=5, mediana 2281.60,
  rango 1738.79–2649.60 → ambos plausibles; confirmar con la factura del PO.
  Valores descartados guardados en dup_valores_descartados. Total excluidas: 15 → 985 activos.
- Tag vs. hostname: fuera de los duplicados, el número del hostname estándar coincide con el tag en
  612 de 612 (100%). En los 6 tags repetidos en equipos distintos se marca flag_tag_sospechoso en la
  fila que no coincide o no se puede comparar (8 filas; en AST-004809 y AST-005252 no se puede
  saber cuál está bien y se marcan las dos).
- Abierto: hay 7 activos con hostname estándar y sin tag; dado el 100% de coincidencia, el tag se
  podría deducir del hostname (decidir).
- Decidido: los 7 tags vacíos con hostname estándar se deducen del hostname (AST-0 + #####) y se
  marcan (flag_tag_desde_hostname). Ninguno choca con un tag existente. Quedan 5 tags vacíos
  (sin hostname estándar: no se pueden deducir).

## Columnas para filtrar lo que hay que corroborar (2026-11-01)
- notas_limpieza (texto) + revisar_limpieza (bool): decisiones basadas en supuestos o deducciones
  (19 banderas; se agregaron flag_location_desde_hostname (5), flag_country_desde_location (9) y
  flag_hostname_conflicto (23)). 251 de 985 filas.
- contradicciones (texto) + n_contradicciones: chequeos que fallan (de 02_limpiar).
- revisar = revisar_limpieza o alguna contradicción: 592 de 985 filas.
- Salida: elaborado/activos_ti_2026_revision.pkl/.csv y activos_ti_2026_chequeos.csv.

## Hostname repetido: tratamiento (2026-11-01)
- Reutilización (5 nombres): no se borra el equipo viejo. Es historia del activo (auditoría,
  depreciación, certificado de baja); un On Order ni siquiera es viejo. Se separa con un filtro.
- Conflicto (11 nombres): se marcan solo los equipos vivos del grupo (flag_hostname_conflicto, 23 filas:
  21 In Use + 2 In Stock; el Retired del grupo no) con nota "revisar cuál conserva el nombre".
- en_inventario_activo: todo salvo Retired, Disposed, Lost/Stolen y On Order. Missing sí cuenta
  (es propio, falta ubicarlo). 826 de 985. Los conteos y costos del análisis usan este filtro.

## Confiabilidad por tema (2026-11-01)
- Para estadísticas no se saca todo 'revisar' (dejaría 40% y sesgado: se van casi todos los
  servidores/red). Tres columnas en activos_ti_2026_revision: confiable_inventario (tag sospechoso,
  hostname en conflicto, dado de baja pero activo, compra posterior al corte, duplicado a confirmar),
  confiable_costos (costo negativo/0, VM con costo, moneda inferida o de otro país, book value que no
  cuadra, duplicado a confirmar), confiable_fechas (fecha ambigua/aproximada/reconstruida, garantía
  antes de compra, visto/auditado antes de compra, last_seen inválido). Hallazgos (ej. sin reportar
  90 días) no restan confiabilidad.

## Contradicciones de fechas: decisiones (2026-11-01)
- Garantía = compra + N años con el mismo día y mes en 716 de 843 (sobre las 1.000 filas, antes de excluir duplicados); N por tipo: 1 teléfono/tablet,
  3 desktop/laptop/monitor/impresora, 5 servidor/firewall/switch (moda en los datos).
- Año de compra imposible (2027-2029: 13; 1999: 7) conserva día y mes de la garantía -> el error es
  el año. purchase_date = garantía - N del tipo (15). Sin garantía propia (3 VM, 1 laptop, 1 monitor
  con garantía derivada de esa compra) -> vacío. flag_compra_reconstruida (20).
- Garantía antes de la compra sin el patrón (25): no se sabe cuál fecha está mal -> se deja, marcado.
- last_seen 09:00:00 exacto y anterior a la compra (8): valor por defecto -> vacío, flag_last_seen_invalido.
- Resultado: compra_posterior_al_corte, compra_muy_antigua y visto_antes_de_compra quedan en 0.

## Contradicciones estado vs. persona: decisiones (2026-11-01)
- Nombre 'Apellido Nombre' sin coma (17): se invierte si así coincide exacto con el email
  (flag_nombre_invertido). Fallo de la limpieza superficial, que solo cubría 'Apellido, Nombre'.
- In Use sin persona con email: nombre desde el email (flag_nombre_desde_email). Afecta 19 filas
  (14 equipos de usuario final + compartidos/préstamo); 2 compartidos ganan persona.
- Chequeo email vs. persona: ignora puntuación (O'Brien = obrien). Quedan 16 con email de otra persona
  (incl. 'lcruz' y 'swhite'): no se sabe cuál es correcto -> se dejan y quedan en revisar. Hallazgo:
  equipos reasignados sin actualizar el registro.
- Ex-empleado con equipo In Use (14) y dado de baja que sigue reportando (12; 4 laptops Lost/Stolen):
  hallazgos de seguridad/recupero, sin cambio en datos.
- sin_uso_con_persona (26) y compartido_o_prestamo_con_persona (39): la usuaria no eligió achicarlos;
  se dejan como están (no afectan confiable_*, solo 'revisar').

## Hostname vs. location / asset_type (2026-11-01)
- Location: comparables los 866 con hostname (estándar/servidor/red por código de sitio; 10 legados
  por la ciudad escrita entera): coinciden todos. 5 coinciden por construcción (venían HQ/DC1 y se
  dedujeron del hostname); independientes: 861/861. HQ = Headquarters (sede) -> Blue Bell (1 caso,
  supuesto débil); DC1 = Data Center 1 -> Eagan (4 casos de infraestructura EAG, único data center).
- asset_type: 706 con código de tipo (LT/DT/MB/TB/PR/SW/FW) coinciden todos (2 por construcción:
  'Network'/'Network Device' resueltos con el código); 10 legados coinciden. Los 150 de servidor
  traen función (APP/DC/ESX/FS/SQL/WEB), no tipo: Server o VM indistinto.
- Hallazgo: 15 VMs con nombre ESX (hipervisor = servidor físico), varias en AWS/Azure.
  Chequeo vm_con_nombre_de_hipervisor (solo revisar). Chequeo hostname_legado_no_coincide: 0.

## Completitud real: vacíos donde el dato corresponde (2026-11-01)
- Regla = tipo (campos.APLICA_A) + estado (campos.ESTADOS_DONDE_APLICA) - no exigidos (NO_EXIGIDO_EN).
  Estados: persona/email solo In Use/In Repair; patch_compliance e IP solo In Use (SCCM solo los
  informa para activos; en bodega/baja 100% vacíos); last_seen en lo que llegó a usarse (no In Stock
  ni On Order); hostname/MAC/encrypted en todo salvo On Order.
- Decidido: last_seen no se exige a impresoras (sin agente SCCM; 21/34 vacías) pero se permite.
  purchase_cost solo físicos (VMs por suscripción; las VMs con costo siguen en vm_con_costo).
  book_value no se exige (es fórmula): book_value_calc = costo × (1 − días/1825), mínimo 0, al corte,
  para todos los físicos (837 vs 174 que venían). 5 años porque 149/175 fórmulas del crudo usan 5
  (la fila 2 documentaba 4).
- Resultado: el faltante real más grande es encrypted (29% de los equipos de usuario sin dato =
  estado de cifrado desconocido). Luego business_criticality y department (5.6%), PO (5.5%),
  garantía (4.7%), IP (4.7%). assigned_to/patch_compliance: 0% donde corresponden.
- Salidas: elaborado/activos_ti_2026_completitud.csv y _completitud_por_tipo.csv.

## "Not applicable" en el Excel (2026-11-01)
- Decidido: una sola etiqueta "Not applicable" (inglés, sin abreviar para no confundir con el 'N/A'
  del crudo, que significaba faltante) en las celdas vacías donde el campo no corresponde por tipo o
  por estado (mismas reglas que la completitud). Vacío = faltante real.
- Solo en la pestaña de datos del Excel (campos.marcar_no_aplica); las bases .pkl conservan el vacío
  para no romper columnas numéricas/fecha. Las derivadas siguen a su campo (os_family -> os,
  storage_type -> storage_gb, book_value_calc -> purchase_cost). book_value del crudo no se marca:
  su vacío es "no informado".

## Excel de entrega (2026-11-01)
- codigo/03_entregar/crear_excel_entrega.py + textos_en.py -> elaborado/entrega/Cortes_Fernanda_DataChallenge.xlsx.
- Todo en inglés (decidido): columnas derivadas, notas por fila, chequeos y registro traducidos al
  exportar (textos_en.py); código y notas internas siguen en español.
- Pestañas: Cover, Executive_Summary, Dashboard, Analysis, Data_Description, Data_Quality,
  Validation_Rules, Assumptions, Clean_Data, Change_Log, Original_Data.
- Analysis y Validation_Rules son fórmulas sobre Clean_Data y Assumptions; Clean_Data tiene 7
  columnas con fórmula (USD, edad, días de garantía, refresh, SO sin soporte, días sin reportar).
- Supuestos nuevos (editables en Assumptions): tipos de cambio a USD aproximados (promedio 2025;
  reemplazar por el oficial de fin de sep-2026); edad de refresh por tipo (laptop 4, desktop 5,
  móvil/tablet 3, monitor 6, impresora 5, servidor 5, switch 6, firewall 5); SO sin soporte al corte
  (Windows 7/8.1/10, Server 2008 R2/2012 R2, CentOS 7, ESXi 7); umbral sin reportar 90 días;
  ventana de garantía 180 días. macOS/iOS/Android no se evalúan (sin fechas oficiales).
- Textos del resumen calculados en pandas con la misma lógica que las fórmulas.
- Uso de IA (portada): Claude (Claude Code) como asistente de programación; decisiones de Fernanda.

## Segunda pasada de identificadores (Fernanda, 2026-11-01)
- Tags repetidos en equipos distintos (6 tags / 12 filas) quedaban solo marcados -> resueltos:
  3 toman el número de su hostname estándar (libre): MTL-DT-05003 -> AST-005003, MNL-LT-04856 ->
  AST-004856, CHI-DT-05523 -> AST-005523 (flag_tag_reasignado). 4 vaciados por no poder deducirse
  (laptop NEWYORK-LAPTOP-428, VM EAGAPP30, 2 monitores sin hostname): flag_tag_vaciado = poner
  etiqueta nueva; el tag original queda en el registro. flag_tag_sospechoso queda en 0.
  Resultado: 0 tags repetidos; 9 vacíos (5 sin tag desde el crudo + 4 vaciados).
- Seriales: 'SEE STICKER' / 'TO BE FILLED' (7) se suman a '0' / 'UNKNOWN' como relleno -> vacío +
  flag_serial_placeholder. Resultado: 0 seriales repetidos (regla de validación nueva).
- Marca vs modelo (7): se corrige con la marca más frecuente del modelo (>=80%, n>=3); el SO lo
  confirma (macOS en MacBook). flag_manufacturer_from_model. Tipo vs marca ya era coherente.
- department con código CC-#### (34): department queda VACÍO a propósito para completar;
  el código en cost_center; flag_department_cc entra en notas/revisar. (Había evidencia de
  código -> área en 16 filas por la misma persona, pero se decidió no completar.)
- Hostname en conflicto (23 vivos): cada par tiene serial, modelo, MAC e IP distintos y ambos
  reportan -> son equipos físicos distintos; el problema es el nombre / registro viejo sin baja.
  Siguen fuera de reliable_for_inventory. Indicios de "equipo viejo sin dar de baja": MUCAPP36
  (fila 99 reemplazo por falla; 828 de 2019 sin garantía ni reporte > 90 d), EAGAPP38 (708 de 2020,
  Server 2012 R2, costo 0), EAGDC11 / MEX-FW-CORE08 (uno In Stock pero reporta), EAG-FW-CORE01
  (Cisco 2024 vs Fortinet 2026), EAGFS20 / EAGSQL40 (VM Hyper-V vs VMware: migración).
- Decidido: marcar en cada conflicto quién conserva el nombre y quién parece el registro viejo.
  hostname_conflict_role: keeps the name (In Use más reciente; empate: último en reportar) /
  probably replaced (10: indicio además de ser más viejo: In Stock con el nombre en uso, garantía
  vencida, sin reportar > 90 d, nota de reemplazo, migración de VM) / undetermined (2: solo más viejo,
  EAG-FW-EDGE04 AST-004708 y EAG-SW-CORE06 AST-004719). Evidencia en hostname_conflict_evidence;
  flag_hostname_probable_reemplazado entra en notas de limpieza. No se borra ni se cambia estado.

## Datos donde no corresponden por tipo (Fernanda, 2026-11-01)
- os (7: 6 monitores + 1 impresora, todos 'Windows 10') y ram_gb (6 monitores): imposibles para el
  tipo, parecen datos de la PC conectada -> vacío + flag_dato_tipo_no_aplica (sección 21d).
- encrypted: la regla se amplía a Server y VM (discos de servidor se cifran): se permite pero no se
  exige (NO_EXIGIDO_EN), así el 29% de cifrado desconocido sigue midiéndose solo en equipos de
  usuario. Los 8 en monitor / impresora / firewall -> vacío + misma bandera.
- last_seen en 59 monitores y MAC en 14 monitores + 6 VMs: se conservan y quedan como contradicción
  (posible inventario vía la PC, dock USB-C, MAC virtual). book_value en 3 VMs: se conserva el dato
  del crudo, marcado; book_value_calc no aplica a VMs.
- Nota: la tabla de completitud de la sección 4 de 01_describir es la foto inicial (antes de limpiar
  y solo por tipo); la final (tipo + estado) es la sección 11b de 02_limpiar.
- Decidido (conservador): los 14 equipos con tag deducido (7), reasignado (3) o vaciado (4) salen de
  confiable_inventario además de quedar en revisar_limpieza: el tag sale de un supuesto hasta
  confirmar el sticker.

## Costos: moneda en el monto y montos atípicos (2026-11-01)
- Fila 615 (Sydney): monto '3785 AUD' con currency USD. Decidido: currency queda como viene,
  flag_moneda_monto_distinta y fuera de confiable_costos (no se sabe cuál es la correcta).
- Estadística por asset_type en USD (sección 16 de 01_describir). Montos chicos: fuera de los 0
  (ya en fisico_con_costo_cero) no hay raros; el menor > 0 es ~0.2–0.4 veces la mediana del tipo.
- Montos grandes: 19 equipos > 3 veces la mediana del tipo (laptop US$442,796, firewall US$2.27M).
  Decidido (conservador): monto sin cambiar, flag_costo_atipico, fuera de confiable_costos.
  causa_costo_atipico (supuesto, a confirmar): 11 'moneda local cargada como USD' (USD fuera de
  EE.UU. y en moneda local el monto es normal), 8 'decimal perdido (x100)'. VM no entran (mediana 0).
- FX pasó a config.py (lo usan 01_describir y el Excel).
- Revisión de comas/puntos en purchase_cost: los 89 textos se leen bien; los montos enormes vienen
  como número en el crudo (no es un error de separadores).
- Costo 0 en físicos (15, todos con PO de revendedor, ninguno leasing) → vacío + flag_costo_cero, a
  completar. VM con 0 se deja (no hay compra).
- purchase_cost_usd queda en la base limpia (costo × config.FX); en el Excel es fórmula.
- Decidido (Fernanda): los 19 costos atípicos se QUITAN de purchase_cost (vacío) y de
  purchase_cost_usd, con flag_costo_atipico; no se usan para estadísticas. El monto original queda en
  purchase_cost_descartado para revisarlo; causa probable en causa_costo_atipico.
- Observación (Fernanda): en los 11 'moneda local cargada como USD' la currency está mal anotada
  (dice USD, el monto solo tiene sentido en moneda local). Chequeado: no es práctica del sitio (Brasil
  40 BRL / 5 USD). De 29 compras en USD fuera de EE.UU. (sin VM): 11 este error (Brasil 5/5, México
  2/2, Hungría 1/1, India 3/4), 18 plausibles en USD (AU, CA, DE, UK). Con AUD/CAD/EUR/GBP el error no
  se detecta por el monto: podría haber más. Los 8 de
  'decimal perdido' tienen la moneda local correcta y son otro patrón (monto x100).

## Garantía vacía (2026-11-01)
- Decidido (Fernanda): equipos físicos con compra y sin fin de garantía (vacío / TBD / Expired) →
  warranty_end_date = compra + años estándar del tipo + flag_garantia_estimada, fuera de
  confiable_fechas. 'Lifetime' queda vacía. VM sin garantía.
- 2 'Expired' con compra 2024 y 2026 (filas 383, 784): la estimación daría vigente → contradicción,
  quedan vacías + flag_garantia_contradictoria (fuera de confiable_fechas). La 3ª (661, compra 2020) sí
  cuadra y queda estimada. Total estimadas: 36.

## Revisión final del Excel (2026-11-01)
- Regla A22 (PO de 10 dígitos) contaba los 49 PO vacíos como fallo: error de fórmula, ahora excluye vacíos → PASS.
- El archivo abre en Cover. Traducidos 'todos' / '(no se exige)' (completitud) y 'se conserva la fila N' (Change_Log).
- B09 hostname_conflict cuenta 24 (todo el grupo, incl. 1 Retired) vs 23 vivos marcados: aclarado en la regla.
