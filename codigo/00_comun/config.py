# -*- coding: utf-8 -*-
"""Rutas y constantes compartidas por todas las etapas."""
from datetime import date
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
CRUDO = RAIZ / "crudo"
ELABORADO = RAIZ / "elaborado"
ENTREGA = ELABORADO / "entrega"
LOGS = RAIZ / "logs"

ARCHIVO_CRUDO = CRUDO / "DataChallenge.xlsx"

# Fecha de corte del análisis (ver docs/notas_supuestos.md)
FECHA_CORTE = date(2026, 10, 1)

# Tipos de cambio a USD (1 unidad = x USD), aproximados. Se usan para comparar montos entre países
# (01_describir) y como insumo editable en la hoja Assumptions del Excel (03_entregar).
FX = {"USD": 1.0, "EUR": 1.12, "GBP": 1.30, "CAD": 0.72, "AUD": 0.65, "MXN": 0.053, "BRL": 0.18,
      "INR": 0.0117, "PHP": 0.0175, "HUF": 0.0028}

# Datos que van en la portada del Excel y el nombre del archivo de entrega.
# (En este repo público la ciudad va vacía: se completa al generar la entrega real.)
NOMBRE_COMPLETO = "Fernanda Cortés"
APELLIDO = "Cortes"
NOMBRE = "Fernanda"
CIUDAD = ""  # completar al generar la entrega
ARCHIVO_ENTREGA = ENTREGA / f"{APELLIDO}_{NOMBRE}_DataChallenge.xlsx"

# Bases intermedias (elaborado/). .pkl conserva los tipos (fechas, enteros); .csv es para mirarla.
BASE_LIMPIA = ELABORADO / "activos_ti_2026_limpio"           # limpieza superficial (01_describir)
REGISTRO_CAMBIOS = ELABORADO / "activos_ti_2026_registro_cambios.csv"
