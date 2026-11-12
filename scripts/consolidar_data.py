# -*- coding: utf-8 -*-
"""
==============================================================
  RESEGUIMIENTOS AMGEN — Script de Consolidación Principal
==============================================================
Qué hace este script:
  1. Lee "Bd pacientes.xlsx" y extrae las columnas clave de cada caso
  2. Lee "Interacciones.xlsx" y se queda solo con la interacción más
     reciente por caso (basado en la columna "Inicio" datetime)
  3. Expande la columna "Recomendaciones" (key:valor por salto de línea)
     en columnas individuales normalizando claves con typos/variantes
  4. Hace el cruce (LEFT JOIN) por "Número del caso" con Interacciones
  5. Lee "Aplicaciones.xlsx" y se queda con la aplicación más reciente
     por caso; calcula "Próxima Aplicación" según la columna Frecuencia
  6. Hace el cruce (LEFT JOIN) por "Número del caso" con Aplicaciones
  7. Valida la calidad de los seguimientos (ver validaciones.py)
  8. Exporta → Reseguimientos.xlsx (Resumen, Reseguimientos, Hallazgos)

Se usa desde la línea de comandos (ejecutar.bat, archivos en CARPETA) o
desde la página web (app.py), que llama a generar_reporte() con los
archivos que sube el usuario.

Ruta fuente  : C:\\Users\\julian.cobos\\Downloads\\reseguimientos\\
Ruta destino : C:\\Users\\julian.cobos\\Downloads\\reseguimientos\\Reseguimientos.xlsx
"""

import io
import re
import difflib
import unicodedata
import warnings
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import date, datetime
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import RowDimension
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo

import validaciones as val

# ─────────────────────────────────────────────────────────────────────────────
#  CONFIG
# ─────────────────────────────────────────────────────────────────────────────
CARPETA = Path(r"C:\Users\julian.cobos\Downloads\reseguimientos")

ARCHIVO_PACIENTES     = CARPETA / "Bd pacientes.xlsx"
ARCHIVO_INTERACCIONES = CARPETA / "Interacciones.xlsx"
ARCHIVO_APLICACIONES  = CARPETA / "Aplicaciones.xlsx"
ARCHIVO_SALIDA        = CARPETA / "Reseguimientos.xlsx"

# Columnas que traemos de BD Pacientes (nombre exacto en el archivo)
COLS_PACIENTES = [
    "Número del caso",
    "Zona",
    "Fecha de inscripción",
    "Regimen",
    "EPS",
    "Logistic Operator",
    "Institución",
    "Estado",
    "Razón del estado",
    "Enfermero(a) de campo",
    "Programa",
]

# Columna de cruce
COL_CASO = "Número del caso"

# Columna datetime para ordenar interacciones
COL_INICIO = "Inicio"

# Columna a expandir
COL_RECOMENDACIONES = "Recomendaciones"

# Columnas para el cruce con Aplicaciones
COL_FECHA_APLICACION   = "Fecha de Aplicación"
COL_FRECUENCIA         = "Frecuencia"
COL_PROXIMA_APLICACION = "Próxima Aplicación"
COL_ESTADO_APLICACION  = "Estado de Aplicación"
COL_APLICADO           = "¿Medicamento aplicado?"

# Columnas de Interacciones que se conservan para las validaciones
COL_ID_INTERACCION = "Interacción: Interacciones ID"
COL_CREADO_POR     = "Interacción: Creado por"
COL_EXITOSA        = "Interacción existosa?"

# ─────────────────────────────────────────────────────────────────────────────
#  MAPA DE FRECUENCIAS → OFFSET DE DÍAS / MESES
#  Claves normalizadas (sin acento, minúsculas) → pd.DateOffset
# ─────────────────────────────────────────────────────────────────────────────
FRECUENCIA_OFFSET = {
    "cada 20 dias" : pd.DateOffset(days=20),
    "cada 21 dias" : pd.DateOffset(days=21),
    "cada 3 semanas": pd.DateOffset(days=21),
    "cada 3 meses" : pd.DateOffset(months=3),
    "diaria"       : pd.DateOffset(days=1),
    "mensual"      : pd.DateOffset(months=1),
    "quincenal"    : pd.DateOffset(days=15),
    "semanal"      : pd.DateOffset(days=7),
    "semestral"    : pd.DateOffset(months=6),
    # "otro" y cualquier valor no listado → Frecuencia Errónea
}

# ─────────────────────────────────────────────────────────────────────────────
#  NORMALIZACIÓN DE CLAVES EN "RECOMENDACIONES"
#  Mapeo: texto_normalizado → nombre_columna_final
#  Cubre typos, variantes con/sin acento, abreviaciones, etc.
# ─────────────────────────────────────────────────────────────────────────────
CLAVES_MAP = {
    # CANAL DE CONTACTO
    "canal de contacto"           : "CANAL DE CONTACTO",
    "canal de contaco"            : "CANAL DE CONTACTO",
    "canal de contancto"          : "CANAL DE CONTACTO",
    "canal contacto"              : "CANAL DE CONTACTO",
    "canal"                       : "CANAL DE CONTACTO",
    "vcanal de contacto"          : "CANAL DE CONTACTO",
    "canal de contacto:"          : "CANAL DE CONTACTO",
    "anal de contacto"            : "CANAL DE CONTACTO",
    "anal de contaco"             : "CANAL DE CONTACTO",
    "nal de contacto"             : "CANAL DE CONTACTO",
    "canacanal de contacto"       : "CANAL DE CONTACTO",

    # PARTICIPANTE
    "participante"                : "PARTICIPANTE",
    "participantes"               : "PARTICIPANTE",
    "paticipante"                 : "PARTICIPANTE",
    "rticipante"                  : "PARTICIPANTE",
    "articipante"                 : "PARTICIPANTE",

    # TIPO DE GESTIÓN
    "tipo de gestion"             : "TIPO DE GESTIÓN",
    "tipo de gestin"              : "TIPO DE GESTIÓN",
    "tipo de gestion"             : "TIPO DE GESTIÓN",
    "tipo de gestio"              : "TIPO DE GESTIÓN",
    "tipo de gestion"             : "TIPO DE GESTIÓN",
    "tipo de geston"              : "TIPO DE GESTIÓN",
    "tipo de gesti0n"             : "TIPO DE GESTIÓN",
    "tipode gestion"              : "TIPO DE GESTIÓN",
    "tipo degestion"              : "TIPO DE GESTIÓN",
    "tipo de llamada"             : "TIPO DE GESTIÓN",
    "tipo de lamada"              : "TIPO DE GESTIÓN",

    # TIPO DE CONTACTO (variantes que merecen su propia columna)
    "tipo de contacto"            : "TIPO DE CONTACTO",
    "tipo de contato"             : "TIPO DE CONTACTO",
    "tipo de contact o"           : "TIPO DE CONTACTO",
    "tipo de co ntacto"           : "TIPO DE CONTACTO",
    "tipo de contactotipo de contacto": "TIPO DE CONTACTO",
    "tipo de  contacto"           : "TIPO DE CONTACTO",
    "tipo de cana l call/tipo de contacto": "TIPO DE CONTACTO",
    "tipo de canal call/ tipo de contacto": "TIPO DE CONTACTO",
    "tipo de canal call/tipo de contacto": "TIPO DE CONTACTO",

    # ESTADO DESPUÉS DE CONTACTO
    "estado despues de contacto"  : "ESTADO DESPUÉS DE CONTACTO",
    "estado despues contacto"     : "ESTADO DESPUÉS DE CONTACTO",
    "stado despues de contacto"   : "ESTADO DESPUÉS DE CONTACTO",
    "estado de contacto"          : "ESTADO DESPUÉS DE CONTACTO",
    "estado despues del contacto" : "ESTADO DESPUÉS DE CONTACTO",
    "despues de contacto"         : "ESTADO DESPUÉS DE CONTACTO",
    "despues del contacto"        : "ESTADO DESPUÉS DE CONTACTO",

    # FECHA DE ÚLTIMA APLICACIÓN
    "fecha de ultima aplicacion"  : "FECHA DE ÚLTIMA APLICACIÓN",
    "fecha ultima aplicacion"     : "FECHA DE ÚLTIMA APLICACIÓN",
    "ultima aplicacion"           : "FECHA DE ÚLTIMA APLICACIÓN",
    "fecha de ultima aplicacion ": "FECHA DE ÚLTIMA APLICACIÓN",
    "echa de ultima aplicacion"   : "FECHA DE ÚLTIMA APLICACIÓN",
    "fecha aplicacion"            : "FECHA DE ÚLTIMA APLICACIÓN",

    # FECHA DE CONTROL CON ESPECIALISTA
    "fecha de control con especialista": "FECHA DE CONTROL CON ESPECIALISTA",
    "fecha control especialista"       : "FECHA DE CONTROL CON ESPECIALISTA",
    "fecha de control"                 : "FECHA DE CONTROL CON ESPECIALISTA",

    # FECHA DE PRÓXIMO CONTACTO
    "fecha de proximo contacto"   : "FECHA DE PRÓXIMO CONTACTO",
    "fecha proximo contacto"      : "FECHA DE PRÓXIMO CONTACTO",
    "fecha de proximo contact"    : "FECHA DE PRÓXIMO CONTACTO",
    "proximo contacto"            : "FECHA DE PRÓXIMO CONTACTO",

    # REPORTE FV
    "reporte fv"                  : "REPORTE FV",
    "reportefv"                   : "REPORTE FV",
    "repore fv"                   : "REPORTE FV",
    "reporta fv"                  : "REPORTE FV",
    "reportafv"                   : "REPORTE FV",
    "reprote fv"                  : "REPORTE FV",
    "repporte fv"                 : "REPORTE FV",
    "reporte  fv"                 : "REPORTE FV",
    "se reporta fv"               : "REPORTE FV",
    "se reporte fv"               : "REPORTE FV",
    "reporte ea"                  : "REPORTE FV",
    "repore fv"                   : "REPORTE FV",
    "repooffimedicasrte fv"       : "REPORTE FV",

    # DETALLE
    "detalle"                     : "DETALLE",
    "detalles"                    : "DETALLE",
    "etalle"                      : "DETALLE",

    # OP LOGÍSTICO
    "op logistico"                : "OP LOGÍSTICO",
    "op logistico"                : "OP LOGÍSTICO",
    "opl logistico"               : "OP LOGÍSTICO",
    "opl logistico"               : "OP LOGÍSTICO",
    "opl. logistico"              : "OP LOGÍSTICO",
    "opl"                         : "OP LOGÍSTICO",
    "operador logistico"          : "OP LOGÍSTICO",
    "oplogistico o"               : "OP LOGÍSTICO",
    "op. logistico"               : "OP LOGÍSTICO",
    "op.logistico"                : "OP LOGÍSTICO",
    "op/ logistico"               : "OP LOGÍSTICO",
    "po logistico"                : "OP LOGÍSTICO",
    "po.logistico"                : "OP LOGÍSTICO",
    "p logistico"                 : "OP LOGÍSTICO",
    "pl"                          : "OP LOGÍSTICO",
    "opl"                         : "OP LOGÍSTICO",
    "perador logistico"           : "OP LOGÍSTICO",
    "op.l"                        : "OP LOGÍSTICO",
    "op logistigo"                : "OP LOGÍSTICO",
    "op logisitico"               : "OP LOGÍSTICO",
    "op logistica"                : "OP LOGÍSTICO",
    "logistico"                   : "OP LOGÍSTICO",
    "oplogistico"                 : "OP LOGÍSTICO",
}

# Claves con las que se intenta la coincidencia aproximada (typos no listados).
# Se excluyen las abreviaturas cortas para evitar falsos positivos.
_CLAVES_DIFUSAS = [k for k in CLAVES_MAP if len(k) >= 8]

# Orden de columnas en el Excel final (Recomendaciones expandidas)
ORDEN_RECOMENDACIONES = [
    "CANAL DE CONTACTO",
    "PARTICIPANTE",
    "TIPO DE GESTIÓN",
    "TIPO DE CONTACTO",
    "ESTADO DESPUÉS DE CONTACTO",
    "FECHA DE ÚLTIMA APLICACIÓN",
    "FECHA DE CONTROL CON ESPECIALISTA",
    "FECHA DE PRÓXIMO CONTACTO",
    "REPORTE FV",
    "DETALLE",
    "OP LOGÍSTICO",
]


# ─────────────────────────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def quitar_acentos(texto: str) -> str:
    """Elimina acentos y caracteres especiales."""
    return "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )


def normalizar_clave(clave: str) -> str:
    """Normaliza clave para comparación: sin acento, lowercase, sin espacios dobles."""
    clave = str(clave).strip()
    clave = quitar_acentos(clave)
    clave = clave.lower()
    clave = re.sub(r"[\xa0\u200b\u00a0]+", " ", clave)  # espacios raros → espacio
    clave = re.sub(r"\s+", " ", clave).strip()
    clave = clave.lstrip("|}./\"' ")                      # limpia caracteres basura al inicio
    return clave


_cache_claves: dict[str, str | None] = {}


def mapear_clave(clave_raw: str) -> str | None:
    """Retorna el nombre estandarizado de la clave, o None si no se reconoce."""
    norm = normalizar_clave(clave_raw)
    if norm in CLAVES_MAP:
        return CLAVES_MAP[norm]
    if norm in _cache_claves:
        return _cache_claves[norm]
    if norm.startswith("no "):
        return None  # p.ej. "NO SE REPORTA FV: ..." es texto libre, no la clave REPORTE FV
    # Intento difuso: typos de una o dos letras sobre claves largas
    destino = None
    if len(norm) >= 8:
        cerca = difflib.get_close_matches(norm, _CLAVES_DIFUSAS, n=1, cutoff=0.88)
        if cerca:
            destino = CLAVES_MAP[cerca[0]]
    _cache_claves[norm] = destino
    return destino


def expandir_recomendaciones(serie: pd.Series) -> pd.DataFrame:
    """
    Convierte la columna Recomendaciones en un DataFrame de columnas expandidas.
    Cada salto de línea es una fila del tipo "CLAVE: VALOR".
    """
    registros = []
    for celda in serie:
        fila = {}
        if pd.isna(celda):
            registros.append(fila)
            continue

        lineas = str(celda).splitlines()
        for linea in lineas:
            linea = linea.strip()
            if not linea:
                continue
            if ":" not in linea:
                continue  # línea sin separador, ignorar
            # Divide solo en el primer ":"
            partes = linea.split(":", 1)
            clave_raw = partes[0].strip()
            valor     = partes[1].strip() if len(partes) > 1 else ""

            col_name = mapear_clave(clave_raw)
            if col_name and col_name not in fila:
                fila[col_name] = valor if valor else np.nan

        registros.append(fila)

    df_expandido = pd.DataFrame(registros)

    # Asegura que existan todas las columnas estándar en orden
    for col in ORDEN_RECOMENDACIONES:
        if col not in df_expandido.columns:
            df_expandido[col] = np.nan

    # Ordena: primero las estándar, luego cualquier extra descubierta
    cols_extra = [c for c in df_expandido.columns if c not in ORDEN_RECOMENDACIONES]
    df_expandido = df_expandido[ORDEN_RECOMENDACIONES + cols_extra]

    return df_expandido


def encontrar_columna(df: pd.DataFrame, nombre_buscado: str) -> str | None:
    """Busca una columna por nombre ignorando acentos y mayúsculas."""
    target = normalizar_clave(nombre_buscado)
    for col in df.columns:
        if normalizar_clave(col) == target:
            return col
    return None


# ─────────────────────────────────────────────────────────────────────────────
#  LECTURA (acepta ruta o archivo subido desde la web)
# ─────────────────────────────────────────────────────────────────────────────
def leer_archivo(fuente, nombre: str = "") -> pd.DataFrame:
    """
    Lee un Excel o CSV desde una ruta o un objeto tipo archivo (BytesIO /
    UploadedFile de Streamlit). Usa el motor 'calamine' si está instalado
    (mucho más rápido para archivos grandes).
    """
    nombre = (nombre or getattr(fuente, "name", "") or str(fuente)).lower()
    if hasattr(fuente, "seek"):
        fuente.seek(0)

    if nombre.endswith(".csv"):
        datos = fuente.read() if hasattr(fuente, "read") else Path(fuente).read_bytes()
        for enc in ("utf-8-sig", "latin-1"):
            try:
                return pd.read_csv(io.BytesIO(datos), encoding=enc, sep=None, engine="python")
            except UnicodeDecodeError:
                continue

    try:
        import python_calamine  # noqa: F401
        return pd.read_excel(fuente, engine="calamine")
    except ImportError:
        return pd.read_excel(fuente)


def _parsear_inicio(serie: pd.Series) -> pd.Series:
    """'6/01/2026, 11:15 a. m.' → datetime (día primero)."""
    if pd.api.types.is_datetime64_any_dtype(serie):
        return serie
    limpio = (
        serie.astype(str)
        .str.replace(r'a\.\s*m\.', 'AM', flags=re.IGNORECASE, regex=True)
        .str.replace(r'p\.\s*m\.', 'PM', flags=re.IGNORECASE, regex=True)
        .str.replace(',', '', regex=False)
        .str.strip()
    )
    resultado = pd.to_datetime(limpio, errors="coerce", format="%d/%m/%Y %I:%M %p")
    faltantes = resultado.isna() & serie.notna()
    if faltantes.any():
        resultado[faltantes] = pd.to_datetime(limpio[faltantes], errors="coerce", dayfirst=True)
    return resultado


def preparar_pacientes(df: pd.DataFrame) -> pd.DataFrame:
    """Retorna solo las columnas requeridas de BD pacientes."""
    print(f"  BD Pacientes: {len(df)} filas leídas")

    # Mapear nombres de columna aunque tengan variaciones de encoding
    cols_encontradas = {}
    for col_buscada in COLS_PACIENTES:
        col_real = encontrar_columna(df, col_buscada)
        if col_real:
            cols_encontradas[col_real] = col_buscada
        else:
            print(f"  [!] Columna no encontrada: '{col_buscada}'")

    if COL_CASO not in cols_encontradas.values():
        raise ValueError("No se encontró 'Número del caso' en BD pacientes")

    df_sel = df[list(cols_encontradas.keys())].rename(columns=cols_encontradas)
    df_sel[COL_CASO] = df_sel[COL_CASO].astype(str).str.strip()
    return df_sel


def preparar_interacciones(df: pd.DataFrame) -> pd.DataFrame:
    """
    Retorna TODAS las interacciones con Inicio como datetime y la columna
    Recomendaciones expandida. Se usan completas para validar (la nota
    anterior trae la fecha programada) y la más reciente para el consolidado.
    """
    print(f"  Interacciones: {len(df)} filas leídas, {df.shape[1]} columnas")

    col_caso_int = encontrar_columna(df, COL_CASO)
    col_inicio   = encontrar_columna(df, COL_INICIO)
    col_rec      = encontrar_columna(df, COL_RECOMENDACIONES)

    if not col_caso_int:
        raise ValueError("No se encontró 'Número del caso' en Interacciones")
    if not col_inicio:
        raise ValueError("No se encontró 'Inicio' en Interacciones")

    renombres = {col_caso_int: COL_CASO, col_inicio: COL_INICIO}
    if col_rec:
        renombres[col_rec] = COL_RECOMENDACIONES
    for original, destino in [(COL_ID_INTERACCION, val.COL_ID_INT),
                              (COL_CREADO_POR, val.COL_CREADO_POR),
                              (COL_EXITOSA, val.COL_EXITOSA)]:
        col = encontrar_columna(df, original)
        if col:
            renombres[col] = destino
    df = df.rename(columns=renombres)

    df[COL_CASO]   = df[COL_CASO].astype(str).str.strip()
    df[COL_INICIO] = _parsear_inicio(df[COL_INICIO])

    cols_keep = [c for c in [COL_CASO, val.COL_ID_INT, COL_INICIO, val.COL_CREADO_POR, val.COL_EXITOSA]
                 if c in df.columns]
    df_out = df[cols_keep].reset_index(drop=True)

    if COL_RECOMENDACIONES in df.columns:
        df_rec = expandir_recomendaciones(df[COL_RECOMENDACIONES])
        df_out = pd.concat([df_out, df_rec.reset_index(drop=True)], axis=1)
        print(f"      Columnas expandidas de Recomendaciones: {list(df_rec.columns)}")
    else:
        print("      [!] No se encontro la columna Recomendaciones en Interacciones")
        for col in ORDEN_RECOMENDACIONES:
            df_out[col] = np.nan

    # Más reciente primero dentro de cada caso
    return df_out.sort_values([COL_CASO, COL_INICIO], ascending=[True, False]).reset_index(drop=True)


def interaccion_mas_reciente(df_int: pd.DataFrame) -> pd.DataFrame:
    """Una fila por caso: la interacción con el Inicio más reciente (sin columnas de auditoría)."""
    df_reciente = df_int.drop_duplicates(subset=[COL_CASO], keep="first")
    print(f"  Interacciones (1 por caso): {len(df_reciente)} registros")
    auditoria = (val.COL_ID_INT, val.COL_CREADO_POR, val.COL_EXITOSA)
    return df_reciente[[c for c in df_int.columns if c not in auditoria]]


def preparar_aplicaciones(df: pd.DataFrame) -> pd.DataFrame:
    """Todas las aplicaciones con fecha datetime y columna Aplicado (SI / NO)."""
    print(f"  Aplicaciones: {len(df)} filas leídas, {df.shape[1]} columnas")

    col_caso_ap  = encontrar_columna(df, COL_CASO)
    col_fecha_ap = encontrar_columna(df, COL_FECHA_APLICACION)
    col_frec     = encontrar_columna(df, COL_FRECUENCIA)
    col_aplicado = encontrar_columna(df, COL_APLICADO) or encontrar_columna(df, "Medicamento aplicado?")

    if not col_caso_ap:
        raise ValueError("No se encontró 'Número del caso' en Aplicaciones")
    if not col_fecha_ap:
        raise ValueError(f"No se encontró columna '{COL_FECHA_APLICACION}' en Aplicaciones")

    out = pd.DataFrame({
        COL_CASO            : df[col_caso_ap].astype(str).str.strip(),
        COL_FECHA_APLICACION: pd.to_datetime(df[col_fecha_ap], errors="coerce", dayfirst=True),
    })
    if col_frec:
        out[COL_FRECUENCIA] = df[col_frec]
    else:
        print(f"  [!] Columna '{COL_FRECUENCIA}' no encontrada en Aplicaciones")
    if col_aplicado:
        out["Aplicado"] = df[col_aplicado].map(val.normalizar_texto)
    else:
        print(f"  [!] Columna '{COL_APLICADO}' no encontrada; se asume que todas fueron aplicadas")
        out["Aplicado"] = "SI"
    return out


def aplicacion_mas_reciente(df_apl: pd.DataFrame) -> pd.DataFrame:
    """
    La aplicación más reciente por caso. Las marcadas como NO aplicadas no
    cuentan para calcular la próxima aplicación.
    """
    df = df_apl[df_apl["Aplicado"] != "NO"]
    df = df.sort_values([COL_CASO, COL_FECHA_APLICACION], ascending=[True, False], na_position="last")
    df_reciente = df.drop_duplicates(subset=[COL_CASO], keep="first")
    print(f"  Aplicaciones (1 por caso): {len(df_reciente)} registros")
    cols_keep = [COL_CASO, COL_FECHA_APLICACION] + ([COL_FRECUENCIA] if COL_FRECUENCIA in df.columns else [])
    return df_reciente[cols_keep]


def calcular_estado_aplicacion(proxima_str, hoy=None) -> str:
    """
    Evalúa el string de Próxima Aplicación contra la fecha de corte (hoy por defecto).
      - Fecha pasada        → 'Aplicación Vencida'
      - Fecha futura/hoy   → 'En espera de aplicación'
      - Vacío / NaN        → 'Sin aplicación'
      - 'Frecuencia Errónea' → 'Revisar'
    """
    if pd.isna(proxima_str) or str(proxima_str).strip() == "":
        return "Sin aplicación"

    proxima_str = str(proxima_str).strip()

    if proxima_str == "Frecuencia Errónea":
        return "Revisar"

    try:
        proxima = datetime.strptime(proxima_str, "%d/%m/%Y").date()
        hoy     = hoy or datetime.today().date()
        return "Aplicación Vencida" if proxima < hoy else "En espera de aplicación"
    except Exception:
        return "Sin aplicación"


def calcular_proxima_aplicacion(fecha, frecuencia):
    """
    Calcula la próxima aplicación sumando el offset de la frecuencia.
    Retorna string DD/MM/YYYY o 'Frecuencia Errónea'.
    """
    if pd.isna(fecha):
        return np.nan

    if pd.isna(frecuencia) or str(frecuencia).strip() == "":
        return "Frecuencia Errónea"

    frec_norm = normalizar_clave(str(frecuencia))
    offset    = FRECUENCIA_OFFSET.get(frec_norm)

    if offset is None:
        return "Frecuencia Errónea"

    try:
        return (fecha + offset).strftime("%d/%m/%Y")
    except Exception:
        return "Frecuencia Errónea"


# ─────────────────────────────────────────────────────────────────────────────
#  PROCESO COMPLETO
# ─────────────────────────────────────────────────────────────────────────────
def consolidar(df_pac: pd.DataFrame, df_int: pd.DataFrame, df_apl: pd.DataFrame | None,
               fecha_corte=None) -> pd.DataFrame:
    """Una fila por paciente con su último seguimiento y su última aplicación."""
    df_final = df_pac.merge(interaccion_mas_reciente(df_int), on=COL_CASO, how="left")
    print(f"      OK Resultado: {len(df_final)} filas | {len(df_final.columns)} columnas")

    if df_apl is not None:
        df_final = df_final.merge(aplicacion_mas_reciente(df_apl), on=COL_CASO, how="left")

        # Calcula Próxima Aplicación mientras la fecha aún es Timestamp
        if COL_FRECUENCIA in df_final.columns:
            df_final[COL_PROXIMA_APLICACION] = [
                calcular_proxima_aplicacion(f, fr)
                for f, fr in zip(df_final[COL_FECHA_APLICACION], df_final[COL_FRECUENCIA])
            ]
        else:
            df_final[COL_PROXIMA_APLICACION] = "Frecuencia Errónea"

        df_final[COL_FECHA_APLICACION] = df_final[COL_FECHA_APLICACION].dt.strftime("%d/%m/%Y")
        hoy = pd.Timestamp(fecha_corte).date() if fecha_corte is not None else None
        df_final[COL_ESTADO_APLICACION] = df_final[COL_PROXIMA_APLICACION].apply(
            calcular_estado_aplicacion, hoy=hoy)

        # Reordena: Próxima Aplicación → Estado de Aplicación, ambas tras Fecha de Aplicación
        cols = [c for c in df_final.columns if c not in (COL_PROXIMA_APLICACION, COL_ESTADO_APLICACION)]
        idx = cols.index(COL_FECHA_APLICACION) + 1
        cols[idx:idx] = [COL_PROXIMA_APLICACION, COL_ESTADO_APLICACION]
        df_final = df_final[cols]

    df_final[COL_INICIO] = pd.to_datetime(df_final[COL_INICIO], errors="coerce").dt.strftime("%d/%m/%Y %H:%M")
    return df_final


def generar_reporte(fuente_pacientes, fuente_interacciones, fuente_aplicaciones=None,
                    params: "val.ParametrosValidacion | None" = None, progreso=None) -> dict:
    """
    Ejecuta todo el proceso. Las fuentes pueden ser rutas o archivos subidos.
    `progreso(fraccion, mensaje)` es opcional (lo usa la página web).

    Retorna dict con: consolidado, hallazgos, resumen, fecha_corte, totales.
    """
    avisar = progreso or (lambda f, m: print(f"\n[{int(f * 100):>3}%] {m}"))
    params = params or val.ParametrosValidacion()

    avisar(0.05, "Leyendo BD Pacientes...")
    df_pac = preparar_pacientes(leer_archivo(fuente_pacientes))

    avisar(0.25, "Leyendo Interacciones...")
    df_int = preparar_interacciones(leer_archivo(fuente_interacciones))

    df_apl = None
    if fuente_aplicaciones is not None:
        avisar(0.45, "Leyendo Aplicaciones...")
        df_apl = preparar_aplicaciones(leer_archivo(fuente_aplicaciones))

    avisar(0.70, "Validando seguimientos...")
    df_h, df_res_caso, fecha_corte = val.validar_seguimientos(df_int, df_pac, df_apl, params)

    avisar(0.85, "Consolidando...")
    df_final = consolidar(df_pac, df_int, df_apl, fecha_corte)
    df_final = df_final.merge(df_res_caso, on=COL_CASO, how="left")
    for col in (val.COL_N_CRITERIOS, val.COL_N_HALLAZGOS):
        df_final[col] = df_final[col].astype("Int64")

    avisar(1.0, "Listo")
    return {
        "consolidado": df_final,
        "hallazgos"  : df_h,
        "resumen"    : tabla_resumen(df_h),
        "fecha_corte": fecha_corte,
        "totales"    : {
            "pacientes"          : len(df_pac),
            "seguimientos"       : int(df_int[COL_INICIO].notna().sum()),
            "casos_validados"    : len(df_res_caso),
            "casos_con_hallazgos": int((df_res_caso[val.COL_N_HALLAZGOS] > 0).sum()),
            "hallazgos"          : len(df_h),
            "sin_interaccion"    : int(df_final[COL_INICIO].isna().sum()),
            "con_aplicaciones"   : df_apl is not None,
        },
    }


def tabla_resumen(df_h: pd.DataFrame) -> pd.DataFrame:
    """Conteo de hallazgos por validación y severidad."""
    filas = []
    for v in val.VALIDACIONES:
        sub = df_h[df_h["Validación"] == v]
        filas.append({
            "Validación"     : v,
            "Casos afectados": sub[COL_CASO].nunique(),
            "Hallazgos"      : len(sub),
            val.ALTA         : int((sub["Severidad"] == val.ALTA).sum()),
            val.MEDIA        : int((sub["Severidad"] == val.MEDIA).sum()),
            val.BAJA         : int((sub["Severidad"] == val.BAJA).sum()),
        })
    return pd.DataFrame(filas)


# ─────────────────────────────────────────────────────────────────────────────
#  FORMATO EXCEL
# ─────────────────────────────────────────────────────────────────────────────
# El libro se escribe en modo "solo escritura" de openpyxl: las filas se van
# guardando a medida que se escriben. Con archivos grandes el modo normal llega
# a más de 3 GB de memoria y la versión web (Python dentro del navegador) se
# queda sin memoria.
_BORDE_FINO = Border(left=Side(style="thin"), right=Side(style="thin"),
                     top=Side(style="thin"), bottom=Side(style="thin"))
_FORMATO_FECHA_HORA = "YYYY-MM-DD HH:MM:SS"
_FORMATO_FECHA = "YYYY-MM-DD"


def hallazgos_por_persona(df_h: pd.DataFrame) -> pd.DataFrame:
    """Matriz persona que registró la nota × validación."""
    if df_h.empty or val.COL_CREADO_POR not in df_h.columns:
        return pd.DataFrame()
    tabla = df_h.pivot_table(index=val.COL_CREADO_POR, columns="Validación",
                             values="Hallazgo", aggfunc="count", fill_value=0)
    tabla["Total"] = tabla.sum(axis=1)
    return tabla.sort_values("Total", ascending=False).reset_index()


def exportar_excel(reporte: dict, destino=None) -> bytes | None:
    """
    Escribe el Excel con hojas Resumen, Reseguimientos y Hallazgos.
    destino=None → retorna los bytes (para descargar desde la web).
    """
    df_h = reporte["hallazgos"].copy()
    if "Fecha llamada" in df_h.columns:
        df_h["Fecha llamada"] = pd.to_datetime(df_h["Fecha llamada"]).dt.strftime("%d/%m/%Y %H:%M")
    por_persona = hallazgos_por_persona(reporte["hallazgos"])
    fila_persona = len(reporte["resumen"]) + 8

    wb = Workbook(write_only=True)
    _hoja_resumen(wb, reporte, por_persona, fila_persona)
    _hoja_tabla(wb, reporte["consolidado"], "Reseguimientos", "1F4E79",
                colores_encabezado={COL_ESTADO_APLICACION: "FFD966",
                                    **{v: "F4B183" for v in val.VALIDACIONES},
                                    val.COL_N_CRITERIOS: "F4B183", val.COL_N_HALLAZGOS: "F4B183"})
    _hoja_tabla(wb, df_h, "Hallazgos", "833C0B",
                colores_celda={"Severidad": {val.ALTA: "F8CBAD", val.MEDIA: "FFE699", val.BAJA: "DDEBF7"}},
                anchos={"Hallazgo": 70, "DETALLE": 70})

    buffer = io.BytesIO()
    wb.save(buffer)
    datos = buffer.getvalue()
    if destino is None:
        return datos
    Path(destino).write_bytes(datos)
    return None


def _celda(ws, valor, fill=None, font=None, alignment=None, border=None, number_format=None):
    c = WriteOnlyCell(ws, value=valor)
    if fill is not None:
        c.fill = fill
    if font is not None:
        c.font = font
    if alignment is not None:
        c.alignment = alignment
    if border is not None:
        c.border = border
    if number_format is not None:
        c.number_format = number_format
    return c


def _valores(df: pd.DataFrame, bloque: int = 5000):
    """Filas del DataFrame como listas de valores de Python (vacíos → None), por bloques."""
    for ini in range(0, len(df), bloque):
        parte = df.iloc[ini:ini + bloque].astype(object)
        parte = parte.where(parte.notna(), None)
        for fila in parte.itertuples(index=False, name=None):
            yield list(fila)


def _formato_fecha(v):
    if isinstance(v, datetime):
        return _FORMATO_FECHA_HORA
    if isinstance(v, date):
        return _FORMATO_FECHA
    return None


def _hoja_resumen(wb, reporte: dict, por_persona: pd.DataFrame, fila_persona: int):
    ws = wb.create_sheet("Resumen")
    ws.column_dimensions["A"].width = 38
    for letra in "BCDEFGH":
        ws.column_dimensions[letra].width = 18

    t = reporte["totales"]
    titulo = dict(fill=PatternFill("solid", fgColor="1F4E79"), font=Font(bold=True, color="FFFFFF"),
                  alignment=Alignment(horizontal="center", vertical="top"), border=_BORDE_FINO)
    ws.append([_celda(ws, "RESEGUIMIENTOS AMGEN — Validación de seguimientos",
                      font=Font(bold=True, size=14, color="1F4E79"))])
    ws.append([f"Fecha de corte: {reporte['fecha_corte']:%d/%m/%Y}   |   "
               f"Casos validados: {t['casos_validados']}   |   "
               f"Casos con hallazgos: {t['casos_con_hallazgos']}   |   Hallazgos: {t['hallazgos']}"])
    if not t["con_aplicaciones"]:
        ws.append([_celda(ws, "No se cargó el archivo de Aplicaciones: la validación 3 no se ejecutó.",
                          font=Font(italic=True, color="C00000"))])
    else:
        ws.append([])
    ws.append([])

    fila = 5
    ws.append([_celda(ws, c, **titulo) for c in reporte["resumen"].columns])
    for valores in _valores(reporte["resumen"]):
        ws.append(valores)
        fila += 1
    if not por_persona.empty:
        while fila < fila_persona - 1:
            ws.append([])
            fila += 1
        ws.append([_celda(ws, "Hallazgos por persona que registró la nota", font=Font(bold=True))])
        ws.append([_celda(ws, c, **titulo) for c in por_persona.columns])
        for valores in _valores(por_persona):
            ws.append(valores)


def _hoja_tabla(wb, df: pd.DataFrame, hoja: str, color_encabezado: str,
                colores_encabezado: dict[str, str] | None = None,
                colores_celda: dict[str, dict[str, str]] | None = None,
                anchos: dict[str, float] | None = None):
    """
    Hoja con los datos como tabla de Excel (filas alternas y filtros),
    encabezado de color, ancho automático (máx 45) y paneles fijos en B2.
    """
    ws = wb.create_sheet(hoja)
    columnas = [str(c) for c in df.columns]
    n = len(columnas)

    # En modo solo escritura, anchos, alto y paneles se definen antes de escribir filas
    muestra = list(_valores(df.head(499)))
    for j, nombre in enumerate(columnas):
        max_len = max([len(nombre)] + [len(str(f[j] or "")) for f in muestra])
        ancho = (anchos or {}).get(nombre, min(max(max_len + 4, 10), 45))
        ws.column_dimensions[get_column_letter(j + 1)].width = ancho
    ws.row_dimensions[1] = RowDimension(ws, index=1, ht=32)
    ws.freeze_panes = "B2"

    alineado = Alignment(horizontal="center", vertical="center", wrap_text=True)
    encabezado = []
    for nombre in columnas:
        color = (colores_encabezado or {}).get(nombre)
        fill = PatternFill("solid", fgColor=color or color_encabezado)
        font = Font(bold=True, color="000000" if color else "FFFFFF", size=11, name="Calibri")
        encabezado.append(_celda(ws, nombre, fill=fill, font=font, alignment=alineado, border=_BORDE_FINO))
    ws.append(encabezado)

    # Columnas que necesitan formato por celda: fechas y colores por valor
    fills = {columnas.index(c): {v: PatternFill("solid", fgColor=col) for v, col in m.items()}
             for c, m in (colores_celda or {}).items() if c in columnas}
    fechas = [j for j, c in enumerate(df.columns)
              if pd.api.types.is_datetime64_any_dtype(df[c])
              or (df[c].dtype == object and df[c].map(lambda v: isinstance(v, date)).any())]
    for valores in _valores(df):
        for j in fechas:
            fmt = _formato_fecha(valores[j])
            if fmt:
                valores[j] = _celda(ws, valores[j], number_format=fmt)
        for j, m in fills.items():
            if valores[j] in m:
                valores[j] = _celda(ws, valores[j], fill=m[valores[j]])
        ws.append(valores)

    if len(df) and n:
        ref = f"A1:{get_column_letter(n)}{len(df) + 1}"
        tabla = Table(displayName=hoja, ref=ref, autoFilter=AutoFilter(ref=ref),
                      tableColumns=[TableColumn(id=i + 1, name=c) for i, c in enumerate(columnas)])
        tabla.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
        with warnings.catch_warnings():
            # Avisa que hay que declarar las columnas a mano; ya están declaradas arriba
            warnings.simplefilter("ignore", UserWarning)
            ws.add_table(tabla)


# ─────────────────────────────────────────────────────────────────────────────
#  MAIN (línea de comandos)
# ─────────────────────────────────────────────────────────────────────────────
def main():
    print("=" * 65)
    print("  RESEGUIMIENTOS AMGEN — Consolidación y validación")
    print("=" * 65)

    reporte = generar_reporte(
        ARCHIVO_PACIENTES,
        ARCHIVO_INTERACCIONES,
        ARCHIVO_APLICACIONES if ARCHIVO_APLICACIONES.exists() else None,
    )

    print(f"\nExportando... {ARCHIVO_SALIDA}")
    try:
        exportar_excel(reporte, ARCHIVO_SALIDA)
    except PermissionError:
        print()
        print("ERROR: No se puede guardar el archivo.")
        print("Probablemente tienes Reseguimientos.xlsx abierto en Excel.")
        print("Cierralo y vuelve a ejecutar el script.")
        return

    t = reporte["totales"]
    df_final = reporte["consolidado"]
    print()
    print("=" * 65)
    print("  RESUMEN FINAL")
    print("=" * 65)
    print(f"  Pacientes procesados     : {t['pacientes']}")
    print(f"  Casos SIN interaccion    : {t['sin_interaccion']}")
    if t["con_aplicaciones"]:
        frec_err = (df_final[COL_PROXIMA_APLICACION] == "Frecuencia Errónea").sum()
        print(f"  Casos SIN aplicacion     : {df_final[COL_FECHA_APLICACION].isna().sum()}")
        print(f"  Frecuencia Errónea       : {frec_err}")
    print(f"  Fecha de corte           : {reporte['fecha_corte']:%d/%m/%Y}")
    print(f"  Casos validados          : {t['casos_validados']}")
    print(f"  Casos con hallazgos      : {t['casos_con_hallazgos']}")
    print(f"  Hallazgos                : {t['hallazgos']}")
    print()
    print(reporte["resumen"].to_string(index=False))
    print(f"\n  Archivo generado         : {ARCHIVO_SALIDA}")
    print("=" * 65)


if __name__ == "__main__":
    main()
