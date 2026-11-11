# -*- coding: utf-8 -*-
"""
==============================================================
  RESEGUIMIENTOS AMGEN — Validación de calidad de seguimientos
==============================================================
Revisa cada seguimiento (interacción) y genera un hallazgo por cada
error encontrado. Validaciones:

  1. Nota vs Razón del estado
     - El estado registrado en la nota (ESTADO DESPUÉS DE CONTACTO o el
       inicio del DETALLE) coincide con la "Razón del estado" del paciente
       (solo para el último seguimiento del caso).
     - El estado del campo y el del DETALLE no se contradicen.
     - "Interacción exitosa = Si" no puede tener estado No Contactable.
     - La nota registra algún estado.

  2. Fecha de la llamada vs fecha programada
     - La llamada se hizo en la fecha de próximo contacto que dejó la
       nota anterior (± tolerancia).
     - La nota deja fecha de próximo contacto y no es anterior a la llamada.
     - Pacientes activos: el próximo contacto no está vencido a la fecha
       de corte.

  3. Aplicación mencionada en la nota registrada en Aplicaciones
     - La FECHA DE ÚLTIMA APLICACIÓN de la nota existe en Aplicaciones
       como aplicada (± tolerancia) y no es posterior a la llamada.
     - Si el DETALLE confirma una aplicación, debe existir una aplicación
       registrada (en la fecha mencionada o en el periodo del seguimiento).

  4. Paciente en gestión de cita
     - La nota documenta especialidad/médico y el estado de la cita
       (fecha si está agendada, o el motivo si aún no lo está).
"""

import re
import difflib
import unicodedata
from dataclasses import dataclass, field

import numpy as np
import pandas as pd

COL_CASO = "Número del caso"

# Columnas de la interacción preparada (ver consolidar_data.preparar_interacciones)
COL_INICIO     = "Inicio"
COL_ID_INT     = "ID Interacción"
COL_CREADO_POR = "Creado por"
COL_EXITOSA    = "Interacción exitosa?"

V1 = "1. Nota vs Razón del estado"
V2 = "2. Fecha llamada vs programada"
V3 = "3. Aplicación registrada"
V4 = "4. Información de cita"
VALIDACIONES = [V1, V2, V3, V4]

# Columnas de conteo por paciente en el consolidado
COL_N_CRITERIOS = "Criterios con error (0 a 4)"
COL_N_HALLAZGOS = "N° hallazgos (errores encontrados)"

ALTA, MEDIA, BAJA = "Alta", "Media", "Baja"

# Estados que cierran el seguimiento (no exigen próximo contacto)
ESTADOS_TERMINALES = {
    "FALLECIMIENTO",
    "SERVICIOS PSP FINALIZADOS",
    "TRATAMIENTO FINALIZADO",
}

# Variantes frecuentes en las notas → estado canónico (normalizado)
ALIAS_ESTADOS = {
    "TRATAMIENTO"                         : "EN TRATAMIENTO",
    "EN TRATAMIENTO ACTIVO"               : "EN TRATAMIENTO",
    "NO CONTACTABLE"                      : "PACIENTE NO CONTACTABLE",
    "PACIENTE NO CONTACTADO"              : "PACIENTE NO CONTACTABLE",
    "EN ESPERA DEL MEDICAMENTO"           : "EN ESPERA DEL MEDICAMENTO O COMPRA PARTICULAR",
    "EN ESPERA DE MEDICAMENTO"            : "EN ESPERA DEL MEDICAMENTO O COMPRA PARTICULAR",
    "EN ESPERA DE MEDICAMENTO OPL"        : "EN ESPERA DEL MEDICAMENTO O COMPRA PARTICULAR",
    "EN ESPERA DEL MEDICAMENTO OPL"       : "EN ESPERA DEL MEDICAMENTO O COMPRA PARTICULAR",
    "PENDIENTE REFORMULACION"             : "PENDIENTE DE REFORMULACION",
    "PROBLEMAS ADMINISTRATIVOS OPL"       : "PROBLEMAS ADMINISTRATIVOS OPERADOR LOGISTICO",
    "PROBLEMAS ADMINISTRATIVOS OL"        : "PROBLEMAS ADMINISTRATIVOS OPERADOR LOGISTICO",
    "DECISION PACIENTE SIN CITA MEDICA"   : "DECISION PACIENTE SIN CITA",
    "DECISION MEDICO"                     : "DECISION MEDICA",
    "DECISION PROPIA"                     : "DECISION PACIENTE",
    "PACIENTE REHUSA"                     : "PACIENTE REHUSA MEDICAMENTO",
    "REHUSA MEDICAMENTO"                  : "PACIENTE REHUSA MEDICAMENTO",
}

ESTADOS_BASE = [
    "EN TRATAMIENTO", "DECISION MEDICA", "SERVICIOS PSP FINALIZADOS",
    "DECISION PACIENTE", "PACIENTE NO CONTACTABLE", "PROBLEMAS ADMINISTRATIVOS EPS",
    "FALLECIMIENTO", "CAMBIO DE TRATAMIENTO FARMACOLOGICO", "TRATAMIENTO FINALIZADO",
    "CAMBIO DE TRATAMIENTO", "PACIENTE REHUSA MEDICAMENTO", "NO FORMULADO",
    "PROBLEMAS ADMINISTRATIVOS OPERADOR LOGISTICO",
    "EN ESPERA DEL MEDICAMENTO O COMPRA PARTICULAR", "EVENTO ADVERSO",
    "FALTA CITA MEDICA", "APROBACION EN ESTUDIO", "PENDIENTE DE REFORMULACION",
    "DECISION MEDICA EXAMENES", "DECISION PACIENTE SIN CITA",
    "DECISION MEDICA EVENTO ADVERSO", "CAMBIO DE TRATAMIENTO NO FARMACOLOGICO",
    "DECISION MEDICA OTROS", "OTRAS PATOLOGIAS", "FORMULAS FALLIDAS",
    "PENDIENTE DEFINICION DE CONDUCTA MEDICA", "DECISION PACIENTE COVID 19",
]

ESTADOS_GESTION_CITA_DEFECTO = (
    "Falta Cita Médica",
    "Decisión Paciente - Sin Cita",
    "Pendiente de Reformulación",
)

MESES = {
    "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6,
    "JULIO": 7, "AGOSTO": 8, "SEPTIEMBRE": 9, "SETIEMBRE": 9, "OCTUBRE": 10,
    "NOVIEMBRE": 11, "DICIEMBRE": 12,
}
_MESES_RE = "|".join(MESES)

RE_FECHA_NUM = re.compile(r"(?<![\d/])(\d{1,2})\s?[/\-.]\s?(\d{1,2})(?:\s?[/\-.]\s?(\d{4}|\d{2}))?(?![\d/])")
RE_FECHA_TXT = re.compile(rf"\b(\d{{1,2}})\s+(?:DE\s+)?({_MESES_RE})\b(?:\s*(?:DE|DEL)?\s*(\d{{4}}))?")
RE_SOLO_MES  = re.compile(rf"\b(?:MES\s+DE|PARA|EN|A\s+FINALES?\s+DE|MEDIADOS\s+DE|INICIOS?\s+DE|HASTA)\s+({_MESES_RE})\b|\bFINAL(?:ES)?\s+DE(?:L)?\s+MES\b")

# Aplicación confirmada en el texto (texto ya normalizado: mayúsculas, sin tildes).
# Solo tiempos pasados: "CONFIRMAR", "REALIZAR", "REALIZARA" no cuentan.
RE_APLICACION_CONFIRMADA = re.compile(
    r"\bCONFIRMA(?:N|RON|DO|DA)?\s+(?:(?:LA|SU|EL|DE|QUE|ULTIMA|SE|LE|REALIZO|REALIZARON)\s+){0,4}APLICACI\w*"
    r"|\bSE\s+(?:LE\s+|LA\s+|LO\s+)?(?:REALIZO\s+(?:LA\s+)?)?APLICO\b"
    r"|\b(?:YA\s+)?(?:SE\s+)?(?:LE\s+|LA\s+|LO\s+)?APLICARON\b|\bAPLICO\b"
    r"|\bREALIZ(?:O|ARON|ADA)\s+(?:LA\s+|SU\s+)?APLICACION\b"
    r"|\bAPLICACION\s+(?:SIN\s+NOVEDAD|EXITOSA|REALIZADA|EFECTIVA)"
)
# Lo que sigue a la mención indica que es futura / pendiente, no una aplicación hecha
RE_APLICACION_FUTURA = re.compile(r"^\W*(?:(?:DE\s+)?(?:LA\s+)?(?:PROGRAMAD|PENDIENTE|PROXIMA|PARA\s+(?:EL|LA|ESTE|ESTA|MANANA|EL\s+DIA))|A\s+LAS)")
# Palabras que, entre la mención y una fecha, indican que la fecha es de otra cosa
RE_CORTE_FECHA = re.compile(r"\bCITA\b|PROXIM|FORMUL|PROGRAM|ENTREGA|CONTROL|ORDEN|EXAMEN|\s/|/\s")
RE_NEGACION = re.compile(r"\b(?:NO|NI|SIN|TAMPOCO|AUN\s+NO|TODAVIA\s+NO|NUNCA|PENDIENTE(?:\s+DE)?)\s+(?:\w+\s+){0,2}$")

RE_ESPECIALIDAD = re.compile(
    r"REUMATOLOG|ENDOCRIN|INTERNIST|MEDICINA\s+INTERNA|ORTOPED|HEMATOLOG|ONCOLOG|CARDIOLOG|CARDIOVASCULAR"
    r"|NEFROLOG|NEUROLOG|DERMATOLOG|GASTROENTEROLOG|GINECOLOG|FISIATR|GERIATR|OFTALMOLOG|NEUMOLOG"
    r"|MEDICO\s+(?:GENERAL|TRATANTE|ESPECIALISTA|DE\s+CONTROL|FAMILIAR)|MEDICINA\s+(?:GENERAL|FAMILIAR)"
    r"|ESPECIALISTA|\bDR\.?\s|\bDRA\.?\s|DOCTOR|INTEGRACION\s+VITAL|CLINICA\s+DE\s+OSTEOPOROSIS"
)
RE_CITA_AGENDADA = re.compile(
    r"CITA\s+(?:YA\s+)?(?:AGENDADA|ASIGNADA|PROGRAMADA|CONFIRMADA)|\bAGENDADA\b"
    r"|(?:TIENE|TIENEN|CUENTA\s+CON|CUENTAN\s+CON|LE\s+(?:DIERON|ASIGNARON|PROGRAMARON))\s+(?:LA\s+|UNA\s+)?(?:CITA|CONTROL)"
    r"|CITA\s+(?:\w+\s+){0,6}?(?:PARA\s+)?EL\s+(?:DIA\s+)?\d"
)
RE_NEGACION_CITA = re.compile(r"\b(?:NO|AUN\s+NO|TODAVIA\s+NO|SIN|CUANDO(?:\s+YA)?)\s+(?:\w+\s+){0,2}$")
RE_CITA_SIN_AGENDAR = re.compile(
    r"SIN\s+AGEND|NO\s+HAY\s+AGENDA|SIN\s+AGENDA|NO\s+(?:HAY|TIENE|TIENEN)\s+(?:DISPONIBILIDAD|CONTRATO)"
    r"|NO\s+(?:HA|HAN)\s+(?:SOLICITADO|SACADO|PEDIDO|REALIZADO|AGENDADO|PODIDO|LOGRADO|TENIDO)"
    r"|NO\s+(?:TIENE|CUENTA\s+CON|CUENTAN\s+CON)\s+(?:LA\s+)?CITA|NO\s+(?:LE\s+)?HAN\s+(?:DADO|ASIGNADO)"
    r"|NO\s+LOGR|NO\s+CONSIGU|NO\s+(?:HA|HAN)\s+ENCONTRADO|EN\s+PROCESO|REINICIAR\s+PROCESO|NO\s+(?:HA\s+)?SIDO\s+POSIBLE|PROCESO\s+DE\s+(?:AGENDAMIENTO|ASIGNACION)"
    r"|LISTA\s+DE\s+ESPERA|PENDIENTE\s+(?:DE\s+)?(?:LA\s+)?(?:ASIGNA|AUTORIZA|AGENDA|CITA|EXAMEN|RESULTADO)"
    r"|(?:EN\s+)?ESPERA\s+DE\s+(?:LA\s+|UNA\s+)?(?:CITA|AUTORIZA|AGENDA|EXAMEN|RESULTADO|ASIGNA|DISPONIBILIDAD|CIRUGIA)"
    r"|(?:SOLICITAR|PEDIR|SACAR|SACANDO|GESTIONAR|AGENDAR|PROGRAMAR|APARTAR|BUSCAR)\s+(?:LA\s+|UNA\s+)?CITA"
    r"|VENCI|CANCELA|REPROGRAM|TRAMITE|BUSQUEDA\s+DE\s+(?:LA\s+)?CITA|CUANDO\s+(?:YA\s+)?(?:TENGA|LE\s+DEN)"
    r"|AUTORIZACION|TUTELA|DESACATO|CAMBIO\s+DE\s+EPS|FUERA\s+DE\s+LA\s+CIUDAD|NO\s+DESEA|NO\s+QUIERE"
)
# Fechas relativas o solo con día ("HOY", "EL 29") cuentan como fecha documentada
RE_FECHA_RELATIVA = re.compile(r"\bHOY\b|\bMANANA\b|\bEL\s+(?:DIA\s+)?\d{1,2}\b(?!\s*(?:DE\s+)?(?:DIAS|MESES|ANOS|SEMANAS))")
RE_PLANTILLA_ESPECIALIDAD = re.compile(r"/\s*ESPECIALIDAD\s*/")
RE_INACTIVACION = re.compile(r"INACTIV|DESMONTE|DESCALAMIENTO|DESESCALAMIENTO|SPEECH|SPPECH")
RE_NO_CONTACTO = re.compile(
    r"NO\s+CONTEST|NO\s+RESPOND|SIN\s+RESPUESTA|NO\s+CONTACTABLE|BUZON|APAGADO|FUERA\s+DE\s+SERVICIO"
    r"|NUMERO\s+(?:ERRADO|EQUIVOCADO|INCORRECTO|NO\s+EXISTE)|RECHAZA\s+(?:LA\s+)?LLAMADA|CUELGA"
)


# ─────────────────────────────────────────────────────────────────────────────
#  PARÁMETROS
# ─────────────────────────────────────────────────────────────────────────────
@dataclass
class ParametrosValidacion:
    # Fecha de referencia para "próximo contacto vencido". None → la última
    # fecha de llamada encontrada en el archivo de Interacciones.
    fecha_corte: pd.Timestamp | None = None
    tolerancia_llamada_dias: int = 3       # ± días entre fecha programada y llamada real
    tolerancia_aplicacion_dias: int = 3    # ± días entre fecha en nota y fecha en Aplicaciones
    ventana_aplicacion_dias: int = 35      # si la nota no trae fecha, buscar aplicación en estos días previos
    solo_ultimo_seguimiento: bool = True   # False → valida todos los seguimientos del archivo
    estados_gestion_cita: tuple = field(default_factory=lambda: ESTADOS_GESTION_CITA_DEFECTO)


# ─────────────────────────────────────────────────────────────────────────────
#  HELPERS DE TEXTO Y FECHAS
# ─────────────────────────────────────────────────────────────────────────────
def normalizar_texto(texto) -> str:
    """Mayúsculas, sin tildes y con espacios simples (conserva / y dígitos)."""
    if texto is None or (isinstance(texto, float) and np.isnan(texto)):
        return ""
    texto = unicodedata.normalize("NFD", str(texto))
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    texto = re.sub(r"[\xa0​]+", " ", texto.upper())
    return re.sub(r"\s+", " ", texto).strip()


def _norm_estado(texto) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[^A-Z0-9 ]", " ", normalizar_texto(texto))).strip()


def parse_fecha_nota(serie: pd.Series) -> pd.Series:
    """Convierte fechas DD/MM/AAAA escritas en la nota. 01/01/1999 (vacío) → NaT."""
    extraida = serie.astype("string").str.extract(r"(\d{1,2})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(\d{4})")
    fechas = pd.to_datetime(
        extraida[2] + "-" + extraida[1].str.zfill(2) + "-" + extraida[0].str.zfill(2),
        format="%Y-%m-%d", errors="coerce",
    )
    return fechas.where(fechas.dt.year > 2000)


def _inferir_anio(dia: int, mes: int, ref: pd.Timestamp) -> pd.Timestamp | None:
    """Elige el año que deja la fecha más cerca de la fecha de referencia."""
    candidatas = []
    for anio in (ref.year - 1, ref.year, ref.year + 1):
        try:
            candidatas.append(pd.Timestamp(year=anio, month=mes, day=dia))
        except ValueError:
            pass
    return min(candidatas, key=lambda f: abs(f - ref)) if candidatas else None


def fechas_en_texto(texto: str, ref: pd.Timestamp) -> list[pd.Timestamp]:
    """Fechas completas mencionadas en el texto (ya normalizado)."""
    fechas = []
    for d, m, a in RE_FECHA_NUM.findall(texto):
        d, m = int(d), int(m)
        if not (1 <= d <= 31 and 1 <= m <= 12):
            continue
        if a:
            anio = int(a) + (2000 if len(a) == 2 else 0)
            try:
                fechas.append(pd.Timestamp(year=anio, month=m, day=d))
            except ValueError:
                pass
        elif pd.notna(ref):
            f = _inferir_anio(d, m, ref)
            if f is not None:
                fechas.append(f)
    for d, mes, a in RE_FECHA_TXT.findall(texto):
        d, m = int(d), MESES[mes]
        if not 1 <= d <= 31:
            continue
        try:
            fechas.append(pd.Timestamp(year=int(a), month=m, day=d) if a else _inferir_anio(d, m, ref))
        except (ValueError, TypeError):
            pass
    return [f for f in fechas if f is not None and f.year > 2000]


def fecha_tras_mencion(texto: str, desde: int, ref: pd.Timestamp, max_dist: int = 30) -> list[pd.Timestamp]:
    """
    Fechas que aparecen justo después de una mención (p.ej. 'CONFIRMA APLICACION
    EL 27 DE MARZO'), sin que en medio haya otra cosa ('... CITA EL 22/07').
    """
    candidatas = []
    for patron in (RE_FECHA_NUM, RE_FECHA_TXT):
        for m in patron.finditer(texto, desde):
            if m.start() - desde > max_dist:
                break
            if not RE_CORTE_FECHA.search(texto[desde:m.start()]):
                candidatas.append((m.start(), m.group(0)))
    fechas = []
    for _, fragmento in sorted(candidatas):
        fechas.extend(fechas_en_texto(fragmento, ref))
    return fechas


def _fmt(fecha) -> str:
    return fecha.strftime("%d/%m/%Y") if pd.notna(fecha) else ""


class CanonEstados:
    """Lleva cualquier texto de estado a su forma canónica (normalizada)."""

    def __init__(self, razones_conocidas):
        canon = set(ESTADOS_BASE) | {_norm_estado(r) for r in razones_conocidas if pd.notna(r)}
        canon.discard("")
        self.canon = canon
        self.lookup = {c: c for c in canon}
        self.lookup.update(ALIAS_ESTADOS)
        self.claves_largo = sorted(self.lookup, key=len, reverse=True)
        self._cache = {}

    def __call__(self, texto) -> str | None:
        n = _norm_estado(texto)
        n = re.sub(r"^(?:NO\s+)?SE\s+REPORTA\s+FV\s*", "", n)
        n = re.sub(r"^(?:INACTIVO|INACTIVA|INACTIVACION|PACIENTE\s+INACTIVO)\s+(?=\w)", "", n)
        n = re.sub(r"\bDESICION\b", "DECISION", n).strip()
        if not n:
            return None
        if n in self._cache:
            return self._cache[n]
        res = self.lookup.get(n)
        if res is None:
            for k in self.claves_largo:
                if n.startswith(k + " "):
                    res = self.lookup[k]
                    break
        if res is None and len(n) >= 8:
            cerca = difflib.get_close_matches(n, self.claves_largo, n=1, cutoff=0.86)
            if cerca:
                res = self.lookup[cerca[0]]
        self._cache[n] = res
        return res


def estados_compatibles(a: str, b: str) -> bool:
    """Igual, o uno es la versión general del otro (p.ej. DECISION MEDICA vs DECISION MEDICA EXAMENES)."""
    return a == b or a.startswith(b + " ") or b.startswith(a + " ")


def estado_desde_detalle(detalle: str, canon: CanonEstados) -> str | None:
    """El DETALLE suele iniciar con 'ESTADO / ...'. Revisa los dos primeros segmentos."""
    segmentos = [s for s in re.split(r"/+", detalle) if s.strip()][:2]
    for seg in segmentos:
        est = canon(seg)
        if est:
            return est
    return None


# ─────────────────────────────────────────────────────────────────────────────
#  ÍNDICE DE APLICACIONES
# ─────────────────────────────────────────────────────────────────────────────
class IndiceAplicaciones:
    """Fechas de aplicación por caso (aplicadas = Si / No), ordenadas."""

    def __init__(self, df_apl: pd.DataFrame | None):
        self.si, self.no = {}, {}
        self.disponible = df_apl is not None and len(df_apl) > 0
        if not self.disponible:
            return
        df = df_apl[[COL_CASO, "Fecha de Aplicación", "Aplicado"]].dropna(subset=["Fecha de Aplicación"])
        df = df.sort_values("Fecha de Aplicación")
        for (caso, aplicado), g in df.groupby([COL_CASO, "Aplicado"], sort=False):
            destino = self.si if aplicado == "SI" else self.no
            destino[caso] = g["Fecha de Aplicación"].values.astype("datetime64[D]")

    @staticmethod
    def _en_rango(arr, ini, fin):
        if arr is None or len(arr) == 0:
            return None
        ini, fin = np.datetime64(ini.date()), np.datetime64(fin.date())
        i = np.searchsorted(arr, ini, side="left")
        if i < len(arr) and arr[i] <= fin:
            return pd.Timestamp(arr[i])
        return None

    def aplicada_en(self, caso, ini, fin):
        return self._en_rango(self.si.get(caso), ini, fin)

    def no_aplicada_en(self, caso, ini, fin):
        return self._en_rango(self.no.get(caso), ini, fin)

    def ultima_aplicada_antes(self, caso, fecha):
        arr = self.si.get(caso)
        if arr is None or len(arr) == 0:
            return None
        i = np.searchsorted(arr, np.datetime64(fecha.date()), side="right")
        return pd.Timestamp(arr[i - 1]) if i > 0 else None


# ─────────────────────────────────────────────────────────────────────────────
#  VALIDACIONES
# ─────────────────────────────────────────────────────────────────────────────
def _hallazgo(row, validacion, severidad, mensaje):
    return {
        COL_CASO          : row[COL_CASO],
        COL_ID_INT        : row.get(COL_ID_INT),
        "Fecha llamada"   : row[COL_INICIO],
        COL_CREADO_POR    : row.get(COL_CREADO_POR),
        "Validación"      : validacion,
        "Severidad"       : severidad,
        "Hallazgo"        : mensaje,
        "Razón del estado": row.get("Razón del estado"),
        "DETALLE"         : row.get("DETALLE"),
    }


def _validar_estado(row, ctx, out):
    """V1 — Coherencia entre la nota y la razón del estado."""
    canon = ctx["canon"]
    detalle = row["_detalle_n"]
    if not detalle and pd.isna(row.get("ESTADO DESPUÉS DE CONTACTO")):
        out.append(_hallazgo(row, V1, ALTA, "El seguimiento no tiene nota (Recomendaciones vacías)"))
        return

    est_campo = canon(row.get("ESTADO DESPUÉS DE CONTACTO")) if pd.notna(row.get("ESTADO DESPUÉS DE CONTACTO")) else None
    est_det   = estado_desde_detalle(detalle, canon)
    est_nota  = est_campo or est_det

    if est_nota is None:
        out.append(_hallazgo(row, V1, MEDIA, "La nota no registra un estado reconocible del paciente"))
    elif est_campo and est_det and not estados_compatibles(est_campo, est_det):
        out.append(_hallazgo(row, V1, MEDIA,
            f"Estado contradictorio en la nota: campo '{est_campo}' vs detalle '{est_det}'"))

    if row["_es_ultimo"]:
        razon = row.get("Razón del estado")
        razon_c = canon(razon) if pd.notna(razon) else None
        estados_nota = [e for e in (est_campo, est_det) if e]
        if razon_c is None:
            out.append(_hallazgo(row, V1, MEDIA, "El paciente no tiene Razón del estado registrada"))
        elif estados_nota and not any(estados_compatibles(e, razon_c) for e in estados_nota):
            if RE_INACTIVACION.search(detalle) and normalizar_texto(row.get("Estado")) == "INACTIVO":
                out.append(_hallazgo(row, V1, BAJA,
                    f"La nota registra inactivación con estado '{est_nota}'; verificar que la "
                    f"Razón del estado '{razon}' sea la correcta"))
            else:
                out.append(_hallazgo(row, V1, ALTA,
                    f"Razón del estado '{razon}' no coincide con el estado de la nota '{est_nota}'"))

    exitosa = normalizar_texto(row.get(COL_EXITOSA))
    if est_nota == "PACIENTE NO CONTACTABLE" and exitosa == "SI":
        out.append(_hallazgo(row, V1, MEDIA,
            "Interacción marcada como exitosa pero la nota indica Paciente No Contactable"))
    elif est_nota and est_nota != "PACIENTE NO CONTACTABLE" and exitosa == "NO" and RE_NO_CONTACTO.search(detalle):
        out.append(_hallazgo(row, V1, MEDIA,
            f"La nota describe que no hubo contacto pero registra el estado '{est_nota}'"))


def _validar_fechas(row, ctx, out):
    """V2 — Fecha de la llamada vs fecha programada."""
    p = ctx["params"]
    tol = p.tolerancia_llamada_dias
    fecha_llamada = row[COL_INICIO].normalize()
    prox_prev = row["_prox_prev"]
    # Varios intentos el mismo día: la fecha programada aplica al primero
    mismo_dia = pd.notna(row["_inicio_prev"]) and row["_inicio_prev"].normalize() == fecha_llamada

    if pd.notna(prox_prev) and not mismo_dia:
        dif = (fecha_llamada - prox_prev).days
        if dif > tol:
            out.append(_hallazgo(row, V2, ALTA,
                f"Llamada realizada {dif} días DESPUÉS de la fecha programada ({_fmt(prox_prev)})"))
        elif dif < -tol:
            out.append(_hallazgo(row, V2, MEDIA,
                f"Llamada realizada {-dif} días ANTES de la fecha programada ({_fmt(prox_prev)})"))

    prox = row["_prox"]
    terminal = row["_estado_nota"] in ESTADOS_TERMINALES or (
        row["_es_ultimo"] and normalizar_texto(row.get("Estado")) == "INACTIVO")
    if pd.isna(prox):
        if not terminal:
            out.append(_hallazgo(row, V2, ALTA, "La nota no registra FECHA DE PRÓXIMO CONTACTO"))
    elif prox < fecha_llamada:
        out.append(_hallazgo(row, V2, ALTA,
            f"FECHA DE PRÓXIMO CONTACTO ({_fmt(prox)}) es anterior a la fecha de la llamada"))
    elif row["_es_ultimo"] and not terminal and prox < ctx["fecha_corte"] - pd.Timedelta(days=tol):
        out.append(_hallazgo(row, V2, ALTA,
            f"Seguimiento vencido: próximo contacto programado para {_fmt(prox)} "
            f"y no hay llamada posterior (corte {_fmt(ctx['fecha_corte'])})"))


def _validar_aplicacion(row, ctx, out):
    """V3 — Si la nota menciona aplicación, debe estar registrada en Aplicaciones."""
    idx: IndiceAplicaciones = ctx["apl"]
    if not idx.disponible:
        return
    p = ctx["params"]
    tol = pd.Timedelta(days=p.tolerancia_aplicacion_dias)
    caso = row[COL_CASO]
    fecha_llamada = row[COL_INICIO].normalize()
    fua = row["_ult_apl"]

    # a) Fecha de última aplicación diligenciada en la nota
    if pd.notna(fua) and (not ctx["todos"] or row["_ult_apl_nueva"]):
        if fua > fecha_llamada + pd.Timedelta(days=1):
            out.append(_hallazgo(row, V3, ALTA,
                f"FECHA DE ÚLTIMA APLICACIÓN ({_fmt(fua)}) es posterior a la fecha de la llamada"))
        elif idx.aplicada_en(caso, fua - tol, fua + tol) is None:
            if idx.no_aplicada_en(caso, fua - tol, fua + tol) is not None:
                msg = f"La aplicación del {_fmt(fua)} está registrada en Aplicaciones como NO aplicada"
            else:
                ultima = idx.ultima_aplicada_antes(caso, fecha_llamada)
                msg = (f"La aplicación del {_fmt(fua)} indicada en la nota no está registrada en Aplicaciones"
                       + (f" (última registrada: {_fmt(ultima)})" if ultima is not None else " (sin aplicaciones registradas)"))
            out.append(_hallazgo(row, V3, ALTA, msg))

    # b) El DETALLE confirma una aplicación
    detalle = row["_detalle_n"]
    m = None
    for cand in RE_APLICACION_CONFIRMADA.finditer(detalle):
        if RE_NEGACION.search(detalle[max(0, cand.start() - 25):cand.start()]):
            continue
        if RE_APLICACION_FUTURA.search(detalle[cand.end():cand.end() + 30]):
            continue
        m = cand
        break
    if m is None:
        return

    fechas = [f for f in fecha_tras_mencion(detalle, m.end(), fecha_llamada)
              if f <= fecha_llamada + pd.Timedelta(days=1)]
    if fechas:
        f = fechas[0]
        if pd.notna(fua) and abs(f - fua) <= tol:
            return  # ya validada en (a)
        if idx.aplicada_en(caso, f - tol, f + tol) is None:
            out.append(_hallazgo(row, V3, ALTA,
                f"La nota confirma aplicación del {_fmt(f)} pero no está registrada en Aplicaciones"))
        return

    if pd.notna(fua):
        return  # la fecha del campo ya respalda la mención
    ini = row["_inicio_prev"] if pd.notna(row["_inicio_prev"]) else fecha_llamada - pd.Timedelta(days=p.ventana_aplicacion_dias)
    ini = max(ini.normalize(), fecha_llamada - pd.Timedelta(days=p.ventana_aplicacion_dias)) - tol
    if idx.aplicada_en(caso, ini, fecha_llamada + pd.Timedelta(days=1)) is None:
        out.append(_hallazgo(row, V3, ALTA,
            f"La nota confirma una aplicación pero no hay aplicación registrada entre "
            f"{_fmt(ini)} y {_fmt(fecha_llamada)}"))
    else:
        out.append(_hallazgo(row, V3, BAJA,
            "La nota confirma aplicación pero no diligencia FECHA DE ÚLTIMA APLICACIÓN"))


def _validar_cita(row, ctx, out):
    """V4 — Paciente en gestión de cita: la nota debe documentar la cita."""
    estados_cita = ctx["estados_cita"]
    estado = row["_estado_nota"]
    if row["_es_ultimo"]:
        razon_c = ctx["canon"](row.get("Razón del estado")) if pd.notna(row.get("Razón del estado")) else None
        en_gestion = (razon_c in estados_cita) or (estado in estados_cita)
    else:
        en_gestion = estado in estados_cita
    if not en_gestion:
        return

    detalle = row["_detalle_n"]
    fecha_llamada = row[COL_INICIO].normalize()
    fecha_control = row["_fecha_control"]
    fecha_en_detalle = bool(fechas_en_texto(detalle, fecha_llamada) or RE_FECHA_RELATIVA.search(detalle))
    tiene_fecha = fecha_en_detalle or pd.notna(fecha_control)
    solo_mes = bool(RE_SOLO_MES.search(detalle))
    agendada = any(not RE_NEGACION_CITA.search(detalle[max(0, m.start() - 20):m.start()])
                   for m in RE_CITA_AGENDADA.finditer(detalle))
    sin_agendar = bool(RE_CITA_SIN_AGENDAR.search(detalle))

    if not RE_ESPECIALIDAD.search(RE_PLANTILLA_ESPECIALIDAD.sub("/", detalle)):
        msg = ("La nota deja el texto de plantilla 'ESPECIALIDAD' sin diligenciar"
               if RE_PLANTILLA_ESPECIALIDAD.search(detalle)
               else "La nota no indica la especialidad o médico de la cita")
        out.append(_hallazgo(row, V4, MEDIA, msg))

    if agendada and not tiene_fecha:
        if solo_mes:
            out.append(_hallazgo(row, V4, BAJA, "Cita agendada sin fecha exacta (solo indica el mes)"))
        else:
            out.append(_hallazgo(row, V4, ALTA, "La nota indica cita agendada pero no registra la fecha"))
    elif not (agendada or tiene_fecha or solo_mes or sin_agendar):
        out.append(_hallazgo(row, V4, ALTA,
            "No documenta el estado de la cita (fecha si está agendada, o motivo si no lo está)"))

    if fecha_en_detalle and pd.isna(fecha_control) and agendada:
        out.append(_hallazgo(row, V4, BAJA,
            "La fecha de la cita está en el detalle pero no en FECHA DE CONTROL CON ESPECIALISTA"))
    if pd.notna(fecha_control) and fecha_control < fecha_llamada and agendada and not sin_agendar:
        out.append(_hallazgo(row, V4, BAJA,
            f"FECHA DE CONTROL CON ESPECIALISTA ({_fmt(fecha_control)}) ya pasó; validar resultado de la cita"))


# ─────────────────────────────────────────────────────────────────────────────
#  ORQUESTADOR
# ─────────────────────────────────────────────────────────────────────────────
def validar_seguimientos(df_int: pd.DataFrame, df_pac: pd.DataFrame,
                         df_apl: pd.DataFrame | None,
                         params: ParametrosValidacion | None = None):
    """
    df_int : todas las interacciones preparadas (una fila por seguimiento,
             con columnas de Recomendaciones ya expandidas e Inicio datetime)
    df_pac : BD pacientes (Número del caso, Estado, Razón del estado, ...)
    df_apl : todas las aplicaciones (Número del caso, Fecha de Aplicación, Aplicado)

    Retorna (df_hallazgos, df_resumen_caso, fecha_corte)
    """
    params = params or ParametrosValidacion()

    df = df_int[df_int[COL_INICIO].notna()].sort_values([COL_CASO, COL_INICIO]).copy()
    g = df.groupby(COL_CASO, sort=False)

    df["_prox"]          = parse_fecha_nota(df["FECHA DE PRÓXIMO CONTACTO"])
    df["_ult_apl"]       = parse_fecha_nota(df["FECHA DE ÚLTIMA APLICACIÓN"])
    df["_fecha_control"] = parse_fecha_nota(df["FECHA DE CONTROL CON ESPECIALISTA"])
    df["_prox_prev"]     = g["_prox"].shift()
    df["_inicio_prev"]   = g[COL_INICIO].shift()
    df["_ult_apl_nueva"] = df["_ult_apl"].ne(g["_ult_apl"].shift())
    df["_es_ultimo"]     = ~df.duplicated(subset=[COL_CASO], keep="last")
    df["_detalle_n"]     = df["DETALLE"].map(normalizar_texto)

    cols_pac = [c for c in [COL_CASO, "Estado", "Razón del estado", "Enfermero(a) de campo"] if c in df_pac.columns]
    df = df.merge(df_pac[cols_pac], on=COL_CASO, how="left")

    canon = CanonEstados(df_pac["Razón del estado"].dropna().unique() if "Razón del estado" in df_pac else [])
    df["_estado_nota"] = [
        (canon(c) if pd.notna(c) else None) or estado_desde_detalle(d, canon)
        for c, d in zip(df["ESTADO DESPUÉS DE CONTACTO"], df["_detalle_n"])
    ]

    fecha_corte = pd.Timestamp(params.fecha_corte).normalize() if params.fecha_corte is not None \
        else df[COL_INICIO].max().normalize()

    ctx = {
        "params"      : params,
        "canon"       : canon,
        "apl"         : IndiceAplicaciones(df_apl),
        "fecha_corte" : fecha_corte,
        "todos"       : not params.solo_ultimo_seguimiento,
        "estados_cita": {canon(e) or _norm_estado(e) for e in params.estados_gestion_cita},
    }

    objetivo = df if ctx["todos"] else df[df["_es_ultimo"]]
    hallazgos = []
    for row in objetivo.to_dict("records"):
        _validar_estado(row, ctx, hallazgos)
        _validar_fechas(row, ctx, hallazgos)
        _validar_aplicacion(row, ctx, hallazgos)
        _validar_cita(row, ctx, hallazgos)

    cols_h = [COL_CASO, COL_ID_INT, "Fecha llamada", COL_CREADO_POR, "Enfermero(a) de campo",
              "Validación", "Severidad", "Hallazgo", "Razón del estado", "DETALLE"]
    df_h = pd.DataFrame(hallazgos)
    if df_h.empty:
        df_h = pd.DataFrame(columns=cols_h)
    else:
        if "Enfermero(a) de campo" in df_pac.columns:
            df_h = df_h.merge(df_pac[[COL_CASO, "Enfermero(a) de campo"]], on=COL_CASO, how="left")
        orden_sev = {ALTA: 0, MEDIA: 1, BAJA: 2}
        df_h = (df_h.assign(_s=df_h["Severidad"].map(orden_sev))
                    .sort_values(["_s", "Validación", COL_CASO, "Fecha llamada"])
                    .drop(columns="_s"))
        df_h = df_h[[c for c in cols_h if c in df_h.columns]]

    return df_h.reset_index(drop=True), resumen_por_caso(df_h, df[df["_es_ultimo"]]), fecha_corte


def resumen_por_caso(df_h: pd.DataFrame, df_ultimos: pd.DataFrame) -> pd.DataFrame:
    """Una fila por caso con los hallazgos de cada validación (para la hoja consolidada)."""
    base = pd.DataFrame({COL_CASO: df_ultimos[COL_CASO].unique()})
    if df_h.empty:
        for v in VALIDACIONES:
            base[v] = "OK"
        base[COL_N_CRITERIOS] = 0
        base[COL_N_HALLAZGOS] = 0
        return base
    piv = (df_h.groupby([COL_CASO, "Validación"])["Hallazgo"]
               .agg(lambda s: " | ".join(dict.fromkeys(s)))
               .unstack())
    base = base.merge(piv, left_on=COL_CASO, right_index=True, how="left")
    for v in VALIDACIONES:
        base[v] = base[v].fillna("OK") if v in base.columns else "OK"
    # Un criterio puede generar varios hallazgos: se informan ambos conteos
    conteo = df_h.groupby(COL_CASO).agg(**{COL_N_CRITERIOS: ("Validación", "nunique"),
                                           COL_N_HALLAZGOS: ("Validación", "size")})
    base = base.merge(conteo, left_on=COL_CASO, right_index=True, how="left")
    for col in (COL_N_CRITERIOS, COL_N_HALLAZGOS):
        base[col] = base[col].fillna(0).astype(int)
    return base[[COL_CASO] + VALIDACIONES + [COL_N_CRITERIOS, COL_N_HALLAZGOS]]
