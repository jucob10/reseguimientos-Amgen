# -*- coding: utf-8 -*-
"""
==============================================================
  RESEGUIMIENTOS AMGEN — Página web
==============================================================
Permite subir BD pacientes, Interacciones y Aplicaciones, ejecuta la
consolidación + validación de seguimientos y descarga el Excel.

Ejecutar:  streamlit run app.py      (o ejecutar_web.bat)
"""

import sys
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).parent / "scripts"))
import consolidar_data as cd  # noqa: E402
import validaciones as val    # noqa: E402

st.set_page_config(page_title="Reseguimientos Amgen", page_icon="🩺", layout="wide")

st.markdown("""
<style>
  .block-container { padding-top: 2rem; }
  div[data-testid="stMetricValue"] { font-size: 1.8rem; }
  .nota { color: #9aa7b8; font-size: .9rem; }
</style>
""", unsafe_allow_html=True)

MANUAL = Path(__file__).parent / "Manual_de_usuario_Reseguimientos.pdf"

COLOR_SEVERIDAD = {val.ALTA: "#f8cbad", val.MEDIA: "#ffe699", val.BAJA: "#ddebf7"}  # celdas
COLOR_GRAFICO   = {val.ALTA: "#d9603b", val.MEDIA: "#e3b23c", val.BAJA: "#6b9bd1"}  # barras

# ─────────────────────────────────────────────────────────────────────────────
#  PARÁMETROS (barra lateral)
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Parámetros")

    alcance = st.radio(
        "Seguimientos a validar",
        ["Último seguimiento de cada paciente", "Todos los seguimientos del archivo"],
        help="El último seguimiento se compara contra la Razón del estado actual del paciente. "
             "Con 'Todos' se auditan además las notas anteriores (más hallazgos, tarda más).",
    )

    corte_auto = st.checkbox("Fecha de corte automática", value=True,
                             help="Usa la fecha de la última llamada que aparece en Interacciones.")
    fecha_corte = None if corte_auto else st.date_input("Fecha de corte", value=date.today(),
                                                        format="DD/MM/YYYY")

    st.subheader("Tolerancias (días)")
    tol_llamada = st.number_input("Llamada vs fecha programada", 0, 30, 3,
                                  help="Días de diferencia aceptados entre el próximo contacto "
                                       "programado y la llamada real.")
    tol_apl = st.number_input("Fecha de aplicación nota vs registro", 0, 15, 3,
                              help="Días de diferencia aceptados entre la fecha de aplicación "
                                   "escrita en la nota y la registrada en Aplicaciones.")
    ventana_apl = st.number_input("Ventana para aplicación sin fecha", 7, 120, 35,
                                  help="Si la nota confirma una aplicación sin fecha, se busca una "
                                       "aplicación registrada en estos días previos a la llamada.")

    st.subheader("Gestión de cita")
    estados_cita = st.multiselect(
        "Razones del estado que se consideran 'en gestión de cita'",
        options=sorted({*val.ESTADOS_GESTION_CITA_DEFECTO, "Pendiente definición de Conducta Médica",
                        "Decisión Médica - Exámenes", "Problemas administrativos (EPS)",
                        "Aprobación en estudio"}),
        default=list(val.ESTADOS_GESTION_CITA_DEFECTO),
    )

    st.divider()
    st.subheader("Ayuda")
    if MANUAL.exists():
        st.download_button("📘 Descargar manual de usuario", MANUAL.read_bytes(),
                           file_name=MANUAL.name, mime="application/pdf", width="stretch")

params = val.ParametrosValidacion(
    fecha_corte=pd.Timestamp(fecha_corte) if fecha_corte else None,
    tolerancia_llamada_dias=int(tol_llamada),
    tolerancia_aplicacion_dias=int(tol_apl),
    ventana_aplicacion_dias=int(ventana_apl),
    solo_ultimo_seguimiento=alcance.startswith("Último"),
    estados_gestion_cita=tuple(estados_cita),
)

# ─────────────────────────────────────────────────────────────────────────────
#  CARGA DE ARCHIVOS
# ─────────────────────────────────────────────────────────────────────────────
st.title("Reseguimientos Amgen")
st.markdown(
    '<p class="nota">Sube los reportes, valida la calidad de los seguimientos y descarga el '
    'consolidado con los hallazgos. Los archivos se procesan en memoria y no se guardan.</p>',
    unsafe_allow_html=True,
)

with st.expander("¿Qué se valida?", expanded=False):
    st.markdown("""
1. **Nota vs Razón del estado** — el estado escrito en la nota (campo *ESTADO DESPUÉS DE CONTACTO*
   o inicio del *DETALLE*) coincide con la *Razón del estado* del paciente; el campo y el detalle no
   se contradicen; una interacción exitosa no puede quedar como *No contactable*.
2. **Fecha de la llamada vs fecha programada** — la llamada se hizo en la *fecha de próximo
   contacto* que dejó la nota anterior (± tolerancia); la nota deja próximo contacto y no es anterior
   a la llamada; en pacientes activos el próximo contacto no está vencido a la fecha de corte.
3. **Aplicación registrada** — si la nota indica una fecha de última aplicación o confirma que se
   aplicó el medicamento, esa aplicación existe en *Aplicaciones* como aplicada.
4. **Información de cita** — si el paciente está en gestión de cita, la nota documenta la
   especialidad/médico y la fecha de la cita (o el motivo por el que aún no está agendada).
""")

c1, c2, c3 = st.columns(3)
tipos = ["xlsx", "xls", "csv"]
with c1:
    f_pac = st.file_uploader("1 · BD pacientes", type=tipos)
with c2:
    f_int = st.file_uploader("2 · Interacciones", type=tipos)
with c3:
    f_apl = st.file_uploader("3 · Aplicaciones", type=tipos,
                             help="Opcional, pero sin este archivo no se ejecuta la validación 3 "
                                  "ni el cálculo de próxima aplicación.")

def refrescar_pantalla():
    """
    En la versión web (stlite, Python dentro del navegador) la pantalla no se
    actualiza mientras Python está ocupado. Esta pausa corta le cede el control
    al navegador para que se vea la barra de avance. Localmente no hace nada.
    """
    if sys.platform != "emscripten":
        return
    try:
        import asyncio
        from pyodide.ffi import can_run_sync, run_sync
        if can_run_sync():
            run_sync(asyncio.sleep(0.1))
    except Exception:  # noqa: BLE001 — si no se puede, solo no se ve el avance
        pass


def avance(barra, fraccion, mensaje):
    barra.progress(fraccion, text=mensaje)
    refrescar_pantalla()


listo = f_pac is not None and f_int is not None
if st.button("Procesar", type="primary", disabled=not listo, width="stretch"):
    barra = st.progress(0.0, text="Iniciando...")
    refrescar_pantalla()
    try:
        reporte = cd.generar_reporte(
            f_pac, f_int, f_apl, params,
            progreso=lambda f, m: avance(barra, min(f * 0.8, 0.8), m),
        )
        avance(barra, 0.85, "Generando Excel...")
        reporte["excel"] = cd.exportar_excel(reporte)
        barra.progress(1.0, text="Listo")
        st.session_state["reporte"] = reporte
    except ValueError as e:
        barra.empty()
        st.error(f"No se pudo procesar: {e}. Revisa que subiste cada archivo en su casilla.")
        st.stop()
    except Exception as e:  # noqa: BLE001 — mostrar cualquier fallo al usuario
        barra.empty()
        st.exception(e)
        st.stop()
elif not listo:
    st.info("Sube al menos **BD pacientes** e **Interacciones** para comenzar.")

# ─────────────────────────────────────────────────────────────────────────────
#  RESULTADOS
# ─────────────────────────────────────────────────────────────────────────────
reporte = st.session_state.get("reporte")
if not reporte:
    st.stop()

t = reporte["totales"]
df_h = reporte["hallazgos"]

st.divider()
cab1, cab2 = st.columns([3, 1])
with cab1:
    st.subheader("Resultados")
    st.caption(f"Fecha de corte: {reporte['fecha_corte']:%d/%m/%Y} · "
               f"{t['seguimientos']:,} seguimientos leídos · {t['pacientes']:,} pacientes".replace(",", "."))
with cab2:
    st.download_button(
        "⬇ Descargar Excel", data=reporte["excel"],
        file_name=f"Reseguimientos_{reporte['fecha_corte']:%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary", width="stretch",
    )

if not t["con_aplicaciones"]:
    st.warning("No se cargó Aplicaciones: la validación 3 no se ejecutó.")

m1, m2, m3, m4 = st.columns(4)
pct = t["casos_con_hallazgos"] / t["casos_validados"] * 100 if t["casos_validados"] else 0
m1.metric("Casos validados", f"{t['casos_validados']:,}".replace(",", "."))
m2.metric("Casos con hallazgos", f"{t['casos_con_hallazgos']:,}".replace(",", "."), f"{pct:.1f}%",
          delta_color="off")
m3.metric("Hallazgos", f"{t['hallazgos']:,}".replace(",", "."))
m4.metric("Severidad alta", f"{int((df_h['Severidad'] == val.ALTA).sum()):,}".replace(",", "."))

# key + on_change="rerun": la pestaña elegida se conserva al usar los filtros
tab_res, tab_hall, tab_pers, tab_cons = st.tabs(
    ["Resumen", "Hallazgos", "Por persona", "Consolidado"], key="pestana", on_change="rerun")

with tab_res:
    st.dataframe(reporte["resumen"], hide_index=True, width="stretch")
    if not df_h.empty:
        conteo = (df_h.groupby(["Validación", "Severidad"]).size().unstack(fill_value=0)
                       .reindex(columns=[val.ALTA, val.MEDIA, val.BAJA], fill_value=0))
        st.bar_chart(conteo, color=[COLOR_GRAFICO[s] for s in conteo.columns],
                     horizontal=True, height=260)
        st.markdown("**Hallazgos más frecuentes**")
        tipos_h = (df_h.assign(Tipo=df_h["Hallazgo"]
                               .str.replace(r"\d{1,2}/\d{2}/\d{4}", "dd/mm/aaaa", regex=True)
                               .str.replace(r"\d+ días", "N días", regex=True)
                               .str.replace(r"'[^']*'", "'…'", regex=True))
                       .groupby(["Validación", "Severidad", "Tipo"]).size()
                       .rename("Cantidad").reset_index()
                       .sort_values("Cantidad", ascending=False))
        st.dataframe(tipos_h, hide_index=True, width="stretch")

with tab_hall:
    if df_h.empty:
        st.success("No se encontraron hallazgos.")
    else:
        f1, f2, f3, f4 = st.columns([2, 1, 2, 1])
        sel_val = f1.multiselect("Validación", val.VALIDACIONES, default=val.VALIDACIONES)
        sel_sev = f2.multiselect("Severidad", [val.ALTA, val.MEDIA, val.BAJA],
                                 default=[val.ALTA, val.MEDIA, val.BAJA])
        personas = sorted(df_h[val.COL_CREADO_POR].dropna().unique()) if val.COL_CREADO_POR in df_h else []
        sel_per = f3.multiselect("Registrado por", personas)
        buscar = f4.text_input("Buscar caso / texto")

        vista = df_h[df_h["Validación"].isin(sel_val) & df_h["Severidad"].isin(sel_sev)]
        if sel_per:
            vista = vista[vista[val.COL_CREADO_POR].isin(sel_per)]
        if buscar:
            patron = buscar.strip()
            vista = vista[vista[cd.COL_CASO].str.contains(patron, case=False, regex=False)
                          | vista["Hallazgo"].str.contains(patron, case=False, regex=False)
                          | vista["DETALLE"].fillna("").str.contains(patron, case=False, regex=False)]

        st.caption(f"{len(vista):,} hallazgos".replace(",", "."))
        st.dataframe(
            vista.style.map(lambda s: f"background-color: {COLOR_SEVERIDAD.get(s, '')}; color: #1B2430",
                            subset=["Severidad"]) if len(vista) <= 20000 else vista,
            hide_index=True, width="stretch", height=520,
            column_config={
                "Fecha llamada": st.column_config.DatetimeColumn(format="DD/MM/YYYY HH:mm"),
                "Hallazgo": st.column_config.TextColumn(width="large"),
                "DETALLE": st.column_config.TextColumn(width="large"),
            },
        )
        st.download_button("⬇ Descargar hallazgos filtrados (CSV)",
                           vista.to_csv(index=False).encode("utf-8-sig"),
                           file_name="hallazgos_filtrados.csv", mime="text/csv")

with tab_pers:
    por_persona = cd.hallazgos_por_persona(df_h)
    if por_persona.empty:
        st.info("Sin hallazgos para agrupar.")
    else:
        st.caption("Hallazgos por la persona que registró la nota (Interacción: Creado por).")
        st.dataframe(por_persona, hide_index=True, width="stretch", height=520)

with tab_cons:
    df_c = reporte["consolidado"]
    solo_err = st.checkbox("Mostrar solo pacientes con hallazgos")
    if solo_err:
        df_c = df_c[df_c[val.COL_N_HALLAZGOS].fillna(0) > 0]
    st.caption(f"{len(df_c):,} pacientes".replace(",", "."))
    st.dataframe(df_c, hide_index=True, width="stretch", height=520)
