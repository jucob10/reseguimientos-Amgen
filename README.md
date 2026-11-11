# Reseguimientos Amgen

Consolida BD pacientes + Interacciones + Aplicaciones y detecta errores en los seguimientos.

## Validaciones

| # | Validación | Qué revisa |
|---|---|---|
| 1 | Nota vs Razón del estado | El estado de la nota (campo *ESTADO DESPUÉS DE CONTACTO* o inicio del *DETALLE*) coincide con la *Razón del estado*; campo y detalle no se contradicen; interacción exitosa ≠ *No contactable*. |
| 2 | Fecha llamada vs programada | La llamada se hizo en la *fecha de próximo contacto* de la nota anterior (± tolerancia); la nota deja próximo contacto y no es anterior a la llamada; pacientes activos sin próximo contacto vencido. |
| 3 | Aplicación registrada | La *fecha de última aplicación* de la nota, o la aplicación que el detalle confirma, existe en Aplicaciones como aplicada. |
| 4 | Información de cita | Pacientes en gestión de cita: la nota indica especialidad/médico y la fecha de la cita (o el motivo por el que no está agendada). |

Cada hallazgo tiene severidad **Alta / Media / Baja**. Las reglas están en [scripts/validaciones.py](scripts/validaciones.py).

## Manual de usuario

[Manual_de_usuario_Reseguimientos.pdf](Manual_de_usuario_Reseguimientos.pdf). La fuente está en `docs/manual/manual.html`;
para regenerar el PDF: `python docs/manual/generar_pdf.py` (requiere `pip install playwright`).
Los datos ficticios de las capturas están en `docs/datos_ejemplo/` (se regeneran con
`python docs/generar_datos_ejemplo.py`) y sirven para capacitar sin usar datos reales.

## Uso

```
pip install -r scripts/requirements.txt
```

- **Página web**: `ejecutar_web.bat` (o `streamlit run app.py`) → http://localhost:8501.
  Se suben los 3 archivos, se ajustan los parámetros en la barra lateral y se descarga el Excel.
- **Línea de comandos**: `ejecutar.bat` lee los archivos de `C:\Users\julian.cobos\Downloads\reseguimientos\`
  y escribe `Reseguimientos.xlsx` en la misma carpeta.

El Excel trae tres hojas: **Resumen** (por validación y por persona), **Reseguimientos**
(una fila por paciente + columnas de validación) y **Hallazgos** (un registro por error).

## Compartir la página

Los archivos contienen datos de pacientes: publicarla solo en la red interna o en un servidor
de la empresa, nunca en un servicio público (p. ej. Streamlit Community Cloud).

- **Red local**: `ejecutar_web.bat` ya escucha en `0.0.0.0:8501`; los demás entran a
  `http://NOMBRE-DEL-EQUIPO:8501` (puede requerir abrir el puerto 8501 en el firewall de Windows).
- **Servidor**: instalar requirements y ejecutar `streamlit run app.py --server.address 0.0.0.0`
  detrás del proxy/autenticación corporativa.
