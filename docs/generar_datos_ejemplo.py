# -*- coding: utf-8 -*-
"""
Genera archivos de EJEMPLO (datos ficticios) con la misma estructura de los
reportes reales, para capacitación y para las capturas del manual.

    python docs/generar_datos_ejemplo.py  ->  docs/datos_ejemplo/*.xlsx
"""

import random
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

random.seed(7)
SALIDA = Path(__file__).parent / "datos_ejemplo"
SALIDA.mkdir(exist_ok=True)

ENFERMEROS = ["Ana Torres", "Luis Gómez", "Marta Ríos", "Carlos Peña"]
PROGRAMAS  = ["Osteoporosis Post Menopausica", "Hipercolesterolemia Familiar", "Evenity"]
EPS        = ["EPS Ejemplo A", "EPS Ejemplo B", "EPS Ejemplo C"]
CORTE      = datetime(2026, 7, 30)


def f(d):  # fecha de nota
    return d.strftime("%d/%m/%Y") if d else "01/01/1999"


def inicio(d):  # formato del reporte de Salesforce
    h = d.hour % 12 or 12
    return f"{d.day}/{d.month:02d}/{d.year}, {h}:{d.minute:02d} {'a.' if d.hour < 12 else 'p.'}\xa0m."


def nota(estado, detalle, prox, ult_apl=None, control=None):
    return "\n".join([
        "CANAL DE CONTACTO: LLAMADA", "PARTICIPANTE: PACIENTE", "TIPO DE GESTION: SEGUIMIENTO",
        f"ESTADO DESPUES DEL CONTACTO: {estado}", f"FECHA DE ULTIMA APLICACION: {f(ult_apl)}",
        f"FECHA DE CONTROL CON ESPECIALISTA: {f(control)}", f"FECHA DE PROXIMO CONTACTO: {f(prox)}",
        "REPORTE FV: NO", f"DETALLE: {detalle}", "OP LOGISTICO: OPERADOR EJEMPLO",
    ])


pacientes, interacciones, aplicaciones = [], [], []
n_int = n_apl = 0


def paciente(i, razon, estado="Activo"):
    caso = f"PSP-90{i:04d}"
    pacientes.append({
        "Número del caso": caso, "Zona": random.choice(["Centro", "Norte", "Sur"]),
        "Fecha de inscripción": datetime(2025, random.randint(1, 12), random.randint(1, 28)),
        "Regimen": random.choice(["Contributivo", "Subsidiado"]), "EPS": random.choice(EPS),
        "Logistic Operator": "Operador Ejemplo", "Institución": "IPS Ejemplo",
        "Estado": estado, "Razón del estado": razon,
        "Enfermero(a) de campo": ENFERMEROS[i % 4], "Programa": PROGRAMAS[i % 3],
    })
    return caso


def llamada(caso, i, fecha, texto, exitosa="Si"):
    global n_int
    n_int += 1
    interacciones.append({
        "Número del caso": caso, "Programa": PROGRAMAS[i % 3],
        "Fecha de ingreso al programa": datetime(2025, 6, 1),
        "Interacción: Interacciones ID": f"INT-{800000 + n_int}",
        "Interacción: Tipo de registro": "Enfermero(a) de campo",
        "Interacción: Creado por": ENFERMEROS[i % 4], "Interacción existosa?": exitosa,
        "Inicio": inicio(fecha), "Fin": inicio(fecha + timedelta(minutes=10)),
        "Recomendaciones": texto,
        "Interacción: Fecha de la última modificación": fecha, "Duración (Hrs)": 0.17,
    })


def aplicacion(caso, fecha, aplicado="Si", frecuencia="Mensual"):
    global n_apl
    n_apl += 1
    aplicaciones.append({
        "Número del caso": caso, "Nombre de Aprobación": f"AP-{500000 + n_apl}",
        "ID de administración de medicaciones": f"MA-{700000 + n_apl}",
        "¿Medicamento aplicado?": aplicado, "Dosis": "60 MG", "Frecuencia": frecuencia,
        "Fecha de entrega programada": fecha, "Fecha real de entrega": fecha,
        "Fecha de aplicación": fecha, "Lugar de la aplicación": "Institución Hospitalaria (IPS)",
    })


D = lambda m, d, h=9: datetime(2026, m, d, h, random.choice([0, 15, 30, 45]))

for i in range(1, 41):
    escenario = i % 14
    if escenario in (0, 1, 2):  # correcto
        c = paciente(i, "En Tratamiento")
        aplicacion(c, D(5, 20)); aplicacion(c, D(6, 20))
        llamada(c, i, D(6, 1), nota("EN TRATAMIENTO", "EN TRATAMIENTO / PACIENTE SIN NOVEDAD / TIPO CANAL CALL", D(6, 22), D(5, 20)))
        llamada(c, i, D(6, 22), nota("EN TRATAMIENTO", "EN TRATAMIENTO / PACIENTE CONFIRMA APLICACION DEL 20/06/2026 SIN NOVEDAD / TIPO CANAL CALL", D(8, 5), D(6, 20)))
    elif escenario == 3:  # V1 razón distinta a la nota
        c = paciente(i, "Falta Cita Médica")
        aplicacion(c, D(6, 10))
        llamada(c, i, D(7, 2), nota("EN TRATAMIENTO", "EN TRATAMIENTO / PACIENTE REFIERE QUE CONTINUA CON EL MEDICAMENTO / TIPO CANAL CALL", D(8, 2), D(6, 10)))
    elif escenario == 4:  # V1 exitosa pero no contactable
        c = paciente(i, "Paciente No Contactable")
        llamada(c, i, D(7, 10), nota("PACIENTE NO CONTACTABLE", "PACIENTE NO CONTACTABLE / NO CONTESTA LLAMADAS", D(8, 10)), exitosa="Si")
    elif escenario == 5:  # V1 campo vs detalle
        c = paciente(i, "Problemas administrativos (EPS)")
        llamada(c, i, D(7, 12), nota("EN TRATAMIENTO", "PROBLEMAS ADMINISTRATIVOS (EPS) / EPS NO HA AUTORIZADO EL MEDICAMENTO / TIPO CANAL CALL", D(8, 12)))
    elif escenario == 6:  # V2 llamada tarde
        c = paciente(i, "En Tratamiento")
        aplicacion(c, D(6, 5)); aplicacion(c, D(7, 5))
        llamada(c, i, D(6, 1), nota("EN TRATAMIENTO", "EN TRATAMIENTO / SIN NOVEDAD / TIPO CANAL CALL", D(6, 15)))
        llamada(c, i, D(7, 6), nota("EN TRATAMIENTO", "EN TRATAMIENTO / PACIENTE CONFIRMA APLICACION DEL 05/07/2026 / TIPO CANAL CALL", D(8, 6), D(7, 5)))
    elif escenario == 7:  # V2 sin próximo contacto + llamada antes
        c = paciente(i, "En Tratamiento")
        llamada(c, i, D(6, 3), nota("EN TRATAMIENTO", "EN TRATAMIENTO / SIN NOVEDAD", D(7, 20)))
        llamada(c, i, D(7, 8), nota("EN TRATAMIENTO", "EN TRATAMIENTO / SE BRINDA EDUCACION", None))
    elif escenario == 8:  # V2 próximo contacto anterior a la llamada
        c = paciente(i, "En Tratamiento")
        llamada(c, i, D(7, 14), nota("EN TRATAMIENTO", "EN TRATAMIENTO / SIN NOVEDAD / TIPO CANAL CALL", D(7, 1)))
    elif escenario == 9:  # V2 seguimiento vencido
        c = paciente(i, "Problemas administrativos (EPS)")
        llamada(c, i, D(6, 10), nota("PROBLEMAS ADMINISTRATIVOS (EPS)", "PROBLEMAS ADMINISTRATIVOS (EPS) / EN ESPERA DE AUTORIZACION", D(7, 5)))
    elif escenario == 10:  # V3 aplicación de la nota no registrada
        c = paciente(i, "En Tratamiento")
        aplicacion(c, D(5, 2))
        llamada(c, i, D(7, 3), nota("EN TRATAMIENTO", "EN TRATAMIENTO / PACIENTE REFIERE APLICACION SIN NOVEDAD", D(8, 3), D(6, 25)))
    elif escenario == 11:  # V3 fecha futura + confirma sin registro
        c = paciente(i, "En Tratamiento")
        llamada(c, i, D(6, 4), nota("EN TRATAMIENTO", "EN TRATAMIENTO / SIN NOVEDAD / TIPO CANAL CALL", D(7, 4)))
        llamada(c, i, D(7, 4), nota("EN TRATAMIENTO", "EN TRATAMIENTO / FAMILIAR CONFIRMA APLICACION EL 2 DE JULIO / TIPO CANAL CALL", D(8, 4), D(9, 15)))
    elif escenario == 12:  # V4 cita sin especialidad / agendada sin fecha
        c = paciente(i, "Falta Cita Médica")
        llamada(c, i, D(7, 9), nota("FALTA CITA MEDICA", "FALTA CITA MEDICA / PACIENTE INFORMA QUE TIENE CITA AGENDADA / TIPO CANAL CALL", D(8, 9)))
    elif escenario == 13:  # V4 no documenta estado de la cita / correcta
        if i % 2:
            c = paciente(i, "Pendiente de Reformulación")
            llamada(c, i, D(7, 15), nota("PENDIENTE DE REFORMULACION", "PENDIENTE DE REFORMULACION / ENDOCRINOLOGIA / SE BRINDA EDUCACION", D(8, 15)))
        else:
            c = paciente(i, "Falta Cita Médica")
            llamada(c, i, D(7, 15), nota("FALTA CITA MEDICA", "FALTA CITA MEDICA / REUMATOLOGIA / CITA AGENDADA PARA EL 12/08/2026 / TIPO CANAL CALL", D(8, 13), control=D(8, 12)))

# Pacientes sin seguimientos en el periodo (aparecen en el consolidado sin interacción)
for i in range(41, 46):
    paciente(i, random.choice(["Servicios PSP Finalizados", "Fallecimiento"]), "Inactivo")

pd.DataFrame(pacientes).to_excel(SALIDA / "Bd pacientes.xlsx", index=False)
pd.DataFrame(interacciones).to_excel(SALIDA / "Interacciones.xlsx", index=False)
pd.DataFrame(aplicaciones).to_excel(SALIDA / "Aplicaciones.xlsx", index=False)
print(f"{len(pacientes)} pacientes, {len(interacciones)} interacciones, {len(aplicaciones)} aplicaciones -> {SALIDA}")
