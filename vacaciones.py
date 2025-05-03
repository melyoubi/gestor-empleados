import pandas as pd
from datetime import datetime

# Lista de trabajadores (nombre, fecha de ingreso)
trabajadores = [
    {"Nombre": "Juan Pérez", "FechaIngreso": "2019-03-15"},
    {"Nombre": "Laura Gómez", "FechaIngreso": "2021-07-01"},
    {"Nombre": "Carlos Ruiz", "FechaIngreso": "2023-01-10"},
]

# Fecha actual
hoy = datetime.today()

# Función para calcular años y vacaciones
def calcular_vacaciones(fecha_ingreso_str):
    fecha_ingreso = datetime.strptime(fecha_ingreso_str, "%Y-%m-%d")
    antiguedad = (hoy - fecha_ingreso).days / 365.25
    dias_vacaciones = antiguedad * 30
    return round(antiguedad, 2), round(dias_vacaciones)

# Añadir cálculos
for trabajador in trabajadores:
    años, vacaciones = calcular_vacaciones(trabajador["FechaIngreso"])
    trabajador["AñosTrabajados"] = años
    trabajador["DiasVacaciones"] = vacaciones

# Crear DataFrame y guardar en Excel
df = pd.DataFrame(trabajadores)
df.to_excel("VacacionesTrabajadores.xlsx", index=False)

print("✅ Archivo 'VacacionesTrabajadores.xlsx' generado con éxito.")
