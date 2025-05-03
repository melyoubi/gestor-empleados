import tkinter as tk
from tkinter import messagebox
from tkinter import ttk  
import pandas as pd
from datetime import datetime
import os
import re
from dateutil.relativedelta import relativedelta
# Constantes
NOMBRE_ARCHIVO = "EmpleadosActivos.xlsx"
HISTORIAL_ARCHIVO = "HistorialVacaciones.xlsx"
ARCHIVO_BAJAS = "BajasEmpleados.xlsx"

def busqueda_avanzada():
    ventana = tk.Toplevel(root)
    ventana.title("Búsqueda Avanzada")

    filtros = ["Nombre", "NIE/DNI", "Telefono", "Correo", "Genero", "FechaIngreso"]
    entradas = {}

    for i, campo in enumerate(filtros):
        tk.Label(ventana, text=f"{campo}:").grid(row=i, column=0, padx=10, pady=5, sticky="e")
        entry = tk.Entry(ventana, width=30)
        entry.grid(row=i, column=1, padx=10, pady=5)
        entradas[campo] = entry

    def buscar():
        datos = cargar_trabajadores_existentes()
        resultados = []

        for t in datos:
            coincide = True
            for campo, entry in entradas.items():
                valor = entry.get().strip().lower()
                if valor and valor not in str(t.get(campo, "")).lower():
                    coincide = False
                    break
            if coincide:
                resultados.append(t)

        mostrar_resultados(resultados)

    tk.Button(ventana, text="Buscar", command=buscar, bg="lightgreen").grid(row=len(filtros), column=0, columnspan=2, pady=10)

def mostrar_resultados(lista):
    ventana = tk.Toplevel(root)
    ventana.title("Resultados de Búsqueda")

    frame = tk.Frame(ventana)
    frame.pack(fill="both", expand=True)

    columnas = ("#", "Nombre", "NIE/DNI", "Telefono", "Correo", "Genero", "FechaIngreso")
    tree = ttk.Treeview(frame, columns=columnas, show="headings")
    tree.pack(side="left", fill="both", expand=True)

    for col in columnas:
        tree.heading(col, text=col)
        tree.column(col, anchor="center", width=110)

    for i, t in enumerate(lista, start=1):
        tree.insert("", "end", values=(
            i,
            t.get("Nombre", ""),
            t.get("NIE/DNI", ""),
            t.get("Telefono", ""),
            t.get("Correo", ""),
            t.get("Genero", ""),
            t.get("FechaIngreso", "")
        ))

    scroll_y = tk.Scrollbar(frame, orient="vertical", command=tree.yview)
    scroll_y.pack(side="right", fill="y")
    tree.configure(yscrollcommand=scroll_y.set)
def dar_de_baja_empleado(dni_baja):
    # Verificar si el archivo existe
    if not os.path.exists(NOMBRE_ARCHIVO):
        return "No existe el archivo de empleados."

    # Leer datos actuales
    df = pd.read_excel(NOMBRE_ARCHIVO)
    # Buscar al empleado
    empleado = df[df["NIE/DNI"] == dni_baja]

    if empleado.empty:
        return f"No se encontró ningún empleado con DNI: {dni_baja}"

    # Agregar la fecha de baja
    fecha_baja = datetime.today().strftime("%Y-%m-%d")
    empleado["FechaBaja"] = fecha_baja

    # Guardar en archivo de bajas
    if os.path.exists(ARCHIVO_BAJAS):
        df_bajas = pd.read_excel(ARCHIVO_BAJAS)
        df_bajas = pd.concat([df_bajas, empleado], ignore_index=True)
    else:
        df_bajas = empleado

    df_bajas.to_excel(ARCHIVO_BAJAS, index=False)

    # Eliminar del archivo principal
    df = df[df["NIE/DNI"] != dni_baja]
    df.to_excel(NOMBRE_ARCHIVO, index=False)

    return f"Empleado con DNI {dni_baja} dado de baja correctamente."

# Ejemplo de uso (puedes probar con un NIE/DNI real que tengas en el archivo)
# resultado = dar_de_baja_empleado("X1234567A")
# print(resultado)

trabajadores = []

# Funciones de archivos
def cargar_trabajadores_existentes():
    if os.path.exists(NOMBRE_ARCHIVO):
        df = pd.read_excel(NOMBRE_ARCHIVO, dtype={"Telefono": str})
        return df.to_dict(orient='records')
    return []

def guardar_excel():
    if trabajadores:
        df = pd.DataFrame(trabajadores)
        df.to_excel(NOMBRE_ARCHIVO, index=False)

def guardar_historial(nombre, fecha, dias):
    nuevo_registro = pd.DataFrame([{
        "Nombre": nombre,
        "Fecha": fecha,
        "DiasTomados": dias
    }])
    if os.path.exists(HISTORIAL_ARCHIVO):
        df = pd.read_excel(HISTORIAL_ARCHIVO)
        df = pd.concat([df, nuevo_registro], ignore_index=True)
    else:
        df = nuevo_registro
    df.to_excel(HISTORIAL_ARCHIVO, index=False)

# Funciones principales
def mostrar_trabajadores():
    datos = cargar_trabajadores_existentes()
    if not datos:
        messagebox.showinfo("Sin datos", "No hay trabajadores registrados.")
        return

    ventana = tk.Toplevel(root)
    ventana.title("Lista de Trabajadores")

    frame = tk.Frame(ventana)
    frame.pack(fill="both", expand=True)

    columnas = ("#", "Nombre", "FechaIngreso", "Antigüedad", "DiasTotales", "DiasTomados", "DiasRestantes")

    tree = ttk.Treeview(frame, columns=columnas, show="headings")
    tree.pack(side="left", fill="both", expand=True)

    # Definir encabezados
    for col in columnas:
        tree.heading(col, text=col)
        tree.column(col, width=120, anchor="center")

    # Scrollbars
    scrollbar_y = tk.Scrollbar(frame, orient="vertical", command=tree.yview)
    scrollbar_y.pack(side="right", fill="y")
    tree.configure(yscrollcommand=scrollbar_y.set)

    scrollbar_x = tk.Scrollbar(ventana, orient="horizontal", command=tree.xview)
    scrollbar_x.pack(side="bottom", fill="x")
    tree.configure(xscrollcommand=scrollbar_x.set)

    hoy = datetime.today()

    for idx, t in enumerate(datos, start=1):
        fecha_ingreso = datetime.strptime(t["FechaIngreso"], "%Y-%m-%d")
        antiguedad = relativedelta(hoy, fecha_ingreso)
        años = antiguedad.years
        meses = antiguedad.months
        dias = antiguedad.days
        texto_antiguedad = f"{años}a {meses}m {dias}d"

        dias_totales = round((hoy - fecha_ingreso).days * 30 / 365.25)
        dias_tomados = t.get("DiasTomados", 0)
        dias_restantes = max(0, dias_totales - dias_tomados)

        tree.insert("", "end", values=(
            idx,  # Número
            t["Nombre"],
            t["FechaIngreso"],
            texto_antiguedad,
            dias_totales,
            dias_tomados,
            dias_restantes
        ))

    # Mostrar la cantidad total de trabajadores abajo
    cantidad = len(datos)
    tk.Label(ventana, text=f"Total de Trabajadores: {cantidad}", font=("Arial", 12, "bold")).pack(pady=10)
def agregar_trabajador():
    ventana = tk.Toplevel(root)
    ventana.title("Agregar nuevo trabajador")

    labels = ["Nombre", "Fecha ingreso (YYYY-MM-DD)", "NIE/DNI", "Teléfono", "Correo electrónico", "Fecha nacimiento (YYYY-MM-DD)"]
    entries = []
    for i, text in enumerate(labels):
        tk.Label(ventana, text=text+":").grid(row=i, column=0, padx=10, pady=5, sticky="e")
        entry = tk.Entry(ventana, width=30)
        entry.grid(row=i, column=1, padx=10, pady=5)
        entries.append(entry)
    entry_nombre, entry_fecha_ingreso, entry_dni, entry_telefono, entry_correo, entry_fecha_nacimiento = entries

    # Género
    genero_var = tk.StringVar(value="Hombre")
    tk.Label(ventana, text="Género:").grid(row=len(labels), column=0, padx=10, pady=5, sticky="e")
    tk.Radiobutton(ventana, text="Hombre", variable=genero_var, value="Hombre").grid(row=len(labels), column=1, padx=10, pady=5, sticky="w")
    tk.Radiobutton(ventana, text="Mujer", variable=genero_var, value="Mujer").grid(row=len(labels)+1, column=1, padx=10, pady=5, sticky="w")

    def guardar():
        nombre = entry_nombre.get().strip()
        fecha_ingreso = entry_fecha_ingreso.get().strip()
        dni = entry_dni.get().strip()
        telefono = entry_telefono.get().strip()
        correo = entry_correo.get().strip()
        genero = genero_var.get()
        fecha_nacimiento = entry_fecha_nacimiento.get().strip()

        if not all([nombre, fecha_ingreso, dni, telefono, correo, genero, fecha_nacimiento]):
            messagebox.showwarning("Campos vacíos", "Completa todos los campos.")
            return
        try:
            datetime.strptime(fecha_ingreso, "%Y-%m-%d")
            datetime.strptime(fecha_nacimiento, "%Y-%m-%d")
        except ValueError:
            messagebox.showerror("Error", "Formato de fecha incorrecto. Usa YYYY-MM-DD.")
            return
        if not telefono.isdigit() or len(telefono) != 9:
            messagebox.showerror("Error", "Número de teléfono inválido (9 dígitos).")
            return
        if not re.match(r"[^@]+@[^@]+\.[^@]+", correo):
            messagebox.showerror("Error", "Correo electrónico inválido.")
            return
        existentes = cargar_trabajadores_existentes() + trabajadores
        if any(str(r.get("NIE/DNI", "")) == dni for r in existentes):
            messagebox.showerror("Duplicado", f"El NIE/DNI '{dni}' ya existe.")
            return

        trabajadores.append({
            "Nombre": nombre,
            "FechaIngreso": fecha_ingreso,
            "NIE/DNI": dni,
            "Telefono": str(telefono),
            "Correo": correo,
            "Genero": genero,
            "FechaNacimiento": fecha_nacimiento,
            "DiasTomados": 0,
            "FechaRegistro": datetime.today().strftime("%Y-%m-%d")
        })
        guardar_excel()
        ventana.destroy()
        messagebox.showinfo("Éxito", f"Empleado {nombre} agregado correctamente.")

    tk.Button(ventana, text="Guardar", command=guardar, bg="lightgreen").grid(row=len(labels)+2, column=0, columnspan=2, pady=10)

def editar_trabajador():
    ventana = tk.Toplevel(root)
    ventana.title("Editar Trabajador")

    tk.Label(ventana, text="NIE/DNI del trabajador:").grid(row=0, column=0, padx=10, pady=5)
    entry_dni = tk.Entry(ventana, width=30)
    entry_dni.grid(row=0, column=1, padx=10)

    def buscar_y_editar():
        dni = entry_dni.get().strip()
        if not dni:
            messagebox.showwarning("Campo vacío", "Ingresa el NIE/DNI.")
            return
        datos = cargar_trabajadores_existentes()
        for t in datos:
            if t.get("NIE/DNI") == dni:
                ventana_editar = tk.Toplevel(ventana)
                ventana_editar.title(f"Editar {t['Nombre']}")

                labels = ["Nombre", "FechaIngreso", "Telefono", "Correo", "FechaNacimiento"]
                entradas = {}

                for idx, campo in enumerate(labels):
                    tk.Label(ventana_editar, text=campo + ":").grid(row=idx, column=0, padx=10, pady=5, sticky="e")
                    entrada = tk.Entry(ventana_editar, width=30)
                    entrada.insert(0, t.get(campo, ""))
                    entrada.grid(row=idx, column=1, padx=10, pady=5)
                    entradas[campo] = entrada

                def guardar_cambios():
                    for campo, entrada in entradas.items():
                        nuevo_valor = entrada.get().strip()
                        if nuevo_valor:
                            t[campo] = nuevo_valor
                    pd.DataFrame(datos).to_excel(NOMBRE_ARCHIVO, index=False)
                    messagebox.showinfo("Éxito", f"Información de {t['Nombre']} actualizada correctamente.")
                    ventana_editar.destroy()
                    ventana.destroy()

                tk.Button(ventana_editar, text="Guardar Cambios", command=guardar_cambios, bg="lightgreen").grid(row=len(labels), column=0, columnspan=2, pady=10)
                return
        messagebox.showerror("No encontrado", "No se encontró el trabajador con ese NIE/DNI.")

    tk.Button(ventana, text="Buscar y Editar", command=buscar_y_editar, bg="lightblue").grid(row=1, column=0, columnspan=2, pady=10)
def dar_de_baja_empleado():
    ventana = tk.Toplevel(root)
    ventana.title("Dar de Baja a un Empleado")

    tk.Label(ventana, text="NIE/DNI del empleado a dar de baja:").grid(row=0, column=0, padx=10, pady=5)
    entry_dni = tk.Entry(ventana, width=30)
    entry_dni.grid(row=0, column=1, padx=10, pady=5)

    def procesar_baja():
        dni_baja = entry_dni.get().strip()
        if not dni_baja:
            messagebox.showwarning("Campo vacío", "Introduce el NIE/DNI del empleado.")
            return

        if not os.path.exists(NOMBRE_ARCHIVO):
            messagebox.showerror("Error", "No se encontró el archivo de empleados.")
            return

        df = pd.read_excel(NOMBRE_ARCHIVO)
        empleado = df[df["NIE/DNI"] == dni_baja]

        if empleado.empty:
            messagebox.showerror("No encontrado", f"No se encontró ningún empleado con NIE/DNI {dni_baja}.")
            return

        fecha_baja = datetime.today().strftime("%Y-%m-%d")
        empleado["FechaBaja"] = fecha_baja

        # Guardar en Bajas
        if os.path.exists("BajasEmpleados.xlsx"):
            df_bajas = pd.read_excel("BajasEmpleados.xlsx")
            df_bajas = pd.concat([df_bajas, empleado], ignore_index=True)
        else:
            df_bajas = empleado

        df_bajas.to_excel("BajasEmpleados.xlsx", index=False)

        # Eliminar del archivo original
        df = df[df["NIE/DNI"] != dni_baja]
        df.to_excel(NOMBRE_ARCHIVO, index=False)

        messagebox.showinfo("Éxito", f"El empleado con NIE/DNI {dni_baja} ha sido dado de baja.")
        ventana.destroy()

    tk.Button(ventana, text="Dar de Baja", command=procesar_baja, bg="tomato").grid(row=1, column=0, columnspan=2, pady=10)

def actualizar_dias_tomados():
    ventana = tk.Toplevel(root)
    ventana.title("Actualizar días tomados")

    tk.Label(ventana, text="NIE/DNI del trabajador:").grid(row=0, column=0, padx=10, pady=5)
    entry_dni = tk.Entry(ventana, width=30)
    entry_dni.grid(row=0, column=1, padx=10)

    tk.Label(ventana, text="Añadir días tomados:").grid(row=1, column=0, padx=10, pady=5)
    entry_dias = tk.Entry(ventana, width=30)
    entry_dias.grid(row=1, column=1, padx=10)

    def actualizar():
        dni = entry_dni.get().strip()
        dias = entry_dias.get().strip()
        if not dni or not dias:
            messagebox.showwarning("Campos vacíos", "Completa ambos campos.")
            return
        try:
            dias = int(dias)
            if dias < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Días debe ser un número entero positivo.")
            return

        datos = cargar_trabajadores_existentes()
        actualizado = False
        for t in datos:
            if t["NIE/DNI"] == dni:
                t["DiasTomados"] = t.get("DiasTomados", 0) + dias
                guardar_historial(t["Nombre"], datetime.today().strftime("%Y-%m-%d"), dias)
                actualizado = True

        if actualizado:
            pd.DataFrame(datos).to_excel(NOMBRE_ARCHIVO, index=False)
            messagebox.showinfo("Éxito", f"Días actualizados correctamente.")
            ventana.destroy()
        else:
            messagebox.showerror("No encontrado", "No se encontró el trabajador con ese NIE/DNI.")

    tk.Button(ventana, text="Actualizar", command=actualizar, bg="lightblue").grid(row=2, column=0, columnspan=2, pady=10)
def consultar_empleado():
    actualizar_dias_automaticamente()  # 🔄 Aquí se actualiza todo antes de abrir
    ventana = tk.Toplevel(root)
    ventana.title("Consultar Empleado")
    tk.Label(ventana, text="NIE/DNI del trabajador:").grid(row=0, column=0, padx=10, pady=5)
    entry_dni = tk.Entry(ventana, width=30)
    entry_dni.grid(row=0, column=1, padx=10)

    def consultar():
        dni = entry_dni.get().strip()
        if not dni:
            messagebox.showwarning("Campo vacío", "Ingresa el NIE/DNI.")
            return
        datos = cargar_trabajadores_existentes()
        for t in datos:
            if t.get("NIE/DNI") == dni:
                fi = datetime.strptime(t["FechaIngreso"], "%Y-%m-%d")
                hoy = datetime.today()
                antiguedad = relativedelta(hoy, fi)
                dias_totales = round((hoy - fi).days * 30 / 365.25)
                dias_tomados = t.get("DiasTomados", 0)
                dias_restantes = max(0, dias_totales - dias_tomados)

                info = (
                    f"Nombre: {t['Nombre']}\n"
                    f"NIE/DNI: {t['NIE/DNI']}\n"
                    f"Teléfono: {t['Telefono']}\n"
                    f"Correo: {t['Correo']}\n"
                    f"Fecha de Ingreso: {t['FechaIngreso']}\n"
                    f"Fecha de Nacimiento: {t['FechaNacimiento']}\n"
                    f"Género: {t['Genero']}\n"
                    f"Antigüedad: {antiguedad.years} años, {antiguedad.months} meses, {antiguedad.days} días\n"
                    f"Vacaciones Totales: {dias_totales} días\n"
                    f"Días Tomados: {dias_tomados} días\n"
                    f"Días Restantes: {dias_restantes} días"
                )
                
                ventana_info = tk.Toplevel(ventana)
                ventana_info.title(f"Ficha de {t['Nombre']}")
                text = tk.Text(ventana_info, width=60, height=20)
                text.pack(padx=10, pady=10)
                text.insert("1.0", info)
                text.config(state="disabled")
                return

        messagebox.showerror("No encontrado", "No se encontró el trabajador con ese NIE/DNI.")

    tk.Button(ventana, text="Consultar", command=consultar, bg="lightblue").grid(row=1, column=0, columnspan=2, pady=10)
def ver_historial():
    actualizar_dias_automaticamente()
    if not os.path.exists(HISTORIAL_ARCHIVO):
        messagebox.showinfo("Historial", "No hay registros de historial todavía.")
        return
    df = pd.read_excel(HISTORIAL_ARCHIVO)
    ventana = tk.Toplevel(root)
    ventana.title("Historial de Vacaciones")
    text = tk.Text(ventana, wrap="none", width=100, height=20)
    text.pack(padx=10, pady=10)
    text.insert("1.0", df.to_string(index=False))
    text.config(state="disabled")

# Inicio de sesión
def pedir_contraseña():
    ventana_login = tk.Tk()
    ventana_login.title("Iniciar Sesión")
    tk.Label(ventana_login, text="Introduce la contraseña:").pack(padx=20, pady=10)
    entry_contraseña = tk.Entry(ventana_login, show="*", width=30)
    entry_contraseña.pack(pady=5)

    def verificar_contraseña():
        if entry_contraseña.get() == "1234":
            ventana_login.destroy()
            abrir_ventana_principal()
        else:
            messagebox.showerror("Error", "Contraseña incorrecta.")

    tk.Button(ventana_login, text="Entrar", command=verificar_contraseña, bg="lightblue").pack(pady=10)
    ventana_login.mainloop()
def actualizar_dias_automaticamente():
    if not os.path.exists(NOMBRE_ARCHIVO):
        return

    df = pd.read_excel(NOMBRE_ARCHIVO)
    hoy = datetime.today()
    nuevos_restantes = []
    for idx, row in df.iterrows():
        try:
            ingreso = datetime.strptime(str(row["FechaIngreso"]), "%Y-%m-%d")
            dias_totales = round((hoy - ingreso).days * 30 / 365.25)
        except:
            dias_totales = 0

        dias_tomados = int(row.get("DiasTomados", 0))
        dias_restantes = max(0, dias_totales - dias_tomados)
        df.at[idx, "DiasRestantes"] = dias_restantes

    df.to_excel(NOMBRE_ARCHIVO, index=False)
def abrir_ventana_principal():
    actualizar_dias_automaticamente()  # 🔄 Aquí se actualiza todo antes de abrir
    global root
    root = tk.Tk()
    root.title("Gestor de Empleados")
    root.geometry("400x550")
    tk.Label(root, text="Gestor de Empleados", font=("Arial", 18, "bold")).pack(pady=20)

    tk.Button(root, text="📋 Mostrar Trabajadores", command=mostrar_trabajadores, width=25, height=2).pack(pady=5)
    tk.Button(root, text="➕ Agregar Trabajador", command=agregar_trabajador, width=25, height=2, bg="lightgreen").pack(pady=5)
    tk.Button(root, text="✏️ Editar Trabajador", command=editar_trabajador, width=25, height=2, bg="lightcyan").pack(pady=5)
    tk.Button(root, text="🔄 Actualizar Días Tomados", command=actualizar_dias_tomados, width=25, height=2, bg="lightblue").pack(pady=5)
    tk.Button(root, text="🔎 Consultar Empleado", command=consultar_empleado, width=25, height=2, bg="lightyellow").pack(pady=5)
    tk.Button(root, text="📄 Ver Historial", command=ver_historial, width=25, height=2, bg="lightgrey").pack(pady=5)
    tk.Button(root, text="🗑️ Dar de Baja", command=dar_de_baja_empleado, width=25, height=2, bg="tomato").pack(pady=5)
    tk.Button(root, text="🔎 Búsqueda Avanzada", command=busqueda_avanzada, width=25, height=2, bg="lightblue").pack(pady=5)

    root.mainloop()

# Iniciar app
pedir_contraseña()