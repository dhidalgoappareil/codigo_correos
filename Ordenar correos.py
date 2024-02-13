import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import csv
import openpyxl

# Función para cargar datos desde un archivo CSV o Excel
def cargar_datos(desde_excel=False):
    try:
        if desde_excel:
            libro_excel = openpyxl.load_workbook('datos.xlsx')
            hoja_excel = libro_excel.active
            datos = [(fila[0].value, fila[1].value) for fila in hoja_excel.iter_rows(values_only=True)]
        else:
            with open('datos.csv', newline='') as archivo:
                lector_csv = csv.reader(archivo)
                datos = list(lector_csv)
        return datos
    except FileNotFoundError:
        return []

# Función para guardar datos en un archivo CSV o Excel
def guardar_datos(datos, a_excel=False):
    if a_excel:
        libro_excel = openpyxl.Workbook()
        hoja_excel = libro_excel.active
        for dato in datos:
            hoja_excel.append(dato)
        libro_excel.save('datos.xlsx')
    else:
        with open('datos.csv', 'w', newline='') as archivo:
            escritor_csv = csv.writer(archivo)
            escritor_csv.writerows(datos)

# Función para mostrar mensajes de confirmación
def mostrar_mensaje(titulo, mensaje):
    messagebox.showinfo(titulo, mensaje)

# Función para eliminar elemento seleccionado
def eliminar_elemento():
    seleccion = resultado.curselection()
    if seleccion:
        indice = seleccion[0]
        elemento_eliminar = datos_combinados[indice]
        confirmacion = messagebox.askokcancel("Eliminar", f"¿Estás seguro de eliminar: {elemento_eliminar[0]} - {elemento_eliminar[1]}?")
        if confirmacion:
            datos_combinados.pop(indice)
            guardar_datos(datos_combinados)
            mostrar_resultado(datos_combinados)
            mostrar_mensaje("Éxito", "Se eliminó el elemento.")
    else:
        mostrar_mensaje("Error", "Por favor, selecciona un elemento para eliminar.")

# Función para editar elemento seleccionado
def editar_elemento():
    seleccion = resultado.curselection()
    if seleccion:
        indice = seleccion[0]
        elemento_editar = datos_combinados[indice]

        # Crear ventana de edición
        ventana_edicion = tk.Toplevel(ventana)
        ventana_edicion.title("Editar Elemento")

        # Etiquetas y cuadros de entrada para edición
        nombre_label = ttk.Label(ventana_edicion, text="Nuevo Nombre:")
        correo_label = ttk.Label(ventana_edicion, text="Nueva Dirección de Correo:")
        nuevo_nombre_entry = ttk.Entry(ventana_edicion, width=30)
        nueva_correo_entry = ttk.Entry(ventana_edicion, width=30)

        # Configurar valores iniciales en los cuadros de entrada
        nuevo_nombre_entry.insert(0, elemento_editar[0])
        nueva_correo_entry.insert(0, elemento_editar[1])

        def aplicar_edicion():
            nuevo_nombre = nuevo_nombre_entry.get()
            nueva_direccion = nueva_correo_entry.get()

            if nuevo_nombre and nueva_direccion:
                datos_combinados[indice] = (nuevo_nombre, nueva_direccion)
                guardar_datos(datos_combinados)
                mostrar_resultado(datos_combinados)
                mostrar_mensaje("Éxito", "Se aplicaron los cambios.")
                ventana_edicion.destroy()
            else:
                mostrar_mensaje("Error", "Por favor, ingrese nuevo nombre y dirección de correo electrónico.")

        # Botón para aplicar la edición
        boton_aplicar_edicion = ttk.Button(ventana_edicion, text="Aplicar Edición", command=aplicar_edicion)

        # Colocar widgets en la ventana de edición
        nombre_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        correo_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        nuevo_nombre_entry.grid(row=0, column=1, padx=10, pady=5, sticky="w")
        nueva_correo_entry.grid(row=1, column=1, padx=10, pady=5, sticky="w")
        boton_aplicar_edicion.grid(row=2, column=0, columnspan=2, pady=10)
    else:
        mostrar_mensaje("Error", "Por favor, selecciona un elemento para editar.")

# Función para cargar datos desde un archivo Excel
def cargar_datos_desde_excel():
    global datos_combinados
    datos_combinados = cargar_datos(desde_excel=True)
    mostrar_resultado(datos_combinados)
    mostrar_mensaje("Éxito", "Datos importados desde Excel.")

# Función para mostrar resultados en el cuadro de texto
def mostrar_resultado(lista):
    resultado.config(state=tk.NORMAL)
    resultado.delete(1.0, tk.END)
    for nombre, direccion in lista:
        resultado.insert(tk.END, f"{nombre}: {direccion}\n")
    resultado.config(state=tk.DISABLED)

# Función para mostrar todos los datos iniciales
def mostrar_resultado_inicial():
    mostrar_resultado(datos_combinados)

# Función para ordenar los datos alfabéticamente por nombre
def ordenar():
    datos_combinados_ordenados = sorted(datos_combinados, key=lambda x: x[0])
    mostrar_resultado(datos_combinados_ordenados)

# Función para deshacer el orden y mostrar los datos originales
def deshacer_orden():
    mostrar_resultado(datos_combinados)

# Función para agregar un nuevo dato
def agregar_dato():
    nuevo_nombre = nombre_entry.get()
    nueva_direccion = correo_entry.get()

    if nuevo_nombre and nueva_direccion:
        datos_combinados.append((nuevo_nombre, nueva_direccion))
        guardar_datos(datos_combinados)
        mostrar_resultado(datos_combinados)
        mostrar_mensaje("Éxito", "Se agregó el nuevo dato.")
    else:
        mostrar_mensaje("Error", "Por favor, ingrese nuevo nombre y dirección de correo electrónico.")

# Crear la ventana principal
ventana = tk.Tk()
ventana.title("Gestión de Nombres y Correos Electrónicos")

# Estilo ttk
estilo = ttk.Style()
estilo.configure("TButton", padding=5, font=('Helvetica', 10))

# Botones
boton_mostrar_todo = ttk.Button(ventana, text="Mostrar Todo", command=mostrar_resultado_inicial)
boton_ordenar = ttk.Button(ventana, text="Ordenar", command=ordenar)
boton_deshacer = ttk.Button(ventana, text="Deshacer Orden", command=deshacer_orden)
boton_agregar = ttk.Button(ventana, text="Agregar Dato", command=agregar_dato)
boton_eliminar = ttk.Button(ventana, text="Eliminar Seleccionado", command=eliminar_elemento)
boton_editar = ttk.Button(ventana, text="Editar Seleccionado", command=editar_elemento)

# Etiquetas y cuadros de entrada
nombre_label = ttk.Label(ventana, text="Nombre:")
correo_label = ttk.Label(ventana, text="Correo Electrónico:")
nombre_entry = ttk.Entry(ventana)
correo_entry = ttk.Entry(ventana)

# Cuadro de texto para mostrar resultados
resultado = tk.Text(ventana, height=10, width=40)
resultado.config(state=tk.DISABLED)

# Scrollbar para el cuadro de texto
scrollbar = tk.Scrollbar(ventana, orient="vertical", command=resultado.yview)
resultado.config(yscrollcommand=scrollbar.set)

# Colocar widgets en la ventana
boton_mostrar_todo.grid(row=0, column=0, padx=10, pady=5, sticky="w")
boton_ordenar.grid(row=0, column=1, padx=10, pady=5)
boton_deshacer.grid(row=0, column=2, padx=10, pady=5)
boton_eliminar.grid(row=5, column=0, padx=10, pady=5, sticky="w")
boton_editar.grid(row=5, column=1, padx=10, pady=5)
resultado.grid(row=1, column=0, columnspan=3, padx=10, pady=10, sticky="w")
scrollbar.grid(row=1, column=3, pady=10, sticky="ns")
nombre_label.grid(row=2, column=0, padx=10, pady=5, sticky="w")
correo_label.grid(row=3, column=0, padx=10, pady=5, sticky="w")
nombre_entry.grid(row=2, column=1, padx=10, pady=5, sticky="w")
correo_entry.grid(row=3, column=1, padx=10, pady=5, sticky="w")
boton_agregar.grid(row=4, column=0, columnspan=2, pady=5)

# Botón para exportar datos a un archivo Excel
boton_exportar_excel = ttk.Button(ventana, text="Exportar a Excel", command=lambda: guardar_datos(datos_combinados, a_excel=True))
boton_exportar_excel.grid(row=6, column=0, padx=10, pady=5, sticky="w")

# Botón para importar datos desde un archivo Excel
boton_importar_excel = ttk.Button(ventana, text="Importar desde Excel", command=lambda: cargar_datos_desde_excel())
boton_importar_excel.grid(row=6, column=1, padx=10, pady=5)

# Mostrar la ventana
ventana.mainloop()
