import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook, load_workbook
import re
import os

#Comprobar si el archivo ya existe
nombre_archivo = 'datos.xlsx'
if os.path.exists(nombre_archivo):
    wb = load_worbook(nombre_archivo)
    wb = wb.active
else:
#crear el libro de excel
    wb = Workbook()
    ws = wb.active
    ws.append(["Nombre","Edad","Email","Telefono","Direccion"])

#Funcion que guardara los datos en el archivo xlsx
def guardar_datos():
    nombre = entry_nombre.get()
    edad = entry_edad.get()
    email = entry_email.get()
    telefono = entry_telefono.get()
    direccion = entry_direccion.get()
    
    #Validar que todos los campos esten llenos
    if not nombre or not edad or not email or not telefono or not direccion:
        messagebox.showwarning("Advertencia", "Todos los campos son obligatorios")
        return
    
    #Validar que edad y telefono sean numeros    
    try:    
        edad = int(edad)
        telefono = int(telefono)
    except ValueError:
        messagebox.showwarning("Advertencia", "Edad y Telefono deben ser numeros")
        return
    
    #Validar que el formato del email sea correcto
    if not re.match(r"[^@]+@[^@]+\.[^@]+", email):  
        messagebox.showwarning("Advertencia", "El correo electronico no es válido")
        return
    
    #Guardar los datos en el archivo excel con los datos que nosotros ingresemos
    ws.append([nombre, edad, email, telefono, direccion])
    wb.save(nombre_archivo)
    messagebox.showinfo("Información", "Datos guardados con éxito")
    
    #Limpiamos los datos ingresados
    entry_nombre.delete(0, tk.END)
    entry_edad.delete(0, tk.END)
    entry_email.delete(0, tk.END)
    entry_telefono.delete(0, tk.END)
    entry_direccion.delete(0, tk.END)
    

#crear ventana
root = tk.Tk()
root.title("Formulario de Entrada de Datos")
root.configure(bg="#085979")
#Centra la ventana
root.eval('tk::PlaceWindow . center')
label_style = {"bg": "#085979", "fg": "white"}
entry_style = {"bg": "#D3D3D3", "fg": "black"}

#Titulo
label_titulo = tk.Label(root, text="INGRESE LOS DATOS EN EL FORMULARIO", **label_style)
label_titulo.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

#Nombre
label_nombre = tk.Label(root, text="Nombre", **label_style)
label_nombre.grid(row=1, column=0, padx=10, pady=5)
entry_nombre = tk.Entry(root, **entry_style)
entry_nombre.grid(row=1, column=1, padx=10, pady=5)

#Edad
label_edad = tk.Label(root, text="Edad", **label_style)
label_edad.grid(row=2, column=0, padx=10, pady=5)
entry_edad = tk.Entry(root, **entry_style)
entry_edad.grid(row=2, column=1, padx=10, pady=5)

#Email
label_email = tk.Label(root, text="Email", **label_style)
label_email.grid(row=3, column=0, padx=10, pady=5)
entry_email = tk.Entry(root, **entry_style)
entry_email.grid(row=3, column=1, padx=10, pady=5)

#Telefono
label_telefono = tk.Label(root, text="Telefono", **label_style)
label_telefono.grid(row=4, column=0, padx=10, pady=5)
entry_telefono = tk.Entry(root, **entry_style)
entry_telefono.grid(row=4, column=1, padx=10, pady=5)

#Direccion
label_direccion = tk.Label(root, text="Direccion", **label_style)
label_direccion.grid(row=5, column=0, padx=10, pady=5)
entry_direccion = tk.Entry(root, **entry_style)
entry_direccion.grid(row=5, column=1, padx=10, pady=5)

boton_guardar = tk.Button(root, text="Guardar", command=guardar_datos, bg="#6D8299", fg="white", width=20)
boton_guardar.grid(row=6, column=0, columnspan=2, padx=10, pady=10)

root.mainloop() 