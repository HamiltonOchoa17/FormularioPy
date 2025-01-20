import tkinter
import tkinter as tk
from dataclasses import make_dataclass
from tkinter import messagebox
from tokenize import cookie_re
import re

from openpyxl import workbook
from openpyxl.workbook import Workbook

# este codigo cre un libro de excel con la libreria workbook

wb = Workbook()  #Se crea un objeto de workbook
ws = wb.active  #Activa el documento de excel para poder trabajar en el
ws.append(["Nombre", "Edad", "Email","Telefono","Direccion"])

root = tk.Tk()    #Crea una ventana con Tk
root.title("Formulario de datos")
root.configure(bg='#4B6587')   #Agrega color a el backgraund (bg)


label_style = {"bg": '#4B6587',"fg": 'white'}
entry_style = {"bg": '#D3D3D3',"fg": 'black'}


def guardad_datos():
    nombre = entry_nombre.get()
    edad = entry_edad.get()
    email = entry_email.get()
    telefono = entry_telefono.get()
    direccion = entry_direccion.get()

    if not nombre  or not edad or not email or not telefono or not direccion:
        messagebox.showwarning("Advertencia","Todos los campos son ovigatorios")
        return
    try:
        edad = int(edad)
        telefono = int(telefono)
    except ValueError:
        messagebox.showwarning("Advertencia", "edad y telefono tiene que ser un numero")
        return
    # Se usa la libreria re para conparara el texto en los entry

    if not re.match(r"[^@]+@[^@]+\.[^@]",email):  #se usa la libreria r [^@] "Cuaquier cosa menos un @"  +@[^@] "Mas un arroba y cuaquier cosa menos un @ despues"+  \.[^@] "Mas un punto y despues lo que sea meno un @"
        messagebox.showwarning("Advertencia", "El correo electronico no es valido")
        return

    ws.append([nombre,edad,email,telefono,direccion])
    wb.save('datos.xls')
    messagebox.showinfo("Guardado", "Se a guardado la informacion")


label_nombre = tk.Label(root, text ="Nombre", **label_style) # Se usa **Diccionario es una funcion de desempaquedo de diccionarios
label_nombre.grid(row=0, column=0, padx=10, pady=5)
entry_nombre =  tk.Entry(root,**entry_style)
entry_nombre.grid(row=0,column=1,padx=10, pady=5)


label_edad = tk.Label(root, text ="Edad", **label_style) # Se usa **Diccionario es una funcion de desempaquedo de diccionarios
label_edad.grid(row=1, column=0, padx=10, pady=5)
entry_edad =  tk.Entry(root,**entry_style)
entry_edad.grid(row=1,column=1,padx=10, pady=5)

label_email = tk.Label(root, text ="Email", **label_style) # Se usa **Diccionario es una funcion de desempaquedo de diccionarios
label_email.grid(row=2, column=0, padx=10, pady=5)
entry_email =  tk.Entry(root,**entry_style)
entry_email.grid(row=2,column=1,padx=10, pady=5)


label_telefono = tk.Label(root, text ="Telefono", **label_style) # Se usa **Diccionario es una funcion de desempaquedo de diccionarios
label_telefono.grid(row=3, column=0, padx=10, pady=5)
entry_telefono =  tk.Entry(root,**entry_style)
entry_telefono.grid(row=3,column=1,padx=10, pady=5)



label_direccion = tk.Label(root, text ="Direccion", **label_style) # Se usa **Diccionario es una funcion de desempaquedo de diccionarios
label_direccion.grid(row=4, column=0, padx=10, pady=5)
entry_direccion  =  tk.Entry(root,**entry_style)
entry_direccion.grid(row=4,column=1,padx=10, pady=5)


boton_guardar = tk.Button(root, text="Guardar", command=guardad_datos, bg='#6D8299', fg='white')
boton_guardar.grid(row=5, column=0, columnspan=2, padx=10, pady=10)

root.mainloop()