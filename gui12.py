from __future__ import print_function
import tkinter as tk
from tkinter import ttk
from mailmerge import MailMerge
from tkcalendar import DateEntry

root = tk.Tk()

root.resizable(width=False, height=False)
root.title("Selección de Modelos de Recargo de Prestaciones")

#Crear una elección de Modelo
tk.Label(root, text="Elige modelo:", font=('Arial', 12, "bold"),
         justify = tk.CENTER,
         padx = 20).grid(row=2, column=0)
# Crear botones para elegir modelo.
modelo = 1
def numerodelmodelo(value):
    global modelo
    modelo = value

botoncondena = tk.Button(root, text="Condenatoria", command=lambda: [opt.grid_forget(), optlabel.grid_forget(), entry13.grid_forget(), label13.grid_forget(), entry14.grid_forget(), label14.grid_forget(), numerodelmodelo(1)], width=15)
botonabsolucion = tk.Button(root, text="Estimatoria", command=lambda: [opt.grid(row=5, column=2), optlabel.grid(row=5, column=0), entry13.grid_forget(), label13.grid_forget(), entry14.grid_forget(), label14.grid_forget(), numerodelmodelo(2)], width=15)
botonabsolucion2 = tk.Button(root, text="Estimatoria Parcial", command=lambda: [opt.grid_forget(), optlabel.grid_forget(), entry13.grid(row=16, column=2), label13.grid(row=16, column=0), entry14.grid(row=17, column=2), label14.grid(row=17, column=0), numerodelmodelo(3)], width=15)
botoncondena.grid(row=1, column=1)
botonabsolucion.grid(row=2, column=1)
botonabsolucion2.grid(row=3, column=1)

#aqui ponemos fuerza mayor o caso fortuito

OptionList = ["Fuerza Mayor", "Fortuito"]
variable = tk.StringVar()
variable.set(OptionList[0])
opt = tk.OptionMenu(root, variable, *OptionList)
optlabel = ttk.Label(root, text="Fundamento de estimatoria")

# Espacio en blanco.
labelx = ttk.Label(root, text=" ")
labelx.grid(row=4, column=0)

# Crear caja de texto 2, nombre del actor.
entry_var2 = tk.StringVar()
entry2 = ttk.Entry(root, textvariable=entry_var2, justify=tk.CENTER)
label2 = ttk.Label(root, text="Nombre del actor")
label2.grid(row=6, column=0)
# Posicionar la caja 2 en la ventana, nombre del actor.
entry2.grid(row=6, column=2)


# Crear caja de texto 3, nombre del demandado.
entry_var3 = tk.StringVar()
entry3 = ttk.Entry(root, textvariable=entry_var3, justify=tk.CENTER)
label3 = ttk.Label(root, text="Nombre del demandado")
label3.grid(row=7, column=0)
# Posicionar la caja 3 en la ventana, nombre del demandado.
entry3.grid(row=7, column=2)


# Crear caja de fecha 4, fecha de acta.
entry_var4 = tk.StringVar()
entry4 = DateEntry(root, textvariable=entry_var4, width=17, background='gray', foreground='white', borderwidth=1)
label4 = ttk.Label(root, text="Fecha del Acta de Infracción")
label4.grid(row=8, column=0)
# Posicionar la caja 4 en la ventana, nombre del demandado.
entry4.grid(row=8, column=2)


# Crear caja de fecha 5, fecha resolucion impugnada.
entry_var5 = tk.StringVar()
entry5 = DateEntry(root, textvariable=entry_var5, width=17, background='gray', foreground='white', borderwidth=1)
label5 = ttk.Label(root, text="Fecha de la Resolución impugnada")
label5.grid(row=9, column=0)
# Posicionar la caja 5 en la ventana, nombre del demandado.
entry5.grid(row=9, column=2)


# Crear caja de texto 15, porcentaje impuesto por la Inspección.
entry_var15 = tk.StringVar()
entry15 = ttk.Entry(root, textvariable=entry_var15, justify=tk.CENTER)
label15 = ttk.Label(root, text="Porcentaje impuesto por la Inspección")
label15.grid(row=10, column=0)
# Posicionar la caja 10 en la ventana, nombre del demandado.
entry15.grid(row=10, column=2)


# Crear caja de texto 6, artículos infringidos.
entry_var6 = tk.StringVar()
entry6 = ttk.Entry(root, textvariable=entry_var6, justify=tk.CENTER)
label6 = ttk.Label(root, text="Artículos infringidos según acta de Infracción")
label6.grid(row=11, column=0)
# Posicionar la caja 6 en la ventana, nombre del demandado.
entry6.grid(row=11, column=2)


# Crear caja de fecha 7, fecha de la reclamación previa.
entry_var7 = tk.StringVar()
entry7 = DateEntry(root, textvariable=entry_var7, width=17, background='gray', foreground='white', borderwidth=1)
label7 = ttk.Label(root, text="Fecha de la Reclamación Previa")
label7.grid(row=12, column=0)
# Posicionar la caja 7 en la ventana, nombre del demandado.
entry7.grid(row=12, column=2)


# Crear caja de fecha 8, fecha de la resolución de la reclamación previa.
entry_var8 = tk.StringVar()
entry8 = DateEntry(root, textvariable=entry_var8, width=17, background='gray', foreground='white', borderwidth=1)
label8 = ttk.Label(root, text="Fecha de Resolución de la Reclamación Previa")
label8.grid(row=13, column=0)
# Posicionar la caja 8 en la ventana, nombre del demandado.
entry8.grid(row=13, column=2)


# Crear caja de texto 9, pruena practicada.
entry_var9 = tk.StringVar()
entry9 = ttk.Entry(root, textvariable=entry_var9, justify=tk.CENTER)
label9 = ttk.Label(root, text="Prueba practicada aparte de la documental")
label9.grid(row=14, column=0)
# Posicionar la caja 9 en la ventana, nombre del demandado.
entry9.grid(row=14, column=2)


# Crear caja de texto 10, en el presente caso.
entry_var10 = tk.StringVar()
entry10 = ttk.Entry(root, textvariable=entry_var10, justify=tk.CENTER)
label10 = ttk.Label(root, text="En el presente caso...")
label10.grid(row=15, column=0)
# Posicionar la caja 10 en la ventana, nombre del demandado.
entry10.grid(row=15, column=2)


# Crear caja de texto 10, nombre del demandado.
entry_var13 = tk.StringVar()
entry13 = ttk.Entry(root, textvariable=entry_var13, justify=tk.CENTER)
label13 = ttk.Label(root, text="Porcentaje que se estima adecuado")
# Posicionar la caja 10 en la ventana, nombre del demandado.

# Crear caja de texto 10, nombre del demandado.
entry_var14 = tk.StringVar()
entry14 = ttk.Entry(root, textvariable=entry_var14, justify=tk.CENTER)
label14 = ttk.Label(root, text="Justificación del porcentaje que se impone")
# Posicionar la caja 10 en la ventana, nombre del demandado.


# Crear caja de texto 11, numero del procedimiento.
entry_var11 = tk.StringVar()
entry11 = ttk.Entry(root, textvariable=entry_var11, justify=tk.CENTER)
label11 = ttk.Label(root, text="Número del procedimiento")
label11.grid(row=18, column=0)
# Posicionar la caja 11 en la ventana, nombre del demandado.
entry11.grid(row=18, column=2)

# Crear caja de texto 12, año del procedimiento.
entry_var12 = tk.StringVar()
entry12 = ttk.Entry(root, textvariable=entry_var12, justify=tk.CENTER)
label12 = ttk.Label(root, text="Año")
label12.grid(row=19, column=0)
# Posicionar la caja 9 en la ventana, nombre del demandado.
entry12.grid(row=19, column=2)


# Posicionar boton de fin
button = ttk.Button(root, text="Crear Modelo", command=root.destroy)
button.grid(row=20, column=1)
root.mainloop()

if modelo == 1:
    template = "Prueba1.docx"
elif modelo == 2:
    template = "Prueba5.docx"
else:
    template = "Prueba4.docx"

if template == "Prueba1.docx":
    sentido = "desestim"
elif template == "Prueba5.docx":
    sentido = "estim total"
else:
    sentido = "estim parcial"

document = MailMerge(template)

document.merge(
        Actor=str(entry_var2.get()).upper(),
        Demandado=str(entry_var3.get()).title(),
        Fechaactainspeccion=entry_var4.get(),
        fecharesolucionrecargo=entry_var5.get(),
        articulosinfringidos=str(entry_var6.get()) + " de la LPRL",
        fechareclamacionprevia=entry_var7.get(),
        fecharesolucionreclamacionprevia=entry_var8.get(),
        pruebaspracticadas=" así como la " + str(entry_var9.get()),
        Año=entry_var12.get(),
        Número=entry_var11.get(),
        Enelpresentecaso=entry_var10.get(),
        porcentajequeseimpone=entry_var13.get(),
        justificacionporcentaje=entry_var14.get(),
        porcentajerecargo=entry_var15.get())

if variable.get() == "Fuerza Mayor":
    document.merge(fortuitomayor1="por la fuerza mayor operada al tiempo del accidente", fortuitomayor2 = "sino que el hecho acaecido responde a un caso de fuerza mayor, esto es, un evento extraño al círculo o ámbito de la actividad del trabajador y de la empresa, que rompe el nexo de causalidad")
else:
    document.merge(fortuitomayor1="por su acaecimiento fortuito", fortuitomayor2 = "sino que el hecho acaecido responde a un caso fortuito, esto es, un hecho independiente de la voluntad del deudor de seguridad (la empresa), imprevisible e inevitable")

document.write('recargo prestaciones' + " " + entry_var11.get() + "-" + entry_var12.get() + " " + sentido + '.docx')