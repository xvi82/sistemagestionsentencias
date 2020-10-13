from __future__ import print_function
import tkinter as tk
from tkinter import ttk
from mailmerge import MailMerge
from tkcalendar import DateEntry

root = tk.Tk()
root.geometry("350x450")
root.resizable(width=False, height=False)
root.title("Selección de Modelos de Recargo de Prestaciones")

#Crear una elección de Modelo
tk.Label(root, text="Elige modelo",
         justify = tk.CENTER,
         padx = 20).pack()
# Crear botones para elegir modelo.
modelo = 1
def numerodelmodelo(value):
    global modelo
    modelo = value

button1 = tk.Button(root, text="Recargo de Prestaciones Condenatoria", command=lambda *args: numerodelmodelo(1), width=30)
button2 = tk.Button(root, text="Recargo de Prestaciones Estimatoria", command=lambda *args: numerodelmodelo(2), width=30)
button1.place(x=60, y=20)
button2.place(x=60, y=50)


# Crear caja de texto 2, nombre del actor.
entry_var2 = tk.StringVar()
entry2 = ttk.Entry(root, textvariable=entry_var2, justify=tk.CENTER)
label2 = ttk.Label(root, text="Nombre del actor")
label2.place(x=15, y=90)
# Posicionar la caja 2 en la ventana, nombre del actor.
entry2.place(x=210, y=90)


# Crear caja de texto 3, nombre del demandado.
entry_var3 = tk.StringVar()
entry3 = ttk.Entry(root, textvariable=entry_var3, justify=tk.CENTER)
label3 = ttk.Label(root, text="Nombre del demandado")
label3.place(x=15, y=115)
# Posicionar la caja 3 en la ventana, nombre del demandado.
entry3.place(x=210, y=115)


# Crear caja de fecha 4, fecha de acta.
entry_var4 = tk.StringVar()
entry4 = DateEntry(root, textvariable=entry_var4, width=12, background='gray', foreground='white', borderwidth=1)
label4 = ttk.Label(root, text="Fecha Acta de Infracción")
label4.place(x=15, y=140)
# Posicionar la caja 4 en la ventana, nombre del demandado.
entry4.place(x=210, y=140, width=126, height=21)


# Crear caja de fecha 5, fecha resolucion impugnada.
entry_var5 = tk.StringVar()
entry5 = DateEntry(root, textvariable=entry_var5, width=12, background='gray', foreground='white', borderwidth=1)
label5 = ttk.Label(root, text="Fecha Resolución impugnada")
label5.place(x=15, y=165)
# Posicionar la caja 5 en la ventana, nombre del demandado.
entry5.place(x=210, y=165, width=126, height=21)


# Crear caja de texto 6, artículos infringidos.
entry_var6 = tk.StringVar()
entry6 = ttk.Entry(root, textvariable=entry_var6, justify=tk.CENTER)
label6 = ttk.Label(root, text="Artículos infringidos")
label6.place(x=15, y=190)
# Posicionar la caja 6 en la ventana, nombre del demandado.
entry6.place(x=210, y=190)


# Crear caja de fecha 7, fecha de la reclamación previa.
entry_var7 = tk.StringVar()
entry7 = DateEntry(root, textvariable=entry_var7, width=12, background='gray', foreground='white', borderwidth=1)
label7 = ttk.Label(root, text="Fecha Reclamación Previa")
label7.place(x=15, y=215)
# Posicionar la caja 7 en la ventana, nombre del demandado.
entry7.place(x=210, y=215, width=126, height=21)


# Crear caja de fecha 8, fecha de la resolución de la reclamación previa.
entry_var8 = tk.StringVar()
entry8 = DateEntry(root, textvariable=entry_var8, width=12, background='gray', foreground='white', borderwidth=1)
label8 = ttk.Label(root, text="Fecha Resolución Recl. Previa")
label8.place(x=15, y=240)
# Posicionar la caja 8 en la ventana, nombre del demandado.
entry8.place(x=210, y=240, width=126, height=21)


# Crear caja de texto 9, pruena practicada.
entry_var9 = tk.StringVar()
entry9 = ttk.Entry(root, textvariable=entry_var9, justify=tk.CENTER)
label9 = ttk.Label(root, text="Prueba practicada (sin documental)")
label9.place(x=15, y=265)
# Posicionar la caja 9 en la ventana, nombre del demandado.
entry9.place(x=210, y=265)


# Crear caja de texto 10, nombre del demandado.
entry_var10 = tk.StringVar()
entry10 = ttk.Entry(root, textvariable=entry_var10, justify=tk.LEFT)
label10 = ttk.Label(root, text="En el presente caso...")
label10.place(x=15, y=290)
# Posicionar la caja 10 en la ventana, nombre del demandado.
entry10.place(x=210, y=290)

#aqui ponemos fuerza mayor o caso fortuito

OptionList = ["Fuerza Mayor", "Fortuito"]
variable = tk.StringVar()
variable.set(OptionList[0])
opt = tk.OptionMenu(root, variable, *OptionList)
opt.place(x=112, y=320)

# Crear caja de texto 11, numero del procedimiento.
entry_var11 = tk.StringVar()
entry11 = ttk.Entry(root, textvariable=entry_var11, justify=tk.LEFT)
label11 = ttk.Label(root, text="Número")
label11.place(x=80, y=365)
# Posicionar la caja 11 en la ventana, nombre del demandado.
entry11.place(x=140, y=365, width=40, height=20)

# Crear caja de texto 12, año del procedimiento.
entry_var12 = tk.StringVar()
entry12 = ttk.Entry(root, textvariable=entry_var12, justify=tk.LEFT)
label12 = ttk.Label(root, text="Año")
label12.place(x=190, y=365)
# Posicionar la caja 9 en la ventana, nombre del demandado.
entry12.place(x=230, y=365, width=40, height=20)


# Posicionar boton de fin
button = ttk.Button(root, text="Crear Modelo", command=root.destroy)
button.place(x=130, y=400)
root.mainloop()

if modelo == 1:
    template = "Prueba1.docx"
else:
    template = "Prueba5.docx"

if template == "Prueba1.docx":
    sentido = "desestim"
else:
    sentido = "estim total"

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
        Enelpresentecaso=entry_var10.get())

if variable.get() == "Fuerza Mayor":
    document.merge(fortuitomayor1="por la fuerza mayor operada al tiempo del accidente", fortuitomayor2 = "sino que el hecho acaecido responde a un caso de fuerza mayor, esto es, un evento extraño al círculo o ámbito de la actividad del trabajador y de la empresa, que rompe el nexo de causalidad")
else:
    document.merge(fortuitomayor1="por su acaecimiento fortuito", fortuitomayor2 = "sino que el hecho acaecido responde a un caso fortuito, esto es, un hecho independiente de la voluntad del deudor de seguridad (la empresa), imprevisible e inevitable")

document.write('recargo prestaciones' + " " + entry_var11.get() + "-" + entry_var12.get() + " " + sentido + '.docx')


