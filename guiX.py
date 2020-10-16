from __future__ import print_function
import tkinter as tk
from tkinter import ttk
from mailmerge import MailMerge
from tkcalendar import DateEntry
from datetime import date
from Buscador import indemnizacion_despido_prueba
today = date.today()

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

def mostrarrecargocondena():
    label27.grid_forget()
    entry27.grid_forget()
    label2.grid_forget()
    entry2.grid_forget()
    label3.grid_forget()
    entry3.grid_forget()
    label20.grid_forget()
    entry20.grid_forget()
    label21.grid_forget()
    entry21.grid_forget()
    label22.grid_forget()
    entry22.grid_forget()
    label23.grid_forget()
    entry23.grid_forget()
    label24.grid_forget()
    entry24.grid_forget()
    opt2.grid_forget()
    optlabel2.grid_forget()
    entry9.grid_forget()
    label9.grid_forget()
    entry10.grid_forget()
    label10.grid_forget()
    label27.grid_forget()
    entry27.grid_forget()
    opt.grid_forget()
    optlabel.grid_forget()
    entry13.grid_forget()
    label13.grid_forget()
    entry14.grid_forget()
    label14.grid_forget()
    label2.grid(row=6, column=0)
    # Posicionar la caja 2 en la ventana, nombre del actor.
    entry2.grid(row=6, column=2)
    label3.grid(row=7, column=0)
    # Posicionar la caja 3 en la ventana, nombre del demandado.
    entry3.grid(row=7, column=2)
    label4.grid(row=8, column=0)
    # Posicionar la caja 4 en la ventana, nombre del demandado.
    entry4.grid(row=8, column=2)
    label5.grid(row=9, column=0)
    # Posicionar la caja 5 en la ventana, nombre del demandado.
    entry5.grid(row=9, column=2)
    label15.grid(row=10, column=0)
    # Posicionar la caja 10 en la ventana, nombre del demandado.
    entry15.grid(row=10, column=2)
    label6.grid(row=11, column=0)
    # Posicionar la caja 6 en la ventana, nombre del demandado.
    entry6.grid(row=11, column=2)
    label7.grid(row=12, column=0)
    # Posicionar la caja 7 en la ventana, nombre del demandado.
    entry7.grid(row=12, column=2)
    label8.grid(row=13, column=0)
    # Posicionar la caja 8 en la ventana, nombre del demandado.
    entry8.grid(row=13, column=2)
    label9.grid(row=14, column=0)
    # Posicionar la caja 9 en la ventana, nombre del demandado.
    entry9.grid(row=14, column=2)
    label10.grid(row=15, column=0)
    # Posicionar la caja 10 en la ventana, nombre del demandado.
    entry10.grid(row=15, column=2)
    label11.grid(row=16, column=0)
    # Posicionar la caja 9 en la ventana, nombre del demandado.
    entry11.grid(row=16, column=2)
    label12.grid(row=17, column=0)
    # Posicionar la caja 10 en la ventana, nombre del demandado.
    entry12.grid(row=17, column=2)
    button.grid(row=20, column=1)

def mostrarrecargoestim():
    label27.grid_forget()
    entry27.grid_forget()
    label2.grid_forget()
    entry2.grid_forget()
    label3.grid_forget()
    entry3.grid_forget()
    label20.grid_forget()
    entry20.grid_forget()
    label21.grid_forget()
    entry21.grid_forget()
    label22.grid_forget()
    entry22.grid_forget()
    label23.grid_forget()
    entry23.grid_forget()
    label24.grid_forget()
    entry24.grid_forget()
    opt2.grid_forget()
    optlabel2.grid_forget()
    entry9.grid_forget()
    label9.grid_forget()
    entry10.grid_forget()
    label10.grid_forget()
    label27.grid_forget()
    entry27.grid_forget()
    entry13.grid_forget()
    label13.grid_forget()
    entry14.grid_forget()
    label14.grid_forget()
    opt.grid(row=5, column=2)
    optlabel.grid(row=5, column=0)
    label2.grid(row=6, column=0)
    # Posicionar la caja 2 en la ventana, nombre del actor.
    entry2.grid(row=6, column=2)
    label3.grid(row=7, column=0)
    # Posicionar la caja 3 en la ventana, nombre del demandado.
    entry3.grid(row=7, column=2)
    label4.grid(row=8, column=0)
    # Posicionar la caja 4 en la ventana, nombre del demandado.
    entry4.grid(row=8, column=2)
    label5.grid(row=9, column=0)
    # Posicionar la caja 5 en la ventana, nombre del demandado.
    entry5.grid(row=9, column=2)
    label15.grid(row=10, column=0)
    # Posicionar la caja 10 en la ventana, nombre del demandado.
    entry15.grid(row=10, column=2)
    label6.grid(row=11, column=0)
    # Posicionar la caja 6 en la ventana, nombre del demandado.
    entry6.grid(row=11, column=2)
    label7.grid(row=12, column=0)
    # Posicionar la caja 7 en la ventana, nombre del demandado.
    entry7.grid(row=12, column=2)
    label8.grid(row=13, column=0)
    # Posicionar la caja 8 en la ventana, nombre del demandado.
    entry8.grid(row=13, column=2)
    label9.grid(row=14, column=0)
    # Posicionar la caja 9 en la ventana, nombre del demandado.
    entry9.grid(row=14, column=2)
    label10.grid(row=15, column=0)
    # Posicionar la caja 10 en la ventana, nombre del demandado.
    entry10.grid(row=15, column=2)
    label11.grid(row=16, column=0)
    # Posicionar la caja 9 en la ventana, nombre del demandado.
    entry11.grid(row=16, column=2)
    label12.grid(row=17, column=0)
    # Posicionar la caja 10 en la ventana, nombre del demandado.
    entry12.grid(row=17, column=2)
    button.grid(row=20, column=1)

def mostrarrecargoestimparcial():
    label27.grid_forget()
    entry27.grid_forget()
    label2.grid_forget()
    entry2.grid_forget()
    label3.grid_forget()
    entry3.grid_forget()
    label20.grid_forget()
    entry20.grid_forget()
    label21.grid_forget()
    entry21.grid_forget()
    label22.grid_forget()
    entry22.grid_forget()
    label23.grid_forget()
    entry23.grid_forget()
    label24.grid_forget()
    entry24.grid_forget()
    opt2.grid_forget()
    optlabel2.grid_forget()
    entry9.grid_forget()
    label9.grid_forget()
    entry10.grid_forget()
    label10.grid_forget()
    label27.grid_forget()
    entry27.grid_forget()
    opt.grid_forget()
    optlabel.grid_forget()
    entry13.grid(row=16, column=2)
    label13.grid(row=16, column=0)
    entry14.grid(row=17, column=2)
    label14.grid(row=17, column=0)
    label2.grid(row=6, column=0)
    # Posicionar la caja 2 en la ventana, nombre del actor.
    entry2.grid(row=6, column=2)
    label3.grid(row=7, column=0)
    # Posicionar la caja 3 en la ventana, nombre del demandado.
    entry3.grid(row=7, column=2)
    label4.grid(row=8, column=0)
    # Posicionar la caja 4 en la ventana, nombre del demandado.
    entry4.grid(row=8, column=2)
    label5.grid(row=9, column=0)
    # Posicionar la caja 5 en la ventana, nombre del demandado.
    entry5.grid(row=9, column=2)
    label15.grid(row=10, column=0)
    # Posicionar la caja 10 en la ventana, nombre del demandado.
    entry15.grid(row=10, column=2)
    label6.grid(row=11, column=0)
    # Posicionar la caja 6 en la ventana, nombre del demandado.
    entry6.grid(row=11, column=2)
    label7.grid(row=12, column=0)
    # Posicionar la caja 7 en la ventana, nombre del demandado.
    entry7.grid(row=12, column=2)
    label8.grid(row=13, column=0)
    # Posicionar la caja 8 en la ventana, nombre del demandado.
    entry8.grid(row=13, column=2)
    label9.grid(row=14, column=0)
    # Posicionar la caja 9 en la ventana, nombre del demandado.
    entry9.grid(row=14, column=2)
    label10.grid(row=15, column=0)
    # Posicionar la caja 10 en la ventana, nombre del demandado.
    entry10.grid(row=15, column=2)
    label11.grid(row=18, column=0)
    # Posicionar la caja 9 en la ventana, nombre del demandado.
    entry11.grid(row=18, column=2)
    label12.grid(row=19, column=0)
    # Posicionar la caja 10 en la ventana, nombre del demandado.
    entry12.grid(row=19, column=2)
    button.grid(row=20, column=1)

def mostrarextincionretraso():
    opt.grid_forget()
    optlabel.grid_forget()
    entry13.grid_forget()
    label13.grid_forget()
    entry14.grid_forget()
    label14.grid_forget()
    entry4.grid_forget()
    label4.grid_forget()
    entry5.grid_forget()
    label5.grid_forget()
    entry15.grid_forget()
    label15.grid_forget()
    entry6.grid_forget()
    label6.grid_forget()
    entry7.grid_forget()
    label7.grid_forget()
    entry8.grid_forget()
    label8.grid_forget()
    label27.grid_forget()
    entry27.grid_forget()
    label2.grid(row=6, column=0)
    entry2.grid(row=6, column=2)
    label3.grid(row=7, column=0)
    entry3.grid(row=7, column=2)
    label20.grid(row=8, column=0)
    entry20.grid(row=8, column=2)
    label21.grid(row=9, column=0)
    entry21.grid(row=9, column=2)
    label22.grid(row=10, column=0)
    entry22.grid(row=10, column=2)
    label23.grid(row=11, column=0)
    entry23.grid(row=11, column=2)
    label24.grid(row=12, column=0)
    entry24.grid(row=12, column=2)
    opt2.grid(row=13, column=2)
    optlabel2.grid(row=13, column=0)
    entry9.grid(row=14, column=2)
    label9.grid(row=14, column=0)
    entry10.grid(row=15, column=2)
    label10.grid(row=15, column=0)
    label11.grid(row=17, column=0)
    # Posicionar la caja 9 en la ventana, nombre del demandado.
    entry11.grid(row=17, column=2)
    label12.grid(row=18, column=0)
    # Posicionar la caja 10 en la ventana, nombre del demandado.
    entry12.grid(row=18, column=2)
    button.grid(row=20, column=1)

def mostrarextincionfaltapago():
    opt.grid_forget()
    optlabel.grid_forget()
    entry13.grid_forget()
    label13.grid_forget()
    entry14.grid_forget()
    label14.grid_forget()
    entry4.grid_forget()
    label4.grid_forget()
    entry5.grid_forget()
    label5.grid_forget()
    entry15.grid_forget()
    label15.grid_forget()
    entry6.grid_forget()
    label6.grid_forget()
    entry7.grid_forget()
    label7.grid_forget()
    entry8.grid_forget()
    label8.grid_forget()
    label2.grid(row=6, column=0)
    entry2.grid(row=6, column=2)
    label3.grid(row=7, column=0)
    entry3.grid(row=7, column=2)
    label20.grid(row=8, column=0)
    entry20.grid(row=8, column=2)
    label21.grid(row=9, column=0)
    entry21.grid(row=9, column=2)
    label22.grid(row=10, column=0)
    entry22.grid(row=10, column=2)
    label23.grid(row=11, column=0)
    entry23.grid(row=11, column=2)
    label24.grid(row=12, column=0)
    entry24.grid(row=12, column=2)
    opt2.grid(row=13, column=2)
    optlabel2.grid(row=13, column=0)
    entry9.grid(row=14, column=2)
    label9.grid(row=14, column=0)
    entry10.grid(row=15, column=2)
    label10.grid(row=15, column=0)
    label27.grid(row=16, column=0)
    entry27.grid(row=16, column=2)
    label11.grid(row=17, column=0)
    # Posicionar la caja 9 en la ventana, nombre del demandado.
    entry11.grid(row=17, column=2)
    label12.grid(row=18, column=0)
    # Posicionar la caja 10 en la ventana, nombre del demandado.
    entry12.grid(row=18, column=2)
    button.grid(row=20, column=1)

def mostrarextincionmodificacion():
    opt.grid_forget()
    optlabel.grid_forget()
    entry13.grid_forget()
    label13.grid_forget()
    entry14.grid_forget()
    label14.grid_forget()
    entry4.grid_forget()
    label4.grid_forget()
    entry5.grid_forget()
    label5.grid_forget()
    entry15.grid_forget()
    label15.grid_forget()
    entry6.grid_forget()
    label6.grid_forget()
    entry7.grid_forget()
    label7.grid_forget()
    entry8.grid_forget()
    label8.grid_forget()
    label27.grid_forget()
    entry27.grid_forget()
    label2.grid(row=6, column=0)
    entry2.grid(row=6, column=2)
    label3.grid(row=7, column=0)
    entry3.grid(row=7, column=2)
    label20.grid(row=8, column=0)
    entry20.grid(row=8, column=2)
    label21.grid(row=9, column=0)
    entry21.grid(row=9, column=2)
    label22.grid(row=10, column=0)
    entry22.grid(row=10, column=2)
    label23.grid(row=11, column=0)
    entry23.grid(row=11, column=2)
    label24.grid(row=12, column=0)
    entry24.grid(row=12, column=2)
    opt2.grid(row=13, column=2)
    optlabel2.grid(row=13, column=0)
    entry9.grid(row=14, column=2)
    label9.grid(row=14, column=0)
    entry10.grid(row=15, column=2)
    label10.grid(row=15, column=0)
    label11.grid(row=17, column=0)
    # Posicionar la caja 9 en la ventana, nombre del demandado.
    entry11.grid(row=17, column=2)
    label12.grid(row=18, column=0)
    # Posicionar la caja 10 en la ventana, nombre del demandado.
    entry12.grid(row=18, column=2)
    button.grid(row=20, column=1)

botoncondena = tk.Button(root, text="Condenatoria", command=lambda: [mostrarrecargocondena(), numerodelmodelo(1)], width=20)
botonabsolucion = tk.Button(root, text="Estimatoria", command=lambda: [mostrarrecargoestim(), numerodelmodelo(2)], width=20)
botonabsolucion2 = tk.Button(root, text="Estimatoria Parcial", command=lambda: [mostrarrecargoestimparcial(), numerodelmodelo(3)], width=20)

botonretraso = tk.Button(root, text="Retraso en el pago", width="20", command=lambda: [mostrarextincionretraso(), numerodelmodelo(4)])
botonfaltapago = tk.Button(root, text="Falta de pago", width="20", command=lambda: [mostrarextincionfaltapago(), numerodelmodelo(5)])
botonmodificacion = tk.Button(root, text="Modificacion Condiciones", width="20", command=lambda: [mostrarextincionmodificacion(), numerodelmodelo(6)])

def mostrarrecargo():
    botoncondena.grid(row=1, column=2)
    botonabsolucion.grid(row=2, column=2)
    botonabsolucion2.grid(row=3, column=2)
    botonretraso.grid_forget()
    botonfaltapago.grid_forget()
    botonmodificacion.grid_forget()


def mostrarexincion():
    botonretraso.grid(row=1,column=2)
    botonfaltapago.grid(row=2,column=2)
    botonmodificacion.grid(row=3,column=2)
    botoncondena.grid_forget()
    botonabsolucion.grid_forget()
    botonabsolucion2.grid_forget()


botonrecargo = tk.Button(root, text="Recargo de Prestaciones", width="20", command= lambda: [mostrarrecargo()]).grid(row=1,column=1)
botonextinción = tk.Button(root, text="Extinción del Contrato", width="20", command= lambda: [mostrarexincion()]).grid(row=3,column=1)

#aqui ponemos fuerza mayor o caso fortuito



OptionList = ["Fuerza Mayor", "Fortuito"]
variable = tk.StringVar()
variable.set(OptionList[0])
opt = tk.OptionMenu(root, variable, *OptionList)
optlabel = ttk.Label(root, text="Fundamento de estimatoria")

# Espacio en blanco.
labelx = ttk.Label(root, text=" ").grid(row=4, column=0)

# Crear caja de texto 2, nombre del actor.
entry_var2 = tk.StringVar()
entry2 = ttk.Entry(root, textvariable=entry_var2, justify=tk.CENTER)
label2 = ttk.Label(root, text="Nombre del actor")



# Crear caja de texto 3, nombre del demandado.
entry_var3 = tk.StringVar()
entry3 = ttk.Entry(root, textvariable=entry_var3, justify=tk.CENTER)
label3 = ttk.Label(root, text="Nombre del demandado")



# Crear caja de fecha 4, fecha de acta.
entry_var4 = tk.StringVar()
entry4 = DateEntry(root, textvariable=entry_var4, width=17, background='gray', foreground='white', borderwidth=1)
label4 = ttk.Label(root, text="Fecha del Acta de Infracción")

# Crear caja de fecha 5, fecha resolucion impugnada.
entry_var5 = tk.StringVar()
entry5 = DateEntry(root, textvariable=entry_var5, width=17, background='gray', foreground='white', borderwidth=1)
label5 = ttk.Label(root, text="Fecha de la Resolución impugnada")



# Crear caja de texto 15, porcentaje impuesto por la Inspección.
entry_var15 = tk.StringVar()
entry15 = ttk.Entry(root, textvariable=entry_var15, justify=tk.CENTER)
label15 = ttk.Label(root, text="Porcentaje impuesto por la Inspección")



# Crear caja de texto 6, artículos infringidos.
entry_var6 = tk.StringVar()
entry6 = ttk.Entry(root, textvariable=entry_var6, justify=tk.CENTER)
label6 = ttk.Label(root, text="Artículos infringidos según acta de Infracción")



# Crear caja de fecha 7, fecha de la reclamación previa.
entry_var7 = tk.StringVar()
entry7 = DateEntry(root, textvariable=entry_var7, width=17, background='gray', foreground='white', borderwidth=1)
label7 = ttk.Label(root, text="Fecha de la Reclamación Previa")



# Crear caja de fecha 8, fecha de la resolución de la reclamación previa.
entry_var8 = tk.StringVar()
entry8 = DateEntry(root, textvariable=entry_var8, width=17, background='gray', foreground='white', borderwidth=1)
label8 = ttk.Label(root, text="Fecha de Resolución de la Reclamación Previa")



# Crear caja de texto 9, pruena practicada.
entry_var9 = tk.StringVar()
entry9 = ttk.Entry(root, textvariable=entry_var9, justify=tk.CENTER)
label9 = ttk.Label(root, text="Prueba practicada aparte de la documental")



# Crear caja de texto 10, en el presente caso.
entry_var10 = tk.StringVar()
entry10 = ttk.Entry(root, textvariable=entry_var10, justify=tk.CENTER)
label10 = ttk.Label(root, text="En el presente caso...")



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


# Crear caja de texto 12, año del procedimiento.
entry_var12 = tk.StringVar()
entry12 = ttk.Entry(root, textvariable=entry_var12, justify=tk.CENTER)
label12 = ttk.Label(root, text="Año")

# Crear caja de fecha 4, fecha de acta.
entry_var20 = tk.StringVar()
entry20 = ttk.Entry(root, textvariable=entry_var20, justify=tk.CENTER)
label20 = ttk.Label(root, text="Actividad de la Empresa")

# Crear caja de fecha 5, fecha resolucion impugnada.
entry_var21 = tk.StringVar()
entry21 = DateEntry(root, textvariable=entry_var21, width=17, background='gray', foreground='white', borderwidth=1)
label21 = ttk.Label(root, text="Antigüedad")



# Crear caja de texto 15, porcentaje impuesto por la Inspección.
entry_var22 = tk.StringVar()
entry22 = ttk.Entry(root, textvariable=entry_var22, justify=tk.CENTER)
label22 = ttk.Label(root, text="Categoría Profesional")



# Crear caja de texto 6, artículos infringidos.
entry_var23 = tk.StringVar()
entry23 = ttk.Entry(root, textvariable=entry_var23, justify=tk.CENTER)
label23 = ttk.Label(root, text="Salario diario")



# Crear caja de fecha 7, fecha de la reclamación previa.
entry_var24 = tk.StringVar()
entry24 = DateEntry(root, textvariable=entry_var24, width=17, background='gray', foreground='white', borderwidth=1)
label24 = ttk.Label(root, text="Fecha celebración SEMAC")

OptionList2 = ["Sin avenencia", "Sin efecto"]
variable2 = tk.StringVar()
variable2.set(OptionList2[0])
opt2 = tk.OptionMenu(root, variable2, *OptionList2)
optlabel2 = ttk.Label(root, text="Resultado del SEMAC")


# Crear caja de texto 10, nombre del demandado.
entry_var27 = tk.StringVar()
entry27 = ttk.Entry(root, textvariable=entry_var27, justify=tk.CENTER)
label27 = ttk.Label(root, text="Desde cuando no cobra")
# Posicionar la caja 10 en la ventana, nombre del demandado.




# Posicionar boton de fin
button = ttk.Button(root, text="Crear Modelo", command=root.destroy)

root.mainloop()

if modelo == 1:
    template = "Prueba1.docx"
elif modelo == 2:
    template = "Prueba5.docx"
elif modelo == 4:
    template = "Extincion retraso pago.docx"
elif modelo == 5:
    template = "Extincion falta de pago.docx"
elif modelo == 6:
    template = "Extincion modificacion sustancial.docx"
else:
    template = "Prueba4.docx"

if template == "Prueba1.docx":
    sentido = "desestim"
elif template == "Prueba5.docx":
    sentido = "estim total"
elif template == "Extincion retraso pago.docx":
    sentido = "estim"
elif template == "Extincion falta de pago.docx":
    sentido = "estim"
elif template == "Extincion modificacion sustancial.docx":
    sentido = "estim"
else:
    sentido = "estim parcial"

if template == "Prueba1.docx":
    materia = "recargo prestaciones"
elif template == "Prueba5.docx":
    materia = "recargo prestaciones"
elif template == "Extincion retraso pago.docx":
    materia = "extincion contrato"
elif template == "Extincion falta de pago.docx":
    materia = "extincion contrato"
elif template == "Extincion modificacion sustancial.docx":
    materia = "extincion contrato"
else:
    materia = "recargo prestaciones"

document = MailMerge(template)

indemnizacion = indemnizacion_despido_prueba(str(entry_var21.get()),str(entry_var23.get()),str(today))

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
        porcentajerecargo=entry_var15.get(),
        antiguedad=entry_var21.get(),
        desdecuandonocobra=entry_var27.get(),
        fechasemac=entry_var24.get(),
        indemnizacionextincion=indemnizacion,
        categoría=entry_var22.get(),
        salariodia=entry_var23.get(),
        actividadempresa=entry_var20.get())

if variable.get() == "Fuerza Mayor":
    document.merge(fortuitomayor1="por la fuerza mayor operada al tiempo del accidente", fortuitomayor2 = "sino que el hecho acaecido responde a un caso de fuerza mayor, esto es, un evento extraño al círculo o ámbito de la actividad del trabajador y de la empresa, que rompe el nexo de causalidad")
else:
    document.merge(fortuitomayor1="por su acaecimiento fortuito", fortuitomayor2 = "sino que el hecho acaecido responde a un caso fortuito, esto es, un hecho independiente de la voluntad del deudor de seguridad (la empresa), imprevisible e inevitable")

if variable2.get() == "Sin avenencia":
    document.merge(avenenciaoefecto="sin avenencia")
else:
    document.merge(avenenciaoefecto="sin efecto")

document.write(materia + " " + entry_var11.get() + "-" + entry_var12.get() + " " + sentido + '.docx')