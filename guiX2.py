from __future__ import print_function
import tkinter as tk
from tkinter import ttk
from mailmerge import MailMerge
from tkcalendar import DateEntry
from datetime import date
from Buscador import *
import pdftotext
from nltk.tokenize import word_tokenize
import re
from nltk.corpus import stopwords
import nltk
from tkinter import filedialog
from tkinter import *

root = Tk()
root.geometry("400x400")

# Obtener valores de alto y ancho de la ventana.
windowWidth = root.winfo_reqwidth()
windowHeight = root.winfo_reqheight()

# Obtener la mitad del ancho/alto de la pantalla y el ancho/alto de la ventana
positionRight = int(root.winfo_screenwidth() / 2 - windowWidth / 2)
positionDown = int(root.winfo_screenheight() / 2 - windowHeight / 2)

# Posiciona la ventana en el centro.
root.geometry("+{}+{}".format(positionRight, positionDown))

#imagen de fondo
C = Canvas(root, bg="blue", height=250, width=300)
filename = PhotoImage(file = "pantallainicio.png")
background_label = Label(root, image=filename)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

#datos de la ventan
root['bg']= "#ffffff"
root.iconbitmap("iconoaplicacionblanco.ico")
root.resizable(width=False, height=False)
root.title("Selección de Demanda")

# sacar una variable global con la dirección del documento y una función para selecionarlo

document = None
def seleccionarchivo():
    global document
    document = filedialog.askopenfilename(initialdir = "/",title = "Selecciona el PDF de la demanda",filetypes = (("pdf files","*.pdf"),("all files","*.*")))

# boton de cargar

botoncarga = tk.Button(root, text="Selecciona la demanda", width="20", bg="#ffffff", command=lambda: [seleccionarchivo(), root.destroy()]).place(x=120, y=360)

root.mainloop()

# cargar el pdf y convertirlo en texto tratable

pdfname = str(document)

with open(pdfname, "rb") as f:
    pdf = pdftotext.PDF(f)
    document = "\n\n".join(pdf)

# limpiar el texto de ruido

cleaned = re.sub(r'[^A-ZñÑa-zá-úä-ü0-9.,\/]+', ' ', document)

# tokenizar el texto

tokenized = word_tokenize(cleaned.lower())

# limpiar el texto de stopwords y tener el texto definitivo de trabajo

stopwords = set(stopwords.words('spanish'))
textprov = [word for word in tokenized if word not in stopwords]

t = list()
for i in textprov:
    if re.match('\,', i):
        t.append(i)
    if re.match('\.', i):
        t.append(i)
textdef = [word for word in textprov if word not in t]

# convierte el texto tokenizado en un texto tratable
text = nltk.Text(textdef)

# variables vacias



# expresiones regulares para el salario y la fecha de efectos

if "extinci" in textdef:
    salarioregex = re.compile('(\d+(\,|\.)?\d+?(\,|\.)?\d+?|\d+)')
    fechaefectosregex = re.compile('\d{1,2}(\/|\.|\-)\d{1,2}(\/|\.|\-)\d{1,4}')

    fechas = [word for word in textdef if re.match('\d{1,2}(\/|\.|\-)\d{1,2}(\/|\.|\-)\d{1,4}', word)]
    try:
        despid = fechas[-1]
    except:
        despid = None

    try:
        antig = fechas[0]
    except:
        antig = None
    salario = None
    demandante = None
    demandado = None
    actainfraccion = None
    resolucionrecargo = None
    reclamacionprevia = None
    resolucionreclamacion = None
    recargo = None
    try:
        for word in textdef:
            if word == "salario":
                if re.match(salarioregex, textdef[(textdef.index(word)) + 2]):
                    salario = textdef[(textdef.index(word)) + 2]
                elif re.match(salarioregex, textdef[(textdef.index(word)) + 1]):
                    salario = textdef[(textdef.index(word)) + 1]
            if word == "representación":
                repreindex = textdef.index(word)
            if word == "mayor":
                mayorindex = textdef.index(word)
                demandante = (' '.join(textdef[repreindex + 1:mayorindex - 1])).title()
            if word == "empresa":
                empresaindex = textdef.index(word)
            if word == "cif" or word == "nif":
                cifindex = textdef.index(word)
                demandado = (' '.join(textdef[empresaindex + 1:cifindex - 1])).upper()
    except:
        salario = None
        demandante = None
        demandado = None
elif "recargo" in textdef:
    salario = None
    antig = None
    despid = None
    fechas = [word for word in textdef if re.match('\d{1,2}(\/|\.|\-)\d{1,2}(\/|\.|\-)\d{1,4}', word)]
    try:

        actainfraccion = fechas[1]
        resolucionrecargo = fechas[2]
        reclamacionprevia = fechas[3]
        resolucionreclamacion = fechas[4]
    except:
        actainfraccion = None
        resolucionrecargo = None
        reclamacionprevia = None
        resolucionreclamacion = None
    try:
        for word in textdef:
            if word == "representación":
                repreindex = textdef.index(word)
            if word == "cif":
                mayorindex = textdef.index(word)
                demandante = (' '.join(textdef[repreindex + 1:mayorindex - 1])).title()
            if word == "interesado":
                empresaindex = textdef.index(word)
            if word == "dni":
                cifindex = textdef.index(word)
                demandado = (' '.join(textdef[empresaindex + 1:cifindex - 1])).upper()
            if word == "incrementadas":
                recargoindex = textdef.index(word)
                recargo = textdef[recargoindex + 1]
    except:
        demandante = None
        demandado = None
        recargo = None

today = date.today()

root = tk.Tk()
root['bg']= "#ffffff"
# Crear una elección de Modelo
tk.Label(root, text="Elige modelo:", font=('Arial', 12, "bold"),
         justify=tk.CENTER,
         padx=20, background="#ffffff").grid(row=2, column=0)

# Obtener valores de alto y ancho de la ventana.
windowWidth = root.winfo_reqwidth()
windowHeight = root.winfo_reqheight()

# Obtener la mitad del ancho/alto de la pantalla y el ancho/alto de la ventana
positionRight = int(root.winfo_screenwidth() / 2 - windowWidth / 2)
positionDown = int(root.winfo_screenheight() / 2 - windowHeight / 2)

# Posiciona la ventana en el centro.
root.geometry("+{}+{}".format(positionRight, positionDown))

#datos de la ventana

root.iconbitmap("iconoaplicacionblanco.ico")
root.resizable(width=False, height=False)
root.title("Selección de Modelos de Recargo de Prestaciones")

# Crear botones para elegir modelo.
modelo = 1
def numerodelmodelo(value):
    global modelo
    modelo = value

# crear funciones de ocultar y mostrar los distintos cuadros de texto y labels

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
    cbox.grid_forget()
    cboxlabel.grid_forget()
    entry13.grid_forget()
    label13.grid_forget()
    entry14.grid_forget()
    label14.grid_forget()
    label2.grid(row=6, column=0, pady=4)
    entry2.grid(row=6, column=2, pady=4)
    label3.grid(row=7, column=0, pady=4)
    entry3.grid(row=7, column=2, pady=4)
    label4.grid(row=8, column=0, pady=4)
    entry4.grid(row=8, column=2, pady=4)
    entry5.grid(row=9, column=2, pady=4)
    label5.grid(row=9, column=0, pady=4)
    label15.grid(row=10, column=0, pady=4)
    entry15.grid(row=10, column=2, pady=4)
    label6.grid(row=11, column=0, pady=4)
    entry6.grid(row=11, column=2, pady=4)
    label7.grid(row=12, column=0, pady=4)
    entry7.grid(row=12, column=2, pady=4)
    label8.grid(row=13, column=0, pady=4)
    entry8.grid(row=13, column=2, pady=4)
    label9.grid(row=14, column=0, pady=4)
    entry9.grid(row=14, column=2, pady=4)
    label10.grid(row=15, column=0, pady=4)
    entry10.grid(row=15, column=2, pady=4)
    label11.grid(row=16, column=0, pady=4)
    entry11.grid(row=16, column=2, pady=4)
    label12.grid(row=17, column=0, pady=4)
    entry12.grid(row=17, column=2, pady=4)
    button.grid(row=20, column=1, pady=4)

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
    cbox.grid(row=5, column=2, pady=4)
    cboxlabel.grid(row=5, column=0, pady=4)
    label2.grid(row=6, column=0, pady=4)
    entry2.grid(row=6, column=2, pady=4)
    label3.grid(row=7, column=0, pady=4)
    entry3.grid(row=7, column=2, pady=4)
    label4.grid(row=8, column=0, pady=4)
    entry4.grid(row=8, column=2, pady=4)
    label5.grid(row=9, column=0, pady=4)
    entry5.grid(row=9, column=2, pady=4)
    label15.grid(row=10, column=0, pady=4)
    entry15.grid(row=10, column=2, pady=4)
    label6.grid(row=11, column=0, pady=4)
    entry6.grid(row=11, column=2, pady=4)
    label7.grid(row=12, column=0, pady=4)
    # Posicionar la caja 7 en la ventana, nombre del demandado.
    entry7.grid(row=12, column=2, pady=4)
    label8.grid(row=13, column=0, pady=4)
    # Posicionar la caja 8 en la ventana, nombre del demandado.
    entry8.grid(row=13, column=2, pady=4)
    label9.grid(row=14, column=0, pady=4)
    # Posicionar la caja 9 en la ventana, nombre del demandado.
    entry9.grid(row=14, column=2, pady=4)
    label10.grid(row=15, column=0, pady=4)
    # Posicionar la caja 10 en la ventana, nombre del demandado.
    entry10.grid(row=15, column=2, pady=4)
    label11.grid(row=16, column=0, pady=4)
    # Posicionar la caja 9 en la ventana, nombre del demandado.
    entry11.grid(row=16, column=2, pady=4)
    label12.grid(row=17, column=0, pady=4)
    # Posicionar la caja 10 en la ventana, nombre del demandado.
    entry12.grid(row=17, column=2, pady=4)
    button.grid(row=20, column=1, pady=4)

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
    cbox.grid_forget()
    cboxlabel.grid_forget()
    entry13.grid(row=16, column=2, pady=4)
    label13.grid(row=16, column=0, pady=4)
    entry14.grid(row=17, column=2, pady=4)
    label14.grid(row=17, column=0, pady=4)
    label2.grid(row=6, column=0, pady=4)
    entry2.grid(row=6, column=2, pady=4)
    label3.grid(row=7, column=0, pady=4)
    entry3.grid(row=7, column=2, pady=4)
    label4.grid(row=8, column=0, pady=4)
    entry4.grid(row=8, column=2, pady=4)
    label5.grid(row=9, column=0, pady=4)
    entry5.grid(row=9, column=2, pady=4)
    label15.grid(row=10, column=0, pady=4)
    entry15.grid(row=10, column=2, pady=4)
    label6.grid(row=11, column=0, pady=4)
    entry6.grid(row=11, column=2, pady=4)
    label7.grid(row=12, column=0, pady=4)
    entry7.grid(row=12, column=2, pady=4)
    label8.grid(row=13, column=0, pady=4)
    entry8.grid(row=13, column=2, pady=4)
    label9.grid(row=14, column=0, pady=4)
    entry9.grid(row=14, column=2, pady=4)
    label10.grid(row=15, column=0, pady=4)
    entry10.grid(row=15, column=2, pady=4)
    label11.grid(row=18, column=0, pady=4)
    entry11.grid(row=18, column=2, pady=4)
    label12.grid(row=19, column=0, pady=4)
    entry12.grid(row=19, column=2, pady=4)
    button.grid(row=20, column=1, pady=4)

def mostrarextincionretraso():
    cbox.grid_forget()
    cboxlabel.grid_forget()
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
    label2.grid(row=6, column=0, pady=4)
    entry2.grid(row=6, column=2, pady=4)
    label3.grid(row=7, column=0, pady=4)
    entry3.grid(row=7, column=2, pady=4)
    label20.grid(row=8, column=0, pady=4)
    entry20.grid(row=8, column=2, pady=4)
    label21.grid(row=9, column=0, pady=4)
    entry21.grid(row=9, column=2, pady=4)
    label22.grid(row=10, column=0, pady=4)
    entry22.grid(row=10, column=2, pady=4)
    label23.grid(row=11, column=0, pady=4)
    entry23.grid(row=11, column=2, pady=4)
    label24.grid(row=12, column=0, pady=4)
    entry24.grid(row=12, column=2, pady=4)
    opt2.grid(row=13, column=2, pady=4)
    optlabel2.grid(row=13, column=0, pady=4)
    entry9.grid(row=14, column=2, pady=4)
    label9.grid(row=14, column=0, pady=4)
    entry10.grid(row=15, column=2, pady=4)
    label10.grid(row=15, column=0, pady=4)
    label11.grid(row=17, column=0, pady=4)
    entry11.grid(row=17, column=2, pady=4)
    label12.grid(row=18, column=0, pady=4)
    entry12.grid(row=18, column=2, pady=4)
    button.grid(row=20, column=1, pady=4)

def mostrarextincionfaltapago():
    cbox.grid_forget()
    cboxlabel.grid_forget()
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
    label2.grid(row=6, column=0, pady=4)
    entry2.grid(row=6, column=2, pady=4)
    label3.grid(row=7, column=0, pady=4)
    entry3.grid(row=7, column=2, pady=4)
    label20.grid(row=8, column=0, pady=4)
    entry20.grid(row=8, column=2, pady=4)
    label21.grid(row=9, column=0, pady=4)
    entry21.grid(row=9, column=2, pady=4)
    label22.grid(row=10, column=0, pady=4)
    entry22.grid(row=10, column=2, pady=4)
    label23.grid(row=11, column=0, pady=4)
    entry23.grid(row=11, column=2, pady=4)
    label24.grid(row=12, column=0, pady=4)
    entry24.grid(row=12, column=2, pady=4)
    opt2.grid(row=13, column=2, pady=4)
    optlabel2.grid(row=13, column=0, pady=4)
    entry9.grid(row=14, column=2, pady=4)
    label9.grid(row=14, column=0, pady=4)
    entry10.grid(row=15, column=2, pady=4)
    label10.grid(row=15, column=0, pady=4)
    label27.grid(row=16, column=0, pady=4)
    entry27.grid(row=16, column=2, pady=4)
    label11.grid(row=17, column=0, pady=4)
    entry11.grid(row=17, column=2, pady=4)
    label12.grid(row=18, column=0, pady=4)
    entry12.grid(row=18, column=2)
    button.grid(row=20, column=1)

def mostrarextincionmodificacion():
    cbox.grid_forget()
    cboxlabel.grid_forget()
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
    label2.grid(row=6, column=0, pady=4)
    entry2.grid(row=6, column=2, pady=4)
    label3.grid(row=7, column=0, pady=4)
    entry3.grid(row=7, column=2, pady=4)
    label20.grid(row=8, column=0, pady=4)
    entry20.grid(row=8, column=2, pady=4)
    label21.grid(row=9, column=0, pady=4)
    entry21.grid(row=9, column=2, pady=4)
    label22.grid(row=10, column=0, pady=4)
    entry22.grid(row=10, column=2, pady=4)
    label23.grid(row=11, column=0, pady=4)
    entry23.grid(row=11, column=2, pady=4)
    label24.grid(row=12, column=0, pady=4)
    entry24.grid(row=12, column=2, pady=4)
    opt2.grid(row=13, column=2, pady=4)
    optlabel2.grid(row=13, column=0, pady=4)
    entry9.grid(row=14, column=2, pady=4)
    label9.grid(row=14, column=0, pady=4)
    entry10.grid(row=15, column=2, pady=4)
    label10.grid(row=15, column=0, pady=4)
    label11.grid(row=17, column=0, pady=4)
    # Posicionar la caja 9 en la ventana, nombre del demandado.
    entry11.grid(row=17, column=2, pady=4)
    label12.grid(row=18, column=0, pady=4)
    # Posicionar la caja 10 en la ventana, nombre del demandado.
    entry12.grid(row=18, column=2, pady=4)
    button.grid(row=20, column=1, pady=4)

# crear botones de seleccion de los modelos

botoncondena = tk.Button(root, text="Condenatoria", command=lambda: [mostrarrecargocondena(), numerodelmodelo(1)],
                         width=20, pady=2, padx=10)
botonabsolucion = tk.Button(root, text="Estimatoria", command=lambda: [mostrarrecargoestim(), numerodelmodelo(2)],
                            width=20, pady=2, padx=10)
botonabsolucion2 = tk.Button(root, text="Estimatoria Parcial",
                             command=lambda: [mostrarrecargoestimparcial(), numerodelmodelo(3)], width=20, pady=2,
                             padx=10)

botonretraso = tk.Button(root, text="Retraso en el pago", width="20",
                         command=lambda: [mostrarextincionretraso(), numerodelmodelo(4)], pady=2, padx=10)
botonfaltapago = tk.Button(root, text="Falta de pago", width="20",
                           command=lambda: [mostrarextincionfaltapago(), numerodelmodelo(5)], pady=2, padx=10)
botonmodificacion = tk.Button(root, text="Modificacion Condiciones", width="20",
                              command=lambda: [mostrarextincionmodificacion(), numerodelmodelo(6)], pady=2, padx=10)


# crear funciones de ocultar y mostrar los botones de seleccion de los modelos

def mostrarrecargo():
    botoncondena.grid(row=1, column=2)
    botonabsolucion.grid(row=2, column=2)
    botonabsolucion2.grid(row=3, column=2)
    botonretraso.grid_forget()
    botonfaltapago.grid_forget()
    botonmodificacion.grid_forget()


def mostrarexincion():
    botonretraso.grid(row=1, column=2)
    botonfaltapago.grid(row=2, column=2)
    botonmodificacion.grid(row=3, column=2)
    botoncondena.grid_forget()
    botonabsolucion.grid_forget()
    botonabsolucion2.grid_forget()


# crear botones principales

botonrecargo = tk.Button(root, text="Recargo de Prestaciones", width="20", command=lambda: [mostrarrecargo()]).grid(
    row=1, column=1, pady=4, padx=5)
botonextinción = tk.Button(root, text="Extinción del Contrato", width="20", command=lambda: [mostrarexincion()]).grid(
    row=3, column=1, pady=4, padx=5)

# aqui ponemos combobox de fuerza mayor o caso fortuito

mayorofortuito = None


def check_cbox(event):
    global mayorofortuito
    if cbox.get() == 'Fuerza Mayor':
        mayorofortuito = cbox.get()  # this will assign the variable c the value of cbox
    if cbox.get() == 'Fortuito':
        mayorofortuito = cbox.get()


cbox = ttk.Combobox(root, values=["Fuerza Mayor", "Fortuito"], width=17)
cboxlabel = ttk.Label(root, text="Fundamento de estimatoria", background="#ffffff")
cbox.bind("<<ComboboxSelected>>", check_cbox)

# Espacio en blanco.
labelx = ttk.Label(root, text=" ", background="#ffffff").grid(row=4, column=0)

# Crear caja de texto 2, nombre del actor.
entry_var2 = tk.StringVar()
entry2 = ttk.Entry(root, textvariable=entry_var2, justify=tk.CENTER)
label2 = ttk.Label(root, text="Nombre del actor", background="#ffffff")
if demandante:
    entry2.insert(0, demandante)

# Crear caja de texto 3, nombre del demandado.
entry_var3 = tk.StringVar()
entry3 = ttk.Entry(root, textvariable=entry_var3, justify=tk.CENTER)
label3 = ttk.Label(root, text="Nombre del demandado", background="#ffffff")
if demandado:
    entry3.insert(0, demandado)

# Crear caja de fecha 4, fecha de acta.
entry_var4 = tk.StringVar()
entry4 = DateEntry(root, textvariable=entry_var4, date_pattern='dd/mm/Y', width=17, background='gray',
                   foreground='white', borderwidth=1)
label4 = ttk.Label(root, text="Fecha del Acta de Infracción", background="#ffffff")
if actainfraccion:
    entry4.set_date(actainfraccion)

# Crear caja de fecha 5, fecha resolucion impugnada.
entry_var5 = tk.StringVar()
entry5 = DateEntry(root, textvariable=entry_var5, date_pattern='dd/mm/Y', width=17, background='gray',
                   foreground='white', borderwidth=1)
label5 = ttk.Label(root, text="Fecha de la Resolución impugnada", background="#ffffff")
if resolucionrecargo:
    entry5.set_date(resolucionrecargo)

# Crear caja de texto 15, porcentaje impuesto por la Inspección.
entry_var15 = tk.StringVar()
entry15 = ttk.Entry(root, textvariable=entry_var15, justify=tk.CENTER)
label15 = ttk.Label(root, text="Porcentaje impuesto por la Inspección", background="#ffffff")
if recargo:
    entry15.insert(0, recargo)

# Crear caja de texto 6, artículos infringidos.
entry_var6 = tk.StringVar()
entry6 = ttk.Entry(root, textvariable=entry_var6, justify=tk.CENTER)
label6 = ttk.Label(root, text="Artículos infringidos según acta de Infracción", background="#ffffff")

# Crear caja de fecha 7, fecha de la reclamación previa.
entry_var7 = tk.StringVar()
entry7 = DateEntry(root, textvariable=entry_var7, date_pattern='dd/mm/Y', width=17, background='gray',
                   foreground='white', borderwidth=1)
label7 = ttk.Label(root, text="Fecha de la Reclamación Previa", background="#ffffff")
if reclamacionprevia:
    entry7.set_date(reclamacionprevia)

# Crear caja de fecha 8, fecha de la resolución de la reclamación previa.
entry_var8 = tk.StringVar()
entry8 = DateEntry(root, textvariable=entry_var8, date_pattern='dd/mm/Y', width=17, background='gray',
                   foreground='white', borderwidth=1)
label8 = ttk.Label(root, text="Fecha de Resolución de la Reclamación Previa", background="#ffffff")
if resolucionreclamacion:
    entry8.set_date(resolucionreclamacion)

# Crear caja de texto 9, pruena practicada.
entry_var9 = tk.StringVar()
entry9 = ttk.Entry(root, textvariable=entry_var9, justify=tk.CENTER)
label9 = ttk.Label(root, text="Prueba practicada aparte de la documental", background="#ffffff")

# Crear caja de texto 10, en el presente caso.
entry_var10 = tk.StringVar()
entry10 = ttk.Entry(root, textvariable=entry_var10, justify=tk.CENTER)
label10 = ttk.Label(root, text="En el presente caso...", background="#ffffff")

# Crear caja de texto 10, nombre del demandado.
entry_var13 = tk.StringVar()
entry13 = ttk.Entry(root, textvariable=entry_var13, justify=tk.CENTER)
label13 = ttk.Label(root, text="Porcentaje que se estima adecuado", background="#ffffff")
# Posicionar la caja 10 en la ventana, nombre del demandado.

# Crear caja de texto 10, nombre del demandado.
entry_var14 = tk.StringVar()
entry14 = ttk.Entry(root, textvariable=entry_var14, justify=tk.CENTER)
label14 = ttk.Label(root, text="Justificación del porcentaje que se impone", background="#ffffff")
# Posicionar la caja 10 en la ventana, nombre del demandado.

# Crear caja de texto 11, numero del procedimiento.
entry_var11 = tk.StringVar()
entry11 = ttk.Entry(root, textvariable=entry_var11, justify=tk.CENTER)
label11 = ttk.Label(root, text="Número del procedimiento", background="#ffffff")

# Crear caja de texto 12, año del procedimiento.
entry_var12 = tk.StringVar()
entry12 = ttk.Entry(root, textvariable=entry_var12, justify=tk.CENTER)
label12 = ttk.Label(root, text="Año", background="#ffffff")

# Crear caja de fecha 4, fecha de acta.
entry_var20 = tk.StringVar()
entry20 = ttk.Entry(root, textvariable=entry_var20, justify=tk.CENTER)
label20 = ttk.Label(root, text="Actividad de la Empresa", background="#ffffff")

# Crear caja de fecha 5, fecha resolucion impugnada.
entry_var21 = tk.StringVar()
entry21 = DateEntry(root, textvariable=entry_var21, date_pattern='dd/mm/Y', width=17, background='gray',
                    foreground='white', borderwidth=1)
label21 = ttk.Label(root, text="Antigüedad", background="#ffffff")
if antig:
    entry21.set_date(antig)

# Crear caja de texto 15, porcentaje impuesto por la Inspección.
entry_var22 = tk.StringVar()
entry22 = ttk.Entry(root, textvariable=entry_var22, justify=tk.CENTER)
label22 = ttk.Label(root, text="Categoría Profesional", background="#ffffff")

# Crear caja de texto 6, artículos infringidos.
entry_var23 = tk.StringVar()
entry23 = ttk.Entry(root, textvariable=entry_var23, justify=tk.CENTER)
label23 = ttk.Label(root, text="Salario diario", background="#ffffff")
if salario:
    entry23.insert(0, salario)

# Crear caja de fecha 7, fecha de la reclamación previa.
entry_var24 = tk.StringVar()
entry24 = DateEntry(root, textvariable=entry_var24, date_pattern='dd/mm/Y', width=17, background='gray',
                    foreground='white', borderwidth=1)
label24 = ttk.Label(root, text="Fecha celebración SEMAC", background="#ffffff")

avenenciaoefecto = None


def check_cbox(event):
    global avenenciaoefecto
    if opt2.get() == 'Sin avenencia':
        avenenciaoefecto = opt2.get()  # this will assign the variable c the value of cbox
    if opt2.get() == 'Sin efecto':
        avenenciaoefecto = opt2.get()


opt2 = ttk.Combobox(root, values=["Sin avenencia", "Sin efecto"], width=17)
optlabel2 = ttk.Label(root, text="Resultado del SEMAC", background="#ffffff")
opt2.bind("<<ComboboxSelected>>", check_cbox)

# Crear caja de texto 10, nombre del demandado.
entry_var27 = tk.StringVar()
entry27 = ttk.Entry(root, textvariable=entry_var27, justify=tk.CENTER)
label27 = ttk.Label(root, text="Desde cuando no cobra", background="#ffffff")
# Posicionar la caja 10 en la ventana, nombre del demandado.

# Posicionar boton de fin
button = ttk.Button(root, text="Crear Modelo", command=root.destroy)

# mainloop de la ventana

root.mainloop()

# selección del modelo a usar en función de los botones pulsados

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

# creacion de variables para el nombre del archivo

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

# uso del template elegido

document = MailMerge(template)


# creación de función para ver si la opción de indemnización ha sido seleccionada y rellenada, en caso contrario no hace nada

def indemnizacionsiono():
    if entry_var23.get():
        return indemnizacion_despido_prueba(str(entry_var21.get()), str(entry_var23.get()), str(today))
    else:
        valor0 = 0
        return valor0


# crear funcion de mayusculas y minusculas actor

def actorcapitalizacion():
    if template == "Prueba1.docx":
        return str(entry_var2.get()).upper()
    elif template == "Prueba5.docx":
        return str(entry_var2.get()).upper()
    elif template == "Extincion retraso pago.docx":
        return str(entry_var2.get()).title()
    elif template == "Extincion falta de pago.docx":
        return str(entry_var2.get()).title()
    elif template == "Extincion modificacion sustancial.docx":
        return str(entry_var2.get()).title()


# crear funcion de mayusculas y minusculas demandado

def demandadocapitalizacion():
    if template == "Prueba1.docx":
        return str(entry_var3.get()).title()
    elif template == "Prueba5.docx":
        return str(entry_var3.get()).title()
    elif template == "Extincion retraso pago.docx":
        return str(entry_var3.get()).upper()
    elif template == "Extincion falta de pago.docx":
        return str(entry_var3.get()).upper()
    elif template == "Extincion modificacion sustancial.docx":
        return str(entry_var3.get()).upper()


# crear funcion tipo de extincion

tipoextin = None

if template == "Extincion retraso pago.docx":
    tipoextin = "retraso pago"
elif template == "Extincion falta de pago.docx":
    tipoextin = "falta de pago"
elif template == "Extincion modificacion sustancial.docx":
    tipoextin = "modificacion sustancial"
else:
    tipoextin = ""

# meter todos los datos en el modelo

document.merge(
    Actor=actorcapitalizacion(),
    Demandado=demandadocapitalizacion(),
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
    indemnizacionextincion=indemnizacionsiono(),
    categoría=entry_var22.get(),
    salariodia=entry_var23.get(),
    actividadempresa=entry_var20.get())

if mayorofortuito == "Fuerza Mayor":
    document.merge(fortuitomayor1="por la fuerza mayor operada al tiempo del accidente",
                   fortuitomayor2="sino que el hecho acaecido responde a un caso de fuerza mayor, esto es, un evento extraño al círculo o ámbito de la actividad del trabajador y de la empresa, que rompe el nexo de causalidad")
else:
    document.merge(fortuitomayor1="por su acaecimiento fortuito",
                   fortuitomayor2="sino que el hecho acaecido responde a un caso fortuito, esto es, un hecho independiente de la voluntad del deudor de seguridad (la empresa), imprevisible e inevitable")

if avenenciaoefecto == "Sin avenencia":
    document.merge(avenenciaoefecto="sin avenencia")
else:
    document.merge(avenenciaoefecto="sin efecto")

# creación del archivo de texto

document.write(materia + " " + tipoextin + " " + entry_var11.get() + "-" + entry_var12.get() + " " + sentido + '.docx')
