from __future__ import print_function
import tkinter as tk
from tkinter import ttk
from mailmerge import MailMerge
from tkcalendar import DateEntry
from datetime import date
from Buscador import indemnizacion_despido_prueba
from tkinter import *
from tkinter import filedialog
from FuncionesGestor import *

#fecha de hoy
today = date.today()

#ventana de inicio

inicio = Tk()
inicio.geometry("400x400")

# Obtener valores de alto y ancho de la ventana.
windowWidth = inicio.winfo_reqwidth()
windowHeight = inicio.winfo_reqheight()

# Obtener la mitad del ancho/alto de la pantalla y el ancho/alto de la ventana
positionRight = int(inicio.winfo_screenwidth() / 2 - windowWidth / 2)
positionDown = int(inicio.winfo_screenheight() / 2 - windowHeight / 2)

# Posiciona la ventana en el centro.
inicio.geometry("+{}+{}".format(positionRight, positionDown))

#imagen de fondo
C = Canvas(inicio, bg="blue", height=250, width=300)
filename = PhotoImage(file = "pantallainicio.png")
background_label = Label(inicio, image=filename)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

#datos de la ventana
inicio['bg']= "#ffffff"
inicio.iconbitmap("iconoaplicacionblanco.ico")
inicio.resizable(width=False, height=False)
inicio.title("Selección de Demanda")

# sacar una variable global con la dirección del documento y una función para selecionarlo

def seleccionarchivo():
    global document
    document = filedialog.askopenfilename(initialdir = "/",title = "Selecciona el PDF de la demanda",filetypes = (("pdf files","*.pdf"),("all files","*.*")))

# boton de cargar

botoncarga = tk.Button(inicio, text="Selecciona la demanda", width="20", bg="#ffffff", command=lambda: [seleccionarchivo(), inicio.destroy()]).place(x=120, y=346)
botoncierre = tk.Button(inicio, text="Modelo vacío", width="20", bg="#ffffff", command=lambda: inicio.destroy()).place(x=120, y=373)

inicio.mainloop()

# cargar el pdf y convertirlo en texto tratable

try:
    pdfname = str(document)
    documento = abrirpdf(pdfname)
    textdef = preparar_texto(documento)

    #distinguir la extraccion de datos según se trate de un modelo de extinción o de recargo

    if "extinci" in textdef:
        antig = encontrarfechaantiguedad(textdef)
        salario = encontrarsalario(textdef)
        demandante = encontrardemandante(textdef)
        demandado = encontrardemandado(textdef)
        categoria = str(encontrarcategoria(textdef)).title()
        actividad = encontraractividad(textdef)
        actainfraccion = None
        resolucionrecargo = None
        reclamacionprevia = None
        resolucionreclamacion = None
        recargo = None
    elif "recargo" in textdef:
        actainfraccion = encontrarfechacercadeunapalabra(textdef, "acta")
        resolucionrecargo = encontrarfechacercadeunapalabra(textdef, "resolución")
        reclamacionprevia = encontrarfechacercadeunapalabra(textdef, "reclamación")
        resolucionreclamacion = encontrarfechacercadeunapalabra(textdef, "provincial")
        demandante = encontrardemandanterecargo(textdef)
        demandado = encontrardemandadorecargo(textdef)
        recargo = encontrarporcentajerecargo(textdef)
        salario = None
        categoria = None
        actividad = None
except:
    salario = None
    categoria = None
    actividad = None
    despid = None
    demandante = None
    demandado = None
    recargo = None
    resolucionreclamacion = None
    reclamacionprevia = None
    resolucionrecargo = None
    actainfraccion = None
    pass

#ventana principal de selección de modelos

principal = tk.Tk()
principal['bg']= "#ffffff"

# Texto de elegir modelo

tk.Label(principal, text="Elige modelo:", font=('Arial', 12, "bold"),
         justify=tk.CENTER,
         padx=20, background="#ffffff").grid(row=2, column=0)

# Obtener valores de alto y ancho de la ventana.
windowWidth = principal.winfo_reqwidth()
windowHeight = principal.winfo_reqheight()

# Obtener la mitad del ancho/alto de la pantalla y el ancho/alto de la ventana
positionRight = int(principal.winfo_screenwidth() / 2 - windowWidth / 2)
positionDown = int(principal.winfo_screenheight() / 2 - windowHeight / 2)

# Posiciona la ventana en el centro.
principal.geometry("+{}+{}".format(positionRight, positionDown))

# datos de la ventana

principal.iconbitmap("iconoaplicacionblanco.ico")
principal.resizable(width=False, height=False)
principal.title("Selección de Modelos de Recargo de Prestaciones")

# Crear botones para elegir modelo.

def numerodelmodelo(value):
    global modelo
    modelo = value

# crear funciones de ocultar y mostrar los distintos cuadros de texto y labels

def mostrarrecargocondena():
    label_desdecuandonocobra.grid_forget()
    entry_desdecuandonocobra.grid_forget()
    label_actor.grid_forget()
    entry_actor.grid_forget()
    label_demandado.grid_forget()
    entry_demandado.grid_forget()
    label_activempres.grid_forget()
    entry_activempres.grid_forget()
    label_antig.grid_forget()
    entry_antig.grid_forget()
    label_categoria.grid_forget()
    entry_categoria.grid_forget()
    label_salario.grid_forget()
    entry_salario.grid_forget()
    label_fechasemac.grid_forget()
    entry_fechasemac.grid_forget()
    resultsemac.grid_forget()
    label_resultsemac.grid_forget()
    entry_prueba.grid_forget()
    label_prueba.grid_forget()
    entry_presentecaso.grid_forget()
    label_presentecaso.grid_forget()
    label_desdecuandonocobra.grid_forget()
    entry_desdecuandonocobra.grid_forget()
    cboxmayorofortuito.grid_forget()
    cboxmayorofortuitolabel.grid_forget()
    entry_porcentajeadecuado.grid_forget()
    label_porcentajeadecuado.grid_forget()
    entry_justifporcentaje.grid_forget()
    label_justifporcentaje.grid_forget()
    label_actor.grid(row=6, column=0, pady=4)
    entry_actor.grid(row=6, column=2, pady=4)
    label_demandado.grid(row=7, column=0, pady=4)
    entry_demandado.grid(row=7, column=2, pady=4)
    label_actainfraccion.grid(row=8, column=0, pady=4)
    entry_actainfraccion.grid(row=8, column=2, pady=4)
    entry_resolucionrecargo.grid(row=9, column=2, pady=4)
    label_resolucionrecargo.grid(row=9, column=0, pady=4)
    label_porcentajerecargo.grid(row=10, column=0, pady=4)
    entry_porcentajerecargo.grid(row=10, column=2, pady=4)
    label_articulosinfringidos.grid(row=11, column=0, pady=4)
    entry_articulosinfringidos.grid(row=11, column=2, pady=4)
    label_reclamprevia.grid(row=12, column=0, pady=4)
    entry_reclamprevia.grid(row=12, column=2, pady=4)
    label_resolreclam.grid(row=13, column=0, pady=4)
    entry_resolreclam.grid(row=13, column=2, pady=4)
    label_prueba.grid(row=14, column=0, pady=4)
    entry_prueba.grid(row=14, column=2, pady=4)
    label_presentecaso.grid(row=15, column=0, pady=4)
    entry_presentecaso.grid(row=15, column=2, pady=4)
    label_numero.grid(row=16, column=0, pady=4)
    entry_numero.grid(row=16, column=2, pady=4)
    label_year.grid(row=17, column=0, pady=4)
    entry_year.grid(row=17, column=2, pady=4)
    button.grid(row=20, column=1, pady=4)

def mostrarrecargoestim():
    label_desdecuandonocobra.grid_forget()
    entry_desdecuandonocobra.grid_forget()
    label_actor.grid_forget()
    entry_actor.grid_forget()
    label_demandado.grid_forget()
    entry_demandado.grid_forget()
    label_activempres.grid_forget()
    entry_activempres.grid_forget()
    label_antig.grid_forget()
    entry_antig.grid_forget()
    label_categoria.grid_forget()
    entry_categoria.grid_forget()
    label_salario.grid_forget()
    entry_salario.grid_forget()
    label_fechasemac.grid_forget()
    entry_fechasemac.grid_forget()
    resultsemac.grid_forget()
    label_resultsemac.grid_forget()
    entry_prueba.grid_forget()
    label_prueba.grid_forget()
    entry_presentecaso.grid_forget()
    label_presentecaso.grid_forget()
    label_desdecuandonocobra.grid_forget()
    entry_desdecuandonocobra.grid_forget()
    entry_porcentajeadecuado.grid_forget()
    label_porcentajeadecuado.grid_forget()
    entry_justifporcentaje.grid_forget()
    label_justifporcentaje.grid_forget()
    cboxmayorofortuito.grid(row=5, column=2, pady=4)
    cboxmayorofortuitolabel.grid(row=5, column=0, pady=4)
    label_actor.grid(row=6, column=0, pady=4)
    entry_actor.grid(row=6, column=2, pady=4)
    label_demandado.grid(row=7, column=0, pady=4)
    entry_demandado.grid(row=7, column=2, pady=4)
    label_actainfraccion.grid(row=8, column=0, pady=4)
    entry_actainfraccion.grid(row=8, column=2, pady=4)
    label_resolucionrecargo.grid(row=9, column=0, pady=4)
    entry_resolucionrecargo.grid(row=9, column=2, pady=4)
    label_porcentajerecargo.grid(row=10, column=0, pady=4)
    entry_porcentajerecargo.grid(row=10, column=2, pady=4)
    label_articulosinfringidos.grid(row=11, column=0, pady=4)
    entry_articulosinfringidos.grid(row=11, column=2, pady=4)
    label_reclamprevia.grid(row=12, column=0, pady=4)
    entry_reclamprevia.grid(row=12, column=2, pady=4)
    label_resolreclam.grid(row=13, column=0, pady=4)
    entry_resolreclam.grid(row=13, column=2, pady=4)
    label_prueba.grid(row=14, column=0, pady=4)
    entry_prueba.grid(row=14, column=2, pady=4)
    label_presentecaso.grid(row=15, column=0, pady=4)
    entry_presentecaso.grid(row=15, column=2, pady=4)
    label_numero.grid(row=16, column=0, pady=4)
    entry_numero.grid(row=16, column=2, pady=4)
    label_year.grid(row=17, column=0, pady=4)
    entry_year.grid(row=17, column=2, pady=4)
    button.grid(row=20, column=1, pady=4)

def mostrarrecargoestimparcial():
    label_desdecuandonocobra.grid_forget()
    entry_desdecuandonocobra.grid_forget()
    label_actor.grid_forget()
    entry_actor.grid_forget()
    label_demandado.grid_forget()
    entry_demandado.grid_forget()
    label_activempres.grid_forget()
    entry_activempres.grid_forget()
    label_antig.grid_forget()
    entry_antig.grid_forget()
    label_categoria.grid_forget()
    entry_categoria.grid_forget()
    label_salario.grid_forget()
    entry_salario.grid_forget()
    label_fechasemac.grid_forget()
    entry_fechasemac.grid_forget()
    resultsemac.grid_forget()
    label_resultsemac.grid_forget()
    entry_prueba.grid_forget()
    label_prueba.grid_forget()
    entry_presentecaso.grid_forget()
    label_presentecaso.grid_forget()
    label_desdecuandonocobra.grid_forget()
    entry_desdecuandonocobra.grid_forget()
    cboxmayorofortuito.grid_forget()
    cboxmayorofortuitolabel.grid_forget()
    entry_porcentajeadecuado.grid(row=16, column=2, pady=4)
    label_porcentajeadecuado.grid(row=16, column=0, pady=4)
    entry_justifporcentaje.grid(row=17, column=2, pady=4)
    label_justifporcentaje.grid(row=17, column=0, pady=4)
    label_actor.grid(row=6, column=0, pady=4)
    entry_actor.grid(row=6, column=2, pady=4)
    label_demandado.grid(row=7, column=0, pady=4)
    entry_demandado.grid(row=7, column=2, pady=4)
    label_actainfraccion.grid(row=8, column=0, pady=4)
    entry_actainfraccion.grid(row=8, column=2, pady=4)
    label_resolucionrecargo.grid(row=9, column=0, pady=4)
    entry_resolucionrecargo.grid(row=9, column=2, pady=4)
    label_porcentajerecargo.grid(row=10, column=0, pady=4)
    entry_porcentajerecargo.grid(row=10, column=2, pady=4)
    label_articulosinfringidos.grid(row=11, column=0, pady=4)
    entry_articulosinfringidos.grid(row=11, column=2, pady=4)
    label_reclamprevia.grid(row=12, column=0, pady=4)
    entry_reclamprevia.grid(row=12, column=2, pady=4)
    label_resolreclam.grid(row=13, column=0, pady=4)
    entry_resolreclam.grid(row=13, column=2, pady=4)
    label_prueba.grid(row=14, column=0, pady=4)
    entry_prueba.grid(row=14, column=2, pady=4)
    label_presentecaso.grid(row=15, column=0, pady=4)
    entry_presentecaso.grid(row=15, column=2, pady=4)
    label_numero.grid(row=18, column=0, pady=4)
    entry_numero.grid(row=18, column=2, pady=4)
    label_year.grid(row=19, column=0, pady=4)
    entry_year.grid(row=19, column=2, pady=4)
    button.grid(row=20, column=1, pady=4)

def mostrarextincionretraso():
    cboxmayorofortuito.grid_forget()
    cboxmayorofortuitolabel.grid_forget()
    entry_porcentajeadecuado.grid_forget()
    label_porcentajeadecuado.grid_forget()
    entry_justifporcentaje.grid_forget()
    label_justifporcentaje.grid_forget()
    entry_actainfraccion.grid_forget()
    label_actainfraccion.grid_forget()
    entry_resolucionrecargo.grid_forget()
    label_resolucionrecargo.grid_forget()
    entry_porcentajerecargo.grid_forget()
    label_porcentajerecargo.grid_forget()
    entry_articulosinfringidos.grid_forget()
    label_articulosinfringidos.grid_forget()
    entry_reclamprevia.grid_forget()
    label_reclamprevia.grid_forget()
    entry_resolreclam.grid_forget()
    label_resolreclam.grid_forget()
    label_desdecuandonocobra.grid_forget()
    entry_desdecuandonocobra.grid_forget()
    label_actor.grid(row=6, column=0, pady=4)
    entry_actor.grid(row=6, column=2, pady=4)
    label_demandado.grid(row=7, column=0, pady=4)
    entry_demandado.grid(row=7, column=2, pady=4)
    label_activempres.grid(row=8, column=0, pady=4)
    entry_activempres.grid(row=8, column=2, pady=4)
    label_antig.grid(row=9, column=0, pady=4)
    entry_antig.grid(row=9, column=2, pady=4)
    label_categoria.grid(row=10, column=0, pady=4)
    entry_categoria.grid(row=10, column=2, pady=4)
    label_salario.grid(row=11, column=0, pady=4)
    entry_salario.grid(row=11, column=2, pady=4)
    label_fechasemac.grid(row=12, column=0, pady=4)
    entry_fechasemac.grid(row=12, column=2, pady=4)
    resultsemac.grid(row=13, column=2, pady=4)
    label_resultsemac.grid(row=13, column=0, pady=4)
    entry_prueba.grid(row=14, column=2, pady=4)
    label_prueba.grid(row=14, column=0, pady=4)
    entry_presentecaso.grid(row=15, column=2, pady=4)
    label_presentecaso.grid(row=15, column=0, pady=4)
    label_numero.grid(row=17, column=0, pady=4)
    entry_numero.grid(row=17, column=2, pady=4)
    label_year.grid(row=18, column=0, pady=4)
    entry_year.grid(row=18, column=2, pady=4)
    button.grid(row=20, column=1, pady=4)

def mostrarextincionfaltapago():
    cboxmayorofortuito.grid_forget()
    cboxmayorofortuitolabel.grid_forget()
    entry_porcentajeadecuado.grid_forget()
    label_porcentajeadecuado.grid_forget()
    entry_justifporcentaje.grid_forget()
    label_justifporcentaje.grid_forget()
    entry_actainfraccion.grid_forget()
    label_actainfraccion.grid_forget()
    entry_resolucionrecargo.grid_forget()
    label_resolucionrecargo.grid_forget()
    entry_porcentajerecargo.grid_forget()
    label_porcentajerecargo.grid_forget()
    entry_articulosinfringidos.grid_forget()
    label_articulosinfringidos.grid_forget()
    entry_reclamprevia.grid_forget()
    label_reclamprevia.grid_forget()
    entry_resolreclam.grid_forget()
    label_resolreclam.grid_forget()
    label_actor.grid(row=6, column=0, pady=4)
    entry_actor.grid(row=6, column=2, pady=4)
    label_demandado.grid(row=7, column=0, pady=4)
    entry_demandado.grid(row=7, column=2, pady=4)
    label_activempres.grid(row=8, column=0, pady=4)
    entry_activempres.grid(row=8, column=2, pady=4)
    label_antig.grid(row=9, column=0, pady=4)
    entry_antig.grid(row=9, column=2, pady=4)
    label_categoria.grid(row=10, column=0, pady=4)
    entry_categoria.grid(row=10, column=2, pady=4)
    label_salario.grid(row=11, column=0, pady=4)
    entry_salario.grid(row=11, column=2, pady=4)
    label_fechasemac.grid(row=12, column=0, pady=4)
    entry_fechasemac.grid(row=12, column=2, pady=4)
    resultsemac.grid(row=13, column=2, pady=4)
    label_resultsemac.grid(row=13, column=0, pady=4)
    entry_prueba.grid(row=14, column=2, pady=4)
    label_prueba.grid(row=14, column=0, pady=4)
    entry_presentecaso.grid(row=15, column=2, pady=4)
    label_presentecaso.grid(row=15, column=0, pady=4)
    label_desdecuandonocobra.grid(row=16, column=0, pady=4)
    entry_desdecuandonocobra.grid(row=16, column=2, pady=4)
    label_numero.grid(row=17, column=0, pady=4)
    entry_numero.grid(row=17, column=2, pady=4)
    label_year.grid(row=18, column=0, pady=4)
    entry_year.grid(row=18, column=2)
    button.grid(row=20, column=1)

def mostrarextincionmodificacion():
    cboxmayorofortuito.grid_forget()
    cboxmayorofortuitolabel.grid_forget()
    entry_porcentajeadecuado.grid_forget()
    label_porcentajeadecuado.grid_forget()
    entry_justifporcentaje.grid_forget()
    label_justifporcentaje.grid_forget()
    entry_actainfraccion.grid_forget()
    label_actainfraccion.grid_forget()
    entry_resolucionrecargo.grid_forget()
    label_resolucionrecargo.grid_forget()
    entry_porcentajerecargo.grid_forget()
    label_porcentajerecargo.grid_forget()
    entry_articulosinfringidos.grid_forget()
    label_articulosinfringidos.grid_forget()
    entry_reclamprevia.grid_forget()
    label_reclamprevia.grid_forget()
    entry_resolreclam.grid_forget()
    label_resolreclam.grid_forget()
    label_desdecuandonocobra.grid_forget()
    entry_desdecuandonocobra.grid_forget()
    label_actor.grid(row=6, column=0, pady=4)
    entry_actor.grid(row=6, column=2, pady=4)
    label_demandado.grid(row=7, column=0, pady=4)
    entry_demandado.grid(row=7, column=2, pady=4)
    label_activempres.grid(row=8, column=0, pady=4)
    entry_activempres.grid(row=8, column=2, pady=4)
    label_antig.grid(row=9, column=0, pady=4)
    entry_antig.grid(row=9, column=2, pady=4)
    label_categoria.grid(row=10, column=0, pady=4)
    entry_categoria.grid(row=10, column=2, pady=4)
    label_salario.grid(row=11, column=0, pady=4)
    entry_salario.grid(row=11, column=2, pady=4)
    label_fechasemac.grid(row=12, column=0, pady=4)
    entry_fechasemac.grid(row=12, column=2, pady=4)
    resultsemac.grid(row=13, column=2, pady=4)
    label_resultsemac.grid(row=13, column=0, pady=4)
    entry_prueba.grid(row=14, column=2, pady=4)
    label_prueba.grid(row=14, column=0, pady=4)
    entry_presentecaso.grid(row=15, column=2, pady=4)
    label_presentecaso.grid(row=15, column=0, pady=4)
    label_numero.grid(row=17, column=0, pady=4)
    # Posicionar la caja 9 en la ventana, nombre del demandado.
    entry_numero.grid(row=17, column=2, pady=4)
    label_year.grid(row=18, column=0, pady=4)
    # Posicionar la caja 10 en la ventana, nombre del demandado.
    entry_year.grid(row=18, column=2, pady=4)
    button.grid(row=20, column=1, pady=4)

# crear botones de seleccion de los modelos
# boton de recargo

botoncondena = tk.Button(principal, text="Condenatoria", command=lambda: [mostrarrecargocondena(), numerodelmodelo(1)],
                         width=20, pady=2, padx=10)
botonabsolucion = tk.Button(principal, text="Estimatoria", command=lambda: [mostrarrecargoestim(), numerodelmodelo(2)],
                            width=20, pady=2, padx=10)
botonabsolucion2 = tk.Button(principal, text="Estimatoria Parcial",
                             command=lambda: [mostrarrecargoestimparcial(), numerodelmodelo(3)], width=20, pady=2,
                             padx=10)

# boton de extinción

botonretraso = tk.Button(principal, text="Retraso en el pago", width="20",
                         command=lambda: [mostrarextincionretraso(), numerodelmodelo(4)], pady=2, padx=10)
botonfaltapago = tk.Button(principal, text="Falta de pago", width="20",
                           command=lambda: [mostrarextincionfaltapago(), numerodelmodelo(5)], pady=2, padx=10)
botonmodificacion = tk.Button(principal, text="Modificacion Condiciones", width="20",
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

botonrecargo = tk.Button(principal, text="Recargo de Prestaciones", width="20", command=lambda: [mostrarrecargo()]).grid(
    row=1, column=1, pady=4, padx=5)
botonextinción = tk.Button(principal, text="Extinción del Contrato", width="20", command=lambda: [mostrarexincion()]).grid(
    row=3, column=1, pady=4, padx=5)

# aqui ponemos combobox de fuerza mayor o caso fortuito

mayorofortuito = None
def check_cboxmayorofortuito(event):
    global mayorofortuito
    if cboxmayorofortuito.get() == 'Fuerza Mayor':
        mayorofortuito = cboxmayorofortuito.get()
    if cboxmayorofortuito.get() == 'Fortuito':
        mayorofortuito = cboxmayorofortuito.get()

cboxmayorofortuito = ttk.Combobox(principal, values=["Fuerza Mayor", "Fortuito"], width=17)
cboxmayorofortuitolabel = ttk.Label(principal, text="Fundamento de estimatoria", background="#ffffff")
cboxmayorofortuito.bind("<<ComboboxSelected>>", check_cboxmayorofortuito)

# Espacio en blanco.
labelx = ttk.Label(principal, text=" ", background="#ffffff").grid(row=4, column=0)

# Crear caja de texto 2, nombre del actor.
entry_var_actor= tk.StringVar()
entry_actor = ttk.Entry(principal, textvariable=entry_var_actor, justify=tk.CENTER)
label_actor = ttk.Label(principal, text="Nombre del actor", background="#ffffff")
if demandante:
    entry_actor.insert(0, demandante)

# Crear caja de texto 3, nombre del demandado.
entry_var_demandado = tk.StringVar()
entry_demandado = ttk.Entry(principal, textvariable=entry_var_demandado, justify=tk.CENTER)
label_demandado = ttk.Label(principal, text="Nombre del demandado", background="#ffffff")
if demandado:
    entry_demandado.insert(0, demandado)

# Crear caja de fecha 4, fecha de acta.
entry_var_actainfraccion = tk.StringVar()
entry_actainfraccion = DateEntry(principal, textvariable=entry_var_actainfraccion, date_pattern='dd/mm/Y', width=17, background='gray',
                   foreground='white', borderwidth=1)
label_actainfraccion = ttk.Label(principal, text="Fecha del Acta de Infracción", background="#ffffff")
if actainfraccion:
    entry_actainfraccion.set_date(actainfraccion)

# Crear caja de fecha 5, fecha resolucion impugnada.
entry_var_resolucionrecargo = tk.StringVar()
entry_resolucionrecargo = DateEntry(principal, textvariable=entry_var_resolucionrecargo, date_pattern='dd/mm/Y', width=17, background='gray',
                   foreground='white', borderwidth=1)
label_resolucionrecargo = ttk.Label(principal, text="Fecha de la Resolución impugnada", background="#ffffff")
if resolucionrecargo:
    entry_resolucionrecargo.set_date(resolucionrecargo)

# Crear caja de texto 15, porcentaje impuesto por la Inspección.
entry_var_porcentajerecargo = tk.StringVar()
entry_porcentajerecargo = ttk.Entry(principal, textvariable=entry_var_porcentajerecargo, justify=tk.CENTER)
label_porcentajerecargo = ttk.Label(principal, text="Porcentaje impuesto por la Inspección", background="#ffffff")
if recargo:
    entry_porcentajerecargo.insert(0, recargo)

# Crear caja de texto 6, artículos infringidos.
entry_var_articulosinfringidos = tk.StringVar()
entry_articulosinfringidos = ttk.Entry(principal, textvariable=entry_var_articulosinfringidos, justify=tk.CENTER)
label_articulosinfringidos = ttk.Label(principal, text="Artículos infringidos según acta de Infracción", background="#ffffff")

# Crear caja de fecha 7, fecha de la reclamación previa.
entry_var_reclamprevia = tk.StringVar()
entry_reclamprevia = DateEntry(principal, textvariable=entry_var_reclamprevia, date_pattern='dd/mm/Y', width=17, background='gray',
                   foreground='white', borderwidth=1)
label_reclamprevia = ttk.Label(principal, text="Fecha de la Reclamación Previa", background="#ffffff")
if reclamacionprevia:
    entry_reclamprevia.set_date(reclamacionprevia)

# Crear caja de fecha 8, fecha de la resolución de la reclamación previa.
entry_var_resolreclam = tk.StringVar()
entry_resolreclam = DateEntry(principal, textvariable=entry_var_resolreclam, date_pattern='dd/mm/Y', width=17, background='gray',
                   foreground='white', borderwidth=1)
label_resolreclam = ttk.Label(principal, text="Fecha de Resolución de la Reclamación Previa", background="#ffffff")
if resolucionreclamacion:
    entry_resolreclam.set_date(resolucionreclamacion)

# Crear caja de texto 9, pruena practicada.
entry_var_prueba = tk.StringVar()
entry_prueba = ttk.Entry(principal, textvariable=entry_var_prueba, justify=tk.CENTER)
label_prueba = ttk.Label(principal, text="Prueba practicada aparte de la documental", background="#ffffff")

# Crear caja de texto 10, en el presente caso.
entry_var_presentecaso = tk.StringVar()
entry_presentecaso = ttk.Entry(principal, textvariable=entry_var_presentecaso, justify=tk.CENTER)
label_presentecaso = ttk.Label(principal, text="En el presente caso...", background="#ffffff")

# Crear caja de texto 10, nombre del demandado.
entry_var_porcentajeadecuado = tk.StringVar()
entry_porcentajeadecuado = ttk.Entry(principal, textvariable=entry_var_porcentajeadecuado, justify=tk.CENTER)
label_porcentajeadecuado = ttk.Label(principal, text="Porcentaje que se estima adecuado", background="#ffffff")
# Posicionar la caja 10 en la ventana, nombre del demandado.

# Crear caja de texto 10, nombre del demandado.
entry_var_justifporcentaje = tk.StringVar()
entry_justifporcentaje = ttk.Entry(principal, textvariable=entry_var_justifporcentaje, justify=tk.CENTER)
label_justifporcentaje = ttk.Label(principal, text="Justificación del porcentaje que se impone", background="#ffffff")
# Posicionar la caja 10 en la ventana, nombre del demandado.

# Crear caja de texto 11, numero del procedimiento.
entry_var_numero = tk.StringVar()
entry_numero = ttk.Entry(principal, textvariable=entry_var_numero, justify=tk.CENTER)
label_numero = ttk.Label(principal, text="Número del procedimiento", background="#ffffff")

# Crear caja de texto 12, año del procedimiento.
entry_var_year = tk.StringVar()
entry_year = ttk.Entry(principal, textvariable=entry_var_year, justify=tk.CENTER)
label_year = ttk.Label(principal, text="Año", background="#ffffff")

# Crear caja de fecha 4, fecha de acta.
entry_var_activempres = tk.StringVar()
entry_activempres = ttk.Entry(principal, textvariable=entry_var_activempres, justify=tk.CENTER)
label_activempres = ttk.Label(principal, text="Actividad de la Empresa", background="#ffffff")
if actividad:
    entry_activempres.insert(0, actividad)

# Crear caja de fecha 5, fecha resolucion impugnada.
entry_var_antig = tk.StringVar()
entry_antig = DateEntry(principal, textvariable=entry_var_antig, date_pattern='dd/mm/Y', width=17, background='gray',
                    foreground='white', borderwidth=1)
label_antig = ttk.Label(principal, text="Antigüedad", background="#ffffff")
if antig:
    entry_antig.set_date(antig)

# Crear caja de texto 15, porcentaje impuesto por la Inspección.
entry_var_categoria = tk.StringVar()
entry_categoria = ttk.Entry(principal, textvariable=entry_var_categoria, justify=tk.CENTER)
label_categoria = ttk.Label(principal, text="Categoría Profesional", background="#ffffff")
if categoria:
    entry_categoria.insert(0, categoria)

# Crear caja de texto 6, artículos infringidos.
entry_var_salario = tk.StringVar()
entry_salario = ttk.Entry(principal, textvariable=entry_var_salario, justify=tk.CENTER)
label_salario = ttk.Label(principal, text="Salario diario", background="#ffffff")
if salario:
    entry_salario.insert(0, salario)

# Crear caja de fecha 7, fecha de la reclamación previa.
entry_var_fechasemac = tk.StringVar()
entry_fechasemac = DateEntry(principal, textvariable=entry_var_fechasemac, date_pattern='dd/mm/Y', width=17, background='gray',
                    foreground='white', borderwidth=1)
label_fechasemac = ttk.Label(principal, text="Fecha celebración SEMAC", background="#ffffff")

avenenciaoefecto = None

def check_cboxmayorofortuito(event):
    global avenenciaoefecto
    if resultsemac.get() == 'Sin avenencia':
        avenenciaoefecto = resultsemac.get()  # this will assign the variable c the value of cboxmayorofortuito
    if resultsemac.get() == 'Sin efecto':
        avenenciaoefecto = resultsemac.get()

resultsemac = ttk.Combobox(principal, values=["Sin avenencia", "Sin efecto"], width=17)
label_resultsemac = ttk.Label(principal, text="Resultado del SEMAC", background="#ffffff")
resultsemac.bind("<<ComboboxSelected>>", check_cboxmayorofortuito)

# Crear caja de texto 10, nombre del demandado.
entry_var_desdecuandonocobra = tk.StringVar()
entry_desdecuandonocobra = ttk.Entry(principal, textvariable=entry_var_desdecuandonocobra, justify=tk.CENTER)
label_desdecuandonocobra = ttk.Label(principal, text="Desde cuando no cobra", background="#ffffff")
# Posicionar la caja 10 en la ventana, nombre del demandado.

# Posicionar boton de fin
button = ttk.Button(principal, text="Crear Modelo", command=principal.destroy)

# mainloop de la ventana

principal.mainloop()

# selección de la plantilla a usar en función de los botones pulsados

if modelo == 1:
    template = "Plantilla Recargo de Prestaciones Desestimatoria.docx"
elif modelo == 2:
    template = "Plantilla Recargo de Prestaciones Estimacion Total.docx"
elif modelo == 4:
    template = "Plantilla Extincion retraso pago.docx"
elif modelo == 5:
    template = "Plantilla Extincion falta de pago.docx"
elif modelo == 6:
    template = "Plantilla Extincion modificacion sustancial.docx"
elif modelo == 3:
    template = "Plantilla Recargo de Prestaciones Estimacion Parcial.docx"

# creacion de variables para el nombre del archivo

if template == "Plantilla Recargo de Prestaciones Desestimatoria.docx":
    sentido = "desestim"
elif template == "Plantilla Recargo de Prestaciones Estimacion Total.docx":
    sentido = "estim total"
elif template == "Plantilla Extincion retraso pago.docx":
    sentido = "estim"
elif template == "Plantilla Extincion falta de pago.docx":
    sentido = "estim"
elif template == "Plantilla Extincion modificacion sustancial.docx":
    sentido = "estim"
else:
    sentido = "estim parcial"

if template == "Plantilla Recargo de Prestaciones Desestimatoria.docx":
    materia = "recargo prestaciones"
elif template == "Plantilla Recargo de Prestaciones Estimacion Total.docx":
    materia = "recargo prestaciones"
elif template == "Plantilla Extincion retraso pago.docx":
    materia = "extincion contrato"
elif template == "Plantilla Extincion falta de pago.docx":
    materia = "extincion contrato"
elif template == "Plantilla Extincion modificacion sustancial.docx":
    materia = "extincion contrato"
else:
    materia = "recargo prestaciones"

# uso del template elegido

document = MailMerge(template)

# creación de función para ver si la opción de indemnización ha sido seleccionada y rellenada, en caso contrario no hace nada

def indemnizacionsiono():
    if entry_var_salario.get():
        return indemnizacion_despido_prueba(str(entry_var_antig.get()), str(entry_var_salario.get()), str(today))
    else:
        valor0 = 0
        return valor0

# crear funcion de mayusculas y minusculas actor

def actorcapitalizacion():
    if template == "Plantilla Recargo de Prestaciones Desestimatoria.docx":
        return str(entry_var_actor.get()).upper()
    elif template == "Plantilla Recargo de Prestaciones Estimacion Total.docx":
        return str(entry_var_actor.get()).upper()
    elif template == "Plantilla Extincion retraso pago.docx":
        return str(entry_var_actor.get()).title()
    elif template == "Plantilla Extincion falta de pago.docx":
        return str(entry_var_actor.get()).title()
    elif template == "Plantilla Extincion modificacion sustancial.docx":
        return str(entry_var_actor.get()).title()

# crear funcion de mayusculas y minusculas demandado

def demandadocapitalizacion():
    if template == "Plantilla Recargo de Prestaciones Desestimatoria.docx":
        return str(entry_var_demandado.get()).title()
    elif template == "Plantilla Recargo de Prestaciones Estimacion Total.docx":
        return str(entry_var_demandado.get()).title()
    elif template == "Plantilla Extincion retraso pago.docx":
        return str(entry_var_demandado.get()).upper()
    elif template == "Plantilla Extincion falta de pago.docx":
        return str(entry_var_demandado.get()).upper()
    elif template == "Plantilla Extincion modificacion sustancial.docx":
        return str(entry_var_demandado.get()).upper()

# crear funcion tipo de extincion

if template == "Plantilla Extincion retraso pago.docx":
    tipoextin = "retraso pago"
elif template == "Plantilla Extincion falta de pago.docx":
    tipoextin = "falta de pago"
elif template == "Plantilla Extincion modificacion sustancial.docx":
    tipoextin = "modificacion sustancial"
else:
    tipoextin = ""

# meter todos los datos en el modelo

document.merge(
    Actor=actorcapitalizacion(),
    Demandado=demandadocapitalizacion(),
    Fechaactainspeccion=entry_var_actainfraccion.get(),
    fecharesolucionrecargo=entry_var_resolucionrecargo.get(),
    articulosinfringidos=str(entry_var_articulosinfringidos.get()) + " de la LPRL",
    fechareclamacionprevia=entry_var_reclamprevia.get(),
    fecharesolucionreclamacionprevia=entry_var_resolreclam.get(),
    pruebaspracticadas=" así como la " + str(entry_var_prueba.get()),
    Año=entry_var_year.get(),
    Número=entry_var_numero.get(),
    Enelpresentecaso=entry_var_presentecaso.get(),
    porcentajequeseimpone=entry_var_porcentajeadecuado.get(),
    justificacionporcentaje=entry_var_justifporcentaje.get(),
    porcentajerecargo=entry_var_porcentajerecargo.get(),
    antiguedad=entry_var_antig.get(),
    desdecuandonocobra=entry_var_desdecuandonocobra.get(),
    fechasemac=entry_var_fechasemac.get(),
    indemnizacionextincion=indemnizacionsiono(),
    categoria=entry_var_categoria.get(),
    salariodia=entry_var_salario.get(),
    actividadempresa=entry_var_activempres.get())

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

document.write(
    materia + " " + tipoextin + " " + entry_var_numero.get() + "-" + entry_var_year.get() + " " + sentido + '.docx')