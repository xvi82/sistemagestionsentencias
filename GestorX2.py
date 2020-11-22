from __future__ import print_function
from mailmerge import MailMerge
import tkinter as tk
from tkinter import ttk, filedialog, Tk, Label, Canvas, PhotoImage
from tkcalendar import DateEntry
from datetime import date, datetime
from FuncionesGestor import *

#fecha de hoy
today = date.today()

now = datetime.now()

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

#extraer los campos del modelo

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
        antig = None
except:
    antig = None
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
principal.title("Selección de Modelos Jurisdicción Social")

# Crear botones para elegir modelo.
def numerodelmodelo(value):
    global modelo
    modelo = value

# crear funciones de ocultar y mostrar los distintos cuadros de texto y labels
def mostrarrecargocondena():
    ocultar_multiple([lista_recargo_estim, lista_recargo_estim_parcial, lista_extincion_modificacion, lista_extincion_falta_pago, lista_extincion_retraso])
    posicion_de_label_y_entries(lista_recargo_condena)
    button.grid(row=99, column=1, pady=4)

def mostrarrecargoestim():
    ocultar_multiple([lista_recargo_condena, lista_recargo_estim_parcial, lista_extincion_modificacion, lista_extincion_falta_pago, lista_extincion_retraso])
    posicion_de_label_y_entries(lista_recargo_estim)
    button.grid(row=99, column=1, pady=4)

def mostrarrecargoestimparcial():
    ocultar_multiple([lista_recargo_condena, lista_recargo_estim, lista_extincion_modificacion, lista_extincion_falta_pago, lista_extincion_retraso])
    posicion_de_label_y_entries(lista_recargo_estim_parcial)
    button.grid(row=99, column=1, pady=4)

def mostrarextincionretraso():
    ocultar_multiple([lista_recargo_condena, lista_recargo_estim, lista_recargo_estim_parcial, lista_extincion_modificacion, lista_extincion_falta_pago])
    posicion_de_label_y_entries(lista_extincion_retraso)
    button.grid(row=99, column=1, pady=4)

def mostrarextincionfaltapago():
    ocultar_multiple([lista_recargo_condena, lista_recargo_estim, lista_recargo_estim_parcial, lista_extincion_modificacion, lista_extincion_retraso])
    posicion_de_label_y_entries(lista_extincion_falta_pago)
    button.grid(row=99, column=1, pady=4)

def mostrarextincionmodificacion():
    ocultar_multiple([lista_recargo_condena, lista_recargo_estim, lista_recargo_estim_parcial, lista_extincion_falta_pago, lista_extincion_retraso])
    posicion_de_label_y_entries(lista_extincion_modificacion)
    button.grid(row=99, column=1, pady=4)

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

# Crear caja de texto, nombre del actor.
entry_var_actor, entry_actor, label_actor = boton_label_variable(principal, "Nombre del actor")
if demandante:
    entry_actor.insert(0, demandante)

# Crear caja de texto, nombre del demandado.
entry_var_demandado, entry_demandado, label_demandado = boton_label_variable(principal, "Nombre del demandado")
if demandado:
    entry_demandado.insert(0, demandado)

# Crear caja de fecha, fecha de acta.
entry_var_fechaactainspeccion, entry_fechaactainspeccion, label_fechaactainspeccion = fecha_label_varible(principal, "Fecha del Acta de Infracción")
if actainfraccion:
    entry_fechaactainspeccion.set_date(actainfraccion)

# Crear caja de fecha, fecha resolucion impugnada.
entry_var_fecharesolucionrecargo, entry_fecharesolucionrecargo, label_fecharesolucionrecargo = fecha_label_varible(principal, "Fecha de la Resolución impugnada")
if resolucionrecargo:
    entry_fecharesolucionrecargo.set_date(resolucionrecargo)

# Crear caja de texto, porcentaje impuesto por la Inspección.
entry_var_porcentajerecargo, entry_porcentajerecargo, label_porcentajerecargo = boton_label_variable(principal, "Porcentaje impuesto por la Inspección")
if recargo:
    entry_porcentajerecargo.insert(0, recargo)

# Crear caja de texto, artículos infringidos.
entry_var_articulosinfringidos, entry_articulosinfringidos, label_articulosinfringidos = boton_label_variable(principal, "Artículos infringidos según acta de Infracción")

# Crear caja de fecha, fecha de la reclamación previa.
entry_var_fechareclamacionprevia, entry_fechareclamacionprevia, label_fechareclamacionprevia = fecha_label_varible(principal, "Fecha de la Reclamación Previa")
if reclamacionprevia:
    entry_fechareclamacionprevia.set_date(reclamacionprevia)

# Crear caja de fecha, fecha de la resolución de la reclamación previa.
entry_var_fecharesolucionreclamacionprevia, entry_fecharesolucionreclamacionprevia, label_fecharesolucionreclamacionprevia = fecha_label_varible(principal, "Fecha de Resolución de la Reclamación Previa")
if resolucionreclamacion:
    entry_fecharesolucionreclamacionprevia.set_date(resolucionreclamacion)

# Crear caja de texto, prueba practicada.
entry_var_pruebaspracticadas, entry_pruebaspracticadas, label_pruebaspracticadas = boton_label_variable(principal, "Prueba practicada aparte de la documental")

# Crear caja de texto de fundamentacion para el caso concreto y valoracion de prueba
entry_var_enelpresentecaso, entry_enelpresentecaso, label_enelpresentecaso = boton_label_variable(principal, "En el presente caso...")

# Crear caja de texto de porcentaje que se impone al trabajador
entry_var_porcentajequeseimpone, entry_porcentajequeseimpone, label_porcentajequeseimpone = boton_label_variable(principal, "Porcentaje que se estima adecuado")

# Crear caja de texto 10, nombre del demandado.
entry_var_justificacionporcentaje, entry_justificacionporcentaje, label_justificacionporcentaje = boton_label_variable(principal, "Justificación del porcentaje que se impone")

# Crear caja de texto, numero del procedimiento.
entry_var_numero, entry_numero, label_numero = boton_label_variable(principal, "Número del procedimiento")

# Crear caja de texto, año del procedimiento.
entry_var_ano, entry_ano, label_ano = boton_label_variable(principal, "Año")

# Crear caja de¨texto de la actividad de la empresa
entry_var_actividad, entry_actividad, label_actividad = boton_label_variable(principal, "Actividad de la Empresa")
if actividad:
    entry_actividad.insert(0, actividad)

# Crear caja de fecha antigüedad del trabajador.
entry_var_antiguedad, entry_antiguedad, label_antiguedad = fecha_label_varible(principal, "Antigüedad")
if antig:
    entry_antiguedad.set_date(antig)

# Crear caja de texto, categoría profesional.
entry_var_categoria, entry_categoria, label_categoria = boton_label_variable(principal, "Categoría Profesional")
if categoria:
    entry_categoria.insert(0, categoria)

# Crear caja de texto de salario diario.
entry_var_salario, entry_salario, label_salario = boton_label_variable(principal, "Salario diario")
if salario:
    entry_salario.insert(0, salario)

# Crear caja de fecha, fecha de celebración del SEMAC.
entry_var_fechasemac, entry_fechasemac, label_fechasemac = fecha_label_varible(principal, "Fecha celebración SEMAC")

# Crear caja de texto, desde cuando no cobra el trabajador.
entry_var_desdecuandonocobra, entry_desdecuandonocobra, label_desdecuandonocobra = boton_label_variable(principal, "Desde cuando no cobra")

#creacion combobox para Resultado del SEMAC

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

# Posicionar boton de fin
button = ttk.Button(principal, text="Crear Modelo", command=principal.destroy)

#lista de entradas y labels como variables
lista_recargo_condena = [label_actor, entry_actor, label_demandado, entry_demandado, label_fechaactainspeccion, entry_fechaactainspeccion, label_fecharesolucionrecargo, entry_fecharesolucionrecargo, label_porcentajerecargo, entry_porcentajerecargo, label_articulosinfringidos, entry_articulosinfringidos, label_fechareclamacionprevia, entry_fechareclamacionprevia, label_fecharesolucionreclamacionprevia, entry_fecharesolucionreclamacionprevia, label_pruebaspracticadas, entry_pruebaspracticadas, label_enelpresentecaso, entry_enelpresentecaso, label_numero, entry_numero, label_ano, entry_ano]
lista_recargo_estim_parcial = [label_actor, entry_actor, label_demandado, entry_demandado, label_fechaactainspeccion, entry_fechaactainspeccion, label_fecharesolucionrecargo, entry_fecharesolucionrecargo, label_porcentajerecargo, entry_porcentajerecargo, label_articulosinfringidos, entry_articulosinfringidos, label_fechareclamacionprevia, entry_fechareclamacionprevia, label_fecharesolucionreclamacionprevia, entry_fecharesolucionreclamacionprevia, label_pruebaspracticadas, entry_pruebaspracticadas, label_enelpresentecaso, entry_enelpresentecaso, label_porcentajequeseimpone, entry_porcentajequeseimpone, label_justificacionporcentaje, entry_justificacionporcentaje, label_numero, entry_numero, label_ano, entry_ano]
lista_recargo_estim = [cboxmayorofortuitolabel, cboxmayorofortuito] + lista_recargo_condena
lista_extincion_retraso = [label_actor, entry_actor, label_demandado, entry_demandado, label_actividad, entry_actividad, label_antiguedad, entry_antiguedad, label_categoria, entry_categoria, label_salario, entry_salario, label_fechasemac, entry_fechasemac, label_resultsemac, resultsemac, label_pruebaspracticadas, entry_pruebaspracticadas, label_enelpresentecaso, entry_enelpresentecaso, label_numero, entry_numero, label_ano, entry_ano]
lista_extincion_falta_pago = [label_actor, entry_actor, label_demandado, entry_demandado, label_actividad, entry_actividad, label_antiguedad, entry_antiguedad, label_categoria, entry_categoria, label_salario, entry_salario, label_fechasemac, entry_fechasemac, label_resultsemac, resultsemac, label_pruebaspracticadas, entry_pruebaspracticadas, label_enelpresentecaso, entry_enelpresentecaso, label_desdecuandonocobra, entry_desdecuandonocobra, label_numero, entry_numero, label_ano, entry_ano]
lista_extincion_modificacion = lista_extincion_retraso

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
# variable del sentido del fallo según la plantilla usada
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

# variable de la materia según la plantilla usada
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
        return indemnizacion_despido(str(entry_var_antiguedad.get()), str(entry_var_salario.get()), str(today))
    else:
        return 0

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
    actor02=actorcapitalizacion(),
    demandado03=demandadocapitalizacion(),
    fechaactainspeccion04=entry_var_fechaactainspeccion.get(),
    fecharesolucionrecargo05=entry_var_fecharesolucionrecargo.get(),
    articulosinfringidos07=str(entry_var_articulosinfringidos.get()) + " de la LPRL",
    fechareclamacionprevia08=entry_var_fechareclamacionprevia.get(),
    fecharesolucionreclamacionprevia09=entry_var_fecharesolucionreclamacionprevia.get(),
    pruebaspracticadas20=" así como la " + str(entry_var_pruebaspracticadas.get()),
    ano98=entry_var_ano.get(),
    numero97=entry_var_numero.get(),
    enelpresentecaso30=entry_var_enelpresentecaso.get(),
    porcentajequeseimpone31=entry_var_porcentajequeseimpone.get(),
    desdecuandonocobra31= entry_var_desdecuandonocobra.get(),
    justificacionporcentaje32=entry_var_justificacionporcentaje.get(),
    porcentajerecargo06=entry_var_porcentajerecargo.get(),
    antiguedad05=entry_var_antiguedad.get(),
    semac08=entry_var_fechasemac.get(),
    indemnizacionextincion40=indemnizacionsiono(),
    categoria06=entry_var_categoria.get(),
    salario07=entry_var_salario.get(),
    actividad04=entry_var_actividad.get(),
    fecha=current_date_format(now))

if mayorofortuito == "Fuerza Mayor":
    document.merge(fortuitomayor101="por la fuerza mayor operada al tiempo del accidente",
                   fortuitomayor201="sino que el hecho acaecido responde a un caso de fuerza mayor, esto es, un evento extraño al círculo o ámbito de la actividad del trabajador y de la empresa, que rompe el nexo de causalidad")
else:
    document.merge(fortuitomayor101="por su acaecimiento fortuito",
                   fortuitomayor201="sino que el hecho acaecido responde a un caso fortuito, esto es, un hecho independiente de la voluntad del deudor de seguridad (la empresa), imprevisible e inevitable")

if avenenciaoefecto == "Sin avenencia":
    document.merge(resultsemac09="sin avenencia")
else:
    document.merge(resultsemac09="sin efecto")

# creación del archivo de texto

document.write(
    materia + " " + tipoextin + " " + entry_var_numero.get() + "-" + entry_var_ano.get() + " " + sentido + '.docx')