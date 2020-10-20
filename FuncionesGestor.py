import re
from datetime import datetime
from nltk.tokenize import word_tokenize
import math
import pdftotext

def abrirpdf(pdf):
    with open(pdf, "rb") as f:
        archivo = pdftotext.PDF(f)
        document = "\n\n".join(archivo)
        return document

def preparar_texto(texto):
    from nltk.corpus import stopwords
    cleaned = re.sub(r'[^A-ZñÑa-zá-úä-ü0-9.,\/]+', ' ', texto)
    tokenized = word_tokenize(cleaned.lower())
    stopwords = set(stopwords.words('spanish'))
    textprov = [word for word in tokenized if word not in stopwords]
    t = list()
    for i in textprov:
        if re.match('\,', i):
            t.append(i)
        if re.match('\.', i):
            t.append(i)
    textdef = [word for word in textprov if word not in t]
    return textdef

def buscador(texto, valor):
    matches = list()
    for index in range(len(texto)):
        if texto[index] == valor:
            matches.append(index)
    if matches:
        return matches
    else:
        raise ValueError("{} no está en {}".format(valor, texto))

def indemnizacion_despido(fecha_antiguedad, salari, fecha_despido):
    from datetime import datetime, timedelta
    fechalimite = datetime.strptime('12/02/2012', '%d/%m/%Y').date()
    antig = datetime.strptime(fecha_antiguedad, '%d/%m/%Y').date()
    despid = datetime.strptime(fecha_despido, '%Y-%m-%d').date()
    salario = float(re.sub(',', '.', salari))
    if antig > fechalimite:
        dif = math.ceil(((despid - antig) / timedelta(days=1))) + 1
        daf = math.ceil(dif / 30.41666667)
        indem = daf * salario * 2.75
        return str(indem)
    else:
        dif1 = math.ceil(((fechalimite - antig) / timedelta(days=1))) + 1
        daf1 = math.ceil((dif1 / 30.41666667))
        indemprev = daf1 * salario * 3.75
        dif2 = math.ceil(((despid - fechalimite) / timedelta(days=1))) + 1
        daf2 = math.ceil(dif2 / 30.41666667)
        indempost = daf2 * salario * 2.75
        indem2 = indemprev + indempost
        return str(indem2)

def encontrarsalario(texto):
    salarioregex = re.compile('(\d+(\,|\.)?\d+?(\,|\.)?\d+?|\d+)')
    for word in texto:
        try:
            if word == "salario":
                if re.match(salarioregex, texto[(texto.index(word)) + 2]):
                    return texto[(texto.index(word)) + 2]
                elif re.match(salarioregex, texto[(texto.index(word)) + 1]):
                    return texto[(texto.index(word)) + 1]
        except:
            return None

def encontrardemandante(texto):
    for word in texto:
        try:
            if word == "representación":
                repreindex = texto.index(word)
            if word == "mayor":
                mayorindex = texto.index(word)
                demandante = (' '.join(texto[repreindex + 1:mayorindex - 1])).title()
                return demandante
        except:
            demandante = None
            return demandante

def encontrardemandado(texto):
    for word in texto:
        try:
            if word == "empresa":
                empresaindex = texto.index(word)
            if word == "cif" or word == "nif":
                cifindex = texto.index(word)
                return (' '.join(texto[empresaindex + 1:cifindex - 1])).upper()
        except:
            return None

def encontrarcategoria(texto):
    for word in texto:
        try:
            if word == "categoría":
                repreindex = texto.index(word)
                if texto[repreindex+1] == "profesional":
                    return texto[repreindex+2]
                else:
                    return texto[repreindex+1]
        except:
            return None

def encontraractividad(texto):
    for word in texto:
        try:
            if word == "actividad":
                repreindex = texto.index(word)
                return texto[repreindex+1]
        except:
            return None

def encontrarfechacercadeunapalabra(texto, palabra):
    try:
        fechas = [word for word in texto if re.match('\d{1,2}(\/|\.|\-)\d{1,2}(\/|\.|\-)\d{1,4}', word)]
        listpalabra = buscador(texto, palabra)
        fechasindex = list()
        for dates in fechas:
            fechasindex.append(texto.index(dates))
        fechaindex = None
        for indexpalabra in listpalabra:
            for indexfechas in fechasindex:
                if indexfechas - indexpalabra > 1 and indexfechas - indexpalabra < 4:
                    fechaindex = fechasindex.index(indexfechas)
        fechabuscada = fechas[fechaindex]
        return fechabuscada
    except:
        return None

def encontrarfechaantiguedad(texto):
    try:
        for item in range(len(texto)):
            if texto[item] == "antigüedad":
                if re.match('(\d+(\,|\.)?\d+?(\,|\.)?\d+?|\d+)', texto[item + 1]):
                    antigu = texto[item+1]
                    fechaclean = re.sub('\,|\.|\-|\/', '', antigu)
                    return datetime.strptime(fechaclean, '%d%m%Y').date()
            else:
                fechas = [word for word in texto if re.match('\d{1,2}(\/|\.|\-)\d{1,2}(\/|\.|\-)\d{1,4}', word)]
                return fechas[0]
    except:
        return None

def encontrardemandanterecargo(texto):
    for word in texto:
        try:
            if word == "representación":
                repreindex = texto.index(word)
            if word == "cif":
                mayorindex = texto.index(word)
                demandante = (' '.join(texto[repreindex + 1:mayorindex - 1])).upper()
                return demandante
        except:
            demandante = None
            return demandante

def encontrardemandadorecargo(texto):
    for word in texto:
        try:
            if word == "interesado":
                empresaindex = texto.index(word)
            if word == "dni":
                cifindex = texto.index(word)
                demandado = (' '.join(texto[empresaindex + 1:cifindex - 1])).title()
                return demandado
        except:
            return None

def encontrarporcentajerecargo(texto):
    for word in texto:
        try:
            if word == "incrementadas":
                recargoindex = texto.index(word)
                recargo = texto[recargoindex + 1]
                return recargo
        except:
            return None
