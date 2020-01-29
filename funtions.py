import pandas as pd
import numpy as np
import math
from scipy import stats
def abrir_archivo(nombre):
    datafreime = pd.read_excel(nombre)
    datos_var = datafreime.columns.tolist()  # lista de las variables
    #datafreime2 = datafreime.to_numpy()
    #descripcion = datafreime.describe().to_numpy()
    tipo_datos = datafreime.dtypes.tolist()
    return list(datos_var),list(tipo_datos),datafreime


def tabla_fercuencia(lista):
    # listafiltrados=[]
    # for i in lista:
    #   if np.isfinite(i):
    #     listafiltrados.append(i)
    # conjunto=list(set(listafiltrados))
    # if len(conjunto)<2 or sum(listafiltrados)==0:
    #    return "no"
    #if np.isinf(sum(arreglo)) or np.isnan(sum(arreglo)):
        #return "no"
    # else:
      #arreglo[np.isnan(arreglo)]=0
      # arreglo=np.array(listafiltrados)
      arreglo=np.array(lista)
      marca_clase = []
      frecuencia = []
      frecuencia_abso = []
      frecuenciare = []
      frecuenciarelaacu = []
      tablafrec = []
      # maximo = np.max(arreglo)
      # minimo = np.min(arreglo)
      # if maximo > 100:
      #     sep = 5
      # else:
      #     sep=math.fabs(round((maximo-minimo)/5,0))
      #     if sep==0:
      #       return "no"
      histo,conjunto=np.histogram(arreglo,5)
      for j in range(len(conjunto)-1):
          if j==(len(conjunto)-2):
              marca_clase.append('[{},{}]'.format(round(conjunto[j],0), round(conjunto[j + 1],0)))
          else:
              marca_clase.append('[{},{})'.format(round(conjunto[j],0), round(conjunto[j + 1]),0))
      frecuencia = histo
      frecuencia_abso = frecuencia.cumsum()
      frecuenciare = list(np.array(frecuencia) / sum(frecuencia))
      frecuenciarelaacu = list(np.array(frecuencia_abso) / sum(frecuencia))
      for j in range(len(marca_clase)):
          tuplita = (marca_clase[j],frecuencia[j],round(frecuencia_abso[j],2),round(frecuenciare[j],2),round(frecuenciarelaacu[j],2))
          tablafrec.append(tuplita)
      return tuple(tablafrec),conjunto,histo


def escritura(varia,var1,var2,var3,var4,var5):
    linea="para el histograma de las variable {} tiene una asimetria de valor {} y una kurtosis de valor {} " \
          "indicando que la variable esta sujeta a -- ademas de constar con 5 marcas de clase en resumen en el histograma" \
          " con respecto al grafico de cajas el rango central en el que estan ubicados es {} el extremo superior es {} y el limite inferior es {}".format(varia,var1,var2,var3,var4,var5)
    return linea



def kolmogorovtest(variable,lista):
    valor=np.array(lista)
    D,p=stats.kstest(valor,'norm')
    texto=""
    if p<0.05:
        texto="es normal"
    else:
        texto="no es normal"
    return variable,round(D,10),round(p,10),texto

def intervaloconfianza(variable,lista):
    arreglo=np.array(lista)
    mean,sigma=np.mean(arreglo),np.std(arreglo)
    conf_int = stats.norm.interval(0.95,loc=mean,scale=sigma/np.sqrt(len(lista)))
    return variable,conf_int[0],conf_int[1]



# variables,tipo,df=abrir_archivo("Datos proyecto-1.xlsx")
# contador=0
# lista=[]
# lista2=[]
# for i in variables:
#      if str(tipo[contador]).count("int")>0 or str(tipo[contador]).count("float")>0:
#           #print(tabla_fercuencia(df[i].tolist()))
#           lista.append(kolmogorovtest(i,df[i].tolist()))
#           lista2.append(intervaloconfianza(i,df[i].tolist()))
#      contador+=1
#
#print(lista2)
