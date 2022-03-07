import os
import pandas as pd

from poligono_ambiente import PoligonoAmbiente
from poligono_fishnet import PoligonoFishnet
from poligono_lote import PoligonoLote
from poligono_muestreo import PoligonoMuestreo
from poligono_parcela import PoligonoParcela
from poligono import Poligono

from vector_rinde_lote import VectorRindeLote

from analisis import Analisis
from pre_muestra import PreMuestra
from analisis_estadistico import AnalisisEstadistico


def procesar_mapa_desde_arcmap(
    ambientes_shp,  # str
    campo_ambiente,  # str
    parcelas_shp,  # str
    campo_parcela,  # str
    lotes_shp,  # str
    campo_lote,  # str
    campo_campo,  # str
    dir_salida,  # str
    puntos_shp,  # str
    campo_rendimiento,  # str
    nombre_testigo=None,  # str
    convertir_campo_parcela=False,  # str | bool
    exportar_poligono_muestreo=True,  # str | bool
    buffer=0,  # str | int
    largo_celda_fishnet=0,  # str | int 
    marco_shp=None,  # str
    excel_rindes=None,  # str 
    tipo_analisis='dca1f',  # str
    exportar_mapa_rinde=True,  # str | bool
    todos_los_ambientes=False,  # str | bool
    perdonar_falta_cuadrados=True,   # str | bool
    repeticiones=0,  # str | int
    nombre_analisis='',  # str
    exportar_intersect=True,  # str | bool
    dca1f_por_ambiente=True  # str | bool
    ):

    ########################
    #                      #
    #        PASOS         #
    #                      #
    ########################

    # Unidades
    # Outliers
    # Factor rinde real
    # Armado de grilla
    # Interseccion de mapa de ambientes con mapa de parcelas
    # Interseccion de grilla con mapa de ambientes_parcelas
    # Calcular superficie de cada poligono
    # Interseccion con el mapa de puntos
    # Listado de todos los puntos de los cuadrados enteros por parcela y ambiente
    # Corroborar si existe la misma cantidad de ambientes por parcela > errs
    # Descartar los cuadrados con muchos o pocos puntos (threshold)
    # Seleccionar cantidad minima de puntos por cuadrado al azar
    # Seleccionar cantidad minima de cuadrados por ambiente al azar
    ## Exportar shp de puntos seleccionados ?
    # Exportar excel de puntos seleccionados
    # Supuestos
    # ANOVA
    # Comparacion de medias

    ma = PoligonoAmbiente()
    ma.cargar(ambientes_shp, campo_ambiente)
    extent = ma.leer_extension()

    mp = PoligonoParcela()    
    mp.cargar(parcelas_shp, campo_parcela, nombre_testigo)
    if convertir_campo_parcela == True or convertir_campo_parcela=='true':
        mp.convertir_campo_parcela() 
    
    if int(largo_celda_fishnet) > 0:
        fn = PoligonoFishnet()
        fn.cargar(extent, int(largo_celda_fishnet))
    else:
        fn = None

    if marco_shp is not None:
        mc = Poligono()
        mc.cargar(marco_shp)
    else:
        mc = None

    mu = PoligonoMuestreo()
    mu.cargar(ma, mp, int(buffer), fn, mc)
    if exportar_poligono_muestreo == True or exportar_poligono_muestreo == 'true':
        mu.exportar(dir_salida + r'\muestreo.shp')
    
    ml = PoligonoLote()
    ml.cargar(lotes_shp, campo_lote, campo_campo)
    lotes = ml.lista_de_lotes()

    mr = VectorRindeLote()
    salida_puntos = dir_salida + r'\mapa_rinde.shp'
    if os.path.exists(salida_puntos):
        mr.cargar_procesado(salida_puntos)
    else:
        mr.cargar_sin_procesar(puntos_shp, campo_rendimiento, ml)
        if excel_rindes is not None:
            rindes = pd.read_excel(excel_rindes)
            rindes['ID_Cultivo'] = rindes['ID_Cultivo'].apply(str)
            mr.corregir_rinde(lotes, rindes)
        if exportar_mapa_rinde == True or exportar_mapa_rinde =='true':
            mr.exportar(salida_puntos)

    tipo_analisis = Analisis.get_item(tipo_analisis)
    pre_muestra = PreMuestra(mr, mu, (todos_los_ambientes == True or todos_los_ambientes=='true'), (perdonar_falta_cuadrados == True or perdonar_falta_cuadrados=='true'), int(repeticiones), tipo_analisis, nombre_analisis)

    muestra = pre_muestra.muestrear(dir_salida, (exportar_intersect == True or exportar_intersect=='true'))

    AnalisisEstadistico((dca1f_por_ambiente == True or dca1f_por_ambiente=='true'), muestra)