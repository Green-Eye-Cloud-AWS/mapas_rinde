

from muestra import Muestra
from vector import Vector, OUT_TEMP

import arcpy

from vector_rinde_lote import VectorRindeLote

class VectorMuestra(Vector):

    def __init__(self, vector_rinde_lote, poligono_muestreo):
        super(VectorMuestra, self).__init__()

        self.campo_idRinde = None
        self.campo_rendimiento = None
        self.campo_lote = None
        self.campo_campo = None
        self.campo_parcela = None
        self.nombre_testigo = None
        self.campo_ambiente = None
        self.area_fishnet = None
        self.largo_buffer = None
        self.campo_area = None
        self.campo_idFishnet = None

        self._intersectar_muestreo(vector_rinde_lote, poligono_muestreo)

        self.borrar_otros_campos()


    def _intersectar_muestreo(self, vector_rinde_lote, poligono_muestreo):

        out_intersect_2 = OUT_TEMP + r"\intersect_2"
        arcpy.Intersect_analysis([vector_rinde_lote.mem_shp, poligono_muestreo.mem_shp], out_intersect_2)
        self.mem_shp = out_intersect_2

        self.cargar_atributos(*vector_rinde_lote.atributos)
        self.campo_idRinde = vector_rinde_lote.campo_idRinde
        self.campo_rendimiento = vector_rinde_lote.campo_rendimiento
        self.campo_lote = vector_rinde_lote.campo_lote
        self.campo_campo = vector_rinde_lote.campo_campo

        self.cargar_atributos(*poligono_muestreo.atributos)
        self.campo_parcela = poligono_muestreo.campo_parcela
        self.nombre_testigo = poligono_muestreo.nombre_testigo
        self.campo_ambiente = poligono_muestreo.campo_ambiente
        self.area_fishnet = poligono_muestreo.area_fishnet
        self.largo_buffer = poligono_muestreo.largo_buffer
        self.campo_area = poligono_muestreo.campo_area
        self.campo_idFishnet = poligono_muestreo.campo_idFishnet


    def exportar(self, muestra):

        new_out_intersect = muestra.out_dir + r"\muestra_" + muestra.timestamp + ".shp"
        if len(muestra.ids) > 0:
            arcpy.Select_analysis(self.mem_shp, new_out_intersect, "\"" + self.campo_idRinde + "\" IN (" + ",".join([str(i) for i in muestra.ids ])  + ")") 
            arcpy.AddMessage("Intersect exportado")
        else:
            arcpy.AddMessage("No hay puntos que exportar!")