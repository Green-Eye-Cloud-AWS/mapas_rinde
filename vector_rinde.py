import os
import numpy as np

from poligono_lote import CAMPO_CAMPO, CAMPO_LOTE

from vector import Vector, OUT_TEMP

import arcpy

CAMPO_RENDIMIENTO = "src_Rinde"
CAMPO_ID_RINDE = "ID_MAPA_RI"
OUT_NAME = "mapa_rinde.shp"

class VectorRinde(Vector):
    
    DISTANCIA = 20

    def __init__(self):
        super(VectorRinde, self).__init__()

        self.campo_idRinde = None
        self.campo_rendimiento = None


    def cargar_procesado(self, shapefile):

        if os.path.exists(shapefile):
            self.mem_shp = OUT_TEMP + r"\puntos_shp"
            self.cargar_en_memoria(shapefile, self.mem_shp)

            self.campo_rendimiento = CAMPO_RENDIMIENTO
            self.campo_idRinde = CAMPO_ID_RINDE
            self.campo_lote = CAMPO_LOTE
            self.campo_campo = CAMPO_CAMPO
            self.cargar_atributos(CAMPO_RENDIMIENTO, CAMPO_ID_RINDE, CAMPO_LOTE, CAMPO_CAMPO)

        else:
            raise Exception("Mapa de rinde inexistente")


    def cargar_sin_procesar(self, shapefile_raw, campo_rendimiento_raw):

        self.mem_shp = OUT_TEMP + r"\puntos_shp"
        self.cargar_en_memoria(shapefile_raw, self.mem_shp)
        self.campo_rendimiento = campo_rendimiento_raw
        
        self.calcular_campo_rinde()
        self.campo_rendimiento = CAMPO_RENDIMIENTO

        self.eliminar_outliers()
        
        self.calcular_campo_id(CAMPO_ID_RINDE)
        self.campo_idRinde = CAMPO_ID_RINDE

        self.borrar_otros_campos()
           

    def unidades_en_toneladas(self):

        i = 0
        suma = 0
        for row in arcpy.SearchCursor(self.mem_shp):
            i += 1
            suma += row.getValue(self.campo_rendimiento)
            if i == 500:
                break
        
        if suma/i < 1000:
            arcpy.AddMessage("Rinde en toneladas")
            arcpy.AddMessage('')
            return True
        else:
            arcpy.AddMessage("Rinde en kilos")
            arcpy.AddMessage('')
            return False


    def calcular_campo_rinde(self):
                
        self.generar_campo(CAMPO_RENDIMIENTO, 'DOUBLE')

        factor = " * 1000" if self.unidades_en_toneladas() else ""

        arcpy.CalculateField_management(self.mem_shp, CAMPO_RENDIMIENTO, "!" + self.campo_rendimiento + "!" + factor, "PYTHON")
        
        self.cargar_atributos(CAMPO_RENDIMIENTO)


    def eliminar_outliers(self):

        out_zero = OUT_TEMP + r"\zero"
        arcpy.AddMessage("Filtrando rendimientos mayores a 0")
        arcpy.Select_analysis(self.mem_shp, out_zero, '"' + self.campo_rendimiento + '" > 0')

        out_min_max = OUT_TEMP + r"\min_max"
        arcpy.AddMessage("Filtrando rendimientos dentro de percentiles")
        puntos = [row.getValue(self.campo_rendimiento) for row in arcpy.SearchCursor(out_zero)]
        rinde_max = np.percentile(puntos, 75) + 1.5 * (np.percentile(puntos, 75) - np.percentile(puntos, 25))
        rinde_min = np.percentile(puntos, 25) - 1.5 * (np.percentile(puntos, 75) - np.percentile(puntos, 25))
        arcpy.Select_analysis(out_zero, out_min_max, '"' + self.campo_rendimiento + '" < ' + str(rinde_max) + ' AND "' + self.campo_rendimiento + '" > ' + str(rinde_min))

        out_outliers = OUT_TEMP + r"\outliers"
        arcpy.AddMessage("Buscando outliers posicionales")
        arcpy.ClustersOutliers_stats(out_min_max, self.campo_rendimiento, out_outliers, "INVERSE_DISTANCE", "EUCLIDEAN_DISTANCE", "NONE", self.DISTANCIA)

        out_filtered = OUT_TEMP + r"\filtered"
        arcpy.AddMessage("Filtrando outliers posicionales")
        arcpy.Select_analysis(out_outliers, out_filtered, "\"COType\" = ' ' OR \"COType\" = 'HH' OR \"COType\" = 'LL'")

        self.mem_shp = out_filtered

        arcpy.AddMessage("")


if __name__ == '__main__':

    mr = VectorRinde() # r'E:\ArcGIS\NewFolder', r'E:\ArcGIS\NewFolder\belver.shp', 'Yld_Mass_W'
    print(mr.mem_shp)
    #mr.cargar_procesado(r'E:\ArcGIS\NewFolder')
    
    mr.cargar_sin_procesar(r'E:\ArcGIS\NewFolder\belver.shp', 'Yld_Mass_W')
    print(mr.atributos)

    #mr.exportar(r'E:\ArcGIS\NewFolder\New Folder\mapa_rinde.shp')
