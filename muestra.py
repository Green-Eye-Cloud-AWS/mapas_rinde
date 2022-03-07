import numpy as np
import pandas as pd
import random
import copy
import time

import arcpy

class Muestra(object):

    def __init__(self, out_dir, repeticiones, puntos_min, pre_muestra):
        
        self.out_dir = out_dir
        self.repeticiones = repeticiones
        self.puntos_min = puntos_min
        self.analisis_objetivo = pre_muestra.analisis_objetivo
        self.nombre_analisis = pre_muestra.nombre_analisis
        self.nombre_campo = pre_muestra.nombre_campo
        self.df_perdonados = pre_muestra.df_perdonados
        self.nombre_testigo = pre_muestra.nombre_testigo

        self.timestamp = str(int(time.time()*1000))
        self.excel_path = self.out_dir +  r"\data_" + self.timestamp + ".xlsx"

        datos_dict_c = copy.deepcopy(pre_muestra.datos_dict)

        self.ids = None
        self.muestra = None

        self._muestrear(datos_dict_c)

        arcpy.AddMessage("Cuadrados por ambiente (repeticiones): " + str(self.repeticiones))
        arcpy.AddMessage("Puntos por cuadrado: 1 (mediana de " + str(self.puntos_min) + " puntos)")

    
    def _muestrear(self, datos_dict_c):

        self._cuadrados_por_ambiente_rnd(datos_dict_c)
        self._puntos_por_cuadrado_rnd(datos_dict_c)
        self._armar_dataframe(datos_dict_c)
        self._exportar_excel()


    def _exportar_excel(self):

        self.muestra.to_excel(self.excel_path, sheet_name="DATA", index=False)
        arcpy.AddMessage("Excel exportado")


    def _armar_dataframe(self, datos_dict_c):

        data = {
            "Parcela": [],
            "Ambiente": [],
            # "ID_fishnet": [],
            "Rinde":[]
        }

        for parcela in datos_dict_c:
            for ambiente in datos_dict_c[parcela]:
                for cuadrado in datos_dict_c[parcela][ambiente]:
                    for rinde in datos_dict_c[parcela][ambiente][cuadrado]:  # Hay uno solo por cuadrado (mediana)

                        data["Parcela"].append(parcela)
                        data["Ambiente"].append(ambiente)
                        # data["ID_fishnet"].append(cuadrado)
                        data["Rinde"].append(rinde)

        self.muestra = pd.DataFrame(data)


    def _puntos_por_cuadrado_rnd(self, datos_dict_c):
        
        ids = []
        for parcela in datos_dict_c:
            for ambiente in datos_dict_c[parcela]:
                for cuadrado in datos_dict_c[parcela][ambiente]:
                    puntos = random.sample(datos_dict_c[parcela][ambiente][cuadrado], self.puntos_min)
                    rindes = [ p[0] for p in puntos ]
                    ids += [ p[1] for p in puntos ]
                    datos_dict_c[parcela][ambiente][cuadrado] = [np.median(rindes)]

        self.ids = ids


    def _cuadrados_por_ambiente_rnd(self, datos_dict_c):

        cells_dict_c = copy.deepcopy(datos_dict_c)
        for parcela in datos_dict_c:
            for ambiente in datos_dict_c[parcela]:
                cuadrados_rnd = random.sample(datos_dict_c[parcela][ambiente], self.repeticiones)
                datos_dict_c[parcela][ambiente] = {}

                for cuadrado in cuadrados_rnd:
                    datos_dict_c[parcela][ambiente][cuadrado] = cells_dict_c[parcela][ambiente][cuadrado]