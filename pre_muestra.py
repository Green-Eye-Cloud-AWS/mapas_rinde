#!/usr/bin/env python
# -*- coding: utf-8 -*-

import copy
import numpy as np
import pandas as pd

from poligono_muestreo import PoligonoMuestreo
from poligono_fishnet import PoligonoFishnet
from poligono_parcela import PoligonoParcela
from poligono_ambiente import PoligonoAmbiente
from poligono_lote import PoligonoLote
from vector_muestra import VectorMuestra
from vector_rinde_lote import VectorRindeLote

from muestra import Muestra
from analisis import Analisis

import arcpy


class PreMuestra(object):

    PUNTOS_MIN = 10
    GL_OBJETIVO = 40

    def __init__(self, vector_rinde_lote, poligono_muestreo, todos_los_ambientes=False, perdonar_falta_cuadrados=True, repeticiones=0, analisis_objetivo=Analisis.dca1f, nombre_analisis=''):
        
        self.analisis_objetivo = analisis_objetivo
        if not isinstance(self.analisis_objetivo, Analisis):
            raise Exception('Análisis desconocido!')

        self.vector_muestra = VectorMuestra(vector_rinde_lote, poligono_muestreo)

        self.todos_los_ambientes = todos_los_ambientes
        self.perdonar_falta_cuadrados = perdonar_falta_cuadrados
        self.nombre_analisis = nombre_analisis
        self.repeticiones = repeticiones

        self.nombre_testigo = self.vector_muestra.nombre_testigo

        lista_campos = self.vector_muestra.listar_atributo(self.vector_muestra.campo_campo)
        if len(lista_campos) > 1:
            raise Exception('No se puede analizar datos de más de un campo!')
        self.nombre_campo = lista_campos[0]

        self.puntos_min = max(self.PUNTOS_MIN, int(self.vector_muestra.area_fishnet * 80))
        
        self.datos_dict = None
        self.parcelas = None
        self.ambientes = None
        self.df_perdonados = None        

        self._preparar_muestra()


    def muestrear(self, out_dir, exportar_shp=False):
        muestra = Muestra(out_dir, self.repeticiones, self.puntos_min, self)

        if exportar_shp:
            self.vector_muestra.exportar(muestra)

        return muestra
        

    def _preparar_muestra(self):

        self._estructurar()  # self.datos_dict
        self._chequear_parcelas()  # self.parcelas
        self._chequear_ambientes()  # self.ambientes

        if len(self.ambientes) == 1:
            self.analisis_objetivo = Analisis.dca1f
        
        self._calcular_repeticiones()  # self.repeticiones
        self._normalizar_puntos_por_cuadrado() # self.puntos_min
        self._perdonar_incompletos()  # self.df_perdonados

        arcpy.AddMessage("")


    def _perdonar_incompletos(self):
        
        self.df_perdonados = pd.DataFrame(columns=["Parcelas","Diferencia_estadistica","Rinde_Parcela","Rinde_Testigo","Nombre","Campo","Ambiente","DMS","Diferencia"])

        if self.perdonar_falta_cuadrados:
            ambientes_flojos = set()
            for parcela in self.datos_dict:
                for ambiente in self.datos_dict[parcela]:
                    if len(self.datos_dict[parcela][ambiente]) < self.repeticiones + 1:
                        ambientes_flojos.add(ambiente)

            rinde_dict = {}
            cells_dict_c = copy.deepcopy(self.datos_dict)
            for parcela in cells_dict_c:
                rinde_dict[parcela] = {}
                for ambiente in cells_dict_c[parcela]:
                    if ambiente in ambientes_flojos:
                        
                        rindes = []
                        for cuadrado in self.datos_dict[parcela][ambiente]:
                            rindes += [ p[0] for p in self.datos_dict[parcela][ambiente][cuadrado] ]
                        rinde = np.median(rindes)
                        
                        rinde_dict[parcela][ambiente] = rinde

            for parcela in cells_dict_c:
                for ambiente in cells_dict_c[parcela]:
                    if ambiente in ambientes_flojos:

                        if parcela != self.vector_muestra.nombre_testigo:
                            self.df_perdonados.loc[len(self.df_perdonados)] = {
                                "Parcelas": parcela,
                                "Diferencia_estadistica": -99999,
                                "Rinde_Parcela": rinde_dict[parcela][ambiente],
                                "Rinde_Testigo": rinde_dict[self.vector_muestra.nombre_testigo][ambiente],
                                "Nombre": self.nombre_analisis,
                                "Campo": self.nombre_campo,
                                "Ambiente": ambiente,
                                "DMS": -99999,
                                "Diferencia": rinde_dict[parcela][ambiente] - rinde_dict[self.vector_muestra.nombre_testigo][ambiente]
                            }
                        
                        del self.datos_dict[parcela][ambiente]
            
            for ambiente in ambientes_flojos:
                arcpy.AddMessage("Perdonado: " + ambiente)

        else:

            cuadrados = []
            for parcela in self.datos_dict:
                for ambiente in self.datos_dict[parcela]:
                    cuadrados.append(len(self.datos_dict[parcela][ambiente]))

            cuadrados_min = np.min(cuadrados)

            if cuadrados_min < self.repeticiones + 1:
                raise Exception("La cantidad de cuadrados por ambiente no es suficiente! " + str(cuadrados_min) + "/" + str(self.repeticiones + 1))

        
    def _normalizar_puntos_por_cuadrado(self):

        puntos = []
        for parcela in self.datos_dict:
            for ambiente in self.datos_dict[parcela]:
                for cuadrado in self.datos_dict[parcela][ambiente]:
                    puntos.append(len(self.datos_dict[parcela][ambiente][cuadrado]))

        puntos_max = np.min([ np.percentile(puntos, 75) + 1.5 * (np.percentile(puntos, 75) - np.percentile(puntos, 25)), np.max(puntos) ])
        self.puntos_min = np.max([ np.percentile(puntos, 25) - 1.5 * (np.percentile(puntos, 75) - np.percentile(puntos, 25)), self.puntos_min + 1 ])

        if self.puntos_min > puntos_max:
            self.puntos_min = self.PUNTOS_MIN + 1

        if self.puntos_min > puntos_max:
            raise Exception("La cantidad de puntos por cuadrado no es suficiente! ({p_max}/{p_min})".format(p_max=puntos_max, p_min=self.puntos_min))

        cells_dict_c = copy.deepcopy(self.datos_dict)
        d = 0
        for parcela in cells_dict_c:
            for ambiente in cells_dict_c[parcela]:
                for cuadrado in cells_dict_c[parcela][ambiente]:
                    if len(self.datos_dict[parcela][ambiente][cuadrado]) > puntos_max or len(self.datos_dict[parcela][ambiente][cuadrado]) < self.puntos_min:
                        del self.datos_dict[parcela][ambiente][cuadrado]
                        d += 1
        
        arcpy.AddMessage("{n} cuadrado(s) eliminados en normalización de cantidad de puntos por cuadrado".format(n=d))
        
        self.puntos_min = int(self.puntos_min) - 1


    def _calcular_repeticiones(self):

        min_rep = 0
        if self.analisis_objetivo == Analisis.dca1f:
            min_rep = self.GL_OBJETIVO // len(self.parcelas) + 1
        elif self.analisis_objetivo == Analisis.dbca1f:
            min_rep = self.GL_OBJETIVO // len(self.parcelas)
        elif self.analisis_objetivo == Analisis.dca2f:
            min_rep = self.GL_OBJETIVO // (len(self.parcelas) * len(self.ambientes)) + 1

        self.repeticiones = max(self.repeticiones, min_rep)


    def _chequear_ambientes(self):

        if self.todos_los_ambientes:

            self.ambientes = self.vector_muestra.listar_atributo(self.vector_muestra.campo_ambiente)

            for parcela in self.datos_dict:
                if len(self.datos_dict[parcela].keys()) != len(self.ambientes):
                    raise Exception("Los ambientes del muestreo [{a_dict}] no coinciden con los ambientes del shp de ambientes [{a_shp}]!".format(a_dict=', '.join(self.datos_dict[parcela].keys()), a_shp=', '.join(self.ambientes)))
        
        else:

            ambientes_dict = {}
            for parcela in self.datos_dict:
                for ambiente in self.datos_dict[parcela]:
                    if ambiente not in ambientes_dict:
                        ambientes_dict[ambiente] = 0
                    ambientes_dict[ambiente] += 1
                
            arcpy.AddMessage("Ambientes por parcela: [{a}] ({n})".format(a=', '.join(ambientes_dict), n=len(ambientes_dict)))
            
            ambientes_dict_c = copy.deepcopy(ambientes_dict)
            for ambiente in ambientes_dict_c:
                if ambientes_dict[ambiente] < len(self.parcelas):
                    del ambientes_dict[ambiente]
            
            self.ambientes = ambientes_dict.keys()
            if len(self.ambientes) == 0:
                raise Exception("Todas las parcelas quedan sin ambientes!")
            
            # Limpia el dict de los ambientes que no tienen todas las parcelas
            cells_dict_c = copy.deepcopy(self.datos_dict)
            for parcela in cells_dict_c:
                for ambiente in cells_dict_c[parcela]:
                    if ambiente not in ambientes_dict:
                        del self.datos_dict[parcela][ambiente]
                        arcpy.AddMessage("Ambiente {ambiente} eliminado de {parcela} por falta de contraparte".format(ambiente=ambiente, parcela=parcela))

        arcpy.AddMessage("Ambientes finales por parcela: [{a}] ({n})".format(a=', '.join(self.ambientes), n=len(self.ambientes)))


    def _chequear_parcelas(self):

        self.parcelas = self.vector_muestra.listar_atributo(self.vector_muestra.campo_parcela)
        
        if len(self.datos_dict.keys()) != len(self.parcelas):
            raise Exception("Las parcelas del muestreo [{p_dict}] no coinciden con las parcelas del shp de parcelas [{p_shp}]!".format(p_dict=', '.join(self.datos_dict.keys()), p_shp=', '.join(self.parcelas)))

        arcpy.AddMessage("Parcelas: [{p}] ({n})".format(p=', '.join(self.datos_dict.keys()), n=len(self.parcelas)))


    def _estructurar(self):
        """ 
        {
            "Testigo": {
                "A": {
                    "1": [...],
                    "2": [...]
                },
                "B": {...},
                "C": {...}
            },
            "Ensayo": {...},
        }
        """

        self.datos_dict = {}
        for row in arcpy.SearchCursor(self.vector_muestra.mem_shp):
            area = row.getValue(self.vector_muestra.campo_area)
            parcela = row.getValue(self.vector_muestra.campo_parcela)
            ambiente = row.getValue(self.vector_muestra.campo_ambiente)
            rinde = row.getValue(self.vector_muestra.campo_rendimiento)
            fid = row.getValue(self.vector_muestra.campo_idRinde)

            if round(area, 2) < self.vector_muestra.area_fishnet:
                continue
            
            if parcela not in self.datos_dict:
                self.datos_dict[parcela] = {}

            if ambiente not in self.datos_dict[parcela]:
                self.datos_dict[parcela][ambiente] = {}

            if self.vector_muestra.area_fishnet > 0:
                fishnet_id = row.getValue(self.vector_muestra.campo_idFishnet)
            else:
                fishnet_id = max(len(self.datos_dict[parcela][ambiente].keys()) - 1, 0)

            if fishnet_id not in self.datos_dict[parcela][ambiente]:
                self.datos_dict[parcela][ambiente][fishnet_id] = []
            elif len(self.datos_dict[parcela][ambiente][fishnet_id]) >= self.PUNTOS_MIN + 1 and self.vector_muestra.area_fishnet == 0:
                fishnet_id += 1
                self.datos_dict[parcela][ambiente][fishnet_id] = []
            
            self.datos_dict[parcela][ambiente][fishnet_id].append([rinde,fid])


if __name__ == '__main__':

    ma = PoligonoAmbiente()
    ma.cargar(r'E:\ArcGIS\NewFolder\Belver_Ambientes.shp', 'GRIDCODE')
    extent = ma.leer_extension()

    mp = PoligonoParcela()    
    mp.cargar(r'E:\ArcGIS\NewFolder\Export_Output_2.shp', 'GRIDCODE')
    mp.convertir_campo_parcela() 
    
    fn = PoligonoFishnet()
    fn.cargar(extent, 50)

    mu = PoligonoMuestreo()
    mu.cargar(ma, mp, 0, fn)
    
    ml = PoligonoLote()
    ml.cargar(r'E:\ArcGIS\NewFolder\Export_Output.shp', 'idAlbor', 'Campo')
    lotes = ml.lista_de_lotes()

    rindes = pd.read_excel(r'E:\ArcGIS\NewFolder\rindes.xlsx')
    rindes['ID_Cultivo'] = rindes['ID_Cultivo'].apply(str)

    mr = VectorRindeLote()
    mr.cargar_sin_procesar(r'E:\ArcGIS\NewFolder\belver.shp', 'Yld_Mass_W', ml)
    mr.corregir_rinde(lotes, rindes)

    pre_muestra = PreMuestra(mr, mu, False, True, 20, Analisis.dca1f, 'Test')

    print(pre_muestra.nombre_campo)

    muestra = pre_muestra.muestrear(r'E:\ArcGIS\NewFolder\New Folder', True)
    
    muestra = pre_muestra.muestrear(r'E:\ArcGIS\NewFolder\New Folder', True)

    print(muestra.muestra)