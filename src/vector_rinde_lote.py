#!/usr/bin/env python
# -*- coding: utf-8 -*-

import numpy as np
import pandas as pd

from poligono_muestreo import PoligonoMuestreo
from poligono_fishnet import PoligonoFishnet
from poligono_parcela import PoligonoParcela
from poligono_ambiente import PoligonoAmbiente
from vector_rinde import VectorRinde, OUT_TEMP
from poligono_lote import PoligonoLote

import arcpy


class VectorRindeLote(VectorRinde):

    def __init__(self):
        super(VectorRindeLote, self).__init__()

        self.campo_lote = None
        self.campo_campo = None


    def cargar_sin_procesar(self, shapefile_raw, campo_rendimiento_raw, poligono_lote):

        VectorRinde.cargar_sin_procesar(self, shapefile_raw, campo_rendimiento_raw)
        self._intersectar_lotes(poligono_lote)
                
        self.borrar_otros_campos()

    
    def _intersectar_lotes(self, poligono_lote):

        out_intersect_3 = OUT_TEMP + r"\intersect_3"
        arcpy.Intersect_analysis([poligono_lote.mem_shp, self.mem_shp], out_intersect_3)
        self.mem_shp = out_intersect_3

        self.cargar_atributos(*poligono_lote.atributos)
        self.campo_lote = poligono_lote.campo_lote
        self.campo_campo = poligono_lote.campo_campo
        

    def corregir_rinde(self, lotes, rindes_reales_df):
        
        for lote in lotes:
            rinde_mapa = np.mean( [ row.getValue(self.campo_rendimiento) for row in arcpy.SearchCursor(self.mem_shp) if row.getValue(self.campo_lote) == lote ] )

            rinde_real = self._buscar_rinde_real(rindes_reales_df, lote)

            factor = 1 + ((rinde_real - rinde_mapa) / rinde_mapa)

            rows =  arcpy.UpdateCursor(self.mem_shp)
            for row in rows:
                if row.getValue(self.campo_lote) == lote:
                    row.setValue(self.campo_rendimiento, row.getValue(self.campo_rendimiento) * factor)
                    rows.updateRow(row)
            del rows

            arcpy.AddMessage("Lote: {lote}, Rinde mapa: {rinde_mapa}, Rinde real: {rinde_real} , Factor: {factor}".format(lote=str(lote), rinde_mapa=str(round(rinde_mapa,2)), rinde_real=str(round(rinde_real,2)), factor=str(round(factor,2))))
    
        arcpy.AddMessage("")


    def _buscar_rinde_real(self, df, lote):
        
        rinde_real = 0

        rindes = df[df['ID_Cultivo'] == lote]['Rinde']
        if len(rindes) > 0:
            rinde_real = rindes.iloc[0]

        if rinde_real > 0:
            return rinde_real

        raise Exception("Rinde real no encontrado en excel")

        
if __name__ == '__main__':
   
    ml = PoligonoLote()
    ml.cargar(r'E:\ArcGIS\NewFolder\Export_Output.shp', 'idAlbor', 'Campo')
    lotes = ml.lista_de_lotes()

    rindes = pd.read_excel(r'E:\ArcGIS\NewFolder\rindes.xlsx')
    rindes['ID_Cultivo'] = rindes['ID_Cultivo'].apply(str)

    mr = VectorRindeLote()
    mr.cargar_sin_procesar(r'E:\ArcGIS\NewFolder\belver.shp', 'Yld_Mass_W', ml)
    mr.corregir_rinde(lotes, rindes)