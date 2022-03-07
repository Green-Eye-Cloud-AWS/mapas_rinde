#!/usr/bin/env python
# -*- coding: utf-8 -*-

from src import procesar_mapa_desde_arcmap

import arcpy

out_dir = arcpy.GetParameterAsText(0) #r"E:\ArcGIS\test"
parcelas_shp = arcpy.GetParameterAsText(1) #out_dir + r"\parcelas.shp"
marco_shp = arcpy.GetParameterAsText(2)
ambientes_shp = arcpy.GetParameterAsText(3) #out_dir + r"\ambientes_2.shp"
puntos_shp = arcpy.GetParameterAsText(4) #out_dir + r"\SG1_Merge.shp"
lotes_shp = arcpy.GetParameterAsText(5) #out_dir + r"\lotes.shp"
rep_config = arcpy.GetParameterAsText(6) #20
amend_yield = arcpy.GetParameterAsText(7) #True
rindes_xlsx = arcpy.GetParameterAsText(8) #out_dir +  r"\rindes.xlsx"
use_fishnet = arcpy.GetParameterAsText(9) #True
cell_size = arcpy.GetParameterAsText(10) #50
buffer = arcpy.GetParameterAsText(11) #0
todos_los_ambientes = arcpy.GetParameterAsText(12) #False
campo_idAlbor = arcpy.GetParameterAsText(13) #"idAlbor"
campo_campo = arcpy.GetParameterAsText(14) #"idAlbor"
campo_rendimiento = arcpy.GetParameterAsText(15) #"Rendimient"
campo_parcela = arcpy.GetParameterAsText(16) #"Parcela"
amend_parcelas = arcpy.GetParameterAsText(17) #True
campo_ambiente = arcpy.GetParameterAsText(18) #"Ambiente"
nombre_analisis = arcpy.GetParameterAsText(19) #"Test"
exportar_aux = arcpy.GetParameterAsText(20) #False
exportar_intersect = arcpy.GetParameterAsText(21) #True
exportar_mapa_rinde = arcpy.GetParameterAsText(22) #True
dca1f_por_ambiente = arcpy.GetParameterAsText(23) #True
perdonar_falta_cuadrados = arcpy.GetParameterAsText(24) #False
nombre_testigo = arcpy.GetParameterAsText(25) #"Testigo"

rindes_xlsx = rindes_xlsx if amend_yield == True or amend_yield == 'true' else None
cell_size = cell_size if use_fishnet == True or use_fishnet == 'true' else 0
procesar_mapa_desde_arcmap(
    ambientes_shp, 
    campo_ambiente, 
    parcelas_shp, 
    campo_parcela, 
    lotes_shp, 
    campo_idAlbor, 
    campo_campo, 
    out_dir, 
    puntos_shp, 
    campo_rendimiento, 
    nombre_testigo, 
    amend_parcelas,
    exportar_aux, 
    buffer, 
    cell_size, 
    marco_shp, 
    rindes_xlsx, 
    'dca1f', 
    exportar_mapa_rinde, 
    todos_los_ambientes, 
    perdonar_falta_cuadrados, 
    rep_config, 
    nombre_analisis, 
    exportar_intersect, 
    dca1f_por_ambiente
)