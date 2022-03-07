#!/usr/bin/env python
# -*- coding: utf-8 -*-

from src import procesar_mapa_desde_arcmap

out_dir = r'E:\ArcGIS\NewFolder\New Folder'
parcelas_shp = r'E:\ArcGIS\NewFolder\Export_Output_2.shp'
marco_shp = r'E:\ArcGIS\NewFolder\marco.shp'
ambientes_shp = r'E:\ArcGIS\NewFolder\Belver_Ambientes.shp'
puntos_shp = r'E:\ArcGIS\NewFolder\belver.shp'
lotes_shp = r'E:\ArcGIS\NewFolder\Export_Output.shp'
rep_config = 20
amend_yield = True
rindes_xlsx = r'E:\ArcGIS\NewFolder\rindes.xlsx'
use_fishnet = True
cell_size = 50
buffer = 5
todos_los_ambientes = False
campo_idAlbor = "idAlbor"
campo_campo = "Campo"
campo_rendimiento = 'Yld_Mass_W'
campo_parcela = 'GRIDCODE'
amend_parcelas = True
campo_ambiente = 'GRIDCODE'
nombre_analisis = "Test"
exportar_aux = True
exportar_intersect = True
exportar_mapa_rinde = False
dca1f_por_ambiente = True
perdonar_falta_cuadrados = True
nombre_testigo = "Testigo"

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