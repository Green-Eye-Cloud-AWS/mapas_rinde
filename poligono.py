from vector import Vector, OUT_TEMP

import arcpy

CAMPO_AREA = 'src_Area'

class Poligono(Vector):

    def __init__(self):
        super(Poligono, self).__init__()

        self.campo_area = None

    def cargar(self, shapefile_raw):

        self.mem_shp = OUT_TEMP + r'\poligono'
        self.cargar_en_memoria(shapefile_raw, self.mem_shp)
        
        self.borrar_otros_campos()

    
    def calcular_campo_area(self):
                
        self.generar_campo(CAMPO_AREA, 'DOUBLE')

        arcpy.CalculateField_management(self.mem_shp, CAMPO_AREA, "!shape.area@hectares!", "PYTHON")

        self.campo_area = CAMPO_AREA
        self.cargar_atributos(CAMPO_AREA)