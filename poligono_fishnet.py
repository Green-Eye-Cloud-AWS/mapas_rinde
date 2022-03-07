
from poligono_lote import PoligonoLote
from poligono import OUT_TEMP, Poligono

import arcpy

CAMPO_ID_FISHNET = "ID_FISHNET"

class PoligonoFishnet(Poligono):
    
    def __init__(self):
        super(PoligonoFishnet, self).__init__()

        self.largo_celda = None
        self.campo_idFishnet = None
    

    def cargar(self, extension, largo_celda):
        self.mem_shp = OUT_TEMP + r'\fishnet_shp'
        self.largo_celda = largo_celda

        origin_coord = " ".join([str(extension.XMin),str(extension.YMin)])
        y_axis_coord = " ".join([str(extension.XMin),str(extension.YMax)])
        corner_coord = " ".join([str(extension.XMax),str(extension.YMax)])

        arcpy.CreateFishnet_management(self.mem_shp, origin_coord, y_axis_coord, largo_celda, largo_celda, 0, 0, corner_coord, "NO_LABELS", "#", "POLYGON")

        self.calcular_campo_id(CAMPO_ID_FISHNET)
        self.campo_idFishnet = CAMPO_ID_FISHNET

        self.borrar_otros_campos()
    
    
    def calcular_area_celda(self):
        return round(float(self.largo_celda) * float(self.largo_celda) / 10000, 2)


if __name__ == '__main__':

    ml = PoligonoLote()
    
    ml.cargar(r'E:\ArcGIS\NewFolder\Export_Output.shp', 'idAlbor', 'Campo')
    print(ml.lista_de_lotes())
    extent = ml.leer_extension()
    print(extent)

    fn = PoligonoFishnet()
    fn.cargar(extent, 50)
    fn.exportar(r'E:\ArcGIS\NewFolder\test.shp')