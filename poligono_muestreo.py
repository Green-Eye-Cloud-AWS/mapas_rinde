from poligono_fishnet import PoligonoFishnet
from poligono_parcela import PoligonoParcela
from poligono_ambiente import PoligonoAmbiente
from poligono import OUT_TEMP, Poligono

import arcpy


class PoligonoMuestreo(Poligono):

    def __init__(self):
        super(PoligonoMuestreo, self).__init__()

        self.area_fishnet = None
        self.largo_buffer = None
        self.campo_parcela = None
        self.nombre_testigo = None
        self.campo_ambiente = None
        self.campo_idFishnet = None


    def cargar(self, poligono_ambiente, poligono_parcela, buffer=0, poligono_fishnet=None, poligono_marco=None):

        if poligono_parcela.nombre_testigo == None:
            raise Exception('No hay franja Testigo!')


        self.mem_shp = OUT_TEMP + r"\intersect_0"
        arcpy.Intersect_analysis([poligono_ambiente.mem_shp, poligono_parcela.mem_shp], self.mem_shp)

        self.cargar_atributos(*poligono_ambiente.atributos)
        self.campo_ambiente = poligono_ambiente.campo_ambiente

        self.cargar_atributos(*poligono_parcela.atributos)
        self.campo_parcela = poligono_parcela.campo_parcela
        self.nombre_testigo = poligono_parcela.nombre_testigo

        self.largo_buffer = buffer
        self.area_fishnet = poligono_fishnet.calcular_area_celda() if poligono_fishnet is not None else 0

        if self.largo_buffer > 0:
            out_buffer = OUT_TEMP + r"\buffer"
            arcpy.Buffer_analysis(self.mem_shp, out_buffer, "-" + str(self.largo_buffer) + " Meters")
            self.mem_shp = out_buffer
        elif poligono_fishnet is None:
            raise Exception("Si no se usa fishnet debe ingresar un valor para el buffer mayor a 0!")
        
        if poligono_fishnet is not None:            
            out_intersect_1 = OUT_TEMP + r"\intersect_1"
            arcpy.Intersect_analysis([self.mem_shp, poligono_fishnet.mem_shp], out_intersect_1)
            self.mem_shp = out_intersect_1
            
            self.cargar_atributos(*poligono_fishnet.atributos)
            self.campo_idFishnet = poligono_fishnet.campo_idFishnet

        if poligono_marco is not None:            
            out_intersect_10 = OUT_TEMP + r"\intersect_10"
            arcpy.Intersect_analysis([self.mem_shp, poligono_marco.mem_shp], out_intersect_10)
            self.mem_shp = out_intersect_10

        self.calcular_campo_area()

        self.borrar_otros_campos()


if __name__ == '__main__':

    ma = PoligonoAmbiente()
    ma.cargar(r'E:\ArcGIS\NewFolder\Belver_Ambientes.shp', 'GRIDCODE')
    extent = ma.leer_extension()

    mp = PoligonoParcela()    
    mp.cargar(r'E:\ArcGIS\NewFolder\Export_Output_2.shp', 'GRIDCODE')
    mp.convertir_campo_parcela() 
    
    fn = PoligonoFishnet()
    fn.cargar(extent, 50)

    mc = Poligono()
    mc.cargar(r'E:\ArcGIS\NewFolder\marco.shp')

    mu = PoligonoMuestreo()
    mu.cargar(ma, mp, 0, fn, mc)
    mu.exportar(r'E:\ArcGIS\NewFolder\New Folder\test.shp')

    print(mu.atributos)

    
