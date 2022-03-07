from poligono import OUT_TEMP, Poligono

CAMPO_AMBIENTE = 'src_Ambien'

class PoligonoAmbiente(Poligono):

    def __init__(self):
        super(PoligonoAmbiente, self).__init__()

        self.campo_ambiente = None
    

    def cargar(self, shapefile_raw, campo_ambiente_raw):

        self.mem_shp = OUT_TEMP + r'\ambientes_shp'
        self.cargar_en_memoria(shapefile_raw, self.mem_shp)
        self.campo_ambiente = campo_ambiente_raw

        self.calcular_campo_texto(CAMPO_AMBIENTE, self.campo_ambiente)
        self.campo_ambiente = CAMPO_AMBIENTE
        
        self.borrar_otros_campos()


if __name__ == '__main__':

    ma = PoligonoAmbiente()
    
    ma.cargar(r'E:\ArcGIS\NewFolder\Belver_Ambientes.shp', 'GRIDCODE')
    print(ma.atributos)

    ma.calcular_campo_area()

    ma.exportar(r'E:\ArcGIS\NewFolder\test.shp')