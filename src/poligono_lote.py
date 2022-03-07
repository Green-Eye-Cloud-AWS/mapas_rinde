from poligono import OUT_TEMP, Poligono

CAMPO_LOTE = "src_Lote"
CAMPO_CAMPO = "src_Campo"

class PoligonoLote(Poligono):

    def __init__(self):
        super(PoligonoLote, self).__init__()

        self.campo_lote = None
        self.campo_campo = None
    

    def cargar(self, shapefile_raw, campo_lote_raw, campo_campo_raw):

        self.mem_shp = OUT_TEMP + r"\lotes_shp"
        self.cargar_en_memoria(shapefile_raw, self.mem_shp)
        self.campo_lote = campo_lote_raw
        self.campo_campo = campo_campo_raw

        self.calcular_campo_texto(CAMPO_LOTE, self.campo_lote)
        self.campo_lote = CAMPO_LOTE

        self.calcular_campo_texto(CAMPO_CAMPO, self.campo_campo)
        self.campo_campo = CAMPO_CAMPO

        self.borrar_otros_campos()
 
     
    def lista_de_campos(self):
        return self.listar_atributo((self.campo_campo))


    def lista_de_lotes(self):
        return self.listar_atributo((self.campo_lote))


if __name__ == '__main__':

    ml = PoligonoLote()
    print(ml.atributos)
    
    ml.cargar(r'E:\ArcGIS\NewFolder\Export_Output.shp', 'idAlbor', 'Campo')
    print(ml.lista_de_lotes())