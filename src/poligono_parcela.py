from poligono import OUT_TEMP, Poligono

import arcpy

CAMPO_PARCELA = 'src_Parcel'
NOMBRE_TESTIGO = 'Testigo'
NOMBRE_ENSAYO = 'Ensayo'

class PoligonoParcela(Poligono):

    def __init__(self):
        super(PoligonoParcela, self).__init__()

        self.campo_parcela = None
        self.nombre_testigo = None
    

    def cargar(self, shapefile_raw, campo_parcela_raw, nombre_testigo=None):

        self.mem_shp = OUT_TEMP + r'\parcelas_shp'
        self.cargar_en_memoria(shapefile_raw, self.mem_shp)
        self.campo_parcela = campo_parcela_raw
        
        if nombre_testigo is not None:
            self.nombrar_testigo(nombre_testigo)

        self.calcular_campo_texto(CAMPO_PARCELA, self.campo_parcela)
        self.campo_parcela = CAMPO_PARCELA
        
        self.borrar_otros_campos()
    

    def convertir_campo_parcela(self):
        rows =  arcpy.UpdateCursor(self.mem_shp)

        for row in rows:
            if int(row.getValue(self.campo_parcela)) == 0:
                row.setValue(self.campo_parcela, NOMBRE_TESTIGO)
            elif int(row.getValue(self.campo_parcela)) > 0:
                row.setValue(self.campo_parcela, NOMBRE_ENSAYO)
            else:
                continue
            
            rows.updateRow(row)
        
        del rows

        self.nombrar_testigo()
    

    def nombrar_testigo(self, nombre_testigo=NOMBRE_TESTIGO):
        if nombre_testigo in self.listar_parcelas():
            self.nombre_testigo = nombre_testigo


    def listar_parcelas(self):
        return self.listar_atributo((self.campo_parcela))


if __name__ == '__main__':

    mp = PoligonoParcela()
    
    mp.cargar(r'E:\ArcGIS\NewFolder\Belver_Ambientes.shp', 'GRIDCODE')
    print(mp.atributos)
    mp.exportar(r'E:\ArcGIS\NewFolder\test.shp')

    print(mp.listar_parcelas())

    mp.convertir_campo_parcela()    
    print(mp.listar_parcelas())


