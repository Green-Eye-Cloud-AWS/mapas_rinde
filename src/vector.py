import time
import arcpy

arcpy.env.overwriteOutput = True
arcpy.env.outputCoordinateSystem = arcpy.SpatialReference("WGS 1984 UTM Zone 21S")
arcpy.Delete_management("in_memory")

OUT_TEMP = "in_memory"

class Vector(object):


    def __init__(self):
        self.mem_shp = None
        self.atributos = set() # ['FID', 'Shape', 'OID']


    def cargar_en_memoria(self, origen, destino):
        try:
            arcpy.CopyFeatures_management(origen, destino)
        except Exception as e:
            arcpy.AddMessage('Error al cargar en memoria')


    def generar_campo(self, nombre, tipo):
        if not nombre in self.lista_de_atributos():
            arcpy.AddField_management(self.mem_shp, nombre, tipo)
        
        arcpy.AddMessage('Campo {nombre} agregado'.format(nombre=nombre))
        arcpy.AddMessage('')
    

    def exportar(self, path):
        arcpy.CopyFeatures_management(self.mem_shp, path)


    def generar_nombre_shp(self):
        timestamp = str(int(time.time()*1000))
        return 'shp_' + timestamp

        
    def cargar_atributos(self, *args):
        for arg in args:
            self.atributos.add(arg)


    def calcular_campo_texto(self, nombre_nuevo, nombre_viejo):
        self.generar_campo(nombre_nuevo, 'TEXT')

        arcpy.CalculateField_management(self.mem_shp, nombre_nuevo, "str(!" + nombre_viejo + "!)", "PYTHON")

        self.cargar_atributos(nombre_nuevo)
    
    
    def calcular_campo_id(self, nombre):
        self.generar_campo(nombre, 'LONG')

        arcpy.CalculateField_management(self.mem_shp, nombre, "!OID!","PYTHON")

        self.cargar_atributos(nombre)


    def lista_de_atributos(self):
        return [field.name for field in arcpy.ListFields(self.mem_shp)]


    def listar_atributo(self, campo):
        return list( set( [ row.getValue(campo) for row in arcpy.SearchCursor(self.mem_shp) ] ) )
    

    def borrar_otros_campos(self):
        fields = [field.name for field in arcpy.ListFields(self.mem_shp) if field.name not in self.atributos]
        for field in fields:
            try:
                arcpy.DeleteField_management(self.mem_shp, [field])
            except Exception:
                pass
    

    def leer_extension(self):
        return arcpy.Describe(self.mem_shp).extent
            

if __name__ == '__main__':
    m = Vector()
    print(m.atributos)