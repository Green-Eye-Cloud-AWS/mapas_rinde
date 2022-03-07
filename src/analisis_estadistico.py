#!/usr/bin/env python
# -*- coding: utf-8 -*-

from collections import namedtuple
import math
import openpyxl
import pandas as pd
from scipy import stats
from statsmodels.formula.api import ols

from poligono_ambiente import PoligonoAmbiente
from poligono_fishnet import PoligonoFishnet
from poligono_lote import PoligonoLote
from poligono_muestreo import PoligonoMuestreo
from poligono_parcela import PoligonoParcela
from vector_rinde_lote import VectorRindeLote

from analisis import Analisis
from pre_muestra import PreMuestra
from q import get_Q

import arcpy

# aditividad = out_dir +  r"\aditividad.png"
# resvspred = out_dir +  r"\res_vs_pred.png"
# qqplot = out_dir +  r"\qq_plot.png"

class AnalisisEstadistico(object):


    def __init__(self, dca1f_por_ambiente, muestra):
        self.analisis = muestra.analisis_objetivo
        self.dca1f_por_ambiente = dca1f_por_ambiente
        
        if self.analisis != Analisis.dca1f and self.dca1f_por_ambiente:
            raise Exception('No se puede hacer un DCA por ambiente porque la muestra no fue confeccionada para ese analisis!')

        self._analizar(muestra.muestra, muestra.nombre_analisis, muestra.nombre_campo, muestra.nombre_testigo, muestra.df_perdonados, muestra.excel_path)

    
    def _analizar(self, data, nombre_analisis, nombre_campo, nombre_testigo, aux_forgiveness, out_data_xlsx):
        
        book = openpyxl.load_workbook(out_data_xlsx)
        writer = pd.ExcelWriter(out_data_xlsx, engine='openpyxl')
        writer.book = book

        aux_sheet = pd.DataFrame(columns=["Nombre","Parcelas","Ambiente","Rinde_Parcela","Rinde_Testigo","Diferencia","DMS","Diferencia_estadistica","Campo"])

        if self.dca1f_por_ambiente:
            supuestos = True
            
            for ambiente in data.Ambiente.unique():
                new_data = data[data["Ambiente"] == ambiente]
                arcpy.AddMessage("Ambiente " + ambiente)
                arcpy.AddMessage("")
                supuestos, m_comp_parcelas, dms = self.Pasos(new_data,writer,ambiente) 
                supuestos = supuestos == True and supuestos

                aux_sheet = pd.concat([aux_sheet, self.Resumen(m_comp_parcelas, dms, nombre_analisis, nombre_campo, nombre_testigo, ambiente)], axis=0, sort=False) 

        else:
            supuestos, m_comp_parcelas, dms = self.Pasos(data,writer)
            aux_sheet = pd.concat([aux_sheet, self.Resumen(m_comp_parcelas, dms, nombre_analisis, nombre_testigo, nombre_campo)], axis=0, sort=False) 

        aux_sheet = pd.concat([aux_sheet, aux_forgiveness], axis=0, sort=False)
        aux_sheet.to_excel(writer, "aux", index=False)

        writer.save()
        writer.close()

        if not supuestos:
            raise Exception("No se cumplen los supuestos")


    def DCA1F(self, data, columnas):
        """
        Diseño completo al azar

        Columnas: [y, x]
        """

        arcpy.AddMessage("ANOVA DCA 1F")

        N = len(data[columnas[0]])
        df_a = len(data[columnas[1]].unique()) - 1
        df_e = N - len(data[columnas[1]].unique())
        df_t = N - 1

        print('N: ', N, 'Trat: ', len(data[columnas[1]].unique()))

        grand_mean = data[columnas[0]].mean()
        ssq_a = sum([(data[data[columnas[1]] == i][columnas[0]].mean() - grand_mean) ** 2 for i in data[columnas[1]]])
        ssq_t = sum((data[columnas[0]] - grand_mean) ** 2)
        ssq_e = ssq_t - ssq_a

        ms_a = ssq_a/df_a
        ms_e = ssq_e/df_e

        f_a = ms_a/ms_e

        p_a = stats.f.sf(f_a, df_a, df_e)

        results = {
            "SC": [ssq_a, ssq_e, ssq_t],
            "gl": [df_a, df_e, df_t],
            "CM": [ms_a, ms_e, ""],
            "F": [f_a, "", ""],
            "p-valor": [p_a, "", ""]
        }

        columns = ["SC", "gl", "CM", "F", "p-valor"]

        return pd.DataFrame(results, columns=columns, index=[columnas[1] + "s", "Error", "Total"])


    def DBCA1F(self, data, columna):
        """
        Diseño en bloques completos al azar

        Los bloques son las repeticiones
        """

        arcpy.AddMessage("ANOVA DBCA 1F")

        df_a = len(data.Parcela.unique()) - 1
        df_b = len(data.Ambiente.unique()) - 1
        df_e = df_a * df_b 
        df_t = len(data.Parcela.unique()) * len(data.Ambiente.unique()) - 1

        grand_mean = data[columna].mean()
        ssq_a = sum([(data[data.Parcela == i][columna].mean() - grand_mean) ** 2 for i in data.Parcela])
        ssq_b = sum([(data[data.Ambiente == i][columna].mean() - grand_mean) ** 2 for i in data.Ambiente])
        ssq_t = sum((data[columna] - grand_mean) ** 2)
        ssq_e = ssq_t - ssq_a - ssq_b

        ms_a = ssq_a/df_a
        ms_b = ssq_b/df_b
        ms_e = ssq_e/df_e

        f_a = ms_a/ms_e
        f_b = ms_b/ms_e

        p_a = stats.f.sf(f_a, df_a, df_e)
        p_b = stats.f.sf(f_b, df_b, df_e)

        results = {
            "SC": [ssq_a, ssq_b, ssq_e, ssq_t],
            "gl": [df_a, df_b, df_e, df_t],
            "CM": [ms_a, ms_b, ms_e, ""],
            "F": [f_a, f_b, "", ""],
            "p-valor": [p_a, p_b, "", ""]
        }

        columns = ["SC", "gl", "CM", "F", "p-valor"]

        return pd.DataFrame(results, columns=columns, index=["Parcelas", "Ambientes", "Error", "Total"])


    def DCA2F(self, data, columna):
        """
        Diseño completo al azar con dos factores
        """

        arcpy.AddMessage("ANOVA DCA 2F")
        
        N = len(data[columna])
        df_a = len(data.Parcela.unique()) - 1
        df_b = len(data.Ambiente.unique()) - 1
        df_axb = df_a * df_b
        df_e = N - (len(data.Parcela.unique()) * len(data.Ambiente.unique()))
        df_t = N - 1

        grand_mean = data[columna].mean()
        ssq_a = sum([(data[data.Parcela == i][columna].mean() - grand_mean) ** 2 for i in data.Parcela])
        ssq_b = sum([(data[data.Ambiente == i][columna].mean() - grand_mean) ** 2 for i in data.Ambiente])
        ssq_t = sum((data[columna] - grand_mean) ** 2)

        ssq_e = 0
        for ambiente in data.Ambiente.unique():
            sub_data = data[data.Ambiente == ambiente]
            means = [sub_data[sub_data.Parcela == i][columna].mean() for i in sub_data.Parcela]
            ssq_e += sum((sub_data[columna] - means) ** 2)

        ssq_axb = ssq_t - ssq_a - ssq_b - ssq_e

        ms_a = ssq_a/df_a
        ms_b = ssq_b/df_b
        ms_axb = ssq_axb/df_axb
        ms_e = ssq_e/df_e

        f_a = ms_a/ms_e
        f_b = ms_b/ms_e
        f_axb = ms_axb/ms_e

        p_a = stats.f.sf(f_a, df_a, df_e)
        p_b = stats.f.sf(f_b, df_b, df_e)
        p_axb = stats.f.sf(f_axb, df_axb, df_e)

        results = {
            "SC": [ssq_a, ssq_b, ssq_axb, ssq_e, ssq_t],
            "gl": [df_a, df_b, df_axb, df_e, df_t],
            "CM": [ms_a, ms_b, ms_axb, ms_e, ""],
            "F": [f_a, f_b, f_axb, "", ""],
            "p-valor": [p_a, p_b, p_axb, "", ""]
        }

        columns = ["SC", "gl", "CM", "F", "p-valor"]

        return pd.DataFrame(results, columns=columns, index=["Parcelas", "Ambientes", "Parcelas * Ambientes", "Error", "Total"])


    def DMS(self, k, gl, cme, r):
        """
        Q(alfa, tratamientos (k), gl error) * rtsq(CMe/repeticiones)
        Q(5%,3,8)=4.41
        4.41 * rtsq(0.43333/5) = 1.18
        """

        return get_Q(gl, k) * math.sqrt(cme/r)


    def Supuestos(self, data, writer, ambiente):
        """
        Normalidad
            Residuos absolutos
            Si la cantidad de datos > 50 >> Kolmogorov-Smirnov
            Si la cantidad de datos < 50 >> Shapiro-Wilks
            QQ-plot

        Homocedasticidad
            Levene: ANOVA usando residuos absolutos
            Res vs Pred: Residuos estudentizados versus valores predichos

        Aditividad
            Para bloques o factores
            Promedio de rinde para cada ambiente por tratamiento

        """

        arcpy.AddMessage("SUPUESTOS")
        arcpy.AddMessage("")

        if self.analisis == Analisis.dca1f:
            model = ols("Rinde ~ C(Parcela)", data=data).fit()
        elif self.analisis == Analisis.dbca1f:
            model = ols("Rinde ~ C(Ambiente) + C(Parcela)", data=data).fit() 
        elif self.analisis == Analisis.dca2f:
            model = ols("Rinde ~ C(Ambiente) + C(Parcela) + C(Parcela):C(Ambiente)", data=data).fit() 

        #
        # NORMALIDAD
        #

        arcpy.AddMessage("Normalidad")
        if len(data.Rinde) < 50:
            _, pvalue = stats.shapiro(model.resid)
            pd.DataFrame({"Shapiro-Wilks": [pvalue]}).to_excel(writer, "NORMALIDAD" + ambiente, index=False)
            arcpy.AddMessage(" ".join(["Shapiro-Wilks:",str(pvalue)]))
        else:
            _, pvalue = stats.kstest(model.resid, "norm")
            pd.DataFrame({"Kolmogorov-Smirnov": [pvalue]}).to_excel(writer, "NORMALIDAD" + ambiente, index=False)
            arcpy.AddMessage(" ".join(["Kolmogorov-Smirnov:",str(pvalue)]))

        """
        AGREGAR + ambiente
        #stats.probplot(model.resid, dist="norm", plot=plt)
        #arcpy.AddMessage("Ver gráfico")
        #plt.show()
        #plt.savefig(qqplot)
        """

        arcpy.AddMessage("")

        #
        # HOMOCEDASTICIDAD
        #

        data["Residuos_ABS"] = [abs(i) for i in model.resid]
        data["Residuos_Est"] = model.outlier_test()["student_resid"]
        data["Predichos"] = model.fittedvalues

        if self.analisis == Analisis.dca1f:
            aov = self.DCA1F(data, ["Residuos_ABS","Parcela"])
            h_var = aov["p-valor"]["Parcelas"] > 0.05
        elif self.analisis == Analisis.dbca1f:
            aov = self.DBCA1F(data, "Residuos_ABS")
            h_var = aov["p-valor"]["Parcelas"] > 0.05 and aov["p-valor"]["Ambientes"] > 0.05
        elif self.analisis == Analisis.dca2f:
            aov = self.DCA1F(data, ["Residuos_ABS","Tratamiento"])
            h_var = aov["p-valor"]["Tratamientos"] > 0.05

        aov.to_excel(writer, "HOMOCEDASTICIDAD" + ambiente)
        arcpy.AddMessage("Homocedasticidad")
        arcpy.AddMessage("Levene")
        arcpy.AddMessage(aov)
        
        """
        AGREGAR + ambiente
        #data.plot.scatter("Predichos", "Residuos_Est")
        #arcpy.AddMessage("Ver gráfico")
        #plt.show()
        #plt.savefig(resvspred)
        """

        arcpy.AddMessage("")

        #
        # ADITIVIDAD
        #

        """
        if tipo == "dca2f" or tipo == "dcba1f":
            arcpy.AddMessage("Aditividad")
            data.groupby(["Parcela","Ambiente"])["Rinde"].mean().unstack("Parcela").plot.line()
            arcpy.AddMessage("Ver gráfico")
            #plt.show()
            plt.savefig(aditividad)
            arcpy.AddMessage("")
        """
        
        # arcpy.AddMessage("")

        return pvalue > 0.05 and h_var


    def _ComparacionMedias(self, nombres, medias, nombre, dms):

        m_comp = pd.DataFrame([nombres, medias]).transpose()
        m_comp.columns = [nombre, "Medias"]

        m_comp.sort_values(by=["Medias"], ascending=False, inplace=True)
        m_comp.reset_index(drop=True, inplace=True)

        letras_lst = []

        for x in range(len(m_comp.Medias)):
            letra = []

            for y in range(len(m_comp.Medias)):
                if y < x or m_comp.Medias[x] - m_comp.Medias[y] > dms:
                    letra.append("")
                elif y == x or m_comp.Medias[x] - m_comp.Medias[y] <= dms:
                    letra.append("x")
                
            if len(letras_lst) == 0 or max([i for i in range(len(m_comp.Medias)) if letras_lst[-1][i] == "x"]) != max([i for i in range(len(m_comp.Medias)) if letra[i] == "x"]):
                letras_lst.append(letra)

        letras_df = pd.DataFrame(letras_lst).transpose()

        return pd.concat([m_comp,letras_df], axis=1)

    def ComparacionMedias(self, data, df_e, ms_e, writer, nombre_ambiente):
        """
        Tukey 5%
        """

        arcpy.AddMessage("COMPARACIÓN DE MEDIAS")
        arcpy.AddMessage("")

        parcelas = data.Parcela.unique()
        ambientes = data.Ambiente.unique()

        for parcela in parcelas:
            for ambiente in ambientes:
                r = len(data[(data.Parcela == parcela) & (data.Ambiente == ambiente)].Rinde)
                break
            break

        medias = [data[data.Parcela == i].Rinde.mean() for i in parcelas]
        k = len(parcelas)
        n = len(ambientes) * r
        if k > 1:
            dms = self.DMS(k,df_e,ms_e,n)
            arcpy.AddMessage(" ".join(["k:",str(k),"n:",str(n),"gl:",str(df_e),"CMe:",str(round(ms_e,2)),"DMS:",str(round(dms, 2))]))

            m_comp_parcelas = self._ComparacionMedias(parcelas, medias, "Parcelas", dms)
            m_comp_parcelas.to_excel(writer, "TUKEY_Parcela" + nombre_ambiente, index=False)
            arcpy.AddMessage(m_comp_parcelas)
            arcpy.AddMessage("")

        if self.analisis == Analisis.dca2f or self.analisis == Analisis.dbca1f:
            medias = [data[data.Ambiente == i].Rinde.mean() for i in ambientes]
            n = len(parcelas) * r
            k = len(ambientes)
            if k > 1:
                dms = self.DMS(k,df_e,ms_e,n)
                arcpy.AddMessage(" ".join(["k:",str(k),"n:",str(n),"gl:",str(df_e),"CMe:",str(round(ms_e,2)),"DMS:",str(round(dms, 2))]))

                m_comp = self._ComparacionMedias(ambientes, medias, "Ambientes", dms)
                m_comp.to_excel(writer, "TUKEY_Ambiente", index=False)
                arcpy.AddMessage(m_comp)
                arcpy.AddMessage("")

        if self.analisis == Analisis.dca2f:
            tratamientos = data.Tratamiento.unique()

            medias = [data[(data.Parcela == p) & (data.Ambiente == a)].Rinde.mean() for a in ambientes for p in parcelas]
            n = r
            k = len(parcelas) * len(ambientes)
            if k > 1:
                dms = self.DMS(k,df_e,ms_e,n)
                arcpy.AddMessage(" ".join(["k:",str(k),"n:",str(n),"gl:",str(df_e),"CMe:",str(round(ms_e,2)),"DMS:",str(round(dms, 2))]))

                m_comp = self._ComparacionMedias(tratamientos, medias, "Parcelas * Ambientes", dms)
                m_comp.to_excel(writer, "TUKEY_Interaccion", index=False)
                arcpy.AddMessage(m_comp)
                arcpy.AddMessage("")
        
        arcpy.AddMessage("")

        return m_comp_parcelas, dms

    def Pasos(self, data, writer, ambiente=""):

        if ambiente != "":
            ambiente = "_" + ambiente

        if self.analisis == Analisis.dbca1f:
            data = data.groupby(["Parcela","Ambiente"], as_index=False).mean()[["Parcela","Ambiente","Rinde"]]
        elif self.analisis == Analisis.dca2f:
            data["Tratamiento"] = [ "_".join([str(r.Parcela), str(r.Ambiente)]) for i, r in data.iterrows() ]

        arcpy.AddMessage("CABECERA DATOS")
        arcpy.AddMessage("")
        arcpy.AddMessage(data.head())
        arcpy.AddMessage("")

        #
        # SUPUESTOS
        #

        supuestos = self.Supuestos(data,writer,ambiente)

        #
        # ANOVA
        #

        arcpy.AddMessage("ANOVA")
        arcpy.AddMessage("")

        if self.analisis == Analisis.dca1f:
            aov = self.DCA1F(data, ["Rinde","Parcela"])
        elif self.analisis == Analisis.dbca1f:
            aov = self.DBCA1F(data, "Rinde")
        elif self.analisis == Analisis.dca2f:
            aov = self.DCA2F(data, "Rinde")

        aov.to_excel(writer, "ANOVA" + ambiente)
        arcpy.AddMessage(aov)
        arcpy.AddMessage("")
        arcpy.AddMessage("")

        #
        # COMPARACION DE MEDIAS
        #

        df_e = aov.gl["Error"]
        ms_e = aov.CM["Error"]
        m_comp_parcelas, dms = self.ComparacionMedias(data,df_e,ms_e,writer,ambiente)

        return supuestos, m_comp_parcelas, dms

    def Resumen(self, m_comp_parcelas, dms, nombre_analisis, nombre_campo, nombre_testigo, ambiente=""):

        testigo = m_comp_parcelas[m_comp_parcelas["Parcelas"] == nombre_testigo]
        testigo_mean = testigo["Medias"].mean()

        print(testigo)

        columnas = []
        for column in m_comp_parcelas.columns:
            if column != "Parcelas" and column != "Medias" and testigo.iloc[0][column] == "x":
                columnas.append(column)

        parcelas = m_comp_parcelas[m_comp_parcelas["Parcelas"] != nombre_testigo]

        out = pd.DataFrame(columns=["Parcelas","Diferencia_estadistica","Rinde_Parcela","Rinde_Testigo"])

        for _, row in parcelas.iterrows():
            diferencia = True
            for columna in columnas:
                if row[columna] == "x":
                    diferencia = False
                    break
            if diferencia:
                out.loc[len(out)] = {"Parcelas": row["Parcelas"], "Diferencia_estadistica": row["Medias"] - testigo_mean, "Rinde_Parcela": row["Medias"], "Rinde_Testigo": testigo_mean}
            else:
                out.loc[len(out)] = {"Parcelas": row["Parcelas"], "Diferencia_estadistica": 0, "Rinde_Parcela": row["Medias"], "Rinde_Testigo": testigo_mean}
        
        out["Nombre"] = nombre_analisis
        out["Campo"] = nombre_campo
        out["Ambiente"] = ambiente
        out["DMS"] = dms
        out["Diferencia"] = out["Rinde_Parcela"] - out["Rinde_Testigo"]

        return out


def test():

    data = pd.DataFrame({
        "Rinde": [20,25,22,27,21,30,45,30,35,36,31,30,40,35,30,20,21,20,20,19,25,30,29,28,30,30,29,31,30,30,32,35,30,40,30,23,25,28,30,31,24,28,24,25,30,39,42,36,42,40,41,45,40,40,35,24,25,30,26,23,28,31,26,29,32,40,45,50,45,60,42,50,40,55,45,29,30,28,27,30],
        "Ambiente": ['1','1','1','1','1','2','2','2','2','2','3','3','3','3','3','4','4','4','4','4','1','1','1','1','1','2','2','2','2','2','3','3','3','3','3','4','4','4','4','4','1','1','1','1','1','2','2','2','2','2','3','3','3','3','3','4','4','4','4','4','1','1','1','1','1','2','2','2','2','2','3','3','3','3','3','4','4','4','4','4'],
        "Parcela": ['1','1','1','1','1','1','1','1','1','1','1','1','1','1','1','1','1','1','1','1','2','2','2','2','2','2','2','2','2','2','2','2','2','2','2','2','2','2','2','2','3','3','3','3','3','3','3','3','3','3','3','3','3','3','3','3','3','3','3','3','4','4','4','4','4','4','4','4','4','4','4','4','4','4','4','4','4','4','4','4']
    })

    data.to_excel(r'E:\ArcGIS\NewFolder\New Folder\test.xlsx', sheet_name="DATA", index=False)

    dict = {
        'analisis_objetivo': Analisis.dca1f,
        'muestra': data,
        'nombre_analisis': '',
        'nombre_campo': '',
        'nombre_testigo': '1',
        'df_perdonados': None,
        'excel_path': r'E:\ArcGIS\NewFolder\New Folder\test.xlsx'
    }

    muestra = namedtuple("Muestra", dict.keys())(*dict.values())

    AnalisisEstadistico(True, muestra)


def test_mapa():

    ma = PoligonoAmbiente()
    ma.cargar(r'E:\ArcGIS\NewFolder\Belver_Ambientes.shp', 'GRIDCODE')
    extent = ma.leer_extension()

    mp = PoligonoParcela()    
    mp.cargar(r'E:\ArcGIS\NewFolder\Export_Output_2.shp', 'GRIDCODE')
    mp.convertir_campo_parcela() 
    
    fn = PoligonoFishnet()
    fn.cargar(extent, 50)

    mu = PoligonoMuestreo()
    mu.cargar(ma, mp, 5, fn)
    
    ml = PoligonoLote()
    ml.cargar(r'E:\ArcGIS\NewFolder\Export_Output.shp', 'idAlbor', 'Campo')
    lotes = ml.lista_de_lotes()

    rindes = pd.read_excel(r'E:\ArcGIS\NewFolder\rindes.xlsx')
    rindes['ID_Cultivo'] = rindes['ID_Cultivo'].apply(str)

    mr = VectorRindeLote()
    mr.cargar_sin_procesar(r'E:\ArcGIS\NewFolder\belver.shp', 'Yld_Mass_W', ml)
    mr.corregir_rinde(lotes, rindes)

    pre_muestra = PreMuestra(mr, mu, False, True, 20, Analisis.dca1f, 'Test')

    muestra = pre_muestra.muestrear(r'E:\ArcGIS\NewFolder\NewFolder', True)

    AnalisisEstadistico(True, muestra)


if __name__ == '__main__':

    test_mapa()
    
