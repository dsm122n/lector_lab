import pandas as pd
import numpy as np
import re 
import pdfplumber
from io import StringIO
from datetime import date

from os import listdir
from os.path import isfile, join

archivos_pdf = [f for f in listdir("laboratorio pdf") if isfile(join("laboratorio pdf", f))]
print(archivos_pdf)

laboratorio = pd.DataFrame({"Fecha": pd.Series(dtype = "str"),
                            "Hora": pd.Series(dtype = "str"),
                            "Creatinina": pd.Series(dtype ="str"),
                            "BUN": pd.Series(dtype ="str"),
                            "Sodio ( Na )": pd.Series(dtype ="str"),
                            "Potasio ( K )": pd.Series(dtype ="str"),
                            "Cloro ( Cl )": pd.Series(dtype ="str"),
                            "Calcio Total": pd.Series(dtype ="str"),
                            "Calcio": pd.Series(dtype ="str"),
                            "Magnesio": pd.Series(dtype ="str"),
                            "Glucosa": pd.Series(dtype ="str"),
                            "Proteinas totales": pd.Series(dtype ="str"),
                            "Albumina": pd.Series(dtype ="str"),
                            "Bilirrubina total": pd.Series(dtype ="str"),
                            "Bilirrubina directa": pd.Series(dtype ="str"),
                            "ASAT/GOT": pd.Series(dtype ="str"),
                            "ALAT/GPT": pd.Series(dtype ="str"),
                            "Gama GT": pd.Series(dtype ="str"),
                            "Fosfatasas Alcalinas": pd.Series(dtype ="str"),
                            "LDH": pd.Series(dtype ="str"),
                            "INR": pd.Series(dtype ="str"),
                            "Tiempo de protrombina %": pd.Series(dtype ="str"),
                            "TTPK": pd.Series(dtype ="str"),
                            "Troponina I": pd.Series(dtype ="str"),
                            "proBNP": pd.Series(dtype ="str"),
                            "pH": pd.Series(dtype ="str"),
                            "pCO2": pd.Series(dtype ="str"),
                            "pO2": pd.Series(dtype ="str"),
                            "HCO3-": pd.Series(dtype ="str"),
                            "BE (B)": pd.Series(dtype ="str"),
                            "sO2": pd.Series(dtype ="str"),
                            "Hb": pd.Series(dtype ="str"),
                            "VCM": pd.Series(dtype ="str"),
                            "HCM": pd.Series(dtype ="str"),
                            "CHCM": pd.Series(dtype ="str"),
                            "Recuento de Plaquetas": pd.Series(dtype ="str"),
                            "Recuento Leucocitos": pd.Series(dtype ="str"),
                            "Segmentados": pd.Series(dtype ="str"),
                            "Eosinófilos": pd.Series(dtype ="str"),
                            "Linfocitos": pd.Series(dtype ="str"),
                            "Monocitos": pd.Series(dtype ="str"),
                            "Basófilos": pd.Series(dtype ="str"),
                            "Neutrofilos": pd.Series(dtype ="str"),
                            "Recuento de Hemoglobina": pd.Series(dtype ="str")})

patron = r" [-+]?\d*\,\d+ | \d+ " #patron para encontrar numeros con decimales
for i, archivo in enumerate(archivos_pdf):
    pdf = pdfplumber.open("laboratorio pdf/" + archivo)
    print(f"\n\ntexto extraido de ----> {archivo}\n\n")
    page = pdf.pages[0]
    posicion_fecha = page.extract_text().find("Fecha y Hora Ingreso Solicitud :") + 32
    posicion_hora = page.extract_text().find("Fecha y Hora Ingreso Solicitud :") + 43
    fecha = page.extract_text()[posicion_fecha:posicion_fecha+10]
    laboratorio.loc[i, "Fecha"] = fecha
    hora = page.extract_text()[posicion_hora:posicion_hora+5]
    laboratorio.loc[i, "Hora"] = hora
    print(f"primer ciclo:{i}")
    for j in range(len(pdf.pages)):
        pagina_actual = pdf.pages[j].crop((28.35, 212.6, 583.65, 680)) # las medidas están en pts
        data = pd.DataFrame(pagina_actual.extract_text().split("\n"))
        for k in range(len(data)):
            if data.at[k, 0].find(":") < 0 or data.at[k, 0].find("Fecha y Hora Ingreso Solicitud") > 0 or data.at[k, 0].find("Tipo de Muestra") > 0 or data.at[k, 0].find("Tipo de muestra") > 0 or data.at[k, 0].find("Criterio de rechazo") > 0:
                data = data.drop(k)
        print("\ncorriginedo\n")
        print(data)
        # lab_limpio = data.loc[data[1]==":"].reset_index(drop=False) 
        print(f"segundo ciclo:{j}")
        for k, examen in enumerate(data.loc[:, 0]):
            if examen.find("pH") >= 0:
                laboratorio.at[i, "pH"] = re.search(patron, examen).group()
            if examen.find("pCO2") >= 0:
                laboratorio.at[i, "pCO2"] = re.search(patron, examen).group()
            if examen.find("pO2") >= 0:
                laboratorio.at[i, "pO2"] = re.search(patron, examen).group()
            if examen.find("HCO3-") >= 0:
                laboratorio.at[i, "HCO3-"] = re.search(patron, examen).group()
            if examen.find("BE (B)") >= 0:
                laboratorio.at[i, "BE (B)"] = re.search(patron, examen).group()
            if examen.find("Sodio ( Na )") >= 0:
                laboratorio.at[i, "Na+"] = re.search(patron, examen).group()
            if examen.find("Potasio ( K )") >= 0:
                laboratorio.at[i, "Potasio ( K )"] = re.search(patron, examen).group()
            if examen.find("Cloro ( Cl )") >= 0:
                laboratorio.at[i, "Cloro ( Cl )"] = re.search(patron, examen).group()
            if examen.find("Calcio") >= 0:
                laboratorio.at[i, "Calcio"] = re.search(patron, examen).group()
            if examen.find("Ca++") >= 0:
                laboratorio.at[i, "Ca++"] = re.search(patron, examen).group()
            if examen.find("Magnesio") >= 0:
                laboratorio.at[i, "Magnesio"] = re.search(patron, examen).group()
            if examen.find("Glucosa") >= 0:
                laboratorio.at[i, "Glucosa"] = re.search(patron, examen).group()
            if examen.find("Creatinina") >= 0:
                laboratorio.at[i, "Creatinina"] = re.search(patron, examen).group()
            if examen.find("Urea") >= 0:
                laboratorio.at[i, "Urea"] = re.search(patron, examen).group()
            if examen.find("Bilirrubina Total") >= 0:
                laboratorio.at[i, "Bilirrubina Total"] = re.search(patron, examen).group()
            if examen.find("Bilirrubina Directa") >= 0:
                laboratorio.at[i, "Bilirrubina Directa"] = re.search(patron, examen).group()
            if examen.find("Fosfatasas Alcalinas") >= 0:
                laboratorio.at[i, "Fosfatasas Alcalinas"] = re.search(patron, examen).group()
            if examen.find("ASAT/GOT") >= 0:
                laboratorio.at[i, "ASAT/GOT"] = re.search(patron, examen).group()
            if examen.find("ALAT/GPT") >= 0:
                laboratorio.at[i, "ALAT/GPT"] = re.search(patron, examen).group()
            if examen.find("Gama GT") >= 0:
                laboratorio.at[i, "Gama GT"] = re.search(patron, examen).group()
            if examen.find("LDH") >= 0:
                laboratorio.at[i, "LDH"] = re.search(patron, examen).group()
            if examen.find("INR") >= 0:
                laboratorio.at[i, "INR"] = re.search(patron, examen).group()
            if examen.find("Tiempo de protrombina %") >= 0:
                laboratorio.at[i, "Tiempo de protrombina %"] = re.search(patron, examen).group()
            if examen.find("TTPK") >= 0:
                laboratorio.at[i, "TTPK"] = re.search(patron, examen).group()
print(laboratorio)
final = laboratorio.transpose()
print(final)
laboratorio.to_csv("salida/laboratorio2.csv", index=False)
final.to_csv("salida/final2.csv", index=False)
float(re.search(patron, "HCO3- : 18,8 * mmol/L | 21,0 - 29,0 |").group())
re.search(patron, "HCO3- : 18,8 * mmol/L | 21,0 - 29,0 |").group()
