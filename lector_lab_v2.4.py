import pandas as pd
import re 
import pdfplumber
from io import StringIO
from datetime import datetime
import tkinter as tk
from tkinter import messagebox
import xlwings as xw
import shutil
import sys

from os import listdir
from os.path import isfile, join

archivos_pdf = [f for f in listdir("laboratorio pdf") if isfile(join("laboratorio pdf", f))]
print(archivos_pdf)

laboratorio_dma = pd.DataFrame({"dd": [""],
                            "mm": [""],
                            "yyyy": [""],
                            "Hora": [""],
                            "Creatinina": [""],
                            "BUN": [""],
                            "Ca++": [""],
                            "Calcio Total": [""],
                            "Fosforo": [""],
                            "LDH": [""],
                            "Magnesio": [""],
                            "Sodio ( Na )": [""],
                            "Potasio ( K )": [""],
                            "Cloro ( Cl )": [""],
                            "Bilirrubina Total": [""],
                            "Bilirrubina Directa": [""],
                            "ASAT/GOT": [""],
                            "ALAT/GPT": [""],
                            "Gama GT": [""],
                            "Fosfatasas Alcalinas": [""],
                            "Proteina C Reactiva": [""],
                            "Lactato": [""],
                            "Albumina": [""],
                            "Proteinas totales": [""],
                            "Glucosa": [""],
                            "Trigliceridos": [""],
                            "CK Total": [""],
                            "CK MB": [""],
                            "Troponina I": [""],
                            "proBNP": [""],
                            "Hb": [""],
                            "VCM": [""],
                            "CHCM": [""],
                            "Recuento de Plaquetas": [""],
                            "Recuento Leucocitos": [""],
                            "Segmentados": [""],
                            "Eosinófilos": [""],
                            "Linfocitos": [""],
                            "Monocitos": [""],
                            "Basófilos": [""],
                            "VHS": [""],
                            "INR": [""],
                            "Tiempo de protrombina %": [""],
                            "TTPK": [""],
                            "pH": [""],
                            "pO2": [""],
                            "pCO2": [""],
                            "HCO3-": [""],
                            "BE (B)": [""],
                            "sO2": [""]
                            })


# patron = r" [-+]?\d+\,*\d*" #patron para encontrar numeros con decimales

patron = r" [-+]?\d+(\d*\.\d*)*\,*\d*" #patron para encontrar numeros con decimales y separador de miles


for i, archivo in enumerate(archivos_pdf):
    with pdfplumber.open("laboratorio pdf/" + archivo) as pdf:
        print(f"\n\ntexto extraido de ----> {archivo}\n\n")
        page = pdf.pages[0]
        posicion_fecha = page.extract_text().find("Fecha y Hora Ingreso Solicitud :") + 32
        posicion_hora = page.extract_text().find("Fecha y Hora Ingreso Solicitud :") + 43
        day = page.extract_text()[posicion_fecha:posicion_fecha+2]
        month = page.extract_text()[posicion_fecha+3:posicion_fecha+5]
        year = page.extract_text()[posicion_fecha+6:posicion_fecha+10]
        # fecha = page.extract_text()[posicion_fecha:posicion_fecha+10]
        # fecha = datetime.strptime(fecha, '%d-%m-%Y').date()
        laboratorio_dma.loc[i, "dd"] = day
        laboratorio_dma.loc[i, "mm"] = month
        laboratorio_dma.loc[i, "yyyy"] = year
        # laboratorio.loc[i, "Fecha"] = fecha.strftime('%d-%m-%Y')

        hora = page.extract_text()[posicion_hora:posicion_hora+5]
        hora = datetime.strptime(hora, '%H:%M').time()
        laboratorio_dma.loc[i, "Hora"] = hora
        tipo_muestra = page.extract_text().find("Tipo de Muestra :")
        print(f"primer ciclo:{i}")
        for j in range(len(pdf.pages)):
            pagina_actual = pdf.pages[j]
            #.crop((28.35, 212.6, 583.65, 680)) # las medidas están en pts
            data = pd.DataFrame(pagina_actual.extract_text().split("\n"))
            for k in range(len(data)):
                if data.at[k, 0].find(":") < 0 or data.at[k, 0].find("Fecha y Hora Ingreso Solicitud") > 0 or data.at[k, 0].find("Criterio de rechazo") > 0:
                    data = data.drop(k)                
            # print("\ncorriginedo\n")
            # print(data)
            print(f"segundo ciclo:{j}")
            for l, tipo_muestra in enumerate(data.loc[:, 0]):
                if tipo_muestra.find("Tipo de Muestra : Orina") >= 0:
                    data["Tipo de Muestra"] = ["Orina" for i in range(len(data))]
                    break
                else:
                    data["Tipo de Muestra"] = ["Sangre" for i in range(len(data))]
            for k, examen, tipo_muestra in zip(range(len(data)), data.loc[:, 0], data.loc[:, "Tipo de Muestra"]):
                if examen.find("Creatinina") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Creatinina"] = re.search(patron, examen).group()
                if examen.find("Nitrogeno ureico") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "BUN"] = re.search(patron, examen).group()
                if examen.find("Ca++") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Ca++"] = re.search(patron, examen).group()
                if examen.find("Calcio") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Calcio Total"] = re.search(patron, examen).group()
                if examen.find("Fosforo") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Fosforo"] = re.search(patron, examen).group()
                if examen.find("LDH") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "LDH"] = re.search(patron, examen).group()
                if examen.find("Magnesio") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Magnesio"] = re.search(patron, examen).group()
                if examen.find("Sodio ( Na )") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Sodio ( Na )"] = re.search(patron, examen).group()
                if examen.find("Potasio ( K )") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Potasio ( K )"] = re.search(patron, examen).group()
                if examen.find("Cloro ( Cl )") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Cloro ( Cl )"] = re.search(patron, examen).group()
                if examen.find("Bilirrubina Total") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Bilirrubina Total"] = re.search(patron, examen).group()
                if examen.find("Bilirrubina Directa") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Bilirrubina Directa"] = re.search(patron, examen).group()
                if examen.find("ASAT/GOT") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "ASAT/GOT"] = re.search(patron, examen).group()
                if examen.find("ALAT/GPT") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "ALAT/GPT"] = re.search(patron, examen).group()
                if examen.find("Gama GT") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Gama GT"] = re.search(patron, examen).group()
                if examen.find("Fosfatasas Alcalinas") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Fosfatasas Alcalinas"] = re.search(patron, examen).group()
                if examen.find("Proteina C Reactiva") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Proteina C Reactiva"] = re.search(patron, examen).group()
                if examen.find("Lactato") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Lactato"] = re.search(patron, examen).group()
                if examen.find("Albumina") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Albumina"] = re.search(patron, examen).group()
                if examen.find("Proteinas") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Proteinas totales"] = re.search(patron, examen).group()
                if examen.find("Glucosa") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Glucosa"] = re.search(patron, examen).group()
                if examen.find("Trigliceridos") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Trigliceridos"] = re.search(patron, examen).group()
                if examen.find("CK - Total") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "CK Total"] = re.search(patron, examen).group()
                if examen.find("CK MB") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "CK MB"] = re.search(patron, examen).group()
                if examen.find("Troponina I") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Troponina I"] = re.search(patron, examen).group()
                if examen.find("proBNP") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "proBNP"] = re.search(patron, examen).group()
                if examen.find("Hemoglobina") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Hb"] = re.search(patron, examen).group()
                if examen.find("VCM") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "VCM"] = re.search(patron, examen).group()
                if examen.find("CHCM") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "CHCM"] = re.search(patron, examen).group()
                if examen.find("Recuento de Plaquetas") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Recuento de Plaquetas"] = re.search(patron, examen).group()
                if examen.find("Recuento Leucocitos") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Recuento Leucocitos"] = re.search(patron, examen).group()
                if examen.find("Segmentados") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Segmentados"] = re.search(patron, examen).group()
                if examen.find("Eosinófilos") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Eosinófilos"] = re.search(patron, examen).group()
                if examen.find("Linfocitos") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Linfocitos"] = re.search(patron, examen).group()
                if examen.find("Monocitos") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Monocitos"] = re.search(patron, examen).group()
                if examen.find("Basófilos") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Basófilos"] = re.search(patron, examen).group()
                if examen.find("V.H.S.") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "VHS"] = re.search(patron, examen).group()
                if examen.find("INR") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "INR"] = re.search(patron, examen).group()
                if examen.find("Tiempo de protrombina %") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "Tiempo de protrombina %"] = re.search(patron, examen).group()
                if examen.find("TTPK") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "TTPK"] = re.search(patron, examen).group()
                if examen.find("pH") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "pH"] = re.search(patron, examen).group()
                if examen.find("pO2") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "pO2"] = re.search(patron, examen).group()
                if examen.find("pCO2") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "pCO2"] = re.search(patron, examen).group()
                if examen.find("HCO3-") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "HCO3-"] = re.search(patron, examen).group()
                if examen.find("BE (B)") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "BE (B)"] = re.search(patron, examen).group()
                if examen.find("sO2") >= 0 and bool(re.search(patron, examen)) == True and tipo_muestra == "Sangre":
                    laboratorio_dma.at[i, "sO2"] = re.search(patron, examen).group()
# sort by date yyyy-mm-dd

laboratorio_dma.sort_values(by=['yyyy',
                                'mm',
                                'dd', 
                                'Hora'], inplace=True, ascending=[True, True, True, True])
# laboratorio_dma["Fecha"] = laboratorio_dma["Fecha"].astype(str)
laboratorio_dma["Hora"] = laboratorio_dma["Hora"].astype(str)
# delete seconds from time
laboratorio_dma["Hora"] = laboratorio_dma["Hora"].str.slice(0, 5)

# join year month day columns into one at the beginning
laboratorio_dma["Fecha"] = laboratorio_dma["dd"].astype(str) + "-" + laboratorio_dma["mm"].astype(str) + "-" + laboratorio_dma["yyyy"].astype(str)

columnas_reordenadas = ["Fecha"] + [col for col in laboratorio_dma if col != "Fecha"]
laboratorio_dma = laboratorio_dma[columnas_reordenadas]
print (laboratorio_dma)
# delete year month day columns
laboratorio_dma = laboratorio_dma.drop(["yyyy", "mm", "dd"], axis=1)

final = laboratorio_dma.transpose()
print(final)
final.to_csv("salida/resultados_examenes_con_comas.csv", index=True)
final.to_csv("salida/resultados_examenes_con_puntocomas.csv", index=True, sep=";")

source = "salida/Salida en blanco.xlsx"
dest = "salida/Salida con resultados de examenes.xlsx"
shutil.copyfile(source, dest)

with xw.App(visible=False, add_book=False) as app:
    wb = app.books.open("salida/Salida con resultados de examenes.xlsx")
    ws = wb.sheets["Hoja1"]
    ws.range('B2').options(index=False, header = False).value = final
    wb.save()
    wb.close()



root = tk.Tk()
root.withdraw()
message = "La lectura veloz ha finalizado. Los resultados han sido copiados en el siguiente directorio \"salida/Salida con resultados de examenes.xlsx\"\n\nAbre el archivo para imprimirlo (imprime solo las páginas y columnas que necesites)\n\n Puede que se haya abierto una nueva ventana de excel, cierrala :) \n\nPD: se aceptan sugerencias, avisos de errores y ***propinas***, contacte a dsm122n@gmail.com"
top = tk.Toplevel()
top.title("Lectura finalizada!!!")
top.attributes("-topmost", True)  # asegura que la ventana esté siempre en primer plano
tk.Label(top, text=message).pack(padx=20, pady=20)
tk.Button(top, text="Cerrar", command=top.destroy).pack(pady=10)

top.mainloop()
# sys.exit(0)