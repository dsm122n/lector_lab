# Lector informes de laboratorio
Este programa intenta extraer los valores de cada examen de los informes de laboratorio
DUDAS y **propinas** CONSULTAR A dsm122n@gmail.com

El extraible y los video-tutoriales se encuentran en **https://drive.google.com/drive/folders/1FOo07qSXZFGCe7Yr27_Hqn-yNu6RSnbj?usp=sharing**

Más información y el programa en python en **https://github.com/dsm122n/lector_lab**


## Cómo usar lector_lab_v2.3

En esta versión solo debes:
1) **guardar los informes** de laboratorio formato pdf en la carpeta "laboratorio pdf", luego 
2) **ejecutar** "lector_lab_v2.exe" y listo :) 
3) La salida queda en el documento excel **"salida/00 Imprimir resultados.xlsx"** listo para **imprimir**. 
4) Luego de imprimir el documento, **elimina los pdf** que descargaste previamente de la carpeta "laboratorio pdf" para que puedas descargar más pdfs

### Antes de usar la primera vez
1) Descarga el archivo "lector_lab_v2.3.zip" desde https://drive.google.com/drive/folders/1FOo07qSXZFGCe7Yr27_Hqn-yNu6RSnbj?usp=sharing
2) Descomprime el archivo en una carpeta (de preferencia en el escritorio para que sea más fácil acceder a ella)
3) elimina el único archivo que viene en la carpeta "laboratorio pdf"


### Algunas cosas nuevas:
- Se agrega VHS y PCR
- Ya no debería copiar datos de orina, pero si de liquido pleural o ascítico :( proximamente podría arreglarlo
- Ahora entrega las columnas ordenadas de buena forma
- Los valores con separadores de miles (ej 1.000) ahora se guardan bien en excel 


-------------------
## Cómo usar lector_lab_v1.1
Descarga los pdf de los examenes en la carpeta "laboratorio pdf". 

Ejecuta el archivo "lector_lab" y espera un momento :) se abrirá una ventana negra que mostrará el proceso y cuando esté listo aparecerá una ventana avisando que el proceso está listo. Sea paciente, se demora un tiempo

En la carpeta "salida" encontrarás un archivo .csv con los examenes. Este lo puedes abrir con excel para hacer lo que desees con los datos e imprimirlos (por lo general, el excel lee vien el que dice "con_puntocomas", pero si se ve raro, puedes intentar abrir el otro).

En la carpeta "salida" dejé un archivo excel llamado "Copiar datos aquí para imprimir.xlsx" que tiene formato listo para ser impreso. Destaca algunos resultados anormales con rojo. Está configurado para imprimirlo en formato carta.

Importante destacar que este programita está destinado a registrar los valores de los examenes de sangre y suero, pero no los de otros fluidos. Si guardas pdfs de OC, análisis de liq ascítico o pleural, u otros, puede que registre los valores de estos examenes (ej, registrar la glucosa de liquido ascítico en misma fila de glucosa en sangre)
