# Lector informes de laboratorio
Este programa intenta extraer los valores de cada examen de los informes de laboratorio

## Cómo usar lector_lab_v2

En esta versión solo debes:
1) guardar los archivos en la carpeta "laboratorio pdf", luego 
2) ejecutar "lector_lab_v2.exe" y listo :) 

La salida queda en el documento excel "salida/Salida con resultados de examenes.xlsx" listo para imprimir

### Algunas cosas nuevas:
- Se agrega VHS
- Ya no debería copiar datos de orina, pero si de liquido pleural o ascítico :( proximamente podría arreglarlo

## Cómo usar lector_lab_v1.1
Descarga los pdf de los examenes en la carpeta "laboratorio pdf". 

Ejecuta el archivo "lector_lab" y espera un momento :) se abrirá una ventana negra que mostrará el proceso y cuando esté listo aparecerá una ventana avisando que el proceso está listo. Sea paciente, se demora un tiempo

En la carpeta "salida" encontrarás un archivo .csv con los examenes. Este lo puedes abrir con excel para hacer lo que desees con los datos e imprimirlos (por lo general, el excel lee vien el que dice "con_puntocomas", pero si se ve raro, puedes intentar abrir el otro).

En la carpeta "salida" dejé un archivo excel llamado "Copiar datos aquí para imprimir.xlsx" que tiene formato listo para ser impreso. Destaca algunos resultados anormales con rojo. Está configurado para imprimirlo en formato carta.

Importante destacar que este programita está destinado a registrar los valores de los examenes de sangre y suero, pero no los de otros fluidos. Si guardas pdfs de OC, análisis de liq ascítico o pleural, u otros, puede que registre los valores de estos examenes (ej, registrar la glucosa de liquido ascítico en misma fila de glucosa en sangre)
