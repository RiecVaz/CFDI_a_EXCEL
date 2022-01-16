# CFDI_a_EXCEL
La finalidad es que este pequeño programa sea una herramienta más para ayudar al contador o cualquier persona encargada de hacer declaraciones, automatizando una tarea.
Este proyecto extrae información relevante de todos los CFDI's dentro de una carpeta y la transfiere a excel. Además, agrega un resumen que desglosa los impuestos de cada factura.

### ¿Cómo se usa?
- Seleccione el tipo de operación a realizar (Facturas emitidas o Facturas recibidas).
- Seleccione el el documento de excel donde se generará una nueva hoja con la información deseada.
- Indique un nombre para esa hoja.
- Seleccion en: "Procesar" y espere a que el programa termine de procesar toda la información.

Al final se le mostrará un mensaje que indicará, además, el número de CFDI's procesados.

### Lista de No deducibles
Agregue las razones sociales en lista tal cual aparecen en su factura sin caracteres extra. Esto, le dirá al programa que no desea tomar en cuenta estos impuestos para que estén en su resumen. Ejemplos de razones sociales: 
- HSBC MEXICO, S.A. INSTITUCION DE BANCA MULTIPLE GRUPO FINANCIERO HSBC
- BANCOPPEL, S.A., INSTITUCION DE BANCA MULTIPLE
- BANCO AZTECA SA INSTITUCION DE BANCA MULTIPLE

### Consideraciones
- Asegúrse de que su carpeta con los xml de su interés no tengan nada más que CFDI (.xml)
- No mezcle para una misma operación CFDI de distinta naturaleza, es decir, mezclar emitidas con recibidas. 

En resumen, tenga distintas carpetas para las facturas emitidas y recibidas. Si realiza la operación de "Emitidas" sobre facturas recibidas o viceversa, este programa no funcionará correctamente.

## IMPORTANTE
Este proyecto utiliza los módulos de openpyxl y cfdiclient de Luis Iturrios. Asegúrese de tener dichos módulos.
