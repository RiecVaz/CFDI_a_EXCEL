[![CFDI-a-EXCEL-Miniatura-1.png](https://i.postimg.cc/qRH1nkpn/CFDI-a-EXCEL-Miniatura-1.png)](https://postimg.cc/QBSgZZ9x)

# CFDI_a_EXCEL
La finalidad es que este pequeño programa sea una herramienta más para ayudar al contador o cualquier persona encargada de hacer declaraciones, automatizando una tarea.
Este proyecto extrae información relevante de todos los CFDI's dentro de una carpeta y la transfiere a excel. Además, agrega un resumen que desglosa los impuestos de cada factura.

### Fundamentos y diagrama de flujo:
El diagrama de flujo mostrado abajo, nos muestra el funcionamiento base. Las variables abajo mencionadas como: Al-16%, Al-0%, Subtotal, Exento, Nota. Hacen referencia a columnas que se formarán dentro del excel. Mientras que variables coo Subtotal_factura, Subtotal_calculado, Impuesto_factura, etc. Hacen referencia a variables internas; estos valores fueron calculados de forma interna, o bien, extraídos directamente del CFDI.

[Diagrama-de-Flujo-CFDI.png](https://postimg.cc/5HR0zf6r) <- VER IMAGEN

Existen dos funciones que representan el motor de este programa, la principal es la indicada como "CFDI", al terminar esta función pasará al siguiente CFDI hasta terminar de procesar todos los .xml existentes en la carpeta correspondiente. El proceso llamado: "FORMAR_CONTENIDO", es, en realidad, una función interna que genera un arreglo con la información necesaria para ser insertada en cada fila

La otra función importante, es interna a la función principal y es vital para asignar el valor del impuesto dell CFDI en su correspondiente columna dependiendo el tipo de impuesto y el descuento. 

### ¿Cómo se usa?
[![Showcase.gif](https://i.postimg.cc/FszVYf08/Showcase.gif)](https://postimg.cc/dLMyXVHj) <- Link de Gif para una vista previa.

- Seleccione el tipo de operación a realizar (Facturas emitidas o Facturas recibidas).
- Seleccione el documento de excel donde se generará una nueva hoja con la información deseada.
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
