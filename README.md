<b>Objetivo:</b>
<br>Analizar dos archivos excel Azure vs Maestro en base a los valores indicados a ser comparados

<b>Precondiciones:</b>
<br>Tener como mínimo JAVA versión 11 instalado y configurado en sus variables de entorno
Tener un computador como una capacidad mínima en RAM de 32GB

<b>Paso a paso:</b>
<br>Para ejecutar el paso a paso debe tener claros los archivos que va a comparar, ya que el programa no diferenciará si usted 
elije dos archivos distintos el uno del otro. El programa está diseñado únicamente para compara los valores que le indique a través del proceso.
Para comenzar el proceso se le compartirá un archivo comprimido con los siguientes elementos:
1. Dos archivos ejecutables, uno con extensión .SH(Para sistema operativo Linux), y .BAT(Para sistema operativo Windows)
2. Una carpeta donde se aloja un ejecutable .JAR
3. Una carpeta "documentos" que contendrá dos carpetas llamadas "ArchivosAzure" y "ArchivosMaestro", carpetas que de las cuales se le recomienda hacer
   uso para alojar los archivos a analizar.
<br><b>Recomendación:</b>Para mas comodidad descomprima la carpeta en el area de Documentos de su computadora.

<b>Durante la ejecución:</b>
1. Dar doble clic sobre el archivo ejecutable correspondiente a su sistema operativo .SH(Linux), .BAT(Windows). Esta acción abrirá una cosola de
   comandos donde se verá parte del proceso de la ejecución.
2. Seleccione el archivo Azure según indicación
3. Seleccione el archivo Maestro según indicación
4. En la ventana emergente empareje las hojas a analizar entre archivo Azure y Maestro
   <br>4.1. Si desea solo analizar un número limitado de hojas, deberá seleccionar las hojas a analizar
   <br>4.2. Si alguno de los match de hojas no coincide, y desea eliminarlo deberá seleccionarlo de la lista mostrada en el recuadro de las hojas unidas y dar
   clic en el botón Eliminar Selecciones
5. Al terminar el match entre hojas, debe dar clic en Terminar Selección
6. A continuación deberá seleccionar el encabezado "código" del archivo Azure según solicitud
7. A continuación deberá seleccionar el encabezado de la fecha de corte del archivo Azure según solicitud
8. A continuación deberá realizar las mismas dos acciones pero con respecto al archivo Maestro
  <br>Nota: en caso de que los encabezados seleccionados con anterioridad no se encuentren en las hojas, se mostrará por pantalla que no se encontraron y el nombre de la hoja
9. Cuando de análisis sea completado el aplicativo avisará con un mensaje emergente.

<br><b>Nota:</b> Los valores comparados se guardarán en dos carpetas que el mismo sistema creará en la carpeta de los archivos Azure.
Las dos carpetas se llamarán "errores" y "messages", en las que encontrará dos archivos excel una con las no coincidencias, y otras con las coincidencias respectivamente.
