# gsEvaluarColorearCodigoNET
Evaluar si tiene fallos, Compilar, Ejecutar y Colorear (el código y para HTML) código de Visual Basic y C#<br>
<br>
<br>
<h3>Sobre gsEvaluarColorearCodigoNET</h3>
Este código lo estoy usando con Visual Studio 2019 Preview<br>
<br>
<h4>Descripción con los cambios</h4>
'------------------------------------------------------------------------------<br>
' Evaluar, Compilar código                                          (26/Sep/20)<br>
'<br>
' Uniendo los proyectos: gsEvaluarColorearCodigoNET y gsCompilarEjecutarNET<br>
' sin usar gsColorearNET ni gsCompilarNET usadas en gsCompilarEjecutarNET<br>
'<br>
'------------ Comentarios de los proyectos antes de la unificación ------------<br>
'<br>
' ColorearSyntaxTree con código de Visual Basic                     (20/Sep/20)<br>
'   renombrado a gsEvaluarColorearCodigoNET<br>
'<br>
' Abrir un fichero de Visual Basic o C#<br>
' Si se compila mostrará los errores que haya, si no, colorea el resultado<br>
' Si no se compila colorea el resultado<br>
' Para .NET 5.0 (inicialmente para .NET Core 3.1)<br>
'<br>
' Empecé usando las clases (adaptadas) de CSharpToVB (Visual Basic) de Paul1956:<br>
'           https://github.com/paul1956/CSharpToVB<br>
' Pero las he reducido (dejado de usar) y las que he dejado es con cambios;<br>
'   Para colorear:<br>
'       Con cambios:    ColorizeSupport.vb (en Compilar.vb)<br>
'<br>
' gsCompilarEjecutarNET                                             (08/Sep/20)<br>
' Utilidad para colorear, compilar y ejecutar código de VB o C#<br>
'<br>
'v1.0.0.9   Opciones de Buscar y Reemplazar.<br>
'           Pongo WordWrap del RichTextBox a False para que no corte las líneas.<br>
'v1.0.0.10  Con panel para buscar y reemplazar y<br>
'           funciones para buscar, buscar siguiente, reemplazar y reemplazar todos.<br>
'           También en el menú de edición están las 5 opciones.<br>
'v1.0.0.11  Nueva opción para compilar sin ejecutar y otras mejoras visuales.<br>
'v1.0.0.12  Se puede indicar la versión de los lenguajes.<br>
'           Se usa Latest para VB y Default (9.0) para C#.<br>
'v1.0.0.13  Añado un menú contextual al editor de código con los comandos de edición.<br>
'v1.0.0.14  Quito la clase gsColorearNET y uso los módulos ColorizeSupport, etc.<br>
'<br>
'------------------------------------------------------------------------------<br>
'<br>
'<br>
'v1.0.0.9   27/Sep/20   Haciendo operativa la unión de las dos aplicaciones<br>
'v1.0.0.10  28/Sep/20   Algunos cambios en colorear HTML para quitar los <br>&nbsp;<br>
'v1.0.0.11              Nuevas cosillas<br>
'v1.0.0.12              Acepta Drag & Drop, pone quita comentarios e indentación<br>
'                       arreglo que los nombres de los ficheros no se pusieran al principio<br>
'                       asigno el estado habilitado/deshabilitado de los menús y botones<br>
'v1.0.0.13  29/Sep/20   Usar marcadores en el código<br>
'                       Al buscar/reemplazar, poder buscar tabs y returns (\t y \r)<br>
'                       En la búsqueda/reemplazo se indicarán \t o \r para buscar vbTab y vbCr<br>
'                       Se guarda la visibilidad de las barras de herramientas<br>
'v1.0.0.14              Uso la imagen "bookmark_003_8x10.png" para indicar que hay un marcador<br>
'v1.0.0.15  30/Sep/20   Recortes para lo último copiado y usarlos para pegar (Ctrl+Shift+V)<br>
'v1.0.0.16              Arreglando el bug al reemplazar siguiente y no hay más coincidencias<br>
'v1.0.0.17              Arreglando al seleccionar el lenguaje<br>
'                       Nuevas pestañas en la ventana de opciones y arreglo bug al eliminar/ordenar<br>
'                       usando los botones.<br>
'                       A día de hoy estas son las pestañas y opciones:<br>
'                       General: Cargar al iniciar, colorear al cargar, mostrar líneas al colorear HTML,<br>
'                                Al evaluar colorear y compilar el código.<br>
'                       Ficheros recientes (eliminarlos)<br>
'                       Colores y fuente: Fuente, Tamaño, indentación<br>
'                       Buscar/reemplazar: textos de buscar y reemplazar, Comprobar Case y palabra completa<br>
'                       Edición: recortes de edición (eliminarlos)<br>
'v1.0.0.18              Arreglando que se vaya a otro sitio al escribir, que no quita los comentarios...<br>
'                       Lo de que no se posicionase bien era al poner los marcadores... cambiaba el SelectionStart<br>
'                       Captura doble pulsaciones de teclas: CtrlK+CtrlK, CtrlK+CtrlL, CtrlK+CtrlC, CtrlK+CtrlU<br>
'v1.0.0.19  01/Oct/20   Ajustar la altura del panelHerrmientas según se muestre o no el panel de buscar<br>
'                       por ahora no compruebo el resto de paneles.<br>
'                       No sé si arreglar las posiciones de los bookmarks si cambia el texto...<br>
'                       ya que si se quitan o ponen líneas, se ajustan a la posición que tenían antes, no a la nueva<br>
'v1.0.0.20              Clasificar el texto seleccionado. Captura Shit+AltL, Shit+Alt+S (lo añado al menú Editor)<br>
'                       Añado el menú Editor (con los mismos comandos que toolStripEditor)<br>
'v1.0.0.21              El alto del panel de herramientas se ajusta correctamente<br>
'                       tanto al mostrar/ocultar los paneles como al cambiar el tamaño del Form1<br>
'v1.0.0.22              Añado las opciones de clasificar al formulario de opciones<br>
'<br>
'<br>
' (c) Guillermo (elGuille) Som, 2020<br>
'------------------------------------------------------------------------------<br>
<br>
<br>
Actualizado el 1 de octubre de 2020 a las 03:38<br>
