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
'<br>
'<br>
' (c) Guillermo (elGuille) Som, 2020<br>
'------------------------------------------------------------------------------<br>
<br>
