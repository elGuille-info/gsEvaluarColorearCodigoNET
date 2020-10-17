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
<br>
'v1.0.0.23              Añado las opciones de cambiar de mayúsculas a minúsculas, etc. al menú Editor<br>
'                       Creo métodos de extensión para las funciones del cambio de "case"<br>
'                       En GuardarComo pongo que siempre muestre el cuadro de diálogo de guardar<br>
'                           antes hacía una comprobación de solo mostrarlo si el nombre era distinto al<br>
'                           usado cuando se abrió.<br>
'v1.0.0.24              Arreglado que cambie el lenguaje cuando se selecciona del buttonLenguaje<br>
'                       Al crear nuevo fichero, mostrar los valores de posición y tamaño<br>
'                       Al cerrar el formulario, no preguntar si se guarda si no hay texto<br>
'v1.0.0.25  02/Oct/20   Los valores de comparar los guardo directamente en las propiedades compartidas<br>
'                       de la clase CompararString.<br>
'v1.0.0.26              Cambios en Compilar.ColorearHTML lo simplifico,<br>
'                           pero quito lo de poner números de línea (puesto nuevamente)<br>
'                       Pongo nuevamente que se guarden MaxFicsConfig (50) (se muestran en el combo de los ficheros)<br>
'                       pero se muestren en el menú recientes solo los indicados en MaxFicsMenu (9)<br>
'v1.0.0.27              Cambios en Compilar.ColorearHTML para corregir el fallo en la documentación HTML.<br>
'v1.0.0.28              Muestro el número de palabras que tiene el fichero.<br>
'v1.0.0.29              Añado shortcuts a las opciones de mayúsculas/minúsculas<br>
'v1.0.0.30              Si el texto está todo en mayúsculas, al cambiarlo a Título se pegunta antes.<br>
'                           más que nada porque la función no lo cambia si está todo en mayúsculas<br>
'                           y al pulsar en cambiar a título no hacía nada...<br>
'                           pregunto para que no parezca que no funciona :-)<br>
'v1.0.0.31              Al guardar mostrar solo el nombre si el directorio es Documentos<br>
'v1.0.0.32              Al cerrar si no está guardado, dar la oportunidad de cancelar<br>
'                       y ahora, si tiene nombre, no pregunta por el nombre y lo guarda.<br>
'v1.0.0.33              Opciones para quitar los espacios delante, detrás y ambos.<br>
'v1.0.0.34              Al pulsar Intro poner los espacios de indentación como esté la línea anterior<br>
'                           por ahora añade una línea de más...<br>
'v1.0.0.35              Al pulsar ENTER en el combo de buscar, empezar la búsqueda (no empieza)<br>
'v1.0.0.36              Solucionado (al estilo McGiver) el pulsar ENTER para que no añada 2 líneas.<br>
'v1.0.0.37              Solucionado que al pulsar RNTER en el combo buscar siga buscando<br>
'                           (había que usar el evento KeyUp)<br>
'v1.0.0.38  03/Oct/20   Añado el método Compilar.ColoreaSeleccionRichTextBox para colorear solo<br>
'                       el texto seleccionado, para usar si colorearAlCargar o colorearAlEvaluar, en:<br>
'                           cambiar mayúsculas/minúsculas, poner/quitar indentación y comentarios<br>
'v1.0.0.39              Cambio de sitio la forma de saber si deben estar habilitados los botones visibles<br>
'v1.0.0.40              Arreglando poner y quitar indentación para que vaya bien en todos los casos<br>
'v1.0.0.41              Arreglando poner y quitar comentarios<br>
'v1.0.0.42              Arreglado clasificar selección y quitar espacios, etc.<br>
'v1.0.0.43              Arreglado fallo en Compilar.ColorearHTML usando ReplaceWord en vez de Replace<br>
'                       porque se encontraba con "col2 y cambiaba el col de color en "<span color="<br>
'v1.0.0.44              Sigue fallando si se usa color en minúsculas, lo cambia...<br>
'                           no cambia COLOR, ni Color (o cualquier cosa que no sea "color") ya que<br>
'                           ReplaceWord distingue entre mayúsculas y minúsculas.<br>
'v1.0.0.45              Arreglando si en el código hay &lt;span> o &lt;b><br>
'v1.0.0.46              Algunas veces no se ve el richTextBoxCodigo completo<br>
'                           Ni idea de porqué pasó...<br>
'v1.0.0.47              La fuente de richTextBoxSyntax no se cambia, se deja en el tamaño asignado en el diseñador<br>
'v1.0.0.48              Nuevo formato para mostrar los errores y warnings<br>
'v1.0.0.49              Al pulsar en richTextBoxSyntax averiguar la línea y caracter del error<br>
'v1.0.0.50  04/Oct/20   Seleccionar la línea clickeada<br>
'v1.0.0.51              Al pulsar en richTextBoxSyntax selecciona la posición completa del error<br>
'v1.0.0.52              ReplaceWord no distingue entre mayúsculas y minúsculas: StringComparison.OrdinalIgnoreCase<br>
'v1.0.0.53              Añado a Extensiones.ReemplazarSiNoEsta para reemplazar si no está lo que se quiere reemplazar<br>
'                           Ejemplo: si busca private y quiere cambiar por <span color:#0000FF>private</span><br>
'                                    se comprueba que no esté por lo que se quiere cambiar.<br>
'v1.0.0.54              Al pegar texto, colorear lo pegado.<br>
'v1.0.0.55              Hacer lo mismo con pegar recortes. De paso soluciono una cosa y añado otra :-)<br>
'v1.0.0.56              Cambiar la forma del cursor del richTextBoxSyntax (he puesto Hand)<br>
'v1.0.0.57              Pongo barraHerramientasContext como menú contextual del formulario<br>
'v1.0.0.58              <br>
'v1.0.0.59              Añado la clase DiagClassifSpanInfo y la uso como parte del valor devuelto por<br>
'                       Compilar.EvaluaCodigo tipo: Dictionary(Of String, Dictionary(Of String, List(Of DiagClassifSpanInfo)))<br>
'v1.0.0.60  05/Oct/20   Poner un ListBox en vez de richTextBoxSysntax para acceder a la posición de<br>
'                       los errores o de las definiciones del código.<br>
'v1.0.0.61              Cambio DiagClassifSpanInfo por dos clases: DiagInfo y ClassifiedSpanInfo<br>
'                       Compilar.EvaluaCodigo usa: Dictionary(Of String, Dictionary(Of String, List(Of ClassifSpanInfo)))<br>
'                       DiagInfo para los errores de compilación<br>
'                       ClassifSpanInfo para la evaluación del código<br>
'                       Al pulsar en un elemento en el ListBox (de cualquiera de los dos tipos)<br>
'                       se selecciona el texto relacionado.<br>
'v1.0.0.62              Quito el control richTextBoxSyntax<br>
'v1.0.0.63              Añado menú en Herramientas para ocultar el panel de evaluación<br>
'v1.0.0.64              Al abrir el menú de herramientas no se muestran habilitados correctamente<br>
'v1.0.0.65              menuOcultarEvaluar sirve para ocultar o mostrar el panel de evaluación<br>
'                       no estará habilitado si no hay datos que mostrar<br>
'v1.0.0.66              Al compilar o ejecutar, etc. borrar antes la lista de evaluación<br>
'v1.0.0.67              Al evaluar el código en las claves, no se hace distinción entre mayúsculas y minúsculas<br>
'                       CompararString implementa IEqualityComparer(Of String)<br>
'                       para usar con Contains en Compilar.EvaluaCodigo.<br>
'v1.0.0.68              Quito el drawMode del listbox para que NO dibuje los items,<br>
'                       algunas veces da error de memoria<br>
'v1.0.0.69              Asigno DrawMode a owner draw cuando se carga la lista, después lo pongo en normal<br>
'                       pero esto NO repinta (no llama al método DrawItem) los items<br>
'v1.0.0.70              Pruebo varias cosas a ver si se repintan, pero nada...<br>
'v1.0.0.71              A ver si con un timer...<br>
'                       lo deja pintado durante lo que dura el timer... pero después los pone normal<br>
'v1.0.0.72              Pongo el código de DrawItem dentro de un Try/Catch y quito el timer<br>
'                       parece que va bien...<br>
'v1.0.0.73              Si da error, lo añado al listbox... por comprobar... (parece que va bien así)<br>
'v1.0.0.74              Clasifico las palabras clave mostradas.<br>
'v1.0.0.75              Solo hay un elemento del tipo ClassifSpanInfo en cada palabra clave (uso un List(of ClassifSpanInfo))<br>
'                       si no voy a poner todas las palabras (no vale la pena llenar la lista),<br>
'                       el tipo devuelto por EvaluaCodigo debería ser: Dictionary(Of String, Dictionary(Of String, ClassifSpanInfo))<br>
'v1.0.0.76              El error lo ha dado como DiagnosticSeverity.Hidden<br>
'v1.0.0.77              Pongo el panel de evaluación más pequeño (*.2 en vez de *.3)<br>
'v1.0.0.78              Al comentar líneas, lo hace donde empieza el texto, no al principio de la línea<br>
'v1.0.0.79              Pongo el panel de evaluación a *.25<br>
'v1.0.0.80              Arreglado al indentar (lo mismo tenía TAB) se come cosas de después, selecciona algo más<br>
'v1.0.0.81              Guardar la selección del último fichero y ponerla al abrirlo...<br>
'v1.0.0.82              Al cargar o guardar cambiar los TAB por 8 espacios. Puesto como opción<br>
'v1.0.0.83              <br>
'v1.0.0.84  06/Oct/20   Al pulsar HOME que se vaya al primer carácter no espacio o TAB, no al principio.<br>
'<br>
'v1.0.0.85  08/Oct/20   Pensando cómo hacer lo de tener múltiples ficheros: ¿MDI o ventanas independientes?<br>
'v1.0.0.86              Cambio ident. por palab. al mostrar los caracteres y palabras del texto.<br>
'v1.0.0.87              Mostrar .txt al abrir y guardar, y si se colorea al cargar, no hacerlo si no es .vb o .cs<br>
'                       Las extensiones mostradas para Texto: .txt, .log y .md<br>
'                       Añado nueva opción al botón de lenguaje: txt / Texto<br>
'v1.0.0.88              Si es texto no compilar, evaluar, etc. (las opciones estarán deshabilitadas)<br>
'v1.0.0.89              En Abrir y Guardar ahora se comprueba si la extensión es .vb, .cs u otra<br>
'                       para asignar la imagen y el texto del botón lenguaje.<br>
'                       Ya no se comprueba si el código tiene instrucción de VB o C#, ya que si la extensión<br>
'                       no es la que corresponde, se considera Texto.<br>
'v1.0.0.90              Opciones para poner y quitar texto del final de las líneas.<br>
'                       O las seleccionadas o la actual.<br>
'v1.0.0.91              Al pulsar HOME si está en la posición 1 (segundo carácter) no se va al primero<br>
'v1.0.0.92              Comprobar si en el resto de opciones de poner / quitar se debe usar:<br>
'                       Dim lineas() As String = richTextBoxCodigo.SelectedText.TrimEnd(ChrW(13)).Split(vbCr.ToCharArray)<br>
'                       En lugar de sin el .TrimEnd(ChrW(13))<br>
'<br>
'<br>
'<br>
' (c) Guillermo (elGuille) Som, 2020<br>
'------------------------------------------------------------------------------<br>
<br>
<br>
Actualizado el 17 de octubre de 2020 a las 18:11<br>

