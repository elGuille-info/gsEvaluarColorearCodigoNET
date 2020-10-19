'------------------------------------------------------------------------------
' Evaluar, Compilar código                                          (26/Sep/20)
'
' Uniendo los proyectos: gsEvaluarColorearCodigoNET y gsCompilarEjecutarNET
' sin usar gsColorearNET ni gsCompilarNET usadas en gsCompilarEjecutarNET
'
#Region " Comentarios de los proyectos antes de la unificación "
'
' ColorearSyntaxTree con código de Visual Basic                     (20/Sep/20)
'   renombrado a gsEvaluarColorearCodigoNET
'
' Abrir un fichero de Visual Basic o C#
' Si se compila mostrará los errores que haya, si no, colorea el resultado
' Si no se compila colorea el resultado
' Para .NET 5.0 (inicialmente para .NET Core 3.1)
'
' Empecé usando las clases (adaptadas) de CSharpToVB (Visual Basic) de Paul1956:
'           https://github.com/paul1956/CSharpToVB
' Pero las he reducido (dejado de usar) y las que he dejado es con cambios;
'   Para colorear:
'       Con cambios:    ColorizeSupport.vb (en Compilar.vb)
'
' gsCompilarEjecutarNET                                             (08/Sep/20)
' Utilidad para colorear, compilar y ejecutar código de VB o C#
'
'v1.0.0.9   Opciones de Buscar y Reemplazar.
'           Pongo WordWrap del RichTextBox a False para que no corte las líneas.
'v1.0.0.10  Con panel para buscar y reemplazar y
'           funciones para buscar, buscar siguiente, reemplazar y reemplazar todos.
'           También en el menú de edición están las 5 opciones.
'v1.0.0.11  Nueva opción para compilar sin ejecutar y otras mejoras visuales.
'v1.0.0.12  Se puede indicar la versión de los lenguajes.
'           Se usa Latest para VB y Default (9.0) para C#.
'v1.0.0.13  Añado un menú contextual al editor de código con los comandos de edición.
'v1.0.0.14  Quito la clase gsColorearNET y uso los módulos ColorizeSupport, etc.
'
#End Region
'
'
'v1.0.0.9   27/Sep/20   Haciendo operativa la unión de las dos aplicaciones
'v1.0.0.10  28/Sep/20   Algunos cambios en colorear HTML para quitar los <br>&nbsp;
'v1.0.0.11              Nuevas cosillas
'v1.0.0.12              Acepta Drag & Drop, pone quita comentarios e indentación
'                       arreglo que los nombres de los ficheros no se pusieran al principio
'                       asigno el estado habilitado/deshabilitado de los menús y botones
'v1.0.0.13  29/Sep/20   Usar marcadores en el código
'                       Al buscar/reemplazar, poder buscar tabs y returns (\t y \r)
'                       En la búsqueda/reemplazo se indicarán \t o \r para buscar vbTab y vbCr
'                       Se guarda la visibilidad de las barras de herramientas
'v1.0.0.14              Uso la imagen "bookmark_003_8x10.png" para indicar que hay un marcador
'v1.0.0.15  30/Sep/20   Recortes para lo último copiado y usarlos para pegar (Ctrl+Shift+V)
'v1.0.0.16              Arreglando el bug al reemplazar siguiente y no hay más coincidencias
'v1.0.0.17              Arreglando al seleccionar el lenguaje
'                       Nuevas pestañas en la ventana de opciones y arreglo bug al eliminar/ordenar
'                       usando los botones.
'                       A día de hoy estas son las pestañas y opciones:
'                       General: Cargar al iniciar, colorear al cargar, mostrar líneas al colorear HTML,
'                                Al evaluar colorear y compilar el código.
'                       Ficheros recientes (eliminarlos)
'                       Colores y fuente: Fuente, Tamaño, indentación
'                       Buscar/reemplazar: textos de buscar y reemplazar, Comprobar Case y palabra completa
'                       Edición: recortes de edición (eliminarlos)
'v1.0.0.18              Arreglando que se vaya a otro sitio al escribir, que no quita los comentarios...
'                       Lo de que no se posicionase bien era al poner los marcadores... cambiaba el SelectionStart
'                       Captura doble pulsaciones de teclas: CtrlK+CtrlK, CtrlK+CtrlL, CtrlK+CtrlC, CtrlK+CtrlU
'v1.0.0.19  01/Oct/20   Ajustar la altura del panelHerrmientas según se muestre o no el panel de buscar
'                       por ahora no compruebo el resto de paneles.
'                       No sé si arreglar las posiciones de los bookmarks si cambia el texto...
'                       ya que si se quitan o ponen líneas, se ajustan a la posición que tenían antes, no a la nueva
'v1.0.0.20              Clasificar el texto seleccionado. Captura Shit+AltL, Shit+Alt+S (lo añado al menú Editor)
'                       Añado el menú Editor (con los mismos comandos que toolStripEditor)
'v1.0.0.21              El alto del panel de herramientas se ajusta correctamente
'                       tanto al mostrar/ocultar los paneles como al cambiar el tamaño del Form1
'v1.0.0.22              Añado las opciones de clasificar al formulario de opciones
'v1.0.0.23              Añado las opciones de cambiar de mayúsculas a minúsculas, etc. al menú Editor
'                       Creo métodos de extensión para las funciones del cambio de "case"
'                       En GuardarComo pongo que siempre muestre el cuadro de diálogo de guardar
'                           antes hacía una comprobación de solo mostrarlo si el nombre era distinto al
'                           usado cuando se abrió.
'v1.0.0.24              Arreglado que cambie el lenguaje cuando se selecciona del buttonLenguaje
'                       Al crear nuevo fichero, mostrar los valores de posición y tamaño
'                       Al cerrar el formulario, no preguntar si se guarda si no hay texto
'v1.0.0.25  02/Oct/20   Los valores de comparar los guardo directamente en las propiedades compartidas
'                       de la clase CompararString.
'v1.0.0.26              Cambios en Compilar.ColorearHTML lo simplifico,
'                           pero quito lo de poner números de línea (puesto nuevamente)
'                       Pongo nuevamente que se guarden MaxFicsConfig (50) (se muestran en el combo de los ficheros)
'                       pero se muestren en el menú recientes solo los indicados en MaxFicsMenu (9)
'v1.0.0.27              Cambios en Compilar.ColorearHTML para corregir el fallo en la documentación HTML.
'v1.0.0.28              Muestro el número de palabras que tiene el fichero.
'v1.0.0.29              Añado shortcuts a las opciones de mayúsculas/minúsculas
'v1.0.0.30              Si el texto está todo en mayúsculas, al cambiarlo a Título se pegunta antes.
'                           más que nada porque la función no lo cambia si está todo en mayúsculas
'                           y al pulsar en cambiar a título no hacía nada...
'                           pregunto para que no parezca que no funciona :-)
'v1.0.0.31              Al guardar mostrar solo el nombre si el directorio es Documentos
'v1.0.0.32              Al cerrar si no está guardado, dar la oportunidad de cancelar
'                       y ahora, si tiene nombre, no pregunta por el nombre y lo guarda.
'v1.0.0.33              Opciones para quitar los espacios delante, detrás y ambos.
'v1.0.0.34              Al pulsar Intro poner los espacios de indentación como esté la línea anterior
'                           por ahora añade una línea de más...
'v1.0.0.35              Al pulsar ENTER en el combo de buscar, empezar la búsqueda (no empieza)
'v1.0.0.36              Solucionado (al estilo McGiver) el pulsar ENTER para que no añada 2 líneas.
'v1.0.0.37              Solucionado que al pulsar RNTER en el combo buscar siga buscando
'                           (había que usar el evento KeyUp)
'v1.0.0.38  03/Oct/20   Añado el método Compilar.ColoreaSeleccionRichTextBox para colorear solo
'                       el texto seleccionado, para usar si colorearAlCargar o colorearAlEvaluar, en:
'                           cambiar mayúsculas/minúsculas, poner/quitar indentación y comentarios
'v1.0.0.39              Cambio de sitio la forma de saber si deben estar habilitados los botones visibles
'v1.0.0.40              Arreglando poner y quitar indentación para que vaya bien en todos los casos
'v1.0.0.41              Arreglando poner y quitar comentarios
'v1.0.0.42              Arreglado clasificar selección y quitar espacios, etc.
'v1.0.0.43              Arreglado fallo en Compilar.ColorearHTML usando ReplaceWord en vez de Replace
'                       porque se encontraba con "col2 y cambiaba el col de color en "<span color="
'v1.0.0.44              Sigue fallando si se usa color en minúsculas, lo cambia...
'                           no cambia COLOR, ni Color (o cualquier cosa que no sea "color") ya que
'                           ReplaceWord distingue entre mayúsculas y minúsculas.
'v1.0.0.45              Arreglando si en el código hay &lt;span> o &lt;b>
'v1.0.0.46              Algunas veces no se ve el richTextBoxCodigo completo
'                           Ni idea de porqué pasó...
'v1.0.0.47              La fuente de richTextBoxSyntax no se cambia, se deja en el tamaño asignado en el diseñador
'v1.0.0.48              Nuevo formato para mostrar los errores y warnings
'v1.0.0.49              Al pulsar en richTextBoxSyntax averiguar la línea y caracter del error
'v1.0.0.50  04/Oct/20   Seleccionar la línea clickeada
'v1.0.0.51              Al pulsar en richTextBoxSyntax selecciona la posición completa del error
'v1.0.0.52              ReplaceWord no distingue entre mayúsculas y minúsculas: StringComparison.OrdinalIgnoreCase
'v1.0.0.53              Añado a Extensiones.ReemplazarSiNoEsta para reemplazar si no está lo que se quiere reemplazar
'                           Ejemplo: si busca private y quiere cambiar por <span color:#0000FF>private</span>
'                                    se comprueba que no esté por lo que se quiere cambiar.
'v1.0.0.54              Al pegar texto, colorear lo pegado.
'v1.0.0.55              Hacer lo mismo con pegar recortes. De paso soluciono una cosa y añado otra :-)
'v1.0.0.56              Cambiar la forma del cursor del richTextBoxSyntax (he puesto Hand)
'v1.0.0.57              Pongo barraHerramientasContext como menú contextual del formulario
'v1.0.0.58              
'v1.0.0.59              Añado la clase DiagClassifSpanInfo y la uso como parte del valor devuelto por
'                       Compilar.EvaluaCodigo tipo: Dictionary(Of String, Dictionary(Of String, List(Of DiagClassifSpanInfo)))
'v1.0.0.60  05/Oct/20   Poner un ListBox en vez de richTextBoxSysntax para acceder a la posición de
'                       los errores o de las definiciones del código.
'v1.0.0.61              Cambio DiagClassifSpanInfo por dos clases: DiagInfo y ClassifiedSpanInfo
'                       Compilar.EvaluaCodigo usa: Dictionary(Of String, Dictionary(Of String, List(Of ClassifSpanInfo)))
'                       DiagInfo para los errores de compilación
'                       ClassifSpanInfo para la evaluación del código
'                       Al pulsar en un elemento en el ListBox (de cualquiera de los dos tipos)
'                       se selecciona el texto relacionado.
'v1.0.0.62              Quito el control richTextBoxSyntax
'v1.0.0.63              Añado menú en Herramientas para ocultar el panel de evaluación
'v1.0.0.64              Al abrir el menú de herramientas no se muestran habilitados correctamente
'v1.0.0.65              menuOcultarEvaluar sirve para ocultar o mostrar el panel de evaluación
'                       no estará habilitado si no hay datos que mostrar
'v1.0.0.66              Al compilar o ejecutar, etc. borrar antes la lista de evaluación
'v1.0.0.67              Al evaluar el código en las claves, no se hace distinción entre mayúsculas y minúsculas
'                       CompararString implementa IEqualityComparer(Of String)
'                       para usar con Contains en Compilar.EvaluaCodigo.
'v1.0.0.68              Quito el drawMode del listbox para que NO dibuje los items,
'                       algunas veces da error de memoria
'v1.0.0.69              Asigno DrawMode a owner draw cuando se carga la lista, después lo pongo en normal
'                       pero esto NO repinta (no llama al método DrawItem) los items
'v1.0.0.70              Pruebo varias cosas a ver si se repintan, pero nada...
'v1.0.0.71              A ver si con un timer...
'                       lo deja pintado durante lo que dura el timer... pero después los pone normal
'v1.0.0.72              Pongo el código de DrawItem dentro de un Try/Catch y quito el timer
'                       parece que va bien...
'v1.0.0.73              Si da error, lo añado al listbox... por comprobar... (parece que va bien así)
'v1.0.0.74              Clasifico las palabras clave mostradas.
'v1.0.0.75              Solo hay un elemento del tipo ClassifSpanInfo en cada palabra clave (uso un List(of ClassifSpanInfo))
'                       si no voy a poner todas las palabras (no vale la pena llenar la lista),
'                       el tipo devuelto por EvaluaCodigo debería ser: Dictionary(Of String, Dictionary(Of String, ClassifSpanInfo))
'v1.0.0.76              El error lo ha dado como DiagnosticSeverity.Hidden
'v1.0.0.77              Pongo el panel de evaluación más pequeño (*.2 en vez de *.3)
'v1.0.0.78              Al comentar líneas, lo hace donde empieza el texto, no al principio de la línea
'v1.0.0.79              Pongo el panel de evaluación a *.25
'v1.0.0.80              Arreglado al indentar (lo mismo tenía TAB) se come cosas de después, selecciona algo más
'v1.0.0.81              Guardar la selección del último fichero y ponerla al abrirlo...
'v1.0.0.82              Al cargar o guardar cambiar los TAB por 8 espacios. Puesto como opción
'v1.0.0.83              Publicado en GitHub
'v1.0.0.84  06/Oct/20   Al pulsar HOME que se vaya al primer carácter no espacio o TAB, no al principio.
'
'v1.0.0.85  08/Oct/20   Pensando cómo hacer lo de tener múltiples ficheros: ¿MDI o ventanas independientes?
'v1.0.0.86              Cambio ident. por palab. al mostrar los caracteres y palabras del texto.
'v1.0.0.87              Mostrar .txt al abrir y guardar, y si se colorea al cargar, no hacerlo si no es .vb o .cs
'                       Las extensiones mostradas para Texto: .txt, .log y .md
'                       Añado nueva opción al botón de lenguaje: txt / Texto
'v1.0.0.88              Si es texto no compilar, evaluar, etc. (las opciones estarán deshabilitadas)
'v1.0.0.89              En Abrir y Guardar ahora se comprueba si la extensión es .vb, .cs u otra
'                       para asignar la imagen y el texto del botón lenguaje.
'                       Ya no se comprueba si el código tiene instrucción de VB o C#, ya que si la extensión
'                       no es la que corresponde, se considera Texto.
'v1.0.0.90              Opciones para poner y quitar texto del final de las líneas.
'                       O las seleccionadas o la actual.
'v1.0.0.91              Al pulsar HOME si está en la posición 1 (segundo carácter) no se va al primero
'v1.0.0.92              Comprobar si en el resto de opciones de poner / quitar se debe usar:
'                       Dim lineas() As String = richTextBoxCodigo.SelectedText.TrimEnd(ChrW(13)).Split(vbCr.ToCharArray)
'                       En lugar de sin el .TrimEnd(ChrW(13))
'v1.0.0.93              Usando un formulario MDI
'v1.0.0.94              Si se abre desde MDI.Nuevo no cargar el último fichero
'v1.0.0.95              Nuevas opciones en el menú ventana para maximiar y restaurar el tamaño de las ventanas hijas
'                       Al abrir desde el MDI que elija el fichero a cargar
'v1.0.0.96  09/Oct/20   Asigno el tamaño mínimo para el MDI
'v1.0.0.97              Creo el módulo UtilFormEditor con las opciones de manejar el contenido de los richtextboxs
'                       richTextBoxCodigo y richTextBoxLineas
'v1.0.0.98              Asignando nuevos valores a UtilFormEditor para acceder a:
'                       el form1 activo, el mdi
'v1.0.0.99              Paso casi todos los menús y controles a UtilFormEditor
'v1.0.0.100             Quito el toolbar y casi todos los menús del MDI
'v1.0.0.101             Todos los controles, menús, etc. están en los formularios que corresponden.
'                       Solo queda ir probando...
'v1.0.0.102             Arreglar el tamaño del form1 según se muetren las herramientas.
'v1.0.0.103 10/Oct/20   Paso los eventos de los controles que hay en Form1 al Form1
'v1.0.0.104             Quito de UtilFormEditor las declaraciones de los menús y botones
'                       y cuando se acceden uso CurrentMDI
'v1.0.0.105             Parece que no hace caso a las opciones de mostrar / ocultar los paneles
'v1.0.0.106             Paso todos los eventos del richTextBoxCodigo al Form1
'                       Paso HabilitarBotones al MDI
'v1.0.0.107             Ya hace caso de las opciones de los menús, etc.
'v1.0.0.108             Arreglando el tamaño del form1 al mostrar/ocultar las barras de herramientas
'                       y al maximizarlo
'v1.0.0.109             Al activarse la ventana, ajustar el lenguaje mostrado en buttonLenguaje
'                       y demás valores de las etiquetas... 
'v1.0.0.110             Quito el status del MDI y lo pongo en el form1
'v1.0.0.111             Pongo el Dock: top al panel de herramientas y al maximizar se ven bien los form1
'v1.0.0.112             Arreglar el tamaño y posición del form1 cuando no está maximizado
'v1.0.0.113             Aún hay que arreglar que muestre el nombre del fichero del form1 activo
'v1.0.0.114             Al activar la ventana cambiar el texto de la ventana MDI
'v1.0.0.115             Se resiste el mostrar el Form1 en tamaño normal...
'v1.0.0.116             Vamos mejorando...
'v1.0.0.117             Ya se ajusta bien el tamaño del form1, esté o no maximizado
'v1.0.0.118             Ajustar el texto de la ventana MDI (la he dejado con Editor de código)
'v1.0.0.119             Cuando no estén maximizados los Form1 mostrar el Form1.Text entre corchetes
'                       que es como lo hace cuando están maximizadas
'v1.0.0.120             Al MDI le pongo el título de application.ProductName
'v1.0.0.121             A ver si ahora... (se muestran dos nombres...)
'v1.0.0.122             Al maximizar las ventanas (o ponerlas normal) cambio el título del MDI al predeterminado
'v1.0.0.123 11/Oct/20   Añado el menú de pegar recortes,
'                       pongo un timer para comprobar el portapapeles, se ejecuta cada 30 segundos
'v1.0.0.124             Cambio el intervalo de timerClipBoard a 15 segundos (que se quedan cosas sin copiar)
'v1.0.0.125             Pongo los valores de AutoScaleDimensions a 6, 13
'                       (estaban a 7,15 y en VS no preview cambia la cosa, al menos con FormRecortes)
'v1.0.0.126 13/Oct/20   La ventana de recortes se puede mover pulsando en:
'                       la barra de títulos y la etiqueta Recortes.
'v1.0.0.127 14/Oct/20   Actualizado el Visual Studio y .NET 5.0 RC2
'v1.0.0.128 15/Oct/20   El diseñador de formularios no va... falta la referencia a System.Windows.Forms
'v1.0.0.129             Era que en TargetFramework hay que usar net5.0-windows, pero estaba en net5.0.
'v1.0.0.130 16/Oct/20   Poner en Nuevo que siempre abra una nueva ventana
'v1.0.0.131             Variables a nivel de Form1: codigoNuevo, codigoAnterior y TextoModificado
'v1.0.0.132             Comprobar porqué no se muestra actualizado el nombre de la ventana en
'                       el menú ventanas... solucionado con el evento menuVentana.DropDownOpening
'v1.0.0.133             Ya se muestra el texto en la ventana nueva, era porque se asignaba nombreFichero
'                       y estaba vacío
'v1.0.0.134             Cambio de sitio el menú Editor (estaba después de Editar)
'                       Añado opción de Guardar todos
'v1.0.0.135             Al ejecutar guardar pero no tener en cuenta el nombre de comboFileName
'                       si no el del nombre del fichero del formulario activo.
'v1.0.0.136             Se guardan los nombres de todos los ficheros abiertos
'                       al cargar la aplicación, si cargarUltimo es true, se abren todas las ventanas
'v1.0.0.137             Uso un timer en el evento load del MDI para que se muestre (espero)
'                       el formulario mientras carga los ficheros.
'                       La carga de los últimos formularios la pongo al final de LeerConfig
'v1.0.0.138             Ya no es necesario cargar el último fichero, se carga con nombresUltimos
'v1.0.0.139             Si en la línea de comandos se indica un fichero, si hay varias ventanas abiertas
'                       se abrirá en nueva ventana, si no, en la predeterminada.
'v1.0.0.140             Muestro una ventana de progreso mientras se cargan los últimos ficheros
'                       La ventana siempre estará encima (usando form.TopMost = True)
'v1.0.0.141             Creo el método OnProgreso para llamarlo desde Abrir(fic) y ColorearCodigo
'v1.0.0.142 17/Oct/20   Pongo la ventana de progreso en guardar y guardar todo
'                       Al llamar a OnProcesando se devuelve el porcentaje hecho.
'v1.0.0.143             Mostrar info de la apliciación después de iniciarse.
'v1.0.0.144             Al ejecutar el programa, que la ventana de dotnet se centre en la pantalla
'v1.0.0.145             Pero no hay forma... seguramnente estará minimizada
'v1.0.0.146             Comprobaciones por si al compilar la DLL está siendo usada por otro proceso
'v1.0.0.147             Más pruebas para ver si consigo que se vea la ventana de dotnet
'                       Usando StarInfo en vez de ejecutar directamente.
'                       (cuando tenga de nuevo el monitor comprobaré si es que estaba fuera de rango)
'v1.0.0.148             No muestra los caracteres ni las palabras al cargar desde el inicio
'                       En el último fichero no lo muestra... ???
'v1.0.0.149             Pongo de nuevo el código en Compilar para posicionar la ventana de dotnet.
'                       nada de nada... 
'v1.0.0.150             Quitar el pitido al pulsar ENTER en el editor
'v1.0.0.151             Solucionado que no se mostrara el tamaño y posición de la última ventana
'                       se mostraba los valores "cero" despoués de Inicializar
'v1.0.0.152             Al seleccionar del menú Ficheros>Recientes abrirlo en nueva ventana
'                       si ya está abierto en una de las ventanas, mostrarla.
'v1.0.0.153             El título y el nombre del último fichero de Form1 siempre con el path completo
'v1.0.0.154             Efecto del texto en gris para Buscar y Reemplazar
'v1.0.0.155             Arreglado BUG al buscar, cuando no se encuentra lo buscado
'v1.0.0.156             Se muestra un mensaje de no hallado al buscar o reemplazar
'v1.0.0.157             Se leen y guardan los nombres de los últimos ficheros si no están en blanco
'v1.0.0.158 18/Oct/20   Ajustes en el color del texto de buscar y reemplazar
'                       se muestra Buscar... o Reemplazar... cuando no hay texto, aunque esté en foco
'v1.0.0.159             Al cargar los ficheros, cambia el número del total en la barra de títulos
'                       de la barra de progreso de carga... ¡y no sé porqué!
'v1.0.0.160             Al usar una variable en vez de nombresUltimos.Count va bien
'v1.0.0.161             Pero todos menos el último tienen el mismo nombre (antes también pasaba)
'v1.0.0.162             Cambio la ubicación el fichero de configuración de las ventanas
'                       ahora están en <Ventanas> antes en <Ficheros> y se ve que se liaba con tantos datos
'v1.0.0.163             GuardarConfig lo pongo al cerraqr el MDI (antes estaba al cerrar el Form1)
'v1.0.0.164             Mejora al escribir donde hay texto predeterminado (grisáceo)
'v1.0.0.165             Mostrar por sepado el toolStrip de buscar, no mostrar Reemplazar cuando sea Buscar
'                       Al Reemplazar, mostrar buscar y reemplazar.
'v1.0.0.166 19/Oct/20   Al quitar los comentarios en varias líneas, solo quita a la primera
'                       Es porque ahora en vez de vbCr pone como cambio de línea vbLf
'v1.0.0.167             Creo la función ComprobarFinLinea para saber qué carácter se ha usado.
'v1.0.0.168             Al cerrar una ventana quitar de nombresUltimos el de esa ventana
'v1.0.0.169             Cambio solo seleccionable al comboBoxFileName
'                       Al pulsar ENTER se abre el fichero seleccionado en nueva ventana
'v1.0.0.170             Al pulsar ENTER en Reemplazar, se reemplazará la siguiente
'v1.0.0.171             Quito el comboBoxFileName y uso la colección UltimosFicherosAbiertos
'                       para tener la lista de los últimos ficheros abiertos
'
' POR HACER:
'                       Solucionar lo del pitido al pulsar ENTER

'
'
'
'TODO:
'   17-oct  Usar el combo de los ficheros recientes para mostrar los que están abiertos
'           Al seleccionar uno, se muestra la ventana correspondiente.
'           Para abrir uno nuevo, se usará el menú Ficheros recientes
'
'   16-oct  Arreglar los marcadores, para que cada fichero tenga los suyos.
'   04-oct  Usar plantillas y crearlas para usar con Nuevo...
'           Poder cargar un proyecto o solución con todos los ficheros relacionados
'   
'
'
' (c) Guillermo (elGuille) Som, 2020
'------------------------------------------------------------------------------
Option Strict On
Option Infer On

Imports Microsoft.VisualBasic

Imports System
Imports System.Collections.Generic

Imports System.IO
Imports System.Text
Imports System.Linq
Imports Microsoft.CodeAnalysis
'Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.Emit
Imports Microsoft.CodeAnalysis.CSharp
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports System.Reflection
Imports Microsoft.CodeAnalysis.Text

'Imports gsEvaluarColorearCodigoNET

Public Class Form1

    ''' <summary>
    ''' Nombre del fichero asignado al formulario actual.
    ''' </summary>
    Friend nombreFichero As String

    ''' <summary>
    ''' El nuevo código del editor
    ''' </summary>
    Friend codigoNuevo As String

    ''' <summary>
    ''' El código anterior del editor.
    ''' Usado para comparar si ha habido cambios
    ''' </summary>
    Friend codigoAnterior As String

    ''' <summary>
    ''' Indica si se ha modificado el texto.
    ''' Cuando cambia el texto actual (<see cref="codigoNuevo"/>)
    ''' se comprueba con <see cref="codigoAnterior"/> por si hay cambios.
    ''' </summary>
    Friend Property TextoModificado As Boolean
        Get
            Return richTextBoxCodigo.Modified
        End Get
        Set(value As Boolean)
            If value Then
                labelModificado.Text = "*"
            Else
                labelModificado.Text = " "
            End If
            richTextBoxCodigo.Modified = value
        End Set
    End Property

    '
    ' Los métodos de evento del formulario
    '

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' Asignaciones al richTextBox                               (28/Sep/20)
        ' (por si las cambio por error en el diseñador)
        richTextBoxCodigo.EnableAutoDragDrop = False
        richTextBoxCodigo.AcceptsTab = False

        ' Asignaciones al Form
        Me.KeyPreview = True

        richTextBoxCodigo.ContextMenuStrip = CurrentMDI.rtbCodigoContext

        'AsignaMetodosDeEventos()

        If Me.MdiParent Is CurrentMDI Then
            'Me.CenterToParent()
            Me.BringToFront()
        End If

        'Inicializar()

    End Sub

    Private Sub Form1_FormClosing(sender As Object,
                                  e As FormClosingEventArgs) Handles Me.FormClosing

        ' BUG: Si se pulsa en nuevo, se pega texto y no se guarda   (18/Sep/20)
        ' al cerrar no pregunta si se quiere guardar.
        ' Esto está solucionado en el evento de richTextBoxCodigo.TextChanged
        ' Además de que al llamar a GuardarComo()                   (01/Oct/20)
        ' solo mostraba el diálogo si el nombre era distinto.
        ' GuardarComo debe mostrar siempre el cuadro de diálogo.

        ' Si no hay texto, no comprobar si se debe guardar          (01/Oct/20)
        If richTextBoxCodigo.TextLength > 0 Then
            If TextoModificado OrElse String.IsNullOrEmpty(nombreUltimoFichero) Then
                ' No preguntaba si quería guardar,                      (01/Oct/20)
                ' porque usaba GuadarComo, y si el nombre es el mismo, no mostraba el diálogo

                ' Dar la oportunidad de cancelar para seguir editando   (02/Oct/20)
                Dim res As DialogResult = DialogResult.No
                res = MessageBox.Show($"El texto está modificado,{vbCrLf}¿Quieres guardarlo?{vbCrLf}{vbCrLf}" &
                                      "Pulsa en Cancelar para seguir editando.",
                                      "Texto modificado y no guardado",
                                      MessageBoxButtons.YesNoCancel,
                                      MessageBoxIcon.Question)
                If res = DialogResult.Yes Then
                    ' GuardarComo debe mostrar siempre el cuadro dediálogo.
                    ' Si no tiene nombre, preguntar                 (02/Oct/20)
                    ' Guardar se encarga de llamar a GuardarComo si no tiene nombre
                    Guardar()
                ElseIf res = DialogResult.Cancel Then
                    e.Cancel = True
                    Return
                End If
            End If
            selectionStartAnt = -1
            selectionEndAnt = -1
            If richTextBoxCodigo.SelectionLength > 0 Then
                selectionStartAnt = richTextBoxCodigo.SelectionStart
                selectionEndAnt = richTextBoxCodigo.SelectionLength '+ selectionStartAnt
            End If
        End If

        ' Asignar los valores de la ventana
        'Form1_Resize(Nothing, Nothing)

        'GuardarConfig()

        ' quitar el fichero que tenía abierto esta ventana          (19/Oct/20)
        ' de la lista nombresUltimos
        ' Pero solo si el MDI no se está cerrando
        If Not e.CloseReason = CloseReason.MdiFormClosing Then
            For i = UltimasVentanasAbiertas.Count - 1 To 0 Step -1
                If UltimasVentanasAbiertas(i) = Me.nombreFichero Then
                    UltimasVentanasAbiertas.RemoveAt(i)
                    Exit For
                End If
            Next
        End If

    End Sub

    Private Sub Form1_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        ' Asignar cuál es el Form1 activo
        Form1Activo = Me
        'CurrentMDI.comboBoxFileName.Text = nombreFichero
        If Not String.IsNullOrWhiteSpace(nombreFichero) Then
            Me.Text = nombreFichero
        End If

        If Me.WindowState = FormWindowState.Normal Then
            CurrentMDI.Text = $"{MDIPrincipal.TituloMDI} [{Me.Text}]"
        End If
    End Sub

    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        If inicializando Then Return

        'If Me.WindowState = FormWindowState.Normal Then
        '    tamForm = (Me.Left, Me.Top, Me.Height, Me.Width)
        'Else
        '    tamForm = (Me.RestoreBounds.Left, Me.RestoreBounds.Top,
        '               Me.RestoreBounds.Height, Me.RestoreBounds.Width)
        'End If

        UtilFormEditor.LabelInfo.Text = $"MDIPrincipal: Ancho: {CurrentMDI.Width}, Alto: {CurrentMDI.Height} / ClientSize: Ancho: {CurrentMDI.ClientSize.Width}, Alto: {CurrentMDI.ClientSize.Height} - Form1: Ancho: {Me.Width}, Alto: {Me.Height}"
    End Sub

    ' Para doble pulsación de teclas
    Private CtrlK As Integer
    Private CtrlC As Integer
    Private CtrlU As Integer
    Private CtrlL As Integer
    Private ShiftAltL As Integer
    Private ShiftAltS As Integer

    Public Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.Control AndAlso e.Shift Then
            'If e.KeyCode = Keys.V Then
            '    e.Handled = True
            '    MostrarRecortes()
            'End If
        ElseIf e.Shift AndAlso e.Alt Then
            ' si se ha pulsado Shit y Alt
            If e.KeyCode = Keys.L Then
                ShiftAltL += 1
            ElseIf e.KeyCode = Keys.S Then
                ShiftAltS += 1
            End If
            If ShiftAltL = 1 AndAlso ShiftAltS = 1 Then
                ClasificarSeleccion()
            End If
        ElseIf e.Control AndAlso Not e.Shift AndAlso Not e.Alt Then
            ' Solo se ha pulsado la tecla Ctrl
            ' comprobar el resto de combinaciones
            If e.KeyCode = Keys.K Then
                CtrlK += 1
            ElseIf e.KeyCode = Keys.C Then
                CtrlC += 1
            ElseIf e.KeyCode = Keys.U Then
                CtrlU += 1
            ElseIf e.KeyCode = Keys.L Then
                CtrlL += 1
            End If
            If CtrlK = 1 AndAlso CtrlC = 1 Then
                ' Ctrl+K, Ctrl+C
                CtrlK = 0
                CtrlC = 0
                PonerComentarios(richTextBoxCodigo)

            ElseIf CtrlK = 1 AndAlso CtrlU = 1 Then
                ' Ctrl+K, Ctrl+U
                CtrlK = 0
                CtrlU = 0
                QuitarComentarios(richTextBoxCodigo)

            ElseIf CtrlK = 1 AndAlso CtrlL = 1 Then
                ' Ctrl+K, Ctrl+L
                CtrlK = 0
                CtrlL = 0
                ' preguntar
                CurrentMDI.buttonEditorMarcadorQuitarTodos.PerformClick()

            ElseIf CtrlK = 2 Then
                ' Ctrl+K, Ctrl+K
                CtrlK = 0
                MarcadorPonerQuitar()

            End If
        ElseIf e.Control = False AndAlso e.Alt = False AndAlso e.Shift = False Then
            ' Si no tiene control, etc.

            ' Comprobar la tecla HOME                               (06/Oct/20)
            ' Esto NO hacerla en el eveto del RichTextBoxCodigo
            If e.KeyCode = Keys.Home Then
                e.Handled = True

                inicializando = True
                Dim selStart = richTextBoxCodigo.SelectionStart
                Dim ln = richTextBoxCodigo.GetLineFromCharIndex(richTextBoxCodigo.SelectionStart)
                If richTextBoxCodigo.Lines.Length < 1 Then Return
                Dim col As Integer

                If richTextBoxCodigo.Lines(ln).StartsWith(" ") Then
                    If Not String.IsNullOrWhiteSpace(richTextBoxCodigo.Lines(ln)) Then
                        col = richTextBoxCodigo.Lines(ln).IndexOf(richTextBoxCodigo.Lines(ln).TrimStart().Substring(0, 1))
                    End If
                    Dim colLine = richTextBoxCodigo.GetFirstCharIndexFromLine(ln)
                    If richTextBoxCodigo.SelectionStart > col + colLine Then
                        richTextBoxCodigo.SelectionStart = colLine + col
                        labelPos.Text = $"Lín: {ln + 1}  Car: {col}"
                    Else
                        richTextBoxCodigo.SelectionStart = colLine
                        labelPos.Text = $"Lín: {ln + 1}  Car: {1}"
                    End If
                Else
                    Dim colLine = richTextBoxCodigo.GetFirstCharIndexFromLine(ln)
                    richTextBoxCodigo.SelectionStart = colLine
                    labelPos.Text = $"Lín: {ln + 1}  Car: {1}"
                End If
                inicializando = False
            End If

        Else
            CtrlK = 0
            CtrlC = 0
            CtrlU = 0
            ShiftAltL = 0
            ShiftAltS = 0

            ' Otras pulsaciones
            ' seguramente captura los TAB del editor
            'richTextBoxCodigo_KeyUp(sender, e)
        End If
    End Sub

    Public Sub Form1_DragEnter(sender As Object, e As DragEventArgs) Handles Me.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) OrElse
                e.Data.GetDataPresent(DataFormats.Text) OrElse
                e.Data.GetDataPresent("System.String") Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Public Sub Form1_DragDrop(sender As Object, e As DragEventArgs) Handles Me.DragDrop
        If e.Data.GetDataPresent("System.String") Then
            Dim fic As String = CType(e.Data.GetData("System.String"), String)
            ' Comprobar que sea una URL o un fichero
            If fic.IndexOf("http") > -1 OrElse fic.IndexOf(":\") > -1 Then
                Abrir(fic)
            End If
        End If
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim fic As String = CType(e.Data.GetData(DataFormats.FileDrop, True), String())(0)
            Abrir(fic)
        ElseIf e.Data.GetDataPresent(DataFormats.Text) Then
            richTextBoxCodigo.SelectedText = e.Data.GetData("System.String", True).ToString()
        End If
    End Sub

    '
    ' Métodos de evento de lstSyntax
    '

    Public Sub lstSyntax_MouseClick(sender As Object, e As MouseEventArgs) Handles lstSyntax.MouseClick
        Dim i = lstSyntax.SelectedIndex
        If i = -1 Then Return
        Dim lst = TryCast(sender, ListBox)

        Dim dcsI = TryCast(lst.Items(i), DiagInfo)
        If dcsI IsNot Nothing Then
            Dim startPosition = dcsI.StartLinePosition
            Dim endPosition = dcsI.EndLinePosition

            Dim pos = richTextBoxCodigo.GetFirstCharIndexFromLine(startPosition.Line)
            richTextBoxCodigo.SelectionStart = pos + startPosition.Character
            pos = richTextBoxCodigo.GetFirstCharIndexFromLine(endPosition.Line)
            richTextBoxCodigo.SelectionLength = (pos + endPosition.Character - richTextBoxCodigo.SelectionStart)
        Else
            Dim csI = TryCast(lst.Items(i), ClassifSpanInfo)
            If csI IsNot Nothing Then
                Dim charStart = csI.StartPosition
                Dim charEnd = csI.EndPosition
                richTextBoxCodigo.SelectionStart = charStart
                richTextBoxCodigo.SelectionLength = charEnd - charStart
            End If
        End If

    End Sub

    Public Sub lstSyntax_MouseMove(sender As Object, e As MouseEventArgs) Handles lstSyntax.MouseMove
        If lstSyntax.Items.Count = 0 Then Return

        Dim point = New Point(e.X, e.Y)
        Dim i = lstSyntax.IndexFromPoint(point)
        If i = -1 Then Return

        If i < lstSyntax.Items.Count Then
            lstSyntax.SelectedIndex = i
        End If

        If lstSyntax.Items(i).ToString.Length > 30 Then
            CurrentMDI.toolTip1.SetToolTip(lstSyntax, lstSyntax.Items(i).ToString)
        Else
            CurrentMDI.toolTip1.SetToolTip(lstSyntax, "")
        End If

    End Sub

    Public Sub lstSyntax_DrawItem(sender As Object, e As DrawItemEventArgs) Handles lstSyntax.DrawItem
        If inicializando OrElse e.Index < 0 Then Return

        Dim lst = TryCast(sender, ListBox)
        If lst Is Nothing Then Return
        If lst.DrawMode = DrawMode.Normal Then Return

        Try
            e.DrawBackground()
            Dim sItem = lst.Items(e.Index).ToString()
            Dim fc = e.ForeColor
            Dim bc = e.BackColor
            Dim laFuente = e.Font
            Dim g = e.Graphics

            Dim dcsI = TryCast(lst.Items(e.Index), DiagInfo)
            If dcsI IsNot Nothing Then
                fc = Color.Yellow
                If dcsI.Severity = DiagnosticSeverity.Error Then
                    bc = Color.Firebrick
                    laFuente = New Font(e.Font, FontStyle.Bold)
                ElseIf dcsI.Severity = DiagnosticSeverity.Warning Then
                    bc = Color.DarkGreen
                ElseIf dcsI.Severity = DiagnosticSeverity.Info Then
                    bc = Color.Navy
                Else ' el error lo da como Hidden
                    bc = Color.Firebrick
                End If
            Else
                bc = Color.FromArgb(250, 250, 250)
                Dim csI = TryCast(lst.Items(e.Index), ClassifSpanInfo)
                If csI IsNot Nothing Then
                    fc = Compilar.GetColorFromName(csI.ClassificationType)
                    'sItem = $"    {csI.Word}"
                Else
                    fc = Color.Black
                    laFuente = New Font(e.Font, FontStyle.Bold)
                End If
            End If

            ' Usar siempre este color
            ' ya que el DrawMode solo lo pngo al llenar el listbox
            'bc = Color.FromArgb(250, 250, 250) ' Color.AliceBlue
            g.FillRectangle(New SolidBrush(bc), e.Bounds)

            'If e.State = DrawItemState.Selected Then
            '    bc = fc
            '    g.FillRectangle(New SolidBrush(bc), e.Bounds)
            '    fc = Color.Yellow
            'ElseIf e.State = DrawItemState.Focus Then
            '    bc = Color.Azure
            '    g.FillRectangle(New SolidBrush(bc), e.Bounds)
            'Else 'If e.State = DrawItemState.Default Then
            '    bc = Color.FromArgb(250, 250, 250) ' Color.AliceBlue
            '    g.FillRectangle(New SolidBrush(bc), e.Bounds)
            'End If
            Using textBrush As New SolidBrush(fc)
                e.Graphics.DrawString(sItem, laFuente, textBrush, e.Bounds.Location, StringFormat.GenericDefault)
            End Using
            e.DrawFocusRectangle()

        Catch ex As Exception
            Dim diag As New DiagInfo()
            diag.Message = ex.Message
            diag.Severity = DiagnosticSeverity.Info
            diag.StartLinePosition = New LinePosition(0, 0)
            diag.EndLinePosition = New LinePosition(0, 0)
            diag.Id = "ERROR"
            lstSyntax.Items.Add(diag)
        End Try

    End Sub

    '
    ' El spliter
    '

    ''' <summary>
    ''' El úlimo ancho del splitContainer2 (28/Sep/20)
    ''' </summary>
    Public splitPanel2 As Integer

    Public Sub splitContainer1_Resize(sender As Object, e As EventArgs) Handles splitContainer1.Resize
        If CurrentMDI.WindowState = FormWindowState.Minimized Then Return

        If splitContainer1.Panel2.Visible Then
            If splitPanel2 < 100 Then
                splitPanel2 = CInt(splitContainer1.Width * 0.25) 'CInt(splitContainer1.Width * 0.3)
            End If
            'splitPanel2 = CInt(splitContainer1.Width * 0.3)
            splitContainer1.SplitterDistance = splitContainer1.Width - splitPanel2
            splitPanel2 = 0
        Else
            If splitPanel2 = 0 Then
                ' asignar el valor para la siguiente vez
                splitPanel2 = splitContainer1.Width - splitContainer1.SplitterDistance
            End If
            splitContainer1.SplitterDistance = splitContainer1.Width
        End If
    End Sub

    '
    ' Eventos del richTextBoxCodigo
    '

    ''' <summary>
    ''' Sincroniza los scrollbar de richTextBoxCodigo y richTextBoxLineas.
    ''' </summary>
    Public Sub richTextBoxCodigo_VScroll(sender As Object,
                                         e As EventArgs) Handles richTextBoxCodigo.VScroll
        If inicializando Then Return
        If String.IsNullOrEmpty(richTextBoxCodigo.Text) Then Return

        Dim nPos As Integer = GetScrollPos(richTextBoxCodigo.Handle, ScrollBarType.SbVert)
        nPos <<= 16
        Dim wParam As Long = ScrollBarCommands.SB_THUMBPOSITION Or nPos
        SendMessage(richTextBoxLineas.Handle, Message.WM_VSCROLL, wParam, 0)
    End Sub

    Public Sub richTextBoxCodigo_KeyUp(sender As Object,
                                       e As KeyEventArgs) Handles richTextBoxCodigo.KeyUp
        ' Se ve que KeyDown falla                                   (28/Sep/20)

        ' TODO: Pasar este código al evento Form1_KeyDown           (06/Oct/20)
        '       ya que se produce antes.

        ' si se pulsa la tecla TAB
        ' insertar 4 espacios en vez de un tabulador

        ' Hay que detectar antes el Shist, Control y Alt
        ' ya que se producen antes que el resto de teclas
        If e.Shift = True AndAlso
            e.Alt = False AndAlso e.Control = False Then
            If e.KeyCode = Keys.Tab Then
                ' Atrás
                e.Handled = True
                QuitarIndentacion(richTextBoxCodigo)
                MostrarPosicion(e)
            End If
        ElseIf e.Alt = False AndAlso e.Control = False Then
            If e.KeyCode = Keys.Tab Then
                ' Adelante
                e.Handled = True
                PonerIndentacion(richTextBoxCodigo)
                MostrarPosicion(e)
            ElseIf e.KeyCode = Keys.Enter Then
                e.Handled = True

                ' Entra dos veces y no sé porqué...

                Dim selStart = richTextBoxCodigo.SelectionStart
                Dim ln = richTextBoxCodigo.GetLineFromCharIndex(richTextBoxCodigo.SelectionStart)
                If richTextBoxCodigo.Lines.Length < 1 Then Return
                Dim col As Integer

                ' ln es el número de línea actual
                ' Si la línea actual (que es la anterior al intro) no está vacía.
                ' Si ln es menor que 1, salir,
                ' seguramente el intro ha llegado por otro lado...

                If ln < 1 Then Return
                'If e.KeyCode = Keys.Enter Then
                If richTextBoxCodigo.Lines(ln - 1) <> "" Then
                    inicializando = True
                    ' Si al quitarle los espacios es una cadena vacía,
                    ' es que solo hay espacios.
                    If richTextBoxCodigo.Lines(ln - 1).TrimStart() = "" Then
                        col = richTextBoxCodigo.Lines(ln - 1).Length
                    Else
                        ' Averiguar la posición del primer carácter,
                        ' aunque puede que haya TABs
                        col = richTextBoxCodigo.Lines(ln - 1).IndexOf(richTextBoxCodigo.Lines(ln - 1).TrimStart().Substring(0, 1))
                    End If
                    richTextBoxCodigo.SelectionStart = selStart
                    richTextBoxCodigo.SelectedText = vbCr & New String(" "c, col)
                    inicializando = False

                    ' Si se da esta condicón, (creo que siempre se da),
                    ' ir al inicio, borrar e ir al final 
                    If richTextBoxCodigo.GetLineFromCharIndex(richTextBoxCodigo.SelectionStart) > ln Then
                        'Debug.Assert(rtEditor.Lines(ln + 1) = "")
                        SendKeys.Send("{HOME}{BS}{END}")
                    End If
                End If
                'End If
                MostrarPosicion(e)
            End If
        End If

    End Sub

    Public Sub richTextBoxCodigo_TextChanged(sender As Object,
                                             e As EventArgs) Handles richTextBoxCodigo.TextChanged
        If inicializando Then Return

        labelTamaño.Text = $"{richTextBoxCodigo.Text.Length:#,##0} car. ({richTextBoxCodigo.Text.CuantasPalabras:#,##0} palab.)"

        codigoNuevo = richTextBoxCodigo.Text
        If String.IsNullOrEmpty(codigoNuevo) Then Return

        ' Añadir los números de línea               (18/Sep/20)
        ' antes de la comparación de codigoAnterior

        ' Añadir los números de línea               (15/Sep/20)
        AñadirNumerosDeLinea()

        If String.IsNullOrEmpty(codigoAnterior) Then
            ' ya se ha pegado el texto              (18/Sep/20)
            ' y si no hay código anterior no se asigna TextoModificado

            TextoModificado = True
            Return
        End If

        ' En realidad no hace falta quitar los vbCr (18/Sep/20)
        'TextoModificado = (codigoAnterior.Replace(vbCr, "").Replace(vbLf, "") <> codigoNuevo.Replace(vbLf, ""))
        TextoModificado = (codigoAnterior <> codigoNuevo)
    End Sub

    Private Sub richTextBoxCodigo_SelectionChanged(sender As Object,
                                                   e As EventArgs) Handles richTextBoxCodigo.SelectionChanged, richTextBoxCodigo.MouseClick
        If inicializando Then Return
        'menuEditDropDownOpenning()
        CurrentMDI.HabilitarBotones()
        MostrarPosicion(Nothing)
    End Sub

    Private Sub richTextBoxCodigo_FontChanged(sender As Object,
                                              e As EventArgs) Handles richTextBoxCodigo.FontChanged
        richTextBoxLineas.Font = New Font(richTextBoxCodigo.Font.FontFamily, richTextBoxCodigo.Font.Size)
    End Sub
End Class
