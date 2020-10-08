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
'v1.0.0.45              Arreglando si en el código hay <span> o <b>
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
'
'
'
'TODO:
'   Usar plantillas y crearlas para usar con Nuevo...
'   Poder tener abiertos más de un fichero
'       Poder cargar un proyecto o solución con todos los ficheros relacionados
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
    ''' Si se está inicializando.
    ''' Usado para que no se provoquen eventos en cadena
    ''' </summary>
    Private inicializando As Boolean = True

    ''' <summary>
    ''' Para saber que no se debe colorear, ni compilar, etc.
    ''' </summary>
    Private Const ExtensionTexto As String = "Texto"

    ''' <summary>
    ''' Si se deben cambiar los TAB por 8 espacios
    ''' </summary>
    Friend cambiarTabs As Boolean = True

    ''' <summary>
    ''' El inicio de la selección al cerrar
    ''' </summary>
    Private selectionStartAnt As Integer
    ''' <summary>
    ''' El final de la selección al cerrar
    ''' </summary>
    Private selectionEndAnt As Integer

    ''' <summary>
    ''' Colección para los últimos recortes del portapapeles.
    ''' Guardar solo 10 recortes
    ''' </summary>
    Friend ColRecortes As New List(Of String)
    ''' <summary>
    ''' El número máximo de recortes.
    ''' </summary>
    Private Const MaxRecortes As Integer = 10

    ''' <summary>
    ''' El número de caracteres para indentar
    ''' </summary>
    Friend EspaciosIndentar As Integer = 4

    ''' <summary>
    ''' Si se deben mostrar los números de línea en el código coloreado como HTML
    ''' </summary>
    Friend Property mostrarLineasHTML As Boolean
        Get
            Return chkMostrarLineasHTML.Checked
        End Get
        Set(value As Boolean)
            chkMostrarLineasHTML.Checked = value
        End Set
    End Property

    ''' <summary>
    ''' El nombre del fichero de configuración
    ''' </summary>
    Private ficheroConfiguracion As String
    ''' <summary>
    ''' Nombre del último fichero asignado al código
    ''' </summary>
    Private nombreUltimoFichero As String
    ''' <summary>
    ''' El directorio de documentos
    ''' Valor predeterminado para guardar los ficheros
    ''' </summary>
    Private dirDocumentos As String

    '''' <summary>
    '''' El nombre del lenguaje VB (como está definido en Compilation.Language)
    '''' </summary>
    'Private Const LenguajeVisualBasic As String = "Visual Basic"
    '''' <summary>
    '''' El nombre del lenguaje C# (como está definido en Compilation.Language)
    '''' </summary>
    'Private Const LenguajeCSharp As String = "C#"

    ''' <summary>
    ''' Tupla para guardar el tamaño y posición del formulario
    ''' </summary>
    Private tamForm As (L As Integer, T As Integer, H As Integer, W As Integer)

    ''' <summary>
    ''' Si se pulsa Ctrl+F o Ctrl+H para indicar 
    ''' los datos a buscar o reemplazar
    ''' </summary>
    Private esCtrlF As Boolean = True
    ''' <summary>
    ''' La posición desde la que se está buscando
    ''' </summary>
    Private buscarPos As Integer
    ''' <summary>
    ''' La posición en que está la primera coincidencia
    ''' al buscar
    ''' </summary>
    Private buscarPosIni As Integer
    ''' <summary>
    ''' Para indicar que se está buscando
    ''' la primera coincidencia de la búsqueda
    ''' </summary>
    Private buscarPrimeraCoincidencia As Boolean = True
    ''' <summary>
    ''' La cadena a buscar
    ''' </summary>
    Private buscarQueBuscar As String
    ''' <summary>
    ''' La cadena a poner cuando se reemplaza
    ''' </summary>
    Private buscarQueReemplazar As String
    ''' <summary>
    ''' Si se busca teniendo en cuenta las mayúsculas y minúsculas
    ''' </summary>
    Friend buscarMatchCase As Boolean
    ''' <summary>
    ''' Si se busca la palabra completa
    ''' </summary>
    Friend buscarWholeWord As Boolean
    ''' <summary>
    ''' El número máximo de items en buscar y reemplazar
    ''' </summary>
    Private Const BuscarMaxItems As Integer = 20
    ''' <summary>
    ''' El nombre de la fuente a usar en el editor
    ''' </summary>
    Friend fuenteNombre As String = "Consolas" 'gsCol.FuentePre
    ''' <summary>
    ''' El tamaño de la fuente a usar en el editor
    ''' </summary>
    Friend fuenteTamaño As String = "11" 'gsCol.FuenteTamPre

    ''' <summary>
    ''' Si se debe cargar el último fichero abierto 
    ''' cuando se cerróa la aplicación
    ''' </summary>
    Friend cargarUltimo As Boolean
    ''' <summary>
    ''' Si se debe colorear el código al abrir el fichero
    ''' </summary>
    Friend colorearAlCargar As Boolean
    ''' <summary>
    ''' Si se debe colorear el código al evaluar
    ''' </summary>
    Friend Property colorearAlEvaluar As Boolean
        Get
            Return buttonColorearAlEvaluar.Checked
        End Get
        Set(value As Boolean)
            buttonColorearAlEvaluar.Checked = value
        End Set
    End Property
    ''' <summary>
    ''' Si se deben comprobar los errores al evaluar.
    ''' Si es <see cref="True"/> se pre-compila el código.
    ''' </summary>
    Friend Property compilarAlEvaluar As Boolean
        Get
            Return buttonCompilarAlEvaluar.Checked
        End Get
        Set(value As Boolean)
            buttonCompilarAlEvaluar.Checked = value
        End Set
    End Property

    ''' <summary>
    ''' El nuevo código del editor
    ''' </summary>
    Private codigoNuevo As String
    ''' <summary>
    ''' El código anterior del editor.
    ''' Usado para comparar si ha habido cambios
    ''' </summary>
    Private codigoAnterior As String
    ''' <summary>
    ''' El número máximo de ficheros en el menú recientes
    ''' </summary>
    Private Const MaxFicsMenu As Integer = 9
    ''' <summary>
    ''' El número máximo de ficheros a guardar en la configuración
    ''' y a mostrar en el combo de los nombres de los ficheros.
    ''' </summary>
    Private Const MaxFicsConfig As Integer = 50

    ''' <summary>
    ''' Indica si se ha modificado el texto.
    ''' Cuando cambia el texto actual (<see cref="codigoNuevo"/>)
    ''' se comprueba con <see cref="codigoAnterior"/> por si hay cambios.
    ''' </summary>
    Private Property TextoModificado As Boolean
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


#Region " Definiciones de API para sincronizar los scrollbar de los richtextbox (15/Sep/20) "

    ' Adaptado de un ejemplo para C# de este foro:
    ' https://social.msdn.microsoft.com/Forums/windows/en-US/161d1636-aea3-4fee-beb4-52370663d44c/
    ' synchronize-scrolling-in-2-richtextboxes-in-c?forum=winforms
    Public Enum ScrollBarType As Integer
        SbHorz = 0
        SbVert = 1
        SbCtl = 2
        SbBoth = 3
    End Enum

    Public Enum Message As Long
        WM_VSCROLL = &H115
    End Enum

    Public Enum ScrollBarCommands As Long
        SB_THUMBPOSITION = 4
    End Enum

    <System.Runtime.InteropServices.DllImport("User32.dll")>
    Public Shared Function GetScrollPos(hWnd As IntPtr, nBar As Integer) As Integer
    End Function

    <System.Runtime.InteropServices.DllImport("User32.dll")>
    Public Shared Function SendMessage(hWnd As IntPtr, msg As Long, wParam As Long, lParam As Long) As Integer
    End Function

#End Region


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

        AsignaMetodosDeEventos()

        Width = CInt(Screen.PrimaryScreen.Bounds.Width * 0.45)
        Height = CInt(Screen.PrimaryScreen.Bounds.Height * 0.65)

        Me.CenterToScreen()

        Inicializar()

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
        Form1_Resize(Nothing, Nothing)

        GuardarConfig()

    End Sub

    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        If inicializando Then Return

        If Me.WindowState = FormWindowState.Normal Then
            tamForm = (Me.Left, Me.Top, Me.Height, Me.Width)
        Else
            tamForm = (Me.RestoreBounds.Left, Me.RestoreBounds.Top,
                       Me.RestoreBounds.Height, Me.RestoreBounds.Width)
        End If
    End Sub

    ' Para doble pulsación de teclas
    Private CtrlK As Integer
    Private CtrlC As Integer
    Private CtrlU As Integer
    Private CtrlL As Integer
    Private ShiftAltL As Integer
    Private ShiftAltS As Integer

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.Control AndAlso e.Shift Then
            If e.KeyCode = Keys.V Then
                e.Handled = True
                MostrarRecortes()
            End If
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
                buttonEditorMarcadorQuitarTodos.PerformClick()

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
                    col = richTextBoxCodigo.Lines(ln).IndexOf(richTextBoxCodigo.Lines(ln).TrimStart().Substring(0, 1))
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

    Private Sub Form1_DragEnter(sender As Object, e As DragEventArgs) Handles MyBase.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) OrElse
                e.Data.GetDataPresent(DataFormats.Text) OrElse
                e.Data.GetDataPresent("System.String") Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub Form1_DragDrop(sender As Object, e As DragEventArgs) Handles MyBase.DragDrop
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

    ' Las expresiones lambda para asignar con AddHandler en el menú editar
    Private lambdaUndo As EventHandler = Sub(sender As Object, e As EventArgs)
                                             If richTextBoxCodigo.CanUndo Then richTextBoxCodigo.Undo()
                                         End Sub
    Private lambdaRedo As EventHandler = Sub(sender As Object, e As EventArgs)
                                             If richTextBoxCodigo.CanRedo Then richTextBoxCodigo.Redo()
                                         End Sub
    Private lambdaPaste As EventHandler = Sub(sender As Object, e As EventArgs) Pegar()
    Private lambdaCopy As EventHandler = Sub(sender As Object, e As EventArgs)
                                             AñadirRecorte(richTextBoxCodigo.SelectedText)
                                             richTextBoxCodigo.Copy()
                                         End Sub
    Private lambdaCut As EventHandler = Sub(sender As Object, e As EventArgs)
                                            AñadirRecorte(richTextBoxCodigo.SelectedText)
                                            richTextBoxCodigo.Cut()
                                        End Sub
    Private lambdaSelectAll As EventHandler = Sub(sender As Object, e As EventArgs) richTextBoxCodigo.SelectAll()


    ''' <summary>
    ''' Asignar los métodos de evento manualmente en vez de con Handles
    ''' ya que así todos los métodos estarán en la lista desplegable al seleccionar Form1
    ''' en la lista desplegable de las clases
    ''' </summary>
    ''' <remarks>Aparte de que hay que hacerlo manualmente hasta que Visual Studio 2019 no los conecte</remarks>
    Private Sub AsignaMetodosDeEventos()
        ' Los métodos de evento del formulario usarlos con Handles  (26/Sep/20)
        'AddHandler Me.Load, AddressOf Form1_Load
        'AddHandler Me.FormClosing, AddressOf Form1_FormClosing
        'AddHandler Me.KeyDown, AddressOf Form1_KeyDown

        ' métodos lambda para asignar con AddHandler
        'Dim rTBoxCodigo_KeyDown = Sub(s As Object, e As KeyEventArgs) mostrarPosicion(e)


        ' richTextBoxCodigo
        Dim lambdarichTextBoxCodigoSelection = Sub(sender As Object, e As EventArgs)
                                                   If inicializando Then Return
                                                   'menuEditDropDownOpenning()
                                                   HabilitarBotones()
                                                   MostrarPosicion(Nothing)
                                               End Sub

        ' Según parece el KeyUp si funciona, pero no el KeyDown ??? (28/Sep/20)
        AddHandler richTextBoxCodigo.KeyUp, AddressOf richTextBoxCodigo_KeyUp
        AddHandler richTextBoxCodigo.SelectionChanged, lambdarichTextBoxCodigoSelection
        AddHandler richTextBoxCodigo.MouseClick, lambdarichTextBoxCodigoSelection

        AddHandler richTextBoxCodigo.TextChanged, AddressOf richTextBoxCodigo_TextChanged
        AddHandler richTextBoxCodigo.VScroll, AddressOf richTextBoxCodigo_VScroll
        AddHandler richTextBoxCodigo.FontChanged, Sub()
                                                      richTextBoxtLineas.Font = New Font(richTextBoxCodigo.Font.FontFamily, richTextBoxCodigo.Font.Size)
                                                  End Sub
        AddHandler richTextBoxCodigo.DragDrop, AddressOf Form1_DragDrop
        AddHandler richTextBoxCodigo.DragEnter, AddressOf Form1_DragEnter
        '' Para ver si averiguo la línea del fallo                   (03/Oct/20)
        'AddHandler richTextBoxSyntax.Click, AddressOf richTextBoxSyntax_Click

        ' Ficheros: los menús de archivo y botones de la barra ficheros
        AddHandler menuFileAcercaDe.Click, Sub() AcercaDe()
        AddHandler menuFileSalir.Click, Sub() Me.Close()
        AddHandler menuFileSeleccionarAbrir.Click, Sub() Abrir()
        AddHandler buttonSeleccionar.Click, Sub() Abrir()
        AddHandler buttonAbrir.Click, Sub() Abrir(comboBoxFileName.Text)
        AddHandler menuFileGuardar.Click, Sub() Guardar()
        AddHandler buttonGuardar.Click, Sub() Guardar()
        AddHandler menuFileGuardarComo.Click, Sub() GuardarComo()
        AddHandler buttonGuardarComo.Click, Sub() GuardarComo()
        AddHandler menuFileNuevo.Click, Sub() Nuevo()
        AddHandler buttonNuevo.Click, Sub() Nuevo()

        AddHandler menuRecargarFichero.Click, Sub() Recargar()
        AddHandler menuFileRecargar.Click, Sub() Recargar()
        AddHandler buttonRecargar.Click, Sub() Recargar()

        AddHandler menuFileRecientes.DropDownOpening, Sub()
                                                          For i = 0 To menuFileRecientes.DropDownItems.Count - 1
                                                              If menuFileRecientes.DropDownItems(i).Text.IndexOf(nombreUltimoFichero) > 3 Then
                                                                  menuFileRecientes.DropDownItems(i).Select()
                                                                  TryCast(menuFileRecientes.DropDownItems(i), ToolStripMenuItem).Checked = True
                                                              End If
                                                          Next
                                                      End Sub

        AddHandler menuCopiarPath.Click, Sub()
                                             Try
                                                 Clipboard.SetText(nombreUltimoFichero)
                                             Catch ex As Exception
                                             End Try
                                         End Sub

        AddHandler comboBoxFileName.Validating, AddressOf comboBoxFileName_Validating
        AddHandler comboBoxFileName.KeyDown, AddressOf comboBoxFileName_KeyDown

        ' Edición: menús y botones de la barra de edición
        AddHandler menuEdit.DropDownOpening, Sub() menuEditDropDownOpening()
        AddHandler menuEditDeshacer.Click, lambdaUndo
        AddHandler buttonDeshacer.Click, lambdaUndo
        AddHandler menuEditRehacer.Click, lambdaRedo
        AddHandler buttonRehacer.Click, lambdaRedo
        AddHandler menuEditPegar.Click, lambdaPaste
        AddHandler buttonPegar.Click, lambdaPaste
        AddHandler menuEditCopiar.Click, lambdaCopy
        AddHandler buttonCopiar.Click, lambdaCopy
        AddHandler menuEditCortar.Click, lambdaCut
        AddHandler buttonCortar.Click, lambdaCut
        AddHandler menuEditSeleccionarTodo.Click, lambdaSelectAll
        AddHandler buttonSeleccionarTodo.Click, lambdaSelectAll
        ' Recortes
        AddHandler buttonEdicionRecortes.Click, Sub()
                                                    ' Mostrar ventana con los recortes guardados
                                                    MostrarRecortes()
                                                End Sub

        ' Compilar, evaluar, ejecutar 
        AddHandler menuTools.DropDownOpening, Sub()
                                                  Dim b = richTextBoxCodigo.TextLength > 0
                                                  ' Si el Lenguaje no es vb/cs no:      (08/Oct/20)
                                                  '   compilar, ejecutar, colorear, etc.
                                                  Dim b2 = b
                                                  If buttonLenguaje.Text = ExtensionTexto Then
                                                      b2 = False
                                                  End If
                                                  menuEvaluar.Enabled = b2
                                                  menuColorearHTML.Enabled = b2
                                                  menuColorear.Enabled = b2
                                                  menuNoColorear.Enabled = b2
                                                  menuCompilar.Enabled = b2
                                                  menuEjecutar.Enabled = b2

                                                  menuOcultarEvaluar.Enabled = lstSyntax.Items.Count > 0
                                                  If splitContainer1.Panel2.Visible = True Then
                                                      menuOcultarEvaluar.Text = "Ocultar panel de evaluación/error"
                                                      'menuOcultarEvaluar.Checked = True
                                                  Else
                                                      menuOcultarEvaluar.Text = "Mostrar panel de evaluación/error"
                                                      'menuOcultarEvaluar.Checked = False
                                                  End If
                                              End Sub
        AddHandler menuCompilar.Click, Sub() Build()
        AddHandler buttonCompilar.Click, Sub() Build()
        AddHandler menuEjecutar.Click, Sub() CompilarEjecutar()
        AddHandler buttonEjecutar.Click, Sub() CompilarEjecutar()

        AddHandler buttonEvaluar.Click, Sub() Evaluar()
        AddHandler menuEvaluar.Click, Sub() Evaluar()
        AddHandler buttonColorearAlEvaluar.Click, Sub() colorearAlEvaluar = buttonColorearAlEvaluar.Checked
        AddHandler buttonCompilarAlEvaluar.Click, Sub() compilarAlEvaluar = buttonCompilarAlEvaluar.Checked

        AddHandler menuColorear.Click, Sub() ColorearCodigo()
        AddHandler buttonColorear.Click, Sub() ColorearCodigo()
        AddHandler menuNoColorear.Click, Sub() NoColorear()
        AddHandler buttonNoColorear.Click, Sub() NoColorear()

        ' Ocultar el panel de evaluación                            (05/Oct/20)
        AddHandler menuOcultarEvaluar.Click, Sub()
                                                 'If lstSyntax.Items.Count = 0 Then
                                                 '    menuOcultarEvaluar.Checked = False
                                                 '    menuOcultarEvaluar.Text = "No hay datos de evaluación"
                                                 'End If
                                                 If menuOcultarEvaluar.Text = "Ocultar panel de evaluación/error" Then
                                                     splitContainer1.Panel2.Visible = False
                                                     menuOcultarEvaluar.Text = "Mostrar panel de evaluación/error"
                                                 Else
                                                     menuOcultarEvaluar.Text = "Ocultar panel de evaluación/error"
                                                     splitContainer1.Panel2.Visible = True
                                                 End If
                                                 splitContainer1_Resize(Nothing, Nothing)
                                             End Sub

        ' Colorear en HTML
        AddHandler menuColorearHTML.Click, Sub() ColorearHTML()
        AddHandler buttonColorearHTML.Click, Sub() ColorearHTML()
        AddHandler chkMostrarLineasHTML.Click, Sub() mostrarLineasHTML = chkMostrarLineasHTML.Checked

        ' Herramientas; Opciones, Colorear, lenguajes
        AddHandler panelHerramientas.SizeChanged, AddressOf panelHerramientas_SizeChanged

        AddHandler buttonLenguaje.DropDownItemClicked, Sub(sender As Object, e As ToolStripItemClickedEventArgs)
                                                           Dim it = e.ClickedItem
                                                           buttonLenguaje.Text = it.Text
                                                           buttonLenguaje.Image = it.Image
                                                       End Sub

        ' Clasificar la selección
        AddHandler menuEditorClasificarSeleccion.Click, Sub() ClasificarSeleccion()
        AddHandler buttonEditorClasificarSeleccion.Click, Sub() ClasificarSeleccion()

        AddHandler menuOpciones.Click, Sub()
                                           ' Mostrar la ventana de opciones
                                           ' usando el form actual como parámetro      (27/Sep/20)
                                           inicializando = True
                                           Dim opFrm As New FormOpciones(Me)
                                           With opFrm
                                               If .ShowDialog() = DialogResult.OK Then
                                                   ' las asignaciones se hacen en el formulario de opciones
                                                   AsignarRecientes()
                                                   ' Comprobar si ha cambiado la fuente
                                                   If labelFuente.Text <> $"{fuenteNombre}; {fuenteTamaño}" Then
                                                       richTextBoxCodigo.Font = New Font(fuenteNombre, CSng(fuenteTamaño))
                                                       labelFuente.Text = $"{fuenteNombre}; {fuenteTamaño}"
                                                       If colorearAlCargar Then ColorearCodigo()
                                                   End If

                                                   GuardarConfig()
                                               End If
                                           End With
                                           inicializando = False
                                       End Sub


        ' Buscar y reemplazar                                       (17/Sep/20)
        AddHandler menuEditBuscar.Click, Sub() BuscarReemplazar(True)
        AddHandler menuEditReemplazar.Click, Sub() BuscarReemplazar(esBuscar:=False)
        AddHandler buttonBuscarSiguiente.Click, Sub() BuscarSiguiente(esReemplazar:=False)
        AddHandler menuEditBuscarSiguiente.Click, Sub() BuscarSiguiente(esReemplazar:=False)
        AddHandler buttonReemplazarSiguiente.Click, Sub() ReemplazarSiguiente()
        AddHandler menuEditReemplazarSiguiente.Click, Sub() ReemplazarSiguiente()
        AddHandler buttonReemplazarTodo.Click, Sub() ReemplazarTodo()
        AddHandler menuEditReemplazarTodos.Click, Sub() ReemplazarTodo()

        AddHandler comboBoxBuscar.Validating, Sub() comboBoxBuscar_Validating()
        AddHandler comboBoxReemplazar.Validating, Sub() comboBoxReemplazar_Validating()
        ' Si se pulsa Intro, que busque                             (02/Oct/20)
        ' hay que usar el evento KeyUp, no el KeyPress
        AddHandler comboBoxBuscar.KeyUp, Sub(sender As Object, e As KeyEventArgs)
                                             If e.KeyCode = Keys.Enter Then
                                                 e.Handled = True
                                                 BuscarSiguiente(esReemplazar:=False)
                                             End If
                                         End Sub

        ' Para palabras completas y case sensitive                  (17/Sep/20)
        AddHandler buttonMatchCase.Click, Sub() buscarMatchCase = buttonMatchCase.Checked
        AddHandler buttonWholeWord.Click, Sub() buscarWholeWord = buttonWholeWord.Checked

        ' Contenedor
        AddHandler splitContainer1.Resize, AddressOf splitContainer1_Resize

        ' Crear un context menú para el richTextBox del código      (18/Sep/20)
        If richTextBoxCodigo.ContextMenuStrip Is Nothing Then
            richTextBoxCodigo.ContextMenuStrip = rtbCodigoContext
        End If
        CrearContextMenuCodigo()
        AddHandler rtbCodigoContext.Opening, Sub() menuEditDropDownOpening()

        Dim lambdaMenuMostrar = Sub(sender As Object, e As EventArgs)
                                    ' Están en un FlowPanel y se reajusta solo
                                    ' al mostrar/ocultar los controles contenidos
                                    ' panelBuscar es un panel con los ToolStrip Buscar y Reemplazar
                                    ' el resto son del tipo ToolStrip
                                    If sender Is menuMostrar_Ficheros Then
                                        toolStripFicheros.Visible = menuMostrar_Ficheros.Checked
                                    ElseIf sender Is menuMostrar_Buscar Then
                                        MostrarPanelBuscar(menuMostrar_Buscar.Checked)
                                    ElseIf sender Is menuMostrar_Compilar Then
                                        toolStripCompilar.Visible = menuMostrar_Compilar.Checked
                                    ElseIf sender Is menuMostrar_Edicion Then
                                        toolStripEdicion.Visible = menuMostrar_Edicion.Checked
                                    ElseIf sender Is menuMostrar_Editor Then
                                        toolStripEditor.Visible = menuMostrar_Editor.Checked
                                    End If

                                    AjustarAltoPanelHerramientas()
                                End Sub

        AddHandler menuMostrar_Ficheros.Click, lambdaMenuMostrar
        AddHandler menuMostrar_Buscar.Click, lambdaMenuMostrar
        AddHandler menuMostrar_Compilar.Click, lambdaMenuMostrar
        AddHandler menuMostrar_Edicion.Click, lambdaMenuMostrar
        AddHandler menuMostrar_Editor.Click, lambdaMenuMostrar

        ' usar también la visibilidad de los paneles
        AddHandler toolStripFicheros.VisibleChanged, Sub() AjustarAltoPanelHerramientas()
        AddHandler panelBuscar.VisibleChanged, Sub() AjustarAltoPanelHerramientas()
        AddHandler toolStripCompilar.VisibleChanged, Sub() AjustarAltoPanelHerramientas()
        AddHandler toolStripEdicion.VisibleChanged, Sub() AjustarAltoPanelHerramientas()
        AddHandler toolStripEditor.VisibleChanged, Sub() AjustarAltoPanelHerramientas()

        ' Barra y menú de Editor (que no edición)
        AddHandler menuEditor.DropDownOpening, Sub()
                                                   Dim b = richTextBoxCodigo.SelectionLength > 0
                                                   menuEditorQuitarIndentacion.Enabled = b
                                                   menuEditorPonerIndentacion.Enabled = b
                                                   menuEditorQuitarComentarios.Enabled = b
                                                   menuEditorPonerComentarios.Enabled = b

                                                   menuEditorClasificarSeleccion.Enabled = b

                                                   menuEditorPonerTexto.Enabled = richTextBoxCodigo.TextLength > 0

                                                   b = ColMarcadores.Count > 0
                                                   menuEditorMarcador.Enabled = b
                                                   menuEditorMarcadorAnterior.Enabled = b
                                                   menuEditorMarcadorSiguiente.Enabled = b
                                                   menuEditorMarcadorQuitarTodos.Enabled = b

                                               End Sub

        AddHandler buttonEditorQuitarIndentacion.Click, Sub() QuitarIndentacion(richTextBoxCodigo)
        AddHandler menuEditorQuitarIndentacion.Click, Sub() QuitarIndentacion(richTextBoxCodigo)
        AddHandler buttonEditorPonerIndentacion.Click, Sub() PonerIndentacion(richTextBoxCodigo)
        AddHandler menuEditorPonerIndentacion.Click, Sub() PonerIndentacion(richTextBoxCodigo)
        AddHandler buttonEditorQuitarComentarios.Click, Sub() QuitarComentarios(richTextBoxCodigo)
        AddHandler menuEditorQuitarComentarios.Click, Sub() QuitarComentarios(richTextBoxCodigo)
        AddHandler buttonEditorPonerComentarios.Click, Sub() PonerComentarios(richTextBoxCodigo)
        AddHandler menuEditorPonerComentarios.Click, Sub() PonerComentarios(richTextBoxCodigo)
        AddHandler buttonEditorMarcador.Click, Sub() MarcadorPonerQuitar()
        AddHandler menuEditorMarcador.Click, Sub() MarcadorPonerQuitar()
        AddHandler buttonEditorMarcadorAnterior.Click, Sub() MarcadorAnterior()
        AddHandler menuEditorMarcadorAnterior.Click, Sub() MarcadorAnterior()
        AddHandler buttonEditorMarcadorSiguiente.Click, Sub() MarcadorSiguiente()
        AddHandler menuEditorMarcadorSiguiente.Click, Sub() MarcadorSiguiente()

        ' opciones para clasificar (solo en el menú Editor)
        AddHandler menuEditorCambiarMayúsculas.DropDownOpening, Sub()
                                                                    Dim b = richTextBoxCodigo.SelectionLength > 0
                                                                    menuMayúsculas.Enabled = b
                                                                    menuMinúsculas.Enabled = b
                                                                    menuTitulo.Enabled = b
                                                                    menuPrimeraMinúsculas.Enabled = b
                                                                End Sub
        AddHandler menuMayúsculas.Click, Sub() CambiarMayúsculasMinúsculas(CasingValues.Upper)
        AddHandler menuMinúsculas.Click, Sub() CambiarMayúsculasMinúsculas(CasingValues.Lower)
        AddHandler menuTitulo.Click, Sub() CambiarMayúsculasMinúsculas(CasingValues.Title)
        AddHandler menuPrimeraMinúsculas.Click, Sub() CambiarMayúsculasMinúsculas(CasingValues.FirstToLower)

        ' quitar los espacios de delante, detrás o ambos            (02/Oct/20)
        AddHandler menuEditorQuitarEspacios.DropDownOpening, Sub()
                                                                 Dim b = richTextBoxCodigo.SelectionLength > 0
                                                                 menuQuitarEspaciosTrim.Enabled = b
                                                                 menuQuitarEspaciosTrimStart.Enabled = b
                                                                 menuQuitarEspaciosTrimEnd.Enabled = b
                                                                 menuQuitarEspaciosTodos.Enabled = b
                                                             End Sub
        AddHandler menuQuitarEspaciosTrim.Click, Sub() QuitarEspacios()
        AddHandler menuQuitarEspaciosTrimStart.Click, Sub() QuitarEspacios(delante:=True, detras:=False)
        AddHandler menuQuitarEspaciosTrimEnd.Click, Sub() QuitarEspacios(delante:=False, detras:=True)
        AddHandler menuQuitarEspaciosTodos.Click, Sub() QuitarEspacios(todos:=True)

        AddHandler menuEditorPonerTextoAlFinal.Click, Sub() PonerTextoAlFinal()
        AddHandler menuEditorQuitarTextoDelfinal.Click, Sub() QuitarTextoDelFinal()

        Dim lambdaQuitarMarcadores = Sub(sender As Object, e As EventArgs)
                                         If ColMarcadores.Count = 0 Then Return
                                         If MessageBox.Show("¿Seguro que quieres quitar todos los marcadores.",
                                                            "Quitar todos los marcadores",
                                                            MessageBoxButtons.YesNo,
                                                            MessageBoxIcon.Question) = DialogResult.Yes Then

                                             MarcadorQuitarTodos()
                                         End If
                                     End Sub
        AddHandler buttonEditorMarcadorQuitarTodos.Click, lambdaQuitarMarcadores
        AddHandler menuEditorMarcadorQuitarTodos.Click, lambdaQuitarMarcadores


    End Sub

    ''' <summary>
    ''' Ajustar el alto del panel de herramientas
    ''' según estén visibles las barras de herramientas que están
    ''' en la segunda línea.
    ''' </summary>
    Private Sub AjustarAltoPanelHerramientas()
        'If inicializando Then Return

        ' Ajustar el alto del FlowPanel (01/Oct/20)
        ' Comprobar si las barras de herramientas el Top es mayor de 10
        ' en ese caso, el FlowPanel ponerlo a 63, si no será 28
        Dim esVisible = False
        For Each c As Control In panelHerramientas.Controls
            If c.Visible AndAlso c.Top > 10 Then
                esVisible = True
                Exit For
            End If
        Next
        If esVisible Then
            HabilitarBotones()
            panelHerramientas.Height = 63
        Else
            panelHerramientas.Height = 31
        End If
    End Sub

    ''' <summary>
    ''' Inicializar los menús, etc.
    ''' Este método se llama desde el evento Form1_Load.
    ''' </summary>
    Private Sub Inicializar()
        '
        ' Todo esto estaba en el evento Form1_Load                  (28/Sep/20)
        '

        Dim extension = ".appconfig.txt"
        dirDocumentos = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        ficheroConfiguracion = Path.Combine(dirDocumentos, Application.ProductName & extension)

        LeerConfig()

        ' Mostrar los 15 primeros en el menú Recientes
        AsignarRecientes()

        inicializando = False

        ' No mostrar el panel al iniciar
        ' Ponerlo después de inicializando = false                  (30/Sep/20)
        ' para que se ajuste el tamaño de FlowPanel
        MostrarPanelBuscar(False)

        esCtrlF = True

        'labelTamaño.Text = "0 car."
        labelTamaño.Text = $"{0:#,##0} car. ({0:#,##0} palab.)"
        labelInfo.Text = "Selecciona el fichero a abrir."
        labelModificado.Text = " "
        labelPos.Text = "Lín: 0  Car: 0"

        ' Guardar los esquemas de colores
        ' en: MyDocuments/ColorDictionary.csv
        Compilar.WriteColorDictionaryToFile()
        ' Si se cambia y se quiere leer
        'Compilar.UpdateColorDictionaryFromFile()

        splitContainer1.Panel2.Visible = False
        splitContainer1_Resize(Nothing, Nothing)

        Dim tmpCargarUltimo = cargarUltimo
        ' Comprobar si se indica un fichero en la línea de comandos (28/Sep/20)
        If Environment.GetCommandLineArgs.Length > 1 Then
            Dim s As String = Environment.GetCommandLineArgs(1)
            If Not String.IsNullOrWhiteSpace(s) Then
                nombreUltimoFichero = s
                tmpCargarUltimo = True
            End If
        End If

        ' posicionarlo al principio
        richTextBoxCodigo.SelectionStart = 0
        richTextBoxCodigo.SelectionLength = 0

        If tmpCargarUltimo Then
            Abrir(nombreUltimoFichero)

            If selectionStartAnt > -1 Then
                richTextBoxCodigo.SelectionStart = selectionStartAnt
                richTextBoxCodigo.SelectionLength = selectionEndAnt
            End If

            ' Mostrar los números de línea
            If nombreUltimoFichero <> "" Then _
                AñadirNumerosDeLinea()


            ' Iniciar la posición al principio
            MostrarPosicion(New KeyEventArgs(Keys.Home))

        Else
            comboBoxFileName.Text = nombreUltimoFichero
        End If

        '
        ' hasta aquí lo que estaba en el Form1_Load
        '

        buttonLenguaje.DropDownItems(0).Text = Compilar.LenguajeVisualBasic
        buttonLenguaje.DropDownItems(1).Text = Compilar.LenguajeCSharp


        PonerLosMarcadores()

        '' posicionarlo al principio
        'richTextBoxCodigo.SelectionStart = 0
        'richTextBoxCodigo.SelectionLength = 0

    End Sub


    ''' <summary>
    ''' Crear un menú contextual para richTextBoxCodigo
    ''' para los comandos de edición
    ''' </summary>
    ''' <remarks>Usando la extensión Clonar: 27/Sep/20</remarks>
    Private Sub CrearContextMenuCodigo()
        rtbCodigoContext.Items.Clear()
        rtbCodigoContext.Items.AddRange({menuEditDeshacer.Clonar(lambdaUndo),
                                        menuEditRehacer.Clonar(lambdaRedo), Me.toolEditSeparator1,
                                        menuEditCortar.Clonar(lambdaCut), menuEditCopiar.Clonar(lambdaCopy),
                                        menuEditPegar.Clonar(lambdaPaste), Me.toolEditSeparator2,
                                        menuEditSeleccionarTodo.Clonar(lambdaSelectAll)})
        'richTextBoxCodigo.ContextMenuStrip = rtbCodigoContext
    End Sub


    '
    ' Los métodos normales
    '

    ''' <summary>
    ''' Guardar los datos de configuración
    ''' </summary>
    Private Sub GuardarConfig()
        Dim cfg = New Config(ficheroConfiguracion)

        ' Si cargarUltimo es falso no guardar el último fichero     (16/Sep/20)
        cfg.SetValue("Ficheros", "CargarUltimo", cargarUltimo)
        If cargarUltimo Then
            cfg.SetValue("Ficheros", "Ultimo", nombreUltimoFichero)
        Else
            cfg.SetValue("Ficheros", "Ultimo", "")
        End If
        cfg.SetValue("Herramientas", "Lenguaje", buttonLenguaje.Text)
        cfg.SetValue("Herramientas", "Colorear", colorearAlCargar)
        cfg.SetValue("Herramientas", "Mostrar Líneas HTML", chkMostrarLineasHTML.Checked)

        cfg.SetValue("Herramientas", "ColorearEvaluar", colorearAlEvaluar)
        cfg.SetValue("Herramientas", "CompilarEvaluar", compilarAlEvaluar)

        ' Para la forma de clasificar                               (02/Oct/20)
        cfg.SetValue("Clasificar", "IgnoreCase", CompararString.IgnoreCase)
        cfg.SetValue("Clasificar", "CompareOrdinal", CompararString.CompareOrdinal)

        Dim j As Integer

        ' Guardar la lista de últimos ficheros
        ' solo los MaxFicsConfig (50) últimos
        Dim cuantos = comboBoxFileName.Items.Count
        cfg.SetValue("Ficheros", "Count", cuantos)
        For i = 0 To comboBoxFileName.Items.Count - 1 ' To 0 Step -1
            ' No añadir el path si es el directorio de Documentos
            Dim s = comboBoxFileName.Items(i).ToString
            If Path.GetDirectoryName(s) = dirDocumentos Then
                s = Path.GetFileName(s)
            End If
            cfg.SetKeyValue("Ficheros", $"Fichero{j}", s)
            j += 1
            If j = MaxFicsConfig Then Exit For
        Next

        ' El nombre y tamaño de la fuente                           (11/Sep/20)
        cfg.SetValue("Fuente", "Nombre", fuenteNombre)
        cfg.SetValue("Fuente", "Tamaño", fuenteTamaño)

        ' El tamaño y la posición de la ventana
        cfg.SetValue("Ventana", "Left", tamForm.L)
        cfg.SetValue("Ventana", "Top", tamForm.T)
        cfg.SetValue("Ventana", "Height", tamForm.H)
        cfg.SetValue("Ventana", "Width", tamForm.W)

        ' Para buscar y reemplazar y CaseSensitive y WholeWord      (17/Sep/20)
        cfg.SetValue("Buscar", "QueBuscar", buscarQueBuscar)
        cfg.SetValue("Reemplazar", "QueReemplazar", buscarQueReemplazar)
        cfg.SetValue("Buscar", "MatchCase", buttonMatchCase.Checked)
        cfg.SetValue("Buscar", "WholeWord", buttonWholeWord.Checked)
        cfg.SetValue("Buscar", "Buscar-Items-Count", comboBoxBuscar.Items.Count)
        cfg.SetValue("Reemplazar", "Reemplazar-Items-Count", comboBoxReemplazar.Items.Count)
        j = 0
        For i = 0 To comboBoxBuscar.Items.Count - 1
            cfg.SetKeyValue("Buscar", $"Buscar-Items{j}", comboBoxBuscar.Items(i).ToString)
            j += 1
            If j = BuscarMaxItems Then Exit For
        Next
        j = 0
        For i = 0 To comboBoxReemplazar.Items.Count - 1
            cfg.SetKeyValue("Reemplazar", $"Reemplazar-Items{j}", comboBoxReemplazar.Items(i).ToString)
            j += 1
            If j = BuscarMaxItems Then Exit For
        Next

        cfg.SetValue("Marcadores", "Marcadores-Fichero", nombreUltimoFichero)
        cfg.SetValue("Marcadores", "Marcadores-Count", ColMarcadores.Count)
        For i = 0 To ColMarcadores.Count - 1
            cfg.SetKeyValue("Marcadores", $"Marcadores-Items{i}", ColMarcadores.Item(i))
        Next

        ' El estado de los ToolStrip                                (29/Sep/20)
        ' No guardar la visibilidad de buscar
        cfg.SetValue("Herramientas", "Barra-Ficheros", toolStripFicheros.Visible)
        cfg.SetValue("Herramientas", "Barra-Edicion", toolStripEdicion.Visible)
        cfg.GetValue("Herramientas", "Barra-Compilar", toolStripCompilar.Visible)
        cfg.GetValue("Herramientas", "Barra-Editor", toolStripEditor.Visible)

        ' guardar los recortes
        cfg.SetValue("Recortes", "Count", ColRecortes.Count)
        For i = 0 To ColRecortes.Count - 1
            cfg.SetKeyValue("Recortes", $"Recortes-Item{i}", ColRecortes(i))
        Next

        ' Para la selección al cerrar                               (05/Oct/20)
        cfg.SetValue("Selección", "selectionStartAnt", selectionStartAnt)
        cfg.SetValue("Selección", "selectionEndAnt", selectionEndAnt)

        ' para el texto a quitar/poner                              (08/Oct/20)
        cfg.SetValue("Quitar Poner Texto", "PonerTexto", txtPonerTexto.Text)

        cfg.Save()
    End Sub

    ''' <summary>
    ''' Leer el fichero de configuración
    ''' y asignar los valores usados anteriormente.
    ''' </summary>
    Private Sub LeerConfig()
        Dim cfg = New Config(ficheroConfiguracion)

        ' Si cargarUltimo es falso no asignar el último fichero     (16/Sep/20)
        cargarUltimo = cfg.GetValue("Ficheros", "CargarUltimo", False)
        If cargarUltimo Then
            nombreUltimoFichero = cfg.GetValue("Ficheros", "Ultimo", "")
        Else
            nombreUltimoFichero = ""
        End If

        buttonLenguaje.Text = cfg.GetValue("Herramientas", "Lenguaje", Compilar.LenguajeVisualBasic)
        If buttonLenguaje.Text = Compilar.LenguajeVisualBasic Then
            buttonLenguaje.Image = buttonLenguaje.DropDownItems(0).Image
        ElseIf buttonLenguaje.Text = Compilar.LenguajeCSharp Then
            buttonLenguaje.Image = buttonLenguaje.DropDownItems(1).Image
        Else
            buttonLenguaje.Image = buttonLenguaje.DropDownItems(2).Image
        End If
        colorearAlCargar = cfg.GetValue("Herramientas", "Colorear", False)
        chkMostrarLineasHTML.Checked = cfg.GetValue("Herramientas", "Mostrar Líneas HTML", True)

        colorearAlEvaluar = cfg.GetValue("Herramientas", "ColorearEvaluar", False)
        compilarAlEvaluar = cfg.GetValue("Herramientas", "CompilarEvaluar", True)

        ' Para la forma de clasificar                               (02/Oct/20)
        CompararString.IgnoreCase = cfg.GetValue("Clasificar", "IgnoreCase", False)
        CompararString.CompareOrdinal = cfg.GetValue("Clasificar", "CompareOrdinal", False)

        Dim cuantos = cfg.GetValue("Ficheros", "Count", 0)
        comboBoxFileName.Items.Clear()

        For i = 0 To MaxFicsConfig - 1
            If i >= cuantos Then Exit For
            Dim s = cfg.GetValue("Ficheros", $"Fichero{i}", "-1")
            If s = "-1" Then Exit For
            ' No añadir el path si es el directorio de Documentos
            If Path.GetDirectoryName(s) = dirDocumentos Then
                s = Path.GetFileName(s)
            End If
            comboBoxFileName.Items.Add(s)
        Next
        ' No clasificar los elementos

        ' El nombre y tamaño de la fuente                           (11/Sep/20)
        fuenteNombre = cfg.GetValue("Fuente", "Nombre", "Consolas")
        fuenteTamaño = cfg.GetValue("Fuente", "Tamaño", "11")
        labelFuente.Text = $"{fuenteNombre}; {fuenteTamaño}"

        ' Solo asignar si son diferentes                            (02/Oct/20)

        richTextBoxCodigo.Font = New Font(fuenteNombre, CSng(fuenteTamaño))
        richTextBoxtLineas.Font = New Font(fuenteNombre, CSng(fuenteTamaño))
        ' No cambiar la fuente del panel de sintaxis                (03/Oct/20)
        'richTextBoxSyntax.Font = New Font(fuenteNombre, CSng(fuenteTamaño))

        ' El tamaño y la posición de la ventana
        tamForm.L = cfg.GetValue("Ventana", "Left", -1)
        tamForm.T = cfg.GetValue("Ventana", "Top", -1)
        tamForm.H = cfg.GetValue("Ventana", "Height", -1)
        tamForm.W = cfg.GetValue("Ventana", "Width", -1)

        If tamForm.L <> -1 Then Me.Left = tamForm.L
        If tamForm.T <> -1 Then Me.Top = tamForm.T
        If tamForm.H > -1 Then Me.Height = tamForm.H
        If tamForm.W > -1 Then Me.Width = tamForm.W

        ' Los datos de configuración para buscar y reemplazar       (17/Sep/20)
        ' y valores de MatchCase, WholeWord
        buscarQueBuscar = cfg.GetValue("Buscar", "Buscar", "")
        buscarQueReemplazar = cfg.GetValue("Reemplazar", "Reemplazar", "")

        buscarMatchCase = cfg.GetValue("Buscar", "MatchCase", False)
        buttonMatchCase.Checked = buscarMatchCase

        buscarWholeWord = cfg.GetValue("Buscar", "WholeWord", False)
        buttonWholeWord.Checked = buscarWholeWord

        cuantos = cfg.GetValue("Buscar", "Buscar-Items-Count", 0)
        comboBoxBuscar.Items.Clear()
        For i = 0 To BuscarMaxItems - 1
            If i >= cuantos Then Exit For
            Dim s = cfg.GetValue("Buscar", $"Buscar-Items{i}", "-1")
            If s = "-1" Then Exit For
            comboBoxBuscar.Items.Add(s)
        Next
        cuantos = cfg.GetValue("Reemplazar", "Reemplazar-Items-Count", 0)
        comboBoxReemplazar.Items.Clear()
        For i = 0 To BuscarMaxItems - 1
            If i >= cuantos Then Exit For
            Dim s = cfg.GetValue("Reemplazar", $"Reemplazar-Items{i}", "-1")
            If s = "-1" Then Exit For
            comboBoxReemplazar.Items.Add(s)
        Next
        comboBoxBuscar.Text = buscarQueBuscar
        comboBoxReemplazar.Text = buscarQueReemplazar

        ' Solo asignar los valores si el fichero es el mismo
        ColMarcadores.Clear()
        MarcadorFichero = cfg.GetValue("Marcadores", "Marcadores-Fichero", "")
        If Not String.IsNullOrEmpty(nombreUltimoFichero) AndAlso nombreUltimoFichero = MarcadorFichero Then
            cuantos = cfg.GetValue("Marcadores", "Marcadores-Count", 0)
            For j = 0 To cuantos - 1
                ColMarcadores.Add(cfg.GetValue("Marcadores", $"Marcadores-Items{j}", 0))
            Next
        End If

        ' El estado de los ToolStrip                                (29/Sep/20)
        ' La de buscar no se muestra al iniciar el programa
        toolStripFicheros.Visible = cfg.GetValue("Herramientas", "Barra-Ficheros", True)
        toolStripEdicion.Visible = cfg.GetValue("Herramientas", "Barra-Edicion", True)
        toolStripCompilar.Visible = cfg.GetValue("Herramientas", "Barra-Compilar", True)
        toolStripEditor.Visible = cfg.GetValue("Herramientas", "Barra-Editor", True)

        ' leer los recortes
        cuantos = cfg.GetValue("Recortes", "Count", 0)
        ColRecortes.Clear()
        For i = 0 To cuantos - 1
            ColRecortes.Add(cfg.GetValue("Recortes", $"Recortes-Item{i}", ""))
        Next

        ' Para la selección al cerrar                               (05/Oct/20)
        selectionStartAnt = cfg.GetValue("Selección", "selectionStartAnt", -1)
        selectionEndAnt = cfg.GetValue("Selección", "selectionEndAnt", -1)

        ' para el texto a quitar/poner                              (08/Oct/20)
        txtPonerTexto.Text = cfg.GetValue("Quitar Poner Texto", "PonerTexto", "")

    End Sub

    '
    ' Sin clasificar
    '

    ''' <summary>
    ''' Añadir los números de línea
    ''' </summary>
    ''' <remarks>Como método separado 18/Sep/20</remarks>
    Private Sub AñadirNumerosDeLinea()
        If inicializando Then Return
        If String.IsNullOrEmpty(richTextBoxCodigo.Text) Then Return

        Dim lineas = richTextBoxCodigo.Lines.Length
        richTextBoxtLineas.Text = ""
        For i = 1 To lineas
            richTextBoxtLineas.Text &= $" {i.ToString("0").PadLeft(4)}{vbCr}"
        Next
        ' Sincronizar los scrolls
        richTextBoxCodigo_VScroll(Nothing, Nothing)

        PonerLosMarcadores()
    End Sub

    '
    ' Los métodos de ficheros
    '

    ''' <summary>
    ''' Asigna los ficheros al menú recientes
    ''' </summary>
    Private Sub AsignarRecientes()
        menuFileRecientes.DropDownItems.Clear()
        For i = 0 To MaxFicsMenu - 1
            ' Salir si en el combo hay menos ficheros que el contador actual
            ' el bucle lo hace hasta un máximo de MaxFicsMenu (10)
            If (i + 1) > comboBoxFileName.Items.Count Then Exit For

            Dim s = $"{i + 1} - {comboBoxFileName.Items(i).ToString}"
            ' Crear el menú y asignar el método de evento para Click
            Dim m = New ToolStripMenuItem(s)
            AddHandler m.Click, Sub() AbrirReciente(m.Text)
            menuFileRecientes.DropDownItems.Add(m)
        Next
    End Sub

    ''' <summary>
    ''' Abre el fichero seleccionado del menú recientes
    ''' </summary>
    ''' <param name="fic">Path completo del fichero a abrir</param>
    Private Sub AbrirReciente(fic As String)
        ' El nombre está después del guión -
        ' posición máxima será 4: 1 - Nombre
        ' pero si hay 2 cifras, será 5: 10 - Nombre
        ' Tomando a partir del la 4ª posición está bien
        If String.IsNullOrEmpty(fic) Then Return

        Dim ficT = fic.Substring(4).Trim()

        ' Si es el que está abierto, salir
        ' Solo si el texto no está modificado                       (14/Sep/20)
        ' por si se quiere re-abrir
        If ficT = nombreUltimoFichero AndAlso TextoModificado = False Then Return

        ' Si está modificado, preguntar si se quiere guardar
        If TextoModificado Then
            GuardarComo()
        End If

        ' pero no abrir ese...
        ' en nombreUltimoFichero se asigna el seleccionado del menú recientes
        nombreUltimoFichero = ficT
        Abrir(nombreUltimoFichero)
    End Sub


    ''' <summary>
    ''' Muestra el cuadro de diálogos de abrir ficheros
    ''' </summary>
    Private Sub Abrir()
        Using oFD As New OpenFileDialog
            oFD.Title = "Seleccionar fichero de código a abrir"
            Dim sFN = comboBoxFileName.Text
            sFN = If(String.IsNullOrEmpty(sFN),
                Path.Combine(dirDocumentos, $"prueba.{If(buttonLenguaje.Text = Compilar.LenguajeVisualBasic, ".vb", ".cs")}"),
                sFN)
            oFD.FileName = sFN
            oFD.Filter = "Código de Visual Basic y CSharp (*.vb;*.cs)|*.vb;*.cs|Textos (*.txt;*.log;*.md)|*.txt;*.log;*.md|Todos los ficheros (*.*)|*.*"
            oFD.InitialDirectory = Path.GetDirectoryName(sFN)
            oFD.Multiselect = False
            oFD.RestoreDirectory = True
            If oFD.ShowDialog() = DialogResult.OK Then

                ' Abrir el fichero
                Abrir(oFD.FileName)
            End If
        End Using
    End Sub

    ''' <summary>
    ''' Abre nuevamente el último fichero
    ''' desechando los datos realizados
    ''' </summary>
    Private Sub Recargar()
        If nombreUltimoFichero <> "" Then _
            Abrir(nombreUltimoFichero)
    End Sub

    ''' <summary>
    ''' Abre el fichero indicado en el parámetro, 
    ''' si no está en el combo de ficheros, añadirlo al principio.
    ''' De añadirlo al princpio se encarga <see cref="AñadirAlComboBoxFileName"/>.
    ''' </summary>
    ''' <param name="fic">El path completo del fichero a abrir</param>
    ''' <remarks>En el combo se muestra solo el nombre sin el path si el path es el directorio de documentos
    ''' (o el que se haya asignado como predeterminado) en otro caso se muestra el path completo</remarks>
    Private Sub Abrir(fic As String)
        If String.IsNullOrWhiteSpace(fic) Then
            Return
        End If

        Dim sDirFic = Path.GetDirectoryName(fic)
        If Not File.Exists(fic) Then
            If String.IsNullOrWhiteSpace(sDirFic) Then
                fic = Path.Combine(dirDocumentos, fic)
            End If
            If File.Exists(fic) Then
                Abrir(fic)
            End If
            Return
        End If

        If String.IsNullOrWhiteSpace(sDirFic) Then
            fic = Path.Combine(dirDocumentos, fic)
        End If

        labelInfo.Text = $"Abriendo {fic}..."

        ' solo ocultarlo si no es el mismo fichero                  (25/Sep/20)
        If fic <> nombreUltimoFichero Then
            splitContainer1.Panel2.Visible = False
            splitContainer1_Resize(Nothing, Nothing)
            'richTextBoxSyntax.Text = ""
            lstSyntax.Items.Clear()
        End If

        Dim sCodigo = ""
        Using sr As New StreamReader(fic, detectEncodingFromByteOrderMarks:=True, encoding:=Encoding.UTF8)
            sCodigo = sr.ReadToEnd()
        End Using

        ' Si se deben cambiar los TAB por 8 espacios                (05/Oct/20)
        If cambiarTabs Then
            sCodigo = sCodigo.Replace(vbTab, "        ")
        End If

        codigoAnterior = sCodigo

        richTextBoxCodigo.Text = sCodigo
        'richTextBoxCodigo.Modified = False
        richTextBoxCodigo.SelectionStart = 0
        richTextBoxCodigo.SelectionLength = 0
        MostrarPosicion(Nothing)

        ' Comprobar si hay que añadir el fichero a la lista de recientes
        sDirFic = Path.GetDirectoryName(fic)
        If sDirFic = dirDocumentos Then
            fic = Path.GetFileName(fic)
        End If
        nombreUltimoFichero = fic
        comboBoxFileName.Text = fic
        AñadirAlComboBoxFileName(fic)

        If MarcadorFichero <> fic Then
            MarcadorQuitarTodos()
            MarcadorFichero = fic
        End If

        ' Extensiones para usar como ficheros de texto              (08/Oct/20)
        Dim ExtensionesTexto As String() = {".txt", ".md", ".log"}

        Dim extension = Path.GetExtension(fic).ToLower

        ' Asignar el lenguaje en los combos
        ' Solo comprobar vb y cs, el resto se considera texto       (08/Oct/20)
        Dim sLenguaje As String
        If extension = ".cs" Then
            sLenguaje = Compilar.LenguajeCSharp
        ElseIf extension = ".vb" Then
            sLenguaje = Compilar.LenguajeVisualBasic
        Else
            sLenguaje = ExtensionTexto

            'ElseIf ExtensionesTexto.Contains(extension) Then
            '    sLenguaje = ExtensionTexto
            'Else
            '    If sCodigo.Contains("end sub") Then
            '        sLenguaje = Compilar.LenguajeVisualBasic
            '    ElseIf sCodigo.ToLower().Contains(");") Then
            '        sLenguaje = Compilar.LenguajeCSharp
            '    Else
            '        sLenguaje = ExtensionTexto
            '    End If
        End If
        If sLenguaje = Compilar.LenguajeVisualBasic Then
            buttonLenguaje.Image = buttonLenguaje.DropDownItems(0).Image
        ElseIf sLenguaje = Compilar.LenguajeCSharp Then
            buttonLenguaje.Image = buttonLenguaje.DropDownItems(1).Image
        Else
            buttonLenguaje.Image = buttonLenguaje.DropDownItems(2).Image
        End If
        buttonLenguaje.Text = sLenguaje

        ' Mostrar información del fichero
        labelInfo.Text = $"{Path.GetFileName(fic)} ({sDirFic})"
        'labelTamaño.Text = $"{richTextBoxCodigo.Text.Length:#,##0} car."
        labelTamaño.Text = $"{richTextBoxCodigo.Text.Length:#,##0} car. ({richTextBoxCodigo.Text.CuantasPalabras:#,##0} palab.)"

        ' Si hay que colorear el fichero cargado
        If colorearAlCargar Then
            ColorearCodigo()
        End If

        ' Si es texto, deshabilitar los botones que correspondan    (08/Oct/20)
        If sLenguaje = ExtensionTexto Then
            HabilitarBotones()
        End If

        TextoModificado = False

    End Sub

    ''' <summary>
    ''' Guarda el fichero indicado en el parámetro
    ''' </summary>
    ''' <param name="fic">El path completo del fichero a guardar</param>
    Private Sub Guardar(fic As String)
        labelInfo.Text = $"Guardando {fic}..."

        Dim sCodigo = richTextBoxCodigo.Text

        Dim sDirFic = Path.GetDirectoryName(fic)
        If String.IsNullOrWhiteSpace(sDirFic) Then
            fic = Path.Combine(dirDocumentos, fic)
        End If

        ' Si se deben cambiar los TAB por 8 espacios                (05/Oct/20)
        If cambiarTabs Then
            sCodigo = sCodigo.Replace(vbTab, "        ")
        End If

        Using sw As New StreamWriter(fic, append:=False, encoding:=Encoding.UTF8)
            sw.WriteLine(sCodigo)
        End Using
        codigoAnterior = sCodigo

        labelInfo.Text = $"{Path.GetFileName(fic)} ({Path.GetDirectoryName(fic)})"
        'labelTamaño.Text = $"{richTextBoxCodigo.Text.Length:#,##0} car."
        labelTamaño.Text = $"{richTextBoxCodigo.Text.Length:#,##0} car. ({richTextBoxCodigo.Text.CuantasPalabras:#,##0} palab.)"

        TextoModificado = False

        'richTextBoxCodigo.Modified = False
        'labelModificado.Text = " "
        Dim fic2 = fic
        If Path.GetDirectoryName(fic) = dirDocumentos Then
            fic2 = Path.GetFileName(fic)
        End If
        ' para que se muestre solo el nombre, si está en documentos (02/Oct/20)
        nombreUltimoFichero = fic2
        comboBoxFileName.Text = fic2
        AñadirAlComboBoxFileName(fic)

        ' Si es texto, deshabilitar los botones que correspondan    (08/Oct/20)
        ' y asignar el lenguaje en el botón de lenguaje
        Dim extension = Path.GetExtension(fic).ToLower

        ' Asignar el lenguaje en los combos
        ' Solo comprobar vb y cs, el resto se considera texto       (08/Oct/20)
        Dim sLenguaje As String
        If extension = ".cs" Then
            sLenguaje = Compilar.LenguajeCSharp
        ElseIf extension = ".vb" Then
            sLenguaje = Compilar.LenguajeVisualBasic
        Else
            sLenguaje = ExtensionTexto
        End If
        If sLenguaje = Compilar.LenguajeVisualBasic Then
            buttonLenguaje.Image = buttonLenguaje.DropDownItems(0).Image
        ElseIf sLenguaje = Compilar.LenguajeCSharp Then
            buttonLenguaje.Image = buttonLenguaje.DropDownItems(1).Image
        Else
            buttonLenguaje.Image = buttonLenguaje.DropDownItems(2).Image
        End If
        buttonLenguaje.Text = sLenguaje

        If sLenguaje = ExtensionTexto Then
            HabilitarBotones()
        End If


    End Sub

    ''' <summary>
    ''' Borra la ventana de código y deja en blanco el <see cref="nombreUltimoFichero"/>.
    ''' Si se ha modificado el que había, pregunta si lo quieres guardar
    ''' </summary>
    Private Sub Nuevo()
        If TextoModificado Then GuardarComo()

        richTextBoxCodigo.Text = ""
        richTextBoxtLineas.Text = ""
        'richTextBoxSyntax.Text = ""
        lstSyntax.Items.Clear()
        nombreUltimoFichero = ""
        TextoModificado = False
        richTextBoxCodigo.Modified = False
        labelInfo.Text = ""
        labelPos.Text = "Lín: 1  Car: 1"
        'labelTamaño.Text = "0 car."
        labelTamaño.Text = $"{0:#,##0} car. ({0:#,##0} palab.)"
        codigoAnterior = ""
    End Sub

    ''' <summary>
    ''' Muestra el cuadro de diálogo de Guardar como.
    ''' </summary>
    Private Sub GuardarComo()

        Dim fichero = comboBoxFileName.Text

        'If nombreUltimoFichero <> comboBoxFileName.Text Then
        Using sFD As New SaveFileDialog
            sFD.Title = "Seleccionar fichero para guardar el código"
            sFD.FileName = fichero
            sFD.InitialDirectory = dirDocumentos
            sFD.RestoreDirectory = True
            sFD.Filter = "Código de Visual Basic y CSharp (*.vb;*.cs)|*.vb;*.cs|Textos (*.txt;*.log;*.md)|*.txt;*.log;*.md|Todos los ficheros (*.*)|*.*"
            If sFD.ShowDialog = DialogResult.Cancel Then
                Return
            End If
            fichero = sFD.FileName
            nombreUltimoFichero = sFD.FileName
            ' Guardarlo
            Guardar(fichero)
        End Using
        'End If
    End Sub

    ''' <summary>
    ''' Guarda el fichero actual.
    ''' Si no tiene nombre muestra el diálogo de guardar como
    ''' </summary>
    Private Sub Guardar()
        If String.IsNullOrEmpty(nombreUltimoFichero) Then
            GuardarComo()
            Return
        End If
        Guardar(nombreUltimoFichero)
    End Sub


    ''' <summary>
    ''' Añade (si no está) el fichero indicado al principio de la lista de ficheros y
    ''' asigna la lista al menú de los ficheros recientes.
    ''' </summary>
    ''' <param name="fic">El fichero a añadir al <see cref="comboBoxFileName"/></param>
    Private Sub AñadirAlComboBoxFileName(fic As String)
        If String.IsNullOrWhiteSpace(fic) Then Return

        ' No añadir el path si es el directorio de Documentos       (24/Sep/20)
        If Path.GetDirectoryName(fic) = dirDocumentos Then
            fic = Path.GetFileName(fic)
        End If

        ' Clasificar los elementos del combo                        (24/Sep/20)
        inicializando = True


        Dim col As New List(Of String)

        ' Añadirlo al principio                                     (28/Sep/20)
        col.Add(fic)

        ' Usando el método de extensión                             (26/Sep/20)
        col.AddRange(comboBoxFileName.ToList())

        comboBoxFileName.Items.Clear()
        ' quitar los repetidos
        ' hacerlo así para que el último abiero/guardado            (28/Sep/20)
        ' siempre esté el primero de la lista
        comboBoxFileName.Items.AddRange(col.Distinct().ToArray())
        'comboBoxFileName.Items.AddRange(col.ToArray())

        AsignarRecientes()

        inicializando = False
    End Sub


    Private Sub comboBoxFileName_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            Abrir(comboBoxFileName.Text)
        End If
    End Sub

    Private Sub comboBoxFileName_Validating(sender As Object, e As ComponentModel.CancelEventArgs)
        If inicializando Then Return

        Dim fic = comboBoxFileName.Text
        If String.IsNullOrWhiteSpace(fic) Then Return

        AñadirAlComboBoxFileName(fic)
    End Sub

    ''' <summary>
    ''' Muestra la ventana informativa sobre esta utilidad
    ''' y las DLL que utiliza
    ''' </summary>
    Private Sub AcercaDe()
        ' Añadir la versión de esta utilidad                        (15/Sep/20)
        Dim ensamblado As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly
        Dim fvi = System.Diagnostics.FileVersionInfo.GetVersionInfo(ensamblado.Location)
        'Dim vers = $" v{My.Application.Info.Version} ({fvi.FileVersion})"
        Dim versionAttr = ensamblado.GetCustomAttributes(GetType(System.Reflection.AssemblyVersionAttribute), False)
        Dim vers = If(versionAttr.Length > 0,
                                (TryCast(versionAttr(0), System.Reflection.AssemblyVersionAttribute)).Version,
                                "1.0.0.0")
        Dim prodAttr = ensamblado.GetCustomAttributes(GetType(System.Reflection.AssemblyProductAttribute), False)
        Dim producto = If(prodAttr.Length > 0,
                                (TryCast(prodAttr(0), System.Reflection.AssemblyProductAttribute)).Product,
                                "gsEvaluarColorearNET")
        Dim descAttr = ensamblado.GetCustomAttributes(GetType(System.Reflection.AssemblyDescriptionAttribute), False)
        Dim desc = If(descAttr.Length > 0,
                                (TryCast(descAttr(0), System.Reflection.AssemblyDescriptionAttribute)).Description,
                                "(para .NET 5.0 RC1 revisión del 26/Sep/2020)")
        Dim k = desc.IndexOf("(para .NET")
        Dim desc1 = desc.Substring(k)
        Dim descL = desc.Substring(0, k - 1)

        'Dim verColorear = gsCol.Version(completa:=True)
        'Dim verCompilar = Compilar.Version(completa:=True)
        'Dim i = verColorear.LastIndexOf(" (")
        'If i > -1 Then
        '    verColorear = $"{verColorear.Substring(0, i)}{vbCrLf}{verColorear.Substring(i + 1)}"
        'End If
        'i = verCompilar.LastIndexOf(" (")
        'If i > -1 Then
        '    verCompilar = $"{verCompilar.Substring(0, i)}{vbCrLf}{verCompilar.Substring(i + 1)}"
        'End If

        'Dim descL = "Utilidad para .NET 5.0 (.NET Core) para mostrar código de VB o C#, colorearlo y compilarlo."

        MessageBox.Show($"{producto} v{vers} ({fvi.FileVersion})" & vbCrLf & vbCrLf &
                        $"{descL}" & vbCrLf &
                        $"{desc1}", ' & vbCrLf & vbCrLf & "Usando las DLL externas:" & vbCrLf & ' verColorear & vbCrLf & vbCrLf & verCompilar,
                        $"Acerca de {producto}",
                        MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub


    '
    ' Los métodos de edición
    '

    ''' <summary>
    ''' Pega el texto del portapapeles en el editor
    ''' </summary>
    Private Sub Pegar()
        If richTextBoxCodigo.CanPaste(DataFormats.GetFormat("Text")) Then
            richTextBoxCodigo.Paste(DataFormats.GetFormat("Text"))

            Dim selStart = richTextBoxCodigo.SelectionStart
            Dim lin = richTextBoxCodigo.GetLineFromCharIndex(selStart)
            Dim pos = richTextBoxCodigo.GetFirstCharIndexFromLine(lin)
            Dim sLinea = richTextBoxCodigo.Lines(lin)


            ' obligar a poner las líneas                            (18/Sep/20)
            AñadirNumerosDeLinea()

            ' Seleccionar el texto después de pegar                 (04/Oct/20)
            richTextBoxCodigo.SelectionStart = pos
            richTextBoxCodigo.SelectionLength = sLinea.Length
            ' y colorearlo si procede
            ColorearSeleccion()

        End If
    End Sub

    ''' <summary>
    ''' Cuando se abre el menú de edición o se cambia la selección del código
    ''' asignar si están o no habilitados el menú en sí, el contextual y las barras de herramientas.
    ''' </summary>
    Private Sub menuEditDropDownOpening()
        If inicializando Then Return

        inicializando = True

        ' para saber si hay texto en el control
        Dim b = richTextBoxCodigo.TextLength > 0

        menuEditDeshacer.Enabled = richTextBoxCodigo.CanUndo
        menuEditRehacer.Enabled = richTextBoxCodigo.CanRedo
        menuEditCopiar.Enabled = richTextBoxCodigo.SelectionLength > 0
        menuEditCortar.Enabled = menuEditCopiar.Enabled
        menuEditSeleccionarTodo.Enabled = b
        menuEditPegar.Enabled = richTextBoxCodigo.CanPaste(DataFormats.GetFormat("Text"))

        'If Clipboard.GetDataObject.GetDataPresent(DataFormats.Text) Then
        '    menuEditPegar.Enabled = richTextBoxCodigo.CanPaste(DataFormats.GetFormat(DataFormats.Text))
        'ElseIf Clipboard.GetDataObject.GetDataPresent(DataFormats.StringFormat) Then
        '    ' StringFormat                                          (30/Oct/04)
        '    menuEditPegar.Enabled = richTextBoxCodigo.CanPaste(DataFormats.GetFormat(DataFormats.StringFormat))
        'ElseIf Clipboard.GetDataObject.GetDataPresent(DataFormats.Html) Then
        '    menuEditPegar.Enabled = richTextBoxCodigo.CanPaste(DataFormats.GetFormat(DataFormats.Html))
        'ElseIf Clipboard.GetDataObject.GetDataPresent(DataFormats.OemText) Then
        '    menuEditPegar.Enabled = richTextBoxCodigo.CanPaste(DataFormats.GetFormat(DataFormats.OemText))
        'ElseIf Clipboard.GetDataObject.GetDataPresent(DataFormats.UnicodeText) Then
        '    menuEditPegar.Enabled = richTextBoxCodigo.CanPaste(DataFormats.GetFormat(DataFormats.UnicodeText))
        'ElseIf Clipboard.GetDataObject.GetDataPresent(DataFormats.Rtf) Then
        '    menuEditPegar.Enabled = richTextBoxCodigo.CanPaste(DataFormats.GetFormat(DataFormats.Rtf))
        'End If

        menuEditBuscar.Enabled = b
        menuEditReemplazar.Enabled = b

        ' Actualizar también los del menú contextual
        For j = 0 To rtbCodigoContext.Items.Count - 1
            For i = 0 To menuEdit.DropDownItems.Count - 1
                If rtbCodigoContext.Items(j).Name = menuEdit.DropDownItems(i).Name Then
                    rtbCodigoContext.Items(j).Enabled = menuEdit.DropDownItems(i).Enabled
                    Exit For
                End If
            Next
        Next

        inicializando = False
    End Sub

    ''' <summary>
    ''' Para saber si se deben o no habilitar los botones de los toolStrip
    ''' Antes estaba en MenuEdiDropDownOpening
    ''' </summary>
    Private Sub HabilitarBotones()
        If inicializando Then Return

        inicializando = True

        ' para saber si hay texto en el control
        Dim b = richTextBoxCodigo.TextLength > 0

        menuEditDeshacer.Enabled = richTextBoxCodigo.CanUndo
        menuEditRehacer.Enabled = richTextBoxCodigo.CanRedo
        buttonCopiar.Enabled = richTextBoxCodigo.SelectionLength > 0
        buttonCortar.Enabled = buttonCopiar.Enabled
        buttonPegar.Enabled = richTextBoxCodigo.CanPaste(DataFormats.GetFormat("Text"))

        buttonBuscarSiguiente.Enabled = b
        buttonReemplazarSiguiente.Enabled = b
        buttonReemplazarTodo.Enabled = b

        buttonEditorPonerIndentacion.Enabled = b
        buttonEditorQuitarIndentacion.Enabled = b
        buttonEditorPonerComentarios.Enabled = b
        buttonEditorQuitarComentarios.Enabled = b

        ' Si el Lenguaje no es vb/cs no:                            (08/Oct/20)
        '   compilar, ejecutar, colorear, etc.
        Dim b2 = b
        If buttonLenguaje.Text = ExtensionTexto Then
            b2 = False
        End If
        buttonCompilar.Enabled = b2
        buttonEjecutar.Enabled = b2
        buttonEvaluar.Enabled = b2
        buttonColorear.Enabled = b2
        buttonNoColorear.Enabled = b2
        buttonColorearHTML.Enabled = b2
        chkMostrarLineasHTML.Enabled = b2

        buttonEditorMarcador.Enabled = b
        buttonEditorClasificarSeleccion.Enabled = richTextBoxCodigo.SelectionLength > 0

        b = ColMarcadores.Count > 0
        buttonEditorMarcadorAnterior.Enabled = b
        buttonEditorMarcadorSiguiente.Enabled = b
        buttonEditorMarcadorQuitarTodos.Enabled = b

        b = ColRecortes.Count > 0
        buttonEdicionRecortes.Enabled = b

        inicializando = False
    End Sub

    '
    ' Buscar y reemplazar
    '

    ''' <summary>
    ''' Muestra el panel de buscar/reemplazar
    ''' y reinicia la posición de búsqueda
    ''' </summary>
    ''' <param name="esBuscar">True si es Buscar, false si es Reemplazar</param>
    Private Sub BuscarReemplazar(esBuscar As Boolean)
        ' Mostrar el panel de buscar/reemplazar
        MostrarPanelBuscar(True)
        esCtrlF = True
        ' Si hay texto seleccionado, usarlo                         (30/Sep/20)
        If richTextBoxCodigo.SelectionLength > 0 Then
            comboBoxBuscar.Text = richTextBoxCodigo.SelectedText
        End If
        If esBuscar Then
            comboBoxBuscar.Focus()
        Else
            comboBoxReemplazar.Focus()
        End If
    End Sub

    ''' <summary>
    ''' Muestra u oculta el panel de buscar/reemplazar
    ''' </summary>
    ''' <param name="mostrar">True si se debe mostrar, false si se oculta</param>
    Private Sub MostrarPanelBuscar(mostrar As Boolean)
        panelBuscar.Visible = mostrar
        menuMostrar_Buscar.Checked = mostrar
        If mostrar Then
            esCtrlF = True
        End If
    End Sub

    ''' <summary>
    ''' Buscar siguiente coincidencia.
    ''' Devuelve False si no hay más
    ''' </summary>
    Private Function BuscarSiguiente(esReemplazar As Boolean) As Boolean
        ' Buscar en el texto lo indicado
        ' Se empieza desde la posición actual del texto

        buscarQueBuscar = comboBoxBuscar.Text
        ' Cambiar los \t por tabs y los \n por vbCr                 (29/Sep/20)
        buscarQueBuscar = buscarQueBuscar.Replace("\n", vbCr).Replace("\t", vbTab)

        If String.IsNullOrEmpty(buscarQueBuscar) Then
            If panelBuscar.Visible = False Then
                BuscarReemplazar(True)
            End If
            Return False
        End If

        If esCtrlF Then
            buscarPosIni = richTextBoxCodigo.SelectionStart
            buscarPos = buscarPosIni
            buscarPrimeraCoincidencia = True
            esCtrlF = False
        End If

        Dim rtbFinds = RichTextBoxFinds.None
        If buscarMatchCase Then rtbFinds = RichTextBoxFinds.MatchCase
        If buscarWholeWord Then rtbFinds = rtbFinds Or RichTextBoxFinds.WholeWord

        If buscarPos >= richTextBoxCodigo.Text.Length Then
            buscarPos = 0
        ElseIf buscarPos < 0 Then
            buscarPos = 0
        End If

        buscarPos = richTextBoxCodigo.Find(buscarQueBuscar, buscarPos, rtbFinds)
        If buscarPos = -1 Then
            If esReemplazar Then
                If numBuscarReemplazar = 0 Then
                    numBuscarReemplazar += 1
                    buscarPos = 0
                    If Not BuscarSiguiente(esReemplazar) Then
                        MessageBox.Show("No hay más coincidencias.",
                                        "Reemplazar siguiente",
                                        MessageBoxButtons.OK, MessageBoxIcon.Information)
                        esCtrlF = True
                        Return False
                    End If
                End If
                Return False
            End If
            buscarPos = 0
            Return BuscarSiguiente(esReemplazar)
        Else
            If buscarPrimeraCoincidencia Then
                buscarPrimeraCoincidencia = False
                buscarPosIni = buscarPos
            Else
                richTextBoxCodigo.SelectionStart = buscarPos
                richTextBoxCodigo.SelectionLength = buscarQueBuscar.Length

                If buscarPos = buscarPosIni Then
                    MessageBox.Show("No hay más coincidencias, se ha llegado al inicio de la búsqueda.",
                            "Buscar siguiente",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
                    esCtrlF = True
                    Return False
                End If
            End If

            buscarPos += buscarQueBuscar.Length
        End If
        Return True
    End Function

    Private numBuscarReemplazar As Integer

    ''' <summary>
    ''' Reemplaza el texto si la selección es lo buscado y
    ''' sigue buscando
    ''' </summary>
    Private Sub ReemplazarSiguiente()
        ' Cambiar los \t por tabs y los \n por vbCr                 (29/Sep/20)
        buscarQueReemplazar = comboBoxReemplazar.Text
        buscarQueReemplazar = buscarQueReemplazar.Replace("\n", vbCr).Replace("\t", vbTab)
        buscarQueBuscar = comboBoxBuscar.Text
        buscarQueBuscar = buscarQueBuscar.Replace("\n", vbCr).Replace("\t", vbTab)

        If buscarPos < 0 Then
            buscarPos = 0
        Else
            buscarPos = richTextBoxCodigo.SelectionStart + 1
        End If
        ' si hay texto seleccionado y es lo que se busca
        If richTextBoxCodigo.SelectedText = buscarQueBuscar Then
            ' reemplazar la actual
            richTextBoxCodigo.SelectedText = buscarQueReemplazar
            buscarPos += buscarQueReemplazar.Length
        End If
        ' y buscar la siguiente
        'buscarPrimeraCoincidencia = True
        'esCtrlF = True
        numBuscarReemplazar = 0
        If Not BuscarSiguiente(esReemplazar:=True) Then
            Return
        End If

    End Sub

    ''' <summary>
    ''' Reemplaza (cambia) todas las coincidencias
    ''' del texto buscado (<see cref="buscarQueBuscar"/>)
    ''' por el indicado en reemplazar (<see cref="buscarQueReemplazar"/>
    ''' </summary>
    ''' <remarks>Se empieza buscando desde el principio del texto</remarks>
    Private Sub ReemplazarTodo()
        ' Cambiar los \t por tabs y los \n por vbCr                 (29/Sep/20)
        buscarQueReemplazar = comboBoxReemplazar.Text
        buscarQueReemplazar = buscarQueReemplazar.Replace("\n", vbCr).Replace("\t", vbTab)
        buscarQueBuscar = comboBoxBuscar.Text
        buscarQueBuscar = buscarQueBuscar.Replace("\n", vbCr).Replace("\t", vbTab)

        Dim rtbFinds = RichTextBoxFinds.None
        If buscarMatchCase Then rtbFinds = RichTextBoxFinds.MatchCase
        If buscarWholeWord Then rtbFinds = rtbFinds Or RichTextBoxFinds.WholeWord

        Dim t = 0
        buscarPos = 0
        Do While buscarPos > -1
            buscarPos = richTextBoxCodigo.Find(buscarQueBuscar, buscarPos, rtbFinds)
            If buscarPos > -1 Then
                t += 1
                If richTextBoxCodigo.SelectionLength > 0 Then
                    richTextBoxCodigo.SelectedText = buscarQueReemplazar
                    buscarPos += buscarQueReemplazar.Length
                Else
                    Exit Do
                End If
            End If
        Loop
        ' motrar todos los cambios realizados
        If t > 0 Then
            Dim plural = If(t > 1, "s", "")
            MessageBox.Show($"{t} cambio{plural} realizado{plural}.",
                            "Reemplazar todos",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show($"No se encontró el texto especificado:{vbCrLf}{buscarQueBuscar}",
                            "Reemplazar todos",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub

    ''' <summary>
    ''' Se produce cuando cambia el texto o se selecciona
    ''' del combo de reemplazar
    ''' </summary>
    Private Sub comboBoxReemplazar_Validating()
        If inicializando Then Return
        inicializando = True
        ' En reemplazar aceptar cadenas vacías
        ' FindStringExact no diferencia las mayúsculas y minúsculas
        Dim i = comboBoxReemplazar.FindStringExact(comboBoxReemplazar.Text)
        If i = -1 Then
            comboBoxReemplazar.Items.Add(comboBoxReemplazar.Text)
        End If
        inicializando = False
    End Sub

    ''' <summary>
    ''' Se produce cuando cambia el texto o se selecciona
    ''' del combo de buscar
    ''' </summary>
    Private Sub comboBoxBuscar_Validating()
        If inicializando Then Return
        inicializando = True
        ' En buscar no aceptar cadenas vacías
        ' FindStringExact no diferencia las mayúsculas y minúsculas
        Dim i = If(String.IsNullOrEmpty(comboBoxBuscar.Text), -1, comboBoxBuscar.FindStringExact(comboBoxBuscar.Text))
        If i = -1 Then
            comboBoxBuscar.Items.Add(comboBoxBuscar.Text)
        End If
        inicializando = False

        ' Si ha cambiado el texto a buscar                          (18/Sep/20)
        ' será como si se pulsase CtrlF
        If buscarQueBuscar <> comboBoxBuscar.Text Then
            esCtrlF = True
        End If
    End Sub


    '
    ' Los métodos de herramientas, colorear
    '

    ''' <summary>
    ''' Mostrar el código sin colorear
    ''' </summary>
    Private Sub NoColorear()
        Dim modifTemp = richTextBoxCodigo.Modified

        Dim codigo = richTextBoxCodigo.Text
        richTextBoxCodigo.Rtf = ""
        richTextBoxCodigo.Text = codigo

        richTextBoxCodigo.Modified = modifTemp
    End Sub

    ''' <summary>
    ''' Colorea el código en el editor.
    ''' Los lenguajes usados son Visual Basic y C#, se usará el indicado en buttonLenguaje.Text.
    ''' </summary>
    Private Sub ColorearCodigo()
        If richTextBoxCodigo.TextLength = 0 Then Return

        ' Si buttonLenguaje.Text es Texto, no colorear              (08/Oct/20)
        If buttonLenguaje.Text = ExtensionTexto Then Return

        labelInfo.Text = $"Coloreando el código de {buttonLenguaje.Text}..."
        Me.Refresh()

        Dim modif = richTextBoxCodigo.Modified
        inicializando = True
        'richTextBoxCodigo.Visible = False

        Dim colCodigo = Compilar.ColoreaRichTextBox(richTextBoxCodigo, buttonLenguaje.Text)

        MostrarResultadoEvaluar(colCodigo, False)

        'richTextBoxCodigo.Visible = True

        inicializando = False
        richTextBoxCodigo.SelectionStart = 0
        richTextBoxCodigo.SelectionLength = 0
        richTextBoxCodigo.Modified = modif

        labelInfo.Text = $"Código coloreado para {buttonLenguaje.Text}."
        Me.Refresh()
    End Sub

    ''' <summary>
    ''' Colorear el código en formato HTML y mostrarlo en ventana aparte
    ''' </summary>
    Private Sub ColorearHTML()
        If richTextBoxCodigo.TextLength = 0 Then Return

        labelInfo.Text = $"Coloreando en HTML para {buttonLenguaje.Text}..."
        Me.Refresh()

        Dim sHTML = Compilar.ColoreaHTML(richTextBoxCodigo.Text,
                                         buttonLenguaje.Text,
                                         mostrarLineas:=chkMostrarLineasHTML.Checked)

        labelInfo.Text = "Mostrando el visor de HTML..."
        Me.Refresh()

        Dim fHTML As New FormVisorHTML(comboBoxFileName.Text, sHTML)
        fHTML.Show()

        labelInfo.Text = ""
    End Sub

    '
    ' Los métodos de compilar, ejecutar, evaluar
    '

    ''' <summary>
    ''' Compila y ejecuta el código actual.
    ''' Si se producen errores muestra la información con los errores.
    ''' </summary>
    Private Sub CompilarEjecutar()
        If richTextBoxCodigo.TextLength = 0 Then Return

        ' guardar el código antes de compilar y ejecutar
        labelInfo.Text = "Compilando el código..."
        Me.Refresh()

        ' Comprobar si es el mismo fichero que se está editando     (22/Sep/20)
        ' Si no es así, se pregunta si se quiere guardar
        If nombreUltimoFichero <> comboBoxFileName.Text Then
            GuardarComo()
        Else
            ' Esto no estaba y no se guardaban los cambios          (02/Oct/20)
            Guardar()
        End If

        Dim fichero = comboBoxFileName.Text

        nombreUltimoFichero = fichero
        comboBoxFileName.Text = fichero
        AñadirAlComboBoxFileName(fichero)

        ' Si no tiene el path, añadirlo                             (27/Sep/20)
        If String.IsNullOrEmpty(Path.GetDirectoryName(fichero)) Then
            fichero = Path.Combine(dirDocumentos, fichero)
        End If

        Dim res = Compilar.CompileRun(fichero, run:=True)
        If res.Result.Success Then
            labelInfo.Text = $"Se ha compilado y ejecutado satisfactoriamente."

            lstSyntax.Items.Clear()
            splitContainer1.Panel2.Visible = False
            splitContainer1_Resize(Nothing, Nothing)

        Else
            ' Mostrar los errores
            MostrarErrores(res.Result)

        End If

    End Sub

    ''' <summary>
    ''' Compilar el código sin ejecutar
    ''' </summary>
    Private Sub Build()
        If richTextBoxCodigo.TextLength = 0 Then Return

        If TextoModificado Then
            Guardar()
        End If

        Dim filepath = nombreUltimoFichero
        labelInfo.Text = $"Compilando para {buttonLenguaje.Text}..."
        Me.Refresh()

        ' Si no tiene el path, añadirlo                             (27/Sep/20)
        If String.IsNullOrEmpty(Path.GetDirectoryName(filepath)) Then
            filepath = Path.Combine(dirDocumentos, filepath)
        End If

        Dim res = Compilar.CompileRun(filepath, run:=False)

        lstSyntax.Items.Clear()

        If Not res.Result.Success Then
            MostrarErrores(res.Result)
            Return
        End If
        labelInfo.Text = res.OutputPath
    End Sub

    ''' <summary>
    ''' Evalúa el código, 
    ''' si se indica compilar al evaluar, se muestran los posibles errores,
    ''' si se indica colorear al evaluar, se colorea el código.
    ''' Si se colorea o no al evaluar, se muestran los miembros del código:
    ''' clases, métodos, palabras claves, etc.
    ''' </summary>
    Private Sub Evaluar()
        If richTextBoxCodigo.TextLength = 0 Then Return

        If TextoModificado Then
            Guardar()
        End If

        Dim res As EmitResult
        lstSyntax.Items.Clear()

        If compilarAlEvaluar Then
            splitContainer1.Panel2.Visible = False
            splitContainer1_Resize(Nothing, Nothing)
            lstSyntax.Items.Clear()

            labelInfo.Text = $"Compilando para {buttonLenguaje.Text}..."
            Me.Refresh()

            res = Compilar.ComprobarCodigo(richTextBoxCodigo.Text, buttonLenguaje.Text)
            If res.Success = False Then
                MostrarErrores(res)
                Return
            End If
            labelInfo.Text = "Compilado sin error."
        End If

        Dim colCodigo As Dictionary(Of String, Dictionary(Of String, ClassifSpanInfo))

        If colorearAlEvaluar Then
            labelInfo.Text = "Coloreando el código..."
            Me.Refresh()

            Dim modif = richTextBoxCodigo.Modified
            inicializando = True
            'richTextBoxCodigo.Visible = False

            colCodigo = Compilar.ColoreaRichTextBox(richTextBoxCodigo, buttonLenguaje.Text)

            'richTextBoxCodigo.Visible = True

            inicializando = False
            richTextBoxCodigo.SelectionStart = 0
            richTextBoxCodigo.SelectionLength = 0
            richTextBoxCodigo.Modified = modif

        Else
            colCodigo = Compilar.EvaluaCodigo(richTextBoxCodigo.Text, buttonLenguaje.Text)
        End If

        MostrarResultadoEvaluar(colCodigo, True)
    End Sub

    ''' <summary>
    ''' Muestra los errores de la compilación (al evaluar o compilar).
    ''' </summary>
    ''' <param name="res">Un objeto del tipo <see cref="EmitResult"/> con los errores y advertencias.</param>
    Private Sub MostrarErrores(res As EmitResult)
        Dim errors = 0
        Dim warnings = 0

        lstSyntax.Items.Clear()
        ' Ponerlo en ownerdraw para que coloree
        lstSyntax.DrawMode = DrawMode.OwnerDrawFixed

        For Each diag As Diagnostic In res.Diagnostics
            Dim dcsI As New DiagInfo(diag)
            lstSyntax.Items.Add(dcsI)
            'Dim index = lstSyntax.Items.Add(dcsI)
            'lstSyntax.SelectedIndex = index
            'lstSyntax.Refresh()
            'Application.DoEvents()

            If diag.Severity = DiagnosticSeverity.Error Then
                errors += 1
            ElseIf diag.Severity = DiagnosticSeverity.Warning Then
                warnings += 1
            End If
        Next
        'lstSyntax.Refresh()

        ' Ponerlo en normal para que no repinte
        'lstSyntax.DrawMode = DrawMode.Normal
        'timer1.Interval = 5000
        'timer1.Start()

        labelInfo.Text = $"Compilado con {errors} error{If(errors = 1, "", "es")} y {warnings} advertencia{If(warnings = 1, "", "s")}."

        splitContainer1.Panel2.Visible = True
        splitContainer1_Resize(Nothing, Nothing)
    End Sub

    ''' <summary>
    ''' Mostrar el resultado de evaluar el código.
    ''' Las clases, métodos, palabras clave, etc. definidos en el código.
    ''' </summary>
    ''' <param name="colCodigo">Colección del tipo <see cref="Dictionary(Of String, List(Of DiagClassifSpanInfo))"/>
    ''' con la lista con los valores sacados de <see cref="ClassifiedSpan"/> </param>
    Private Sub MostrarResultadoEvaluar(colCodigo As Dictionary(Of String, Dictionary(Of String, ClassifSpanInfo)),
                                        mostrarSyntax As Boolean)
        Dim codTiposCount As (Clases As Integer, Metodos As Integer, Keywords As Integer)
        codTiposCount.Clases = If(colCodigo.Keys.Contains("class name"), colCodigo("class name").Count, 0) + If(colCodigo.Keys.Contains("module name"), colCodigo("module name").Count, 0)
        codTiposCount.Metodos = If(colCodigo.Keys.Contains("method name"), colCodigo("method name").Count, 0)
        codTiposCount.Keywords = If(colCodigo.Keys.Contains("keyword"), colCodigo("keyword").Count, 0) + If(colCodigo.Keys.Contains("keyword - control"), colCodigo("keyword - control").Count, 0)

        labelInfo.Text = $"Código con {codTiposCount.Clases} clase{If(codTiposCount.Clases = 1, "", "s")}, {codTiposCount.Metodos} método{If(codTiposCount.Metodos = 1, "", "s")} y {codTiposCount.Keywords} palabra{If(codTiposCount.Keywords = 1, "", "s")} clave."

        lstSyntax.Items.Clear()
        ' Ponerlo en ownerdraw para que coloree
        lstSyntax.DrawMode = DrawMode.OwnerDrawFixed

        ' Clasificar las claves                                     (25/Sep/20)
        Dim claves = From kv In colCodigo Order By kv.Key Ascending Select kv

        For Each kv In claves
            Dim s1 = kv.Key

            ' No mostrar el contenido de estos símbolos
            If s1 = "comment" OrElse s1.StartsWith("string") OrElse
                s1 = "punctuation" OrElse s1.StartsWith("operator") OrElse
                s1 = "number" OrElse s1.StartsWith("xml") Then
            Else
                lstSyntax.Items.Add(s1)
                Dim claves2 = From kv0 In kv.Value Order By kv0.Key Ascending Select kv0
                For Each kv1 In claves2
                    lstSyntax.Items.Add(kv1.Value)
                Next
            End If
        Next
        'lstSyntax.Refresh()

        ' Ponerlo en normal para que no repinte
        'lstSyntax.DrawMode = DrawMode.Normal
        'timer1.Interval = 5000
        'timer1.Start()

        If mostrarSyntax Then
            splitContainer1.Panel2.Visible = True
            splitContainer1_Resize(Nothing, Nothing)
        End If
    End Sub


    '
    ' Los métodos relacionados con el richTextBoxCodigo
    '

    'Private yaEstoy As Boolean = False

    Private Sub richTextBoxCodigo_KeyUp(sender As Object, e As KeyEventArgs)
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
                    e.Handled = True
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

    ''' <summary>
    ''' Sincroniza los scrollbar de richTextBoxCodigo y richTextBoxLineas.
    ''' </summary>
    Private Sub richTextBoxCodigo_VScroll(sender As Object, e As EventArgs)
        If inicializando Then Return
        If String.IsNullOrEmpty(richTextBoxCodigo.Text) Then Return

        Dim nPos As Integer = GetScrollPos(richTextBoxCodigo.Handle, ScrollBarType.SbVert)
        nPos <<= 16
        Dim wParam As Long = ScrollBarCommands.SB_THUMBPOSITION Or nPos
        SendMessage(richTextBoxtLineas.Handle, Message.WM_VSCROLL, wParam, 0)
    End Sub

    Private Sub richTextBoxCodigo_TextChanged(sender As Object, e As EventArgs)
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

    '
    ' Para la indentación                                           (28/Sep/20)
    '

    ''' <summary>
    ''' Poner indentación
    ''' </summary>
    ''' <param name="rtEditor"></param>
    Private Sub PonerIndentacion(rtEditor As RichTextBox)
        ' Si se selecciona el "retorno" de la última línea, se eliminará
        ' y se juntará con la siguiente.
        If rtEditor.SelectionLength > 0 Then
            Dim selStart As Integer = rtEditor.SelectionStart
            Dim sIndent As String '= New String(" "c, indentar)
            ' No tomar el texto rtf, sino el normal, que colorearemos
            ' El texto que está seleccionado antes de indentar      (05/Oct/20)
            Dim selAnt = rtEditor.SelectedText
            Dim lineas() As String = selAnt.Split(vbCr.ToCharArray)
            sIndent = New String(" "c, EspaciosIndentar)
            Dim sb As New System.Text.StringBuilder
            Dim j As Integer = lineas.Length - 1
            Dim k As Integer

            For i = 0 To j
                If lineas(i).Length > 0 AndAlso i < j Then
                    k += EspaciosIndentar
                End If
                If i = j Then
                    If lineas(i).Length = 0 Then
                        'sb.AppendFormat("{0}{1}", sIndent, lineas(i))
                        'k -= EspaciosIndentar
                        'Else
                        'sb.AppendFormat("{0}{1}", sIndent, lineas(i))
                    End If
                    'sb.AppendFormat("{0}{1}", sIndent, lineas(i).Replace(ChrW(9), sIndent))
                    sb.AppendFormat("{0}{1}", sIndent, lineas(i))
                Else
                    'sb.AppendFormat("{0}{1}{2}", sIndent, lineas(i).Replace(ChrW(9), sIndent), vbCr)
                    sb.AppendFormat("{0}{1}{2}", sIndent, lineas(i), vbCr)
                End If
            Next
            'Dim sbT = sb.ToString
            'If j = 1 Then
            '    sbt = sbt.TrimEnd(ChrW(13))
            'End If

            Dim lenAnt = selAnt.Length

            ' Poner la selección y sustituir el texto
            rtEditor.SelectionStart = selStart
            'rtEditor.SelectionLength = sb.Length - k - EspaciosIndentar 'lenAnt 
            rtEditor.SelectedText = sb.ToString.TrimEnd() & vbCr
            ' si se debe colorear                                   (03/Oct/20)
            rtEditor.SelectionStart = selStart
            rtEditor.SelectionLength = sb.Length - EspaciosIndentar
            ColorearSeleccion()

            ' para que se mantenga seleccionado                     (28/Sep/20)
            rtEditor.SelectionStart = selStart
            rtEditor.SelectionLength = sb.Length - EspaciosIndentar 'lenAnt + k
        Else
            Dim selStart As Integer = rtEditor.SelectionStart
            Dim sIndent As String = New String(" "c, EspaciosIndentar)
            If selStart > 0 Then
                rtEditor.SelectionLength = 0
                rtEditor.SelectedText = sIndent
                rtEditor.SelectionStart = selStart + EspaciosIndentar
                rtEditor.SelectionLength = 0
            Else
                rtEditor.SelectionStart = 0
                rtEditor.SelectionLength = 0
                rtEditor.SelectedText = sIndent
                rtEditor.SelectionStart = EspaciosIndentar
            End If
        End If

    End Sub

    ''' <summary>
    ''' Quitar indentación
    ''' </summary>
    ''' <param name="rtEditor"></param>
    Private Sub QuitarIndentacion(rtEditor As RichTextBox)
        If rtEditor.SelectionLength > 0 Then
            Dim selStart As Integer = rtEditor.SelectionStart
            Dim sIndent As String = New String(" "c, EspaciosIndentar)
            Dim lineas() As String = rtEditor.SelectedText.Split(vbCr.ToCharArray)
            Dim sb As New System.Text.StringBuilder
            Dim j As Integer = lineas.Length - 1

            For i As Integer = 0 To j
                ' si empieza con 4 caracteres, quitarlos
                If lineas(i) <> "" Then
                    If lineas(i).StartsWith(sIndent) Then
                        ' Si solo hay una línea seleccionada,       (26/Nov/05)
                        ' es que no hay que añadir retornos de carro
                        If j = 0 OrElse i = j Then
                            sb.AppendFormat("{0}", lineas(i).Substring(EspaciosIndentar))
                        Else
                            sb.AppendFormat("{0}{1}", lineas(i).Substring(EspaciosIndentar), vbCr)
                        End If
                    ElseIf lineas(i).StartsWith(vbTab) Then
                        ' Considerar los tab como indentación       (22/Nov/05)
                        If i = j Then
                            sb.AppendFormat("{0}", lineas(i).Substring(1))
                        Else
                            sb.AppendFormat("{0}{1}", lineas(i).Substring(1), vbCr)
                        End If
                    Else
                        If i = j Then
                            sb.AppendFormat("{0}", lineas(i))
                        Else
                            sb.AppendFormat("{0}{1}", lineas(i), vbCr)
                        End If
                    End If
                End If
            Next
            Dim sbT = sb.ToString
            If j = 1 Then
                sbT = sbT.TrimEnd(ChrW(13))
            End If
            rtEditor.SelectedText = sbT

            ' Que se mantenga seleccionado                          (22/Nov/05)
            rtEditor.SelectionStart = selStart
            rtEditor.SelectionLength = sbT.Length

            ' si se debe colorear                                   (03/Oct/20)
            ColorearSeleccion()
            rtEditor.SelectionStart = selStart
            rtEditor.SelectionLength = sbT.Length

        Else
            Dim selStart As Integer = rtEditor.SelectionStart
            Dim sIndent As String = New String(" "c, EspaciosIndentar)

            If selStart >= EspaciosIndentar Then
                rtEditor.SelectionStart = selStart - EspaciosIndentar
                rtEditor.SelectionLength = EspaciosIndentar
                If rtEditor.SelectedText = sIndent Then
                    rtEditor.SelectedText = ""
                    rtEditor.SelectionStart = selStart - EspaciosIndentar
                Else
                    rtEditor.SelectionStart = selStart
                End If
                rtEditor.SelectionLength = 0
            Else
                rtEditor.SelectionLength = EspaciosIndentar - rtEditor.SelectionStart
                rtEditor.SelectedText = ""

                rtEditor.SelectionLength = 0
            End If
        End If
    End Sub

    '
    ' Para los comentarios                                          (28/Sep/20)
    '

    ''' <summary>
    ''' Poner comentarios a las líneas seleccionadas,
    ''' si no hay seleccionada, se pone donde esté el cursor
    ''' </summary>
    ''' <param name="rtEditor"></param>
    Private Sub PonerComentarios(rtEditor As RichTextBox)
        ' Se usará según el lenguaje seleccionado

        Dim sSep As String = "'"
        If buttonLenguaje.Text = Compilar.LenguajeCSharp Then
            sSep = "//"
        End If

        If rtEditor.SelectionLength > 0 Then
            Dim selStart As Integer = rtEditor.SelectionStart
            Dim lineas() As String = rtEditor.SelectedText.Split(vbCr.ToCharArray)
            Dim sb As New System.Text.StringBuilder
            Dim j As Integer = lineas.Length - 1
            If String.IsNullOrEmpty(lineas(j)) Then j -= 1
            Dim k As Integer = 0

            For i As Integer = 0 To j
                ' Esto pone el comentario al principio de la línea
                'sb.AppendFormat("{0}{1}{2}", sSep, lineas(i), vbCr)
                ' Poner el comentario después de los espacios inciales  (05/Oct/20)
                Dim p = lineas(i).TrimStart().Length
                Dim esp = lineas(i).Length - p
                'Debug.Assert(esp <> lineas(i).Length)
                If esp = lineas(i).Length Then
                    sb.AppendFormat("{0}{1}", lineas(i), vbCr)
                Else
                    sb.AppendFormat("{0}{1}{2}{3}", lineas(i).Substring(0, esp), sSep, lineas(i).Substring(esp), vbCr)
                End If

                k += 1
            Next
            Dim sbT = sb.ToString
            If j = 1 Then
                sbT = sbT.TrimEnd(ChrW(13))
            End If

            rtEditor.SelectedText = sbT
            rtEditor.SelectionStart = selStart
            rtEditor.SelectionLength = sbT.Length
            ' si se debe colorear                                   (03/Oct/20)
            ColorearSeleccion()

            ' dejar la selección de lo comentado                    (03/Oct/20)
            rtEditor.SelectionStart = selStart
            rtEditor.SelectionLength = sbT.Length
        Else
            ' No hay texto seleccionado
            If rtEditor.Text.Length > 0 Then
                If rtEditor.SelectionStart > 0 Then
                    ' Pierde el valor de SelectionStart al asignar el texto
                    Dim selStart As Integer = rtEditor.SelectionStart
                    Dim lin = rtEditor.GetLineFromCharIndex(selStart)
                    selStart = rtEditor.GetFirstCharIndexFromLine(lin)
                    Dim len = rtEditor.Lines(lin).Length - rtEditor.Lines(lin).TrimStart().Length
                    rtEditor.SelectionStart = selStart + len
                    rtEditor.SelectedText = sSep
                    ' si se debe colorear                           (03/Oct/20)
                    rtEditor.SelectionStart = selStart
                    rtEditor.SelectionLength = rtEditor.Lines(lin).Length
                    ColorearSeleccion()

                    rtEditor.SelectionStart = selStart
                    rtEditor.SelectionLength = 0
                Else
                    rtEditor.Text = sSep & rtEditor.Text.Substring(0)
                End If
            End If
        End If

    End Sub

    ''' <summary>
    ''' Quitar comentarios a las líneas seleccionadas
    ''' </summary>
    ''' <param name="rtEditor"></param>
    Private Sub QuitarComentarios(rtEditor As RichTextBox)
        ' Se usará según el lenguaje seleccionado
        Dim sSep As String = "'"
        If buttonLenguaje.Text = Compilar.LenguajeCSharp Then
            sSep = "//"
        End If

        If rtEditor.SelectionLength > 0 Then
            Dim selStart As Integer = rtEditor.SelectionStart
            Dim lineas() As String = rtEditor.SelectedText.Split(vbCr.ToCharArray)
            Dim sb As New System.Text.StringBuilder
            Dim j As Integer = lineas.Length - 1
            If String.IsNullOrEmpty(lineas(j)) Then j -= 1
            Dim k As Integer = 0

            For i As Integer = 0 To j
                ' Si la línea empieza con un comentario
                If lineas(i).TrimStart.StartsWith(sSep) Then
                    If lineas(i) = sSep Then Continue For

                    Dim p As Integer = lineas(i).IndexOf(sSep)

                    If lineas(i) = sSep Then
                        sb.AppendFormat("{0}{1}", lineas(i), vbCr)
                    Else
                        sb.AppendFormat("{0}{1}{2}", lineas(i).Substring(0, p), lineas(i).Substring(p + sSep.Length), vbCr)
                    End If

                    k += 1
                Else
                    ' si no tiene comentario, dejarlo como está
                    sb.AppendFormat("{0}{1}", lineas(i), vbCr)
                    k += 1
                End If
            Next
            Dim sbT = sb.ToString
            If j = 1 Then
                sbT = sbT.TrimEnd(ChrW(13))
            End If
            rtEditor.SelectedText = sbT
            rtEditor.SelectionStart = selStart
            rtEditor.SelectionLength = sbT.Length
            ' si se debe colorear                                   (03/Oct/20)
            ColorearSeleccion()
            rtEditor.SelectionStart = selStart
            rtEditor.SelectionLength = sbT.Length

        Else
            ' Averiguar en que línea estamos y comprobar si empieza por un comentario
            Dim posA = PosicionActual()
            If rtEditor.Lines(posA.Linea - 1).TrimStart().StartsWith(sSep) Then
                Dim j = rtEditor.Lines(posA.Linea - 1).IndexOf(sSep)
                Dim linLength = rtEditor.Lines(posA.Linea - 1).Length
                Dim selStart = rtEditor.GetFirstCharIndexFromLine(posA.Linea - 1)

                rtEditor.SelectionStart = selStart + j
                rtEditor.SelectionLength = sSep.Length
                rtEditor.SelectedText = ""
                ' si se debe colorear                               (03/Oct/20)
                rtEditor.SelectionStart = selStart
                rtEditor.SelectionLength = linLength

                ColorearSeleccion()
                rtEditor.SelectionStart = selStart
            End If

        End If

    End Sub

    ''' <summary>
    ''' Colorear la selección actual del código
    ''' </summary>
    Private Sub ColorearSeleccion()
        If inicializando Then Return
        ' No colorear si no es vb o c#                              (08/Oct/20)
        If buttonLenguaje.Text = ExtensionTexto Then Return

        inicializando = True
        ' si se debe colorear                                       (03/Oct/20)
        If colorearAlCargar OrElse colorearAlEvaluar Then
            Compilar.ColoreaSeleccionRichTextBox(richTextBoxCodigo, buttonLenguaje.Text)
        End If
        inicializando = False
    End Sub


    '
    ' Para los marcadores / Bookmarks                               (28/Sep/20)
    '

    ''' <summary>
    ''' Colección con los marcadores del código que se está editando.
    ''' </summary>
    Private ColMarcadores As New List(Of Integer)

    ''' <summary>
    ''' El fichero usado con los marcadores.
    ''' </summary>
    Private MarcadorFichero As String

    ''' <summary>
    ''' Poner los marcadores, si hay... (30/Sep/20)
    ''' </summary>
    Private Sub PonerLosMarcadores()
        If ColMarcadores.Count = 0 Then Return

        inicializando = True

        ' Recordar la posición                                      (30/Sep/20)
        Dim selStart = richTextBoxCodigo.SelectionStart

        ColMarcadores.Sort()
        Dim colMarcadorTemp = ColMarcadores.ToList
        ColMarcadores.Clear()
        For i = 0 To colMarcadorTemp.Count - 1
            Dim pos = colMarcadorTemp(i)
            richTextBoxCodigo.SelectionStart = pos '- 1
            MarcadorPonerQuitar()
        Next
        ' Poner la posición en la que estaba antes
        richTextBoxCodigo.SelectionStart = selStart

        inicializando = False
    End Sub


    ''' <summary>
    ''' Poner o quitar el marcador.
    ''' Si está marcado se quita y si no lo está se pone.
    ''' Se guarda la posición del inicio de la línea en la que está el cursor (o la posición dentro del richTextBoxCodigo).
    ''' </summary>
    Private Sub MarcadorPonerQuitar()
        Dim posActual = PosicionActual()
        If ColMarcadores.Contains(posActual.Inicio) Then
            ' quitarlo
            ColMarcadores.Remove(posActual.Inicio)

            richTextBoxCodigo.SelectionStart = If(posActual.Inicio - 1 < 0, 0, posActual.Inicio - 1) ' posActual.Inicio - 1
            richTextBoxCodigo.SelectionLength = 0
            'richTextBoxCodigo.SelectionBullet = False
            Dim fcol = richTextBoxtLineas.GetFirstCharIndexFromLine(posActual.Linea - 1)
            richTextBoxtLineas.SelectionStart = fcol
            richTextBoxtLineas.SelectionLength = 5
            richTextBoxtLineas.SelectionBullet = False
            ' así es como se pone en AñadirNumerosDeLinea
            richTextBoxtLineas.SelectedText = $" {(posActual.Linea).ToString("0").PadLeft(4)}"
        Else
            ColMarcadores.Add(posActual.Inicio)
            ' Poner los marcadores en richTextBoxtLineas
            'richTextBoxtLineas.SelectionBullet = True
            Dim fcol = richTextBoxtLineas.GetFirstCharIndexFromLine(posActual.Linea - 1)
            richTextBoxtLineas.SelectionStart = fcol
            richTextBoxtLineas.SelectionLength = 5
            'richTextBoxtLineas.SelectionBullet = True

            ' Poner delante la imagen del marcador
            ' Usando la imagen bookmark_003_8x10.png
            richTextBoxtLineas.SelectedRtf = $"{picBookmark}{(posActual.Linea).ToString("0").PadLeft(4)}" & "}"
            'richTextBoxtLineas.SelectedText = $"*{(posActual.Linea).ToString("0").PadLeft(4)}"

            'richTextBoxCodigo.SelectionBullet = True
            richTextBoxCodigo.SelectionStart = If(posActual.Inicio - 1 < 0, 0, posActual.Inicio - 1)
            richTextBoxCodigo.SelectionLength = 0
        End If
        ColMarcadores.Sort()
    End Sub

    ''' <summary>
    ''' La imagen a usar cuando se muestra un marcador en richTextBoxtLineas.
    ''' </summary>
    Private picBookmark As String = "{\rtf1\ansi\deff0\nouicompat{\fonttbl{\f0\fnil Consolas;}}
{\colortbl ;\red0\green128\blue128;}
\uc1 
\pard\cf1\f0\fs22\lang9{\pict{\*\picprop}\wmetafile8\picw212\pich265\picwgoal120\pichgoal150 
0100090000037e00000000005500000000000400000003010800050000000b0200000000050000
000c020a000800030000001e000400000007010400040000000701040055000000410b2000cc00
0a000800000000000a0008000000000028000000080000000a0000000100040000000000000000
000000000000000000000000000000000000000000ffffff00424242003f3f3f00404040003737
3700505050003c3c3c003a3a3a0076767600d1d1d1005c5c5c00c8c8c800fbfbfb000000000000
000000bcd11dcb789aa98724566542222332222222222222222222222222222222222222222222
22222222040000002701ffff030000000000
}\f1\lang3082 "


    ''' <summary>
    ''' Ir al marcador anterior.
    ''' Si está antes del primero, ir al último
    ''' </summary>
    Private Sub MarcadorAnterior()
        If ColMarcadores.Count = 0 Then Return
        Dim posActual = PosicionActual()
        Dim res = ColMarcadores.Where(Function(x) x < posActual.Inicio)
        If res.Count > 0 Then
            Dim pos = res.Last
            richTextBoxCodigo.SelectionStart = pos
        Else
            ' si no hay más marcadores, ir al último
            Dim fcol As Integer
            If richTextBoxCodigo.Lines.Count < 2 Then
                fcol = richTextBoxCodigo.TextLength
            Else
                fcol = richTextBoxCodigo.GetFirstCharIndexFromLine(richTextBoxCodigo.Lines.Count - 2)
            End If
            richTextBoxCodigo.SelectionStart = fcol
            richTextBoxCodigo.SelectionLength = 0

            MarcadorAnterior()
        End If
    End Sub

    ''' <summary>
    ''' Ir al marcador siguiente.
    ''' Si está después del último, ir al anterior.
    ''' </summary>
    Private Sub MarcadorSiguiente()
        If ColMarcadores.Count = 0 Then Return

        Dim posActual = PosicionActual()
        Dim res = ColMarcadores.Where(Function(x) x > posActual.Inicio)
        If res.Count > 0 Then
            Dim pos = res.First
            richTextBoxCodigo.SelectionStart = pos
        Else
            ' si no hay más marcadores, ir al anterior
            richTextBoxCodigo.SelectionStart = 0
            MarcadorSiguiente()
        End If
    End Sub

    ''' <summary>
    ''' Quitar todos los marcadores.
    ''' </summary>
    Private Sub MarcadorQuitarTodos()
        If ColMarcadores.Count = 0 Then Return

        ColMarcadores.Clear()
        AñadirNumerosDeLinea()

        MarcadorFichero = ""
    End Sub

    ''' <summary>
    ''' Averigua la línea, columna (y primer caracter de la línea) de la posición actual en richTextBoxCodigo.
    ''' </summary>
    ''' <returns>Una tupla con la Fila, Columna y la posición del primer caracter de la línea (Inicio)</returns>
    Private Function PosicionActual() As (Linea As Integer, Columna As Integer, Inicio As Integer)
        Dim pos As Integer = richTextBoxCodigo.SelectionStart + 1
        Dim lin As Integer = richTextBoxCodigo.GetLineFromCharIndex(pos) + 1
        Dim col As Integer = pos - richTextBoxCodigo.GetFirstCharIndexOfCurrentLine()
        Dim fcol = richTextBoxCodigo.GetFirstCharIndexFromLine(lin - 1)
        If fcol = pos Then lin -= 1

        Return (lin, col, fcol)
    End Function

    ''' <summary>
    '''  Saber la línea y columna de la posición actual en richTextBoxCodigo.
    ''' </summary>
    Private Sub MostrarPosicion(e As KeyEventArgs)
        Dim pos = PosicionActual()
        Dim lin = pos.Linea
        Dim col = pos.Columna

        If e IsNot Nothing Then
            If e.KeyCode = Keys.Tab AndAlso e.Modifiers = Keys.Shift Then
                col = 1
            ElseIf e.KeyCode = Keys.Home Then
                col = 1
            End If
        End If

        labelPos.Text = $"Lín: {lin}  Car: {col}"
    End Sub

    '
    ' Para los recortes
    '

    ''' <summary>
    ''' Añadir la cadena indicada a la colección de recortes.
    ''' </summary>
    ''' <param name="str">La cadena a añadir.</param>
    Private Sub AñadirRecorte(str As String)
        ' Añadirlo al principio y dejar solo MaxRecortes
        Dim col As New List(Of String)

        ' Añadirlo al principio
        col.Add(str)
        col.AddRange(ColRecortes.Take(MaxRecortes - 1))

        ColRecortes.Clear()
        ' quitar los repetidos
        ColRecortes.AddRange(col.Distinct())

    End Sub

    ''' <summary>
    ''' Mostrar la ventana de recortes y pegar el seleccionado.
    ''' </summary>
    Private Sub MostrarRecortes()
        ' Mostrar una ventana emergente con los recortes            (30/Sep/20)
        ' al seleccionar uno de ellos se pegará como si usase Paste.
        If ColRecortes.Count = 0 Then Return

        Dim posRtb = richTextBoxCodigo.GetPositionFromCharIndex(richTextBoxCodigo.SelectionStart)
        posRtb.X += richTextBoxCodigo.Left
        posRtb.Y += richTextBoxCodigo.Top
        Dim posScr = richTextBoxCodigo.PointToScreen(posRtb)
        Dim fClip As New FormRecortes(posScr, ColRecortes, MaxRecortes)
        If fClip.ShowDialog() = DialogResult.OK Then
            Dim s = fClip.TextoSeleccionado

            Dim selStart = richTextBoxCodigo.SelectionStart

            ' pegarlo
            richTextBoxCodigo.SelectedText = s

            Dim lin = richTextBoxCodigo.GetLineFromCharIndex(selStart)
            Dim pos = richTextBoxCodigo.GetFirstCharIndexFromLine(lin)

            ' obligar a poner las líneas
            AñadirNumerosDeLinea()

            ' Seleccionar el texto después de pegar                 (04/Oct/20)
            richTextBoxCodigo.SelectionStart = pos
            richTextBoxCodigo.SelectionLength = s.Length + (selStart - pos)
            ' y colorearlo si procede
            ColorearSeleccion()
            richTextBoxCodigo.SelectionStart = selStart 'pos
        End If

    End Sub

    '
    ' El spliter
    '

    ''' <summary>
    ''' El úlimo ancho del splitContainer2 (28/Sep/20)
    ''' </summary>
    Private splitPanel2 As Integer

    Private Sub splitContainer1_Resize(sender As Object, e As EventArgs)
        If Me.WindowState = FormWindowState.Minimized Then Return

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
    ' Cambio de tamaño del panel de herramientas
    '

    Private Sub panelHerramientas_SizeChanged(sender As Object, e As EventArgs)
        'If inicializando Then Return

        Dim iant = inicializando

        inicializando = True
        Dim tAnt = splitContainer1.Top
        splitContainer1.Top = panelHerramientas.Top + panelHerramientas.Height + 10
        splitContainer1.Height += (tAnt - splitContainer1.Top)

        AjustarAltoPanelHerramientas()

        inicializando = iant
    End Sub

    '
    ' Clasificar
    '

    ''' <summary>
    ''' Clasificar el texto seleccionado
    ''' </summary>
    Private Sub ClasificarSeleccion()
        If richTextBoxCodigo.SelectionLength = 0 Then Return

        Dim selStart = richTextBoxCodigo.SelectionStart
        Dim lineas() As String = richTextBoxCodigo.SelectedText.Split(vbCr.ToCharArray)
        Dim sb As New System.Text.StringBuilder
        Dim j = lineas.Length - 1
        If String.IsNullOrWhiteSpace(lineas(j)) Then j -= 1
        Dim k = 0

        ' Asignar las opciones de clasificación
        ' ya estarán asignadas directamente en la clase
        'CompararString.IgnoreCase = clasif_caseSensitive
        'CompararString.UsarCompareOrdinal = clasif_compareOrdinal

        ' Clasificar el array usando el comparador de CompararString
        Array.Sort(lineas, 0, j + 1, New CompararString)

        For i As Integer = 0 To j
            sb.AppendFormat("{0}{1}", lineas(i), vbCr)
            k += 1
        Next

        ' Poner el nuevo texto
        richTextBoxCodigo.SelectedText = sb.ToString()
        richTextBoxCodigo.SelectionStart = selStart
        richTextBoxCodigo.SelectionLength = sb.ToString().Length
        ColorearSeleccion()
        richTextBoxCodigo.SelectionStart = selStart
        richTextBoxCodigo.SelectionLength = sb.ToString().Length

    End Sub

    '
    ' Cambiar las mayúsculas y minúsculas
    '

    Private Sub CambiarMayúsculasMinúsculas(queCase As CasingValues)
        ' convertir el texto seleccionado a mayúsculas
        If richTextBoxCodigo.SelectionLength = 0 Then Return

        Dim selStart = richTextBoxCodigo.SelectionStart
        Dim selLength = richTextBoxCodigo.SelectionLength

        If queCase = CasingValues.FirstToLower Then
            richTextBoxCodigo.SelectedText = richTextBoxCodigo.SelectedText.ToLowerFirst
        ElseIf queCase = CasingValues.Title OrElse queCase = CasingValues.FirstToUpper Then
            ' si está todo en mayúsculas preguntar              (02/Oct/20)
            ' si no, va a parecer que no hace bien la conversión a título ;-)
            If richTextBoxCodigo.SelectedText = richTextBoxCodigo.SelectedText.ToUpper Then
                If MessageBox.Show("¡El texto está todo en mayúsculas!" & vbCrLf &
                                       "¿Quieres cambiarlo a Título?",
                                       "Cambiar selección a Título",
                                       MessageBoxButtons.YesNo,
                                       MessageBoxIcon.Question) = DialogResult.Yes Then

                    richTextBoxCodigo.SelectedText = richTextBoxCodigo.SelectedText.ToLower().ToTitle
                Else
                    Return
                End If
            Else
                richTextBoxCodigo.SelectedText = richTextBoxCodigo.SelectedText.ToTitle
            End If

        ElseIf queCase = CasingValues.Lower Then
            richTextBoxCodigo.SelectedText = richTextBoxCodigo.SelectedText.ToLower
        ElseIf queCase = CasingValues.Upper Then
            richTextBoxCodigo.SelectedText = richTextBoxCodigo.SelectedText.ToUpper
        Else
            ' nada que hacer
        End If
        ' Solo si se debe colorear                              (03/Oct/20)
        richTextBoxCodigo.SelectionStart = selStart
        richTextBoxCodigo.SelectionLength = selLength
        ColorearSeleccion()

    End Sub

    '
    ' Quitar los espacios
    '

    ''' <summary>
    ''' Quitar los espacios de delante, detrás o ambos
    ''' </summary>
    ''' <param name="delante"></param>
    ''' <param name="detras"></param>
    Private Sub QuitarEspacios(Optional delante As Boolean = True,
                               Optional detras As Boolean = True,
                               Optional todos As Boolean = False)

        ' Quitar los espacios, delante, detrás y ambos              (02/Oct/20)
        If richTextBoxCodigo.SelectionLength = 0 Then Return

        Dim selStart As Integer = richTextBoxCodigo.SelectionStart
        Dim lineas() As String = richTextBoxCodigo.SelectedText.Split(vbCr.ToCharArray)
        Dim sb As New System.Text.StringBuilder
        Dim j As Integer = lineas.Length - 1
        'If vb.Len(lineas(j)) = 0 Then j -= 1
        Dim k As Integer = 0

        For i As Integer = 0 To j
            If lineas(i) <> "" Then
                ' Poner como primera comprobación               (02/Oct/20)
                ' para no tener que asignar los otros dos argumentos
                If todos Then
                    ' quitar todos los espacios
                    sb.AppendFormat("{0}{1}", lineas(i).QuitarTodosLosEspacios, vbCr)
                ElseIf delante AndAlso detras Then
                    sb.AppendFormat("{0}{1}", lineas(i).Trim(), vbCr)
                ElseIf delante Then
                    sb.AppendFormat("{0}{1}", lineas(i).TrimStart(), vbCr)
                ElseIf detras Then
                    sb.AppendFormat("{0}{1}", lineas(i).TrimEnd(), vbCr)
                End If
                k += 1
            Else
                sb.AppendFormat("{0}{1}", lineas(i), vbCr)
            End If
        Next
        richTextBoxCodigo.SelectedText = sb.ToString().TrimEnd(ChrW(13))

        richTextBoxCodigo.SelectionStart = selStart
        richTextBoxCodigo.SelectionLength = sb.ToString().Length
        ColorearSeleccion()

        ' Que se mantenga seleccionado
        richTextBoxCodigo.SelectionStart = selStart
        richTextBoxCodigo.SelectionLength = sb.ToString().Length
    End Sub

    Private Sub PonerTextoAlFinal()
        If String.IsNullOrEmpty(txtPonerTexto.Text) Then Return
        ' Si no hay texto, salir
        If richTextBoxCodigo.Text.Length = 0 Then Return

        ' Poner el texto al final de cada línea seleccionada o de la actual
        Dim sPoner As String = txtPonerTexto.Text

        If richTextBoxCodigo.SelectionLength > 0 Then
            Dim selStart As Integer = richTextBoxCodigo.SelectionStart
            Dim lineas() As String = richTextBoxCodigo.SelectedText.TrimEnd(ChrW(13)).Split(vbCr.ToCharArray)
            Dim sb As New System.Text.StringBuilder
            Dim j As Integer = lineas.Length - 1

            For i As Integer = 0 To j
                'If i = j AndAlso String.IsNullOrEmpty(lineas(i)) Then Exit For
                sb.AppendFormat("{0}{1}{2}", lineas(i), sPoner, vbCr)
            Next
            Dim sbT = sb.ToString()

            richTextBoxCodigo.SelectedText = sbT
            richTextBoxCodigo.SelectionStart = selStart
            richTextBoxCodigo.SelectionLength = sbT.Length
            ' si se debe colorear                                   (03/Oct/20)
            ColorearSeleccion()

            ' dejar la selección de lo comentado                    (03/Oct/20)
            richTextBoxCodigo.SelectionStart = selStart
            richTextBoxCodigo.SelectionLength = sbT.Length
        Else
            ' Pierde el valor de SelectionStart al asignar el texto
            Dim selStart As Integer = richTextBoxCodigo.SelectionStart
            Dim lin = richTextBoxCodigo.GetLineFromCharIndex(selStart)
            selStart = richTextBoxCodigo.GetFirstCharIndexFromLine(lin)
            Dim len = richTextBoxCodigo.Lines(lin).Length
            richTextBoxCodigo.SelectionStart = selStart + len
            richTextBoxCodigo.SelectedText = sPoner
            ' si se debe colorear                           (03/Oct/20)
            richTextBoxCodigo.SelectionStart = selStart
            richTextBoxCodigo.SelectionLength = richTextBoxCodigo.Lines(lin).Length
            ColorearSeleccion()

            richTextBoxCodigo.SelectionStart = selStart
            richTextBoxCodigo.SelectionLength = 0
        End If

    End Sub

    Private Sub QuitarTextoDelFinal()
        If String.IsNullOrEmpty(txtPonerTexto.Text) Then Return
        ' Si no hay texto, salir
        If richTextBoxCodigo.Text.Length = 0 Then Return

        ' Quitar el texto al final de cada línea seleccionada o de la actual
        Dim sQuitar As String = txtPonerTexto.Text

        If richTextBoxCodigo.SelectionLength > 0 Then
            Dim selStart As Integer = richTextBoxCodigo.SelectionStart
            Dim lineas() As String = richTextBoxCodigo.SelectedText.TrimEnd(ChrW(13)).Split(vbCr.ToCharArray)
            Dim sb As New System.Text.StringBuilder
            Dim j As Integer = lineas.Length - 1

            For i As Integer = 0 To j
                If lineas(i).TrimEnd.EndsWith(sQuitar, StringComparison.OrdinalIgnoreCase) Then
                    Dim k = lineas(i).LastIndexOf(sQuitar)
                    If k = 0 Then
                        sb.AppendFormat("{0}{1}", "", vbCr)
                    Else
                        sb.AppendFormat("{0}{1}", lineas(i).Substring(0, k), vbCr)
                    End If
                End If
            Next
            Dim sbT = sb.ToString
            'If j = 1 Then
            '    sbT = sbT.TrimEnd(ChrW(13))
            'End If

            richTextBoxCodigo.SelectedText = sbT
            richTextBoxCodigo.SelectionStart = selStart
            richTextBoxCodigo.SelectionLength = sbT.Length
            ' si se debe colorear                                   (03/Oct/20)
            ColorearSeleccion()

            ' dejar la selección de lo comentado                    (03/Oct/20)
            richTextBoxCodigo.SelectionStart = selStart
            richTextBoxCodigo.SelectionLength = sbT.Length
        Else

            ' Pierde el valor de SelectionStart al asignar el texto
            Dim selStart As Integer = richTextBoxCodigo.SelectionStart
            Dim lin = richTextBoxCodigo.GetLineFromCharIndex(selStart)
            selStart = richTextBoxCodigo.GetFirstCharIndexFromLine(lin)
            If richTextBoxCodigo.Lines(lin).TrimEnd.EndsWith(sQuitar, StringComparison.OrdinalIgnoreCase) = False Then Return

            Dim k = richTextBoxCodigo.Lines(lin).LastIndexOf(sQuitar)
            richTextBoxCodigo.SelectionStart = selStart + k
            richTextBoxCodigo.SelectionLength = sQuitar.Length
            richTextBoxCodigo.SelectedText = ""
            ' si se debe colorear                           (03/Oct/20)
            richTextBoxCodigo.SelectionStart = selStart
            richTextBoxCodigo.SelectionLength = richTextBoxCodigo.Lines(lin).Length
            ColorearSeleccion()

            richTextBoxCodigo.SelectionStart = selStart
            richTextBoxCodigo.SelectionLength = 0
        End If

    End Sub

    '
    ' Métodos de evento con Handles
    '

#Region " Métodos de evento con Handles "

    Private Sub lstSyntax_MouseClick(sender As Object, e As MouseEventArgs) Handles lstSyntax.MouseClick
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

    Private Sub lstSyntax_MouseMove(sender As Object, e As MouseEventArgs) Handles lstSyntax.MouseMove
        If lstSyntax.Items.Count = 0 Then Return

        Dim point = New Point(e.X, e.Y)
        Dim i = lstSyntax.IndexFromPoint(point)
        If i = -1 Then Return

        If i < lstSyntax.Items.Count Then
            lstSyntax.SelectedIndex = i
        End If

        If lstSyntax.Items(i).ToString.Length > 30 Then
            toolTip1.SetToolTip(lstSyntax, lstSyntax.Items(i).ToString)
        Else
            toolTip1.SetToolTip(lstSyntax, "")
        End If

        'Try
        '    toolTip1.SetToolTip(lstSyntax, lstSyntax.Items(i).ToString)
        'Catch ex As Exception
        'End Try

    End Sub

    Private Sub lstSyntax_DrawItem(sender As Object, e As DrawItemEventArgs) Handles lstSyntax.DrawItem
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

    'Private Sub timer1_Tick(sender As Object, e As EventArgs) Handles timer1.Tick
    '    timer1.Stop()

    '    'lstSyntax.Refresh()
    '    lstSyntax.DrawMode = DrawMode.Normal
    'End Sub

#End Region

End Class
