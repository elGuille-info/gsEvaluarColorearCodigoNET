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

'Imports gsEvaluarColorearCodigoNET

Public Class Form1

    ''' <summary>
    ''' Si se está inicializando.
    ''' Usado para que no se provoquen eventos en cadena
    ''' </summary>
    Private inicializando As Boolean = True

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
    ''' <summary>
    ''' El nombre del lenguaje VB (como está definido en Compilation.Language)
    ''' </summary>
    Private Const LenguajeVisualBasic As String = "Visual Basic"
    ''' <summary>
    ''' El nombre del lenguaje C# (como está definido en Compilation.Language)
    ''' </summary>
    Private Const LenguajeCSharp As String = "C#"

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
    Private Const MaxFicsMenu As Integer = 15
    '''' <summary>
    '''' El número máximo de ficheros a guardar en la configuración
    '''' </summary>
    'Private Const MaxFicsConfig As Integer = 50

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
        If TextoModificado OrElse String.IsNullOrEmpty(nombreUltimoFichero) Then
            GuardarComo()
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
            If e.KeyValue = Keys.L Then
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
        Else
            CtrlK = 0
            CtrlC = 0
            CtrlU = 0
            ShiftAltL = 0
            ShiftAltS = 0

            ' Otras pulsaciones
            ' seguramente captura los TAB del editor
            richTextBoxCodigo_KeyUp(sender, e)
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
                                             If richTextBoxCodigo.CanUndo Then richTextBoxCodigo.Redo()
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
                                                   menuEditDropDownOpenning()
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
        AddHandler menuEdit.DropDownOpening, Sub() menuEditDropDownOpenning()
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

        ' Colorear en HTML
        AddHandler menuColorearHTML.Click, Sub() ColorearHTML()
        AddHandler buttonColorearHTML.Click, Sub() ColorearHTML()
        AddHandler chkMostrarLineasHTML.Click, Sub() mostrarLineasHTML = chkMostrarLineasHTML.Checked

        ' Herramientas; Opciones, Colorear, lenguajes
        AddHandler panelHerramientas.SizeChanged, AddressOf panelHerramientas_SizeChanged

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
        AddHandler rtbCodigoContext.Opening, Sub() menuEditDropDownOpenning()

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
            panelHerramientas.Height = 63
        Else
            panelHerramientas.Height = 28
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

        labelTamaño.Text = ""
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

        If tmpCargarUltimo Then
            Abrir(nombreUltimoFichero)

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

        buttonLenguaje.DropDownItems(0).Text = LenguajeVisualBasic
        buttonLenguaje.DropDownItems(1).Text = LenguajeCSharp


        PonerLosMarcadores()

        ' posicionarlo al principio
        richTextBoxCodigo.SelectionStart = 0
        richTextBoxCodigo.SelectionLength = 0

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

        ' Elegir cuál de las 2 formas usar                          (26/Sep/20)
        ' es mejor la de comboBoxFileName
        Dim j As Integer

        ' Guardar la lista de últimos ficheros
        ' solo los maxFicsConfig (50) últimos
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
            If j = MaxFicsMenu Then Exit For
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

        colorearAlEvaluar = cfg.GetValue("Herramientas", "ColorearEvaluar", False)
        compilarAlEvaluar = cfg.GetValue("Herramientas", "CompilarEvaluar", True)

        buttonLenguaje.Text = cfg.GetValue("Herramientas", "Lenguaje", LenguajeVisualBasic)
        If buttonLenguaje.Text = LenguajeVisualBasic Then
            buttonLenguaje.Image = buttonLenguaje.DropDownItems(0).Image
        Else
            buttonLenguaje.Image = buttonLenguaje.DropDownItems(1).Image
        End If
        colorearAlCargar = cfg.GetValue("Herramientas", "Colorear", False)
        chkMostrarLineasHTML.Checked = cfg.GetValue("Herramientas", "Mostrar Líneas HTML", True)

        Dim cuantos = cfg.GetValue("Ficheros", "Count", 0)
        comboBoxFileName.Items.Clear()

        For i = 0 To MaxFicsMenu - 1
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

        'comboFuenteNombre.Text = fuenteNombre
        'comboFuenteTamaño.Text = fuenteTamaño
        richTextBoxCodigo.Font = New Font(fuenteNombre, CSng(fuenteTamaño))
        richTextBoxtLineas.Font = New Font(fuenteNombre, CSng(fuenteTamaño))
        richTextBoxSyntax.Font = New Font(fuenteNombre, CSng(fuenteTamaño))

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
            ' el bucle lo hace hasya un máximo de MaxFicsMenu (15)
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
                Path.Combine(dirDocumentos, $"prueba.{If(buttonLenguaje.Text = LenguajeVisualBasic, ".vb", ".cs")}"),
                sFN)
            oFD.FileName = sFN
            oFD.Filter = "Código de Visual Basic y CSharp (*.vb;*.cs)|*.vb;*.cs|Todos los ficheros (*.*)|*.*"
            oFD.InitialDirectory = Path.GetDirectoryName(sFN)
            oFD.Multiselect = False
            oFD.RestoreDirectory = True
            If oFD.ShowDialog() = DialogResult.OK Then
                'comboBoxFileName.Text = oFD.FileName
                'nombreUltimoFichero = oFD.FileName

                ' Abrir el fichero
                Abrir(oFD.FileName)
            End If
        End Using
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
            richTextBoxSyntax.Text = ""
        End If

        Dim sCodigo = ""
        Using sr As New StreamReader(fic, detectEncodingFromByteOrderMarks:=True, encoding:=Encoding.UTF8)
            sCodigo = sr.ReadToEnd()
        End Using

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

        ' Asignar el lenguaje en los combos
        Dim sLenguaje As String
        If Path.GetExtension(fic).ToLower = ".cs" Then
            sLenguaje = LenguajeCSharp
        ElseIf Path.GetExtension(fic).ToLower = ".vb" Then
            sLenguaje = LenguajeVisualBasic
        Else
            If sCodigo.Contains("end sub") Then
                sLenguaje = LenguajeVisualBasic
            Else 'If sCodigo.ToLower().Contains(");") Then
                sLenguaje = LenguajeCSharp
            End If
        End If
        If sLenguaje = LenguajeVisualBasic Then
            buttonLenguaje.Image = buttonLenguaje.DropDownItems(0).Image
        Else
            buttonLenguaje.Image = buttonLenguaje.DropDownItems(1).Image
        End If
        buttonLenguaje.Text = sLenguaje

        ' Mostrar información del fichero
        labelInfo.Text = $"{Path.GetFileName(fic)} ({sDirFic})"
        labelTamaño.Text = $"{richTextBoxCodigo.Text.Length:#,##0} car."

        ' Si hay que colorear el fichero cargado
        If colorearAlCargar Then
            ColorearCodigo()
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

        Using sw As New StreamWriter(fic, append:=False, encoding:=Encoding.UTF8)
            sw.WriteLine(sCodigo)
        End Using
        codigoAnterior = sCodigo

        labelInfo.Text = $"{Path.GetFileName(fic)} ({Path.GetDirectoryName(fic)})"
        labelTamaño.Text = $"{richTextBoxCodigo.Text.Length:#,##0} car."

        TextoModificado = False

        'richTextBoxCodigo.Modified = False
        'labelModificado.Text = " "

        nombreUltimoFichero = fic
        comboBoxFileName.Text = fic
        AñadirAlComboBoxFileName(fic)

    End Sub

    ''' <summary>
    ''' Borra la ventana de código y deja en blanco el <see cref="nombreUltimoFichero"/>.
    ''' Si se ha modificado el que había, pregunta si lo quieres guardar
    ''' </summary>
    Private Sub Nuevo()
        If TextoModificado Then GuardarComo()

        richTextBoxCodigo.Text = ""
        richTextBoxtLineas.Text = ""
        richTextBoxSyntax.Text = ""
        nombreUltimoFichero = ""
        TextoModificado = False
        labelInfo.Text = ""
        labelPos.Text = ""
        labelTamaño.Text = ""
        codigoAnterior = ""
    End Sub

    ''' <summary>
    ''' Muestra el cuadro de diálogo de Guardar como.
    ''' </summary>
    Private Sub GuardarComo()

        Dim fichero = comboBoxFileName.Text

        If nombreUltimoFichero <> comboBoxFileName.Text Then
            Using sFD As New SaveFileDialog
                sFD.Title = "Seleccionar fichero para guardar el código"
                sFD.FileName = fichero
                sFD.InitialDirectory = dirDocumentos
                sFD.RestoreDirectory = True
                sFD.Filter = "Código de Visual Basic y CSharp (*.vb;*.cs)|*.vb;*.cs|Todos los ficheros (*.*)|*.*"
                If sFD.ShowDialog = DialogResult.Cancel Then
                    Return
                End If
                fichero = sFD.FileName
                nombreUltimoFichero = sFD.FileName
                ' Guardarlo
                Guardar(fichero)
            End Using
        End If
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
    ''' Abre nuevamente el último fichero
    ''' desechando los datos realizados
    ''' </summary>
    Private Sub Recargar()
        If nombreUltimoFichero <> "" Then _
            Abrir(nombreUltimoFichero)
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
            ' BUG: obligar a poner las líneas                       (18/Sep/20)
            AñadirNumerosDeLinea()
        End If
    End Sub

    ''' <summary>
    ''' Cuando se abre el menú de edición o se cambia la selección del código
    ''' asignar si están o no habilitados el menú en sí, el contextual y las barras de herramientas.
    ''' </summary>
    Private Sub menuEditDropDownOpenning()
        If inicializando Then Return

        inicializando = True

        ' para saber si hay texto en el control
        Dim b As Boolean = richTextBoxCodigo.TextLength > 0

        menuEditDeshacer.Enabled = richTextBoxCodigo.CanUndo
        menuEditRehacer.Enabled = richTextBoxCodigo.CanRedo
        menuEditCopiar.Enabled = richTextBoxCodigo.SelectionLength > 0
        menuEditCortar.Enabled = menuEditCopiar.Enabled
        menuEditSeleccionarTodo.Enabled = b
        'menuEditPegar.Enabled = richTextBoxCodigo.CanPaste(DataFormats.GetFormat("Text"))

        If Clipboard.GetDataObject.GetDataPresent(DataFormats.Text) Then
            menuEditPegar.Enabled = richTextBoxCodigo.CanPaste(DataFormats.GetFormat(DataFormats.Text))
        ElseIf Clipboard.GetDataObject.GetDataPresent(DataFormats.StringFormat) Then
            ' StringFormat                                          (30/Oct/04)
            menuEditPegar.Enabled = richTextBoxCodigo.CanPaste(DataFormats.GetFormat(DataFormats.StringFormat))
        ElseIf Clipboard.GetDataObject.GetDataPresent(DataFormats.Html) Then
            menuEditPegar.Enabled = richTextBoxCodigo.CanPaste(DataFormats.GetFormat(DataFormats.Html))
        ElseIf Clipboard.GetDataObject.GetDataPresent(DataFormats.OemText) Then
            menuEditPegar.Enabled = richTextBoxCodigo.CanPaste(DataFormats.GetFormat(DataFormats.OemText))
        ElseIf Clipboard.GetDataObject.GetDataPresent(DataFormats.UnicodeText) Then
            menuEditPegar.Enabled = richTextBoxCodigo.CanPaste(DataFormats.GetFormat(DataFormats.UnicodeText))
        ElseIf Clipboard.GetDataObject.GetDataPresent(DataFormats.Rtf) Then
            menuEditPegar.Enabled = richTextBoxCodigo.CanPaste(DataFormats.GetFormat(DataFormats.Rtf))
        End If

        buttonCopiar.Enabled = menuEditCopiar.Enabled
        buttonCortar.Enabled = menuEditCopiar.Enabled
        buttonPegar.Enabled = menuEditPegar.Enabled
        buttonBuscarSiguiente.Enabled = b ' (richTextBoxCodigo.TextLength > 0)
        menuEditBuscar.Enabled = b
        menuEditReemplazar.Enabled = b
        buttonReemplazarSiguiente.Enabled = b
        buttonReemplazarTodo.Enabled = b

        buttonEditorPonerIndentacion.Enabled = b
        buttonEditorQuitarIndentacion.Enabled = b
        buttonEditorPonerComentarios.Enabled = b
        buttonEditorQuitarComentarios.Enabled = b
        '
        menuEditorPonerIndentacion.Enabled = b
        menuEditorQuitarIndentacion.Enabled = b
        menuEditorPonerComentarios.Enabled = b
        menuEditorQuitarComentarios.Enabled = b

        buttonCompilar.Enabled = b
        buttonEjecutar.Enabled = b
        buttonEvaluar.Enabled = b
        menuEvaluar.Enabled = b
        buttonColorear.Enabled = b
        buttonNoColorear.Enabled = b
        buttonColorearHTML.Enabled = b
        menuColorearHTML.Enabled = b
        menuColorear.Enabled = b
        menuNoColorear.Enabled = b
        menuCompilar.Enabled = b
        menuEjecutar.Enabled = b

        ' Actualizar también los del menú contextual
        For j = 0 To rtbCodigoContext.Items.Count - 1
            For i = 0 To menuEdit.DropDownItems.Count - 1
                If rtbCodigoContext.Items(j).Name = menuEdit.DropDownItems(i).Name Then
                    rtbCodigoContext.Items(j).Enabled = menuEdit.DropDownItems(i).Enabled
                    Exit For
                End If
            Next
        Next

        b = ColMarcadores.Count > 0
        buttonEditorMarcadorAnterior.Enabled = b
        buttonEditorMarcadorSiguiente.Enabled = b
        buttonEditorMarcadorQuitarTodos.Enabled = b
        '
        menuEditorMarcadorAnterior.Enabled = b
        menuEditorMarcadorSiguiente.Enabled = b
        menuEditorMarcadorQuitarTodos.Enabled = b

        b = ColRecortes.Count > 0
        buttonEdicionRecortes.Enabled = b

        b = richTextBoxCodigo.SelectionLength > 0
        menuEditorClasificarSeleccion.Enabled = b
        buttonEditorClasificarSeleccion.Enabled = b

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
        '' Ajustar el alto del FlowPanel                             (01/Oct/20)
        'If mostrar Then
        '    panelHerramientas.Height = 63
        'Else
        '    ' solo hacerlo menos alto si el toolstripEditor no está arriba
        '    If toolStripEditor.Visible AndAlso toolStripEditor.Top = 3 Then
        '        panelHerramientas.Height = 28 '32
        '    End If
        'End If
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
                If richTextBoxCodigo.SelectedText <> "" Then
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
        labelInfo.Text = $"Coloreando el código de {buttonLenguaje.Text}..."
        Me.Refresh()

        Dim modif = richTextBoxCodigo.Modified
        inicializando = True
        'richTextBoxCodigo.Visible = False

        Dim colCodigo = Compilar.ColoreaRichTextBox(richTextBoxCodigo, buttonLenguaje.Text)

        'richTextBoxCodigo.Visible = True

        inicializando = False
        richTextBoxCodigo.SelectionStart = 0
        richTextBoxCodigo.SelectionLength = 0
        richTextBoxCodigo.Modified = modif

        labelInfo.Text = $"Código coloreado para {buttonLenguaje.Text}."
        Me.Refresh()
    End Sub

    Private Sub ColorearTimer_Tick(sender As Object, e As EventArgs) Handles ColorearTimer.Tick

        ColorearTimer.Stop()


    End Sub


    ''' <summary>
    ''' Colorear el código en formato HTML y mostrarlo en ventana aparte
    ''' </summary>
    Private Sub ColorearHTML()
        If String.IsNullOrWhiteSpace(richTextBoxCodigo.Text) Then Return

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
        ' guardar el código antes de compilar y ejecutar
        labelInfo.Text = "Compilando el código..."
        Me.Refresh()

        ' Comprobar si es el mismo fichero que se está editando     (22/Sep/20)
        ' Si no es así, se pregunta si se quiere guardar
        If nombreUltimoFichero <> comboBoxFileName.Text Then
            GuardarComo()
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
            splitContainer1.Panel2.Visible = False
            splitContainer1_Resize(Nothing, Nothing)
            Return
        End If

        ' Mostrar los errores
        Dim sb As New StringBuilder
        splitContainer1.Panel2.Visible = True
        splitContainer1_Resize(Nothing, Nothing)

        labelInfo.Text = $"Error al compilar ({res.Result.Diagnostics.Count} errores)."

        For Each diagnostic In res.Result.Diagnostics
            Dim lin = diagnostic.Location.GetLineSpan().StartLinePosition.Line + 1
            Dim pos = diagnostic.Location.GetLineSpan().StartLinePosition.Character + 1
            sb.AppendFormat("{0}: {1} (en línea {2}, posición {3})",
                                diagnostic.Id,
                                diagnostic.GetMessage(),
                                lin, pos)
            sb.AppendLine()
        Next

        Me.richTextBoxSyntax.Text = sb.ToString

    End Sub

    ''' <summary>
    ''' Compilar el código sin ejecutar
    ''' </summary>
    Private Sub Build()
        If TextoModificado Then
            GuardarComo()
        End If

        Dim filepath = nombreUltimoFichero
        labelInfo.Text = $"Compilando para {buttonLenguaje.Text}..."
        Me.Refresh()

        ' Si no tiene el path, añadirlo                             (27/Sep/20)
        If String.IsNullOrEmpty(Path.GetDirectoryName(filepath)) Then
            filepath = Path.Combine(dirDocumentos, filepath)
        End If

        Dim res = Compilar.CompileRun(filepath, run:=False)

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
        If String.IsNullOrWhiteSpace(richTextBoxCodigo.Text) Then Return

        Dim res As EmitResult

        If compilarAlEvaluar Then
            splitContainer1.Panel2.Visible = False
            splitContainer1_Resize(Nothing, Nothing)

            labelInfo.Text = $"Compilando para {buttonLenguaje.Text}..."
            Me.Refresh()

            res = Compilar.ComprobarCodigo(richTextBoxCodigo.Text, buttonLenguaje.Text)
            If res.Success = False Then
                MostrarErrores(res)
                Return
            End If
            labelInfo.Text = "Compilado sin error."
        End If
        Dim colCodigo As Dictionary(Of String, List(Of String))

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

        MostrarResultadoEvaluar(colCodigo)
    End Sub

    ''' <summary>
    ''' Muestra los errores de la compilación.
    ''' </summary>
    ''' <param name="res">Un objeto del tipo <see cref="EmitResult"/> con los errores y advertencias.</param>
    Private Sub MostrarErrores(res As EmitResult)
        Dim errors = 0
        Dim warnings = 0
        Dim sbErr As New StringBuilder
        For Each d In res.Diagnostics
            sbErr.AppendLine(d.ToString)
            If d.Severity = DiagnosticSeverity.Error Then
                errors += 1
            ElseIf d.Severity = DiagnosticSeverity.Warning Then
                warnings += 1
            End If
        Next
        labelInfo.Text = $"Compilado con {errors} errores y {warnings} advertencias."
        richTextBoxSyntax.Text = sbErr.ToString

        splitContainer1.Panel2.Visible = True
        splitContainer1_Resize(Nothing, Nothing)
    End Sub

    ''' <summary>
    ''' Mostrar el resultado de evaluar el código.
    ''' Las clases, métodos, palabras clave, etc. definidos en el código.
    ''' </summary>
    ''' <param name="colCodigo">Colección del tipo <see cref="Dictionary(Of String, List(Of String))"/>
    ''' con la lista con los valores sacados de <see cref="ClassifiedSpan"/> </param>
    Private Sub MostrarResultadoEvaluar(colCodigo As Dictionary(Of String, List(Of String)))
        Dim codTiposCount As (Clases As Integer, Metodos As Integer, Keywords As Integer)
        codTiposCount.Clases = If(colCodigo.Keys.Contains("class name"), colCodigo("class name").Count, 0) + If(colCodigo.Keys.Contains("module name"), colCodigo("module name").Count, 0)
        codTiposCount.Metodos = If(colCodigo.Keys.Contains("method name"), colCodigo("method name").Count, 0)
        codTiposCount.Keywords = If(colCodigo.Keys.Contains("keyword"), colCodigo("keyword").Count, 0) + If(colCodigo.Keys.Contains("keyword - control"), colCodigo("keyword - control").Count, 0)

        labelInfo.Text = $"Código con {codTiposCount.Clases} clase{If(codTiposCount.Clases = 1, "", "s")}, {codTiposCount.Metodos} método{If(codTiposCount.Metodos = 1, "", "s")} y {codTiposCount.Keywords} palabra{If(codTiposCount.Keywords = 1, "", "s")} clave."

        Dim sb As New StringBuilder

        ' Clasificar las claves                                     (25/Sep/20)
        Dim claves = From kv In colCodigo Order By kv.Key Ascending Select kv

        For Each kv In claves
            Dim s1 = kv.Key

            ' No mostrar el contenido de estos símbolos
            If s1 = "comment" OrElse s1.StartsWith("string") OrElse
                s1 = "punctuation" OrElse s1.StartsWith("operator") OrElse
                s1 = "number" OrElse s1.StartsWith("xml") Then
                'sb.AppendLine("   <Omitidos>")
            Else
                sb.AppendLine(s1)
                kv.Value.Sort()
                For Each s2 In kv.Value
                    sb.AppendLine($"   {s2}")
                Next
                sb.AppendLine()
            End If
        Next

        richTextBoxSyntax.Text = sb.ToString

        splitContainer1.Panel2.Visible = True
        splitContainer1_Resize(Nothing, Nothing)
    End Sub


    '
    ' Los métodos relacionados con el richTextBoxCodigo
    '

    Private Sub richTextBoxCodigo_KeyUp(sender As Object, e As KeyEventArgs)
        ' Se ve que KeyDown falla                                   (28/Sep/20)
        '
        ' si se pulsa la tecla TAB
        ' insertar 4 espacios en vez de un tabulador
        '
        ' Hay que detectar antes el Shist, Control y Alt
        ' ya que se producen antes que el resto de teclas
        If e.Shift = True AndAlso e.Alt = False AndAlso e.Control = False Then
            If e.KeyCode = Keys.Tab Then
                ' Atrás
                e.Handled = True
                QuitarIndentacion(richTextBoxCodigo)
            End If
        ElseIf e.Alt = False AndAlso e.Control = False Then
            If e.KeyCode = Keys.Tab Then
                ' Adelante
                e.Handled = True
                PonerIndentacion(richTextBoxCodigo)
            ElseIf e.KeyCode = Keys.Enter Then
                '
                ' Comprobar esto (es para añadir la indentación)
                '

                '' ln es el número de línea actual (con base 1)
                '' Si la línea actual (que es la anterior al intro)
                '' no está vacía.
                '' Si ln es menor que 1, salir                   (16/Dic/05) 0.40825
                '' seguramente el intro ha llegado por otro lado...
                'Dim rtEditor = richTextBoxCodigo

                'Dim ln As Integer = rtEditor.GetLineFromCharIndex(rtEditor.SelectionStart)
                'Dim col As Integer = rtEditor.SelectionStart - rtEditor.GetFirstCharIndexFromLine(ln)

                'If ln < 1 Then Return
                'If rtEditor.Lines.Length < 1 Then Return
                'If rtEditor.Lines(ln - 1) <> "" Then
                '    ' Si al quitarle los espacios es una cadena vacía,
                '    ' es que solo hay espacios.
                '    If rtEditor.Lines(ln - 1).TrimStart() = "" Then
                '        col = rtEditor.Lines(ln - 1).Length
                '    Else
                '        ' Averiguar la posición del primer carácter,
                '        ' aunque puede que haya TABs
                '        col = rtEditor.Lines(ln - 1).IndexOf(rtEditor.Lines(ln - 1).TrimStart().Substring(0, 1))
                '    End If
                '    e.Handled = True
                '    SendKeys.SendWait(New String(" "c, col))
                'End If
            End If
        End If
        MostrarPosicion(e)
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
    ''' Quitar indentación
    ''' </summary>
    ''' <param name="rtEditor"></param>
    Private Sub QuitarIndentacion(rtEditor As RichTextBox)
        If Not String.IsNullOrEmpty(rtEditor.SelectedText) Then
            Dim selStart As Integer = rtEditor.SelectionStart
            Dim sIndent As String = New String(" "c, EspaciosIndentar)
            Dim lineas() As String = rtEditor.SelectedText.Split(vbCrLf.ToCharArray)
            Dim sb As New System.Text.StringBuilder
            Dim j As Integer = lineas.Length - 1
            If String.IsNullOrEmpty(lineas(j)) Then j -= 1
            Dim k As Integer = 0
            ' El k += 1 solo es necesario si se usa vbCrLf
            '
            For i As Integer = 0 To j
                ' si empieza con 4 caracteres, quitarlos
                If lineas(i) <> "" Then
                    If lineas(i).StartsWith(sIndent) Then
                        ' Si solo hay una línea seleccionada,   (26/Nov/05)
                        ' es que no hay que añadir retornos de carro
                        If j = 0 OrElse i = j Then
                            sb.AppendFormat("{0}", lineas(i).Substring(EspaciosIndentar))
                        Else
                            sb.AppendFormat("{0}{1}", lineas(i).Substring(EspaciosIndentar), vbCr)
                            'k += 1
                        End If
                    ElseIf lineas(i).StartsWith(vbTab) Then
                        ' Considerar los tab como indentación   (22/Nov/05)
                        If i = j Then
                            sb.AppendFormat("{0}", lineas(i).Substring(1))
                        Else
                            sb.AppendFormat("{0}{1}", lineas(i).Substring(1), vbCr)
                        End If
                        'k += 1
                    Else
                        If i = j Then
                            sb.AppendFormat("{0}", lineas(i))
                        Else
                            sb.AppendFormat("{0}{1}", lineas(i), vbCr)
                        End If
                        'k += 1
                    End If
                Else
                    If i = j Then
                        sb.AppendFormat(" {0}", vbCr)
                    Else
                        sb.AppendFormat(" {0}", vbCr)
                    End If
                    'k += 1
                End If
            Next
        Else
            Dim selStart As Integer = rtEditor.SelectionStart
            Dim sIndent As String = New String(" "c, EspaciosIndentar)
            If selStart >= EspaciosIndentar Then
                rtEditor.SelectionLength = 0
                rtEditor.SelectionStart = selStart - EspaciosIndentar
                rtEditor.SelectionLength = EspaciosIndentar
                rtEditor.SelectedText = ""

                rtEditor.SelectionLength = 0
            Else
                rtEditor.SelectionLength = EspaciosIndentar - rtEditor.SelectionStart
                rtEditor.SelectedText = ""

                rtEditor.SelectionLength = 0
            End If
        End If
    End Sub

    ''' <summary>
    ''' Poner indentación
    ''' </summary>
    ''' <param name="rtEditor"></param>
    Private Sub PonerIndentacion(rtEditor As RichTextBox)
        ' Si se selecciona el "retorno" de la última línea, se eliminará
        ' y se juntará con la siguiente.
        If Not String.IsNullOrEmpty(rtEditor.SelectedText) Then
            Dim selStart As Integer = rtEditor.SelectionStart
            Dim sIndent As String '= New String(" "c, indentar)
            ' No tomar el texto rtf, sino el normal, que colorearemos
            Dim lineas() As String = rtEditor.SelectedText.Split(vbCrLf.ToCharArray)
            sIndent = New String(" "c, EspaciosIndentar)
            Dim sb As New System.Text.StringBuilder
            Dim j As Integer = lineas.Length - 1
            Dim k As Integer

            For i = 0 To j
                If i = j Then
                    sb.AppendFormat("{0}{1}", sIndent, lineas(i))
                Else
                    sb.AppendFormat("{0}{1}{2}", sIndent, lineas(i), vbCr)
                End If
                k += EspaciosIndentar
            Next
            ' Poner la selección y sustituir el texto
            rtEditor.SelectionStart = selStart
            rtEditor.SelectionLength = sb.ToString().Length - k
            rtEditor.SelectedText = sb.ToString
            ' para que se mantenga seleccionado                 (28/Sep/20)
            rtEditor.SelectionStart = selStart
            rtEditor.SelectionLength = sb.ToString().Length - EspaciosIndentar
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

    '
    ' Para los comentarios                                          (28/Sep/20)
    '

    ''' <summary>
    ''' Quitar comentarios a las líneas seleccionadas
    ''' </summary>
    ''' <param name="rtEditor"></param>
    Private Sub QuitarComentarios(rtEditor As RichTextBox)
        ' Se usará según el lenguaje seleccionado
        Dim sSep As String = "'"

        If buttonLenguaje.Text = LenguajeCSharp Then
            sSep = "//"
        End If

        If Not String.IsNullOrEmpty(rtEditor.SelectedText) Then
            Dim selStart As Integer = rtEditor.SelectionStart
            Dim lineas() As String = rtEditor.SelectedText.Split(vbCrLf.ToCharArray)
            Dim sb As New System.Text.StringBuilder
            Dim j As Integer = lineas.Length - 1
            If String.IsNullOrEmpty(lineas(j)) Then j -= 1
            Dim k As Integer = 0

            For i As Integer = 0 To j
                ' Si la línea empieza con un comentario
                If lineas(i).TrimStart.StartsWith(sSep) Then
                    If lineas(i) = sSep Then Continue For


                    Dim p As Integer = lineas(i).IndexOf(sSep)
                    ' Esto no se puede dar
                    'If p = -1 Then Continue For

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
            ' Dejar la selección que había                      (22/Nov/05)
            rtEditor.SelectedText = sb.ToString()
            rtEditor.SelectionStart = selStart
            rtEditor.SelectionLength = sb.ToString().Length - k
        Else
            ' Averiguar en que línea estamos y comprobar si empieza por un comentario
            Dim posA = PosicionActual()
            Dim selStart As Integer = rtEditor.SelectionStart
            If rtEditor.Lines(posA.Linea - 1).TrimStart().StartsWith(sSep) Then
                Dim j = rtEditor.Text.IndexOf(sSep, selStart - 1)
                rtEditor.Text = rtEditor.Text.Substring(0, j) & rtEditor.Text.Substring(j + sSep.Length)

                'Dim j = rtEditor.Lines(posA.Linea - 1).IndexOf(sSep)
                ' Esto no lo cambia, lo cambia, pero no lo muestra ???  (30/Sep/20)
                ' Lines es de solo lectura
                'rtEditor.Lines(posA.Linea - 1) = rtEditor.Lines(posA.Linea - 1).Remove(j, sSep.Length)
                'rtEditor.Lines(posA.Linea - 1) = New String((rtEditor.Lines(posA.Linea - 1).Substring(0, j) & rtEditor.Lines(posA.Linea - 1).Substring(j + sSep.Length)).ToCharArray)
            End If
            rtEditor.SelectionStart = selStart
        End If

    End Sub

    ''' <summary>
    ''' Poner comentarios a las líneas seleccionadas,
    ''' si no hay seleccionada, se pone donde esté el cursor
    ''' </summary>
    ''' <param name="rtEditor"></param>
    Private Sub PonerComentarios(rtEditor As RichTextBox)
        ' Se usará según el lenguaje seleccionado

        Dim sSep As String = "'"

        If buttonLenguaje.Text = LenguajeCSharp Then
            sSep = "//"
        End If

        If Not String.IsNullOrEmpty(rtEditor.SelectedText) Then
            Dim selStart As Integer = rtEditor.SelectionStart
            Dim lineas() As String = rtEditor.SelectedText.Split(vbCrLf.ToCharArray)
            Dim sb As New System.Text.StringBuilder
            Dim j As Integer = lineas.Length - 1
            If String.IsNullOrEmpty(lineas(j)) Then j -= 1
            Dim k As Integer = 0

            For i As Integer = 0 To j
                sb.AppendFormat("{0}{1}{2}", sSep, lineas(i), vbCr)
                k += 1
            Next
            ' Dejar la selección que había                      (22/Nov/05)
            rtEditor.SelectedText = sb.ToString()
            rtEditor.SelectionStart = selStart
            rtEditor.SelectionLength = sb.ToString().Length - k
        Else
            ' No hay texto seleccionado
            If rtEditor.Text.Length > 0 Then
                If rtEditor.SelectionStart > 0 Then
                    ' Pierde el valor de SelectionStart al asignar el texto
                    Dim selStart As Integer = rtEditor.SelectionStart - 1
                    rtEditor.SelectedText = sSep
                    rtEditor.SelectionStart = selStart + 1
                    rtEditor.SelectionLength = 0
                Else
                    rtEditor.Text = sSep & rtEditor.Text.Substring(0)
                End If
            End If
        End If

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
        Dim pos = richTextBoxCodigo.PointToScreen(posRtb)
        Dim fClip As New FormRecortes(pos, ColRecortes, MaxRecortes)
        If fClip.ShowDialog() = DialogResult.OK Then
            Dim s = fClip.TextoSeleccionado
            ' pegarlo
            richTextBoxCodigo.SelectedText = s
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
                splitPanel2 = CInt(splitContainer1.Width * 0.3)
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

    Friend clasif_caseSensitive As Boolean
    Friend clasif_compareOrdinal As Boolean

    ''' <summary>
    ''' Clasificar el texto seleccionado
    ''' </summary>
    Private Sub ClasificarSeleccion()
        If richTextBoxCodigo.SelectedText <> "" Then
            Dim selStart = richTextBoxCodigo.SelectionStart
            Dim lineas() As String = richTextBoxCodigo.SelectedText.Split(vbCr.ToCharArray)
            Dim sb As New System.Text.StringBuilder
            Dim j = lineas.Length - 1
            If String.IsNullOrWhiteSpace(lineas(j)) Then j -= 1
            Dim k = 0

            ' Asignar las opciones de clasificación
            CompararString.IgnoreCase = clasif_caseSensitive
            CompararString.UsarCompareOrdinal = clasif_compareOrdinal

            ' Clasificar el array usando el comparador de CompararString
            Array.Sort(lineas, 0, j + 1, New CompararString)

            For i As Integer = 0 To j
                sb.AppendFormat("{0}{1}", lineas(i), vbCr)
                k += 1
            Next

            ' Poner el nuevo texto
            richTextBoxCodigo.SelectedText = sb.ToString()
            richTextBoxCodigo.SelectionStart = selStart
            richTextBoxCodigo.SelectionLength = sb.ToString().Length - k
        End If
    End Sub

    '
    ' Cambiar las mayúsculas y minúsculas
    '

    'Private Sub mnuHerMayMinMayusculas_Click(sender As System.Object, e As System.EventArgs) Handles mnuHerMayMinMayusculas.Click
    '    ' convertir el texto seleccionado a mayúsculas
    '    If rtEditor.SelectedText <> "" Then
    '        rtEditor.SelectedText = rtEditor.SelectedText.ToUpper()
    '    End If
    '    tssLabelInfo.Text = StatusStrip1.Text
    'End Sub
    'Private Sub mnuHerMayMinMayusculas_Select(sender As Object, e As System.EventArgs) Handles mnuHerMayMinMayusculas.DropDownOpening
    '    tssLabelInfo.Text = "Convierte el texto seleccionado en mayúsculas"
    'End Sub
    'Private Sub mnuHerMayMinMinusculas_Click(sender As System.Object, e As System.EventArgs) Handles mnuHerMayMinMinusculas.Click
    '    ' convertir el texto seleccionado a minúsculas
    '    If rtEditor.SelectedText <> "" Then
    '        rtEditor.SelectedText = rtEditor.SelectedText.ToLower()
    '    End If
    '    tssLabelInfo.Text = StatusStrip1.Text
    'End Sub
    'Private Sub mnuHerMayMinMinusculas_Select(sender As Object, e As System.EventArgs) Handles mnuHerMayMinMinusculas.DropDownOpening
    '    tssLabelInfo.Text = "Convierte el texto seleccionado en minúsculas"
    'End Sub

    'Private Sub mnuHerMayMinTitulo_Click(sender As System.Object, e As System.EventArgs) Handles mnuHerMayMinTitulo.Click
    '    If rtEditor.SelectedText <> "" Then
    '        rtEditor.SelectedText = vb.StrConv(rtEditor.SelectedText, vb.VbStrConv.ProperCase)
    '    End If
    '    tssLabelInfo.Text = StatusStrip1.Text
    'End Sub

End Class
