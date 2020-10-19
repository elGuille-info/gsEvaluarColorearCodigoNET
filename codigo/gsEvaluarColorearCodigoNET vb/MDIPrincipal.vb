'------------------------------------------------------------------------------
' MDIPrincipal - formulario para tener varios formularios abiertos  (08/Oct/20)
'
'
' (c) Guillermo (elGuille) Som, 2020
'------------------------------------------------------------------------------
Option Strict On
Option Infer On

Imports System
Imports System.Data
Imports System.Collections.Generic
Imports System.Text
Imports System.Linq
Imports Microsoft.VisualBasic
Imports vb = Microsoft.VisualBasic

Imports System.Windows.Forms
Imports System.IO

Public Class MDIPrincipal

    ''' <summary>
    ''' El último texto del portapapeles asignado.
    ''' </summary>
    Private sClipUltimo As String = ""

    Friend Shared TituloMDI As String = Application.ProductName ' "Editor de código"

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        CurrentMDI = Me
        Form1Activo = New Form1
        Form1Activo.MdiParent = Me

    End Sub

    Private Sub MDIPrincipal_Load(sender As Object, e As EventArgs) Handles Me.Load
        Width = CInt(Screen.PrimaryScreen.Bounds.Width * 0.75)
        Height = CInt(Screen.PrimaryScreen.Bounds.Height * 0.85)

        Me.CenterToScreen()

        ' Activar el temporizador para copia del portapapeles       (11/Oct/20)
        ' cada 15 segundos
        timerClipBoard.Interval = 15 * 1000
        timerClipBoard.Enabled = True

        timerInicio.Interval = 100
        timerInicio.Start()

    End Sub

    Private Sub timerInicio_Tick(sender As Object, e As EventArgs) Handles timerInicio.Tick

        timerInicio.Stop()

        AsignaMetodosDeEventos()

        Inicializar()

        ' Abrir una nueva ventana si no se han abierto
        ' los anteriores
        If Me.MdiChildren.Count = 0 Then
            NuevaVentana()
        Else
            ' Esta es la creada en el constructor de MDIPrincipal   (17/Oct/20)
            Me.MdiChildren(0).Close()
        End If

        Me.Text = $"{MDIPrincipal.TituloMDI} [{Form1Activo.Text}]"

        Dim sCopyR = $"Copyright Guillermo Som (elGuille), 2020{If(Date.Now.Year > 2020, $"-{Date.Now.Year}", "")}"
        sCopyR = $"{Application.ProductName} v{Application.ProductVersion}, {sCopyR}"

        'Dim t = 0
        ' Asignar el texto de la información a todas las ventanas
        For Each frm As Form1 In MdiChildren
            frm.labelInfo.Text = sCopyR
            't += 1
        Next
        'Debug.WriteLine(t)
    End Sub


    Private Sub MDIPrincipal_FormClosing(sender As Object,
                                         e As FormClosingEventArgs) Handles Me.FormClosing
        timerClipBoard.Stop()

        ' Antes estaba en el Form1                                  (18/Oct/20)
        GuardarConfig()
    End Sub

    ''' <summary>
    ''' Asignar los métodos de evento manualmente en vez de con Handles
    ''' ya que así todos los métodos estarán en la lista desplegable al seleccionar Form1
    ''' en la lista desplegable de las clases
    ''' </summary>
    ''' <remarks>Aparte de que hay que hacerlo manualmente hasta que Visual Studio 2019 no los conecte</remarks>
    Public Sub AsignaMetodosDeEventos()

        ' Ficheros: los menús de archivo y botones de la barra ficheros
        AddHandler menuFileAcercaDe.Click, Sub() AcercaDe()
        AddHandler menuFileSalir.Click, Sub() CurrentMDI.Close()
        AddHandler menuFileSeleccionarAbrir.Click, Sub() Abrir()
        AddHandler buttonSeleccionar.Click, Sub() Abrir()
        AddHandler buttonAbrir.Click, Sub() Abrir()
        AddHandler menuFileGuardar.Click, Sub() Guardar()
        AddHandler buttonGuardar.Click, Sub() Guardar()
        AddHandler menuFileGuardarComo.Click, Sub() GuardarComo()
        AddHandler buttonGuardarComo.Click, Sub() GuardarComo()
        AddHandler menuFileNuevo.Click, Sub() Nuevo()
        AddHandler buttonNuevo.Click, Sub() Nuevo()

        AddHandler menuFileRecargar.Click, Sub() Recargar()
        AddHandler buttonRecargar.Click, Sub() Recargar()

        AddHandler menuRecargarFichero.Click, Sub() Recargar()
        AddHandler menuCopiarPath.Click, Sub()
                                             Try
                                                 Clipboard.SetText(nombreUltimoFichero)
                                             Catch ex As Exception
                                             End Try
                                         End Sub

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
        AddHandler buttonEdicionRecortes.Click, AddressOf menuEditPegarRecorte_Click

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

                                                  menuOcultarEvaluar.Enabled = Form1Activo.lstSyntax.Items.Count > 0
                                                  If Form1Activo.splitContainer1.Panel2.Visible = True Then
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
                                                     Form1Activo.splitContainer1.Panel2.Visible = False
                                                     menuOcultarEvaluar.Text = "Mostrar panel de evaluación/error"
                                                 Else
                                                     menuOcultarEvaluar.Text = "Ocultar panel de evaluación/error"
                                                     Form1Activo.splitContainer1.Panel2.Visible = True
                                                 End If
                                                 Form1Activo.splitContainer1_Resize(Nothing, Nothing)
                                             End Sub

        ' Colorear en HTML
        AddHandler menuColorearHTML.Click, Sub() ColorearHTML()
        AddHandler buttonColorearHTML.Click, Sub() ColorearHTML()
        AddHandler chkMostrarLineasHTML.Click, Sub() MostrarLineasHTML = chkMostrarLineasHTML.Checked

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
                                           Dim opFrm As New FormOpciones()
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
        AddHandler menuEditBuscar.Click, Sub() BuscarReemplazar(esBuscar:=True)
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
        AddHandler comboBoxReemplazar.KeyUp, Sub(sender As Object, e As KeyEventArgs)
                                                 If e.KeyCode = Keys.Enter Then
                                                     e.Handled = True
                                                     BuscarSiguiente(esReemplazar:=True)
                                                 End If
                                             End Sub

        ' Para palabras completas y case sensitive                  (17/Sep/20)
        AddHandler buttonMatchCase.Click, Sub() buscarMatchCase = buttonMatchCase.Checked
        AddHandler buttonWholeWord.Click, Sub() buscarWholeWord = buttonWholeWord.Checked

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
                                        MostrarPanelBuscar(menuMostrar_Buscar.Checked, esReemplazar:=False)
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

        ' opciones para clasificar (solo en el menú Editor)
        AddHandler menuMayúsculas.Click, Sub() CambiarMayúsculasMinúsculas(CasingValues.Upper)
        AddHandler menuMinúsculas.Click, Sub() CambiarMayúsculasMinúsculas(CasingValues.Lower)
        AddHandler menuTitulo.Click, Sub() CambiarMayúsculasMinúsculas(CasingValues.Title)
        AddHandler menuPrimeraMinúsculas.Click, Sub() CambiarMayúsculasMinúsculas(CasingValues.FirstToLower)

        ' quitar los espacios de delante, detrás o ambos            (02/Oct/20)
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

    Private Sub menuEditPegarRecorte_Click(sender As Object,
                                           e As EventArgs) Handles menuEditPegarRecorte.Click
        ' Mostrar ventana con los recortes guardados
        MostrarRecortes()
    End Sub



    ''' <summary>
    ''' Inicializar los menús, etc.
    ''' Este método se llama desde el evento Form1_Load.
    ''' </summary>
    Public Sub Inicializar()
        '
        ' Todo esto estaba en el evento Form1_Load                  (28/Sep/20)
        '

        Dim extension = ".appconfig.txt"
        DirDocumentos = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        FicheroConfiguracion = Path.Combine(DirDocumentos, Application.ProductName & extension)

        LeerConfig()

        ' Mostrar los 15 primeros en el menú Recientes
        AsignarRecientes()

        inicializando = False

        ' No mostrar el panel al iniciar
        ' Ponerlo después de inicializando = false                  (30/Sep/20)
        ' para que se ajuste el tamaño de FlowPanel
        MostrarPanelBuscar(False, esReemplazar:=False)

        esCtrlF = True

        ' Quitar esto para que se muestre los datos correctos,      (17/Oct/20)
        ' ya que se cargan los ficheros en LeerConfig
        'labelTamaño.Text = $"{0:#,##0} car. ({0:#,##0} palab.)"
        'labelInfo.Text = "Selecciona el fichero a abrir."
        'labelModificado.Text = " "
        'labelPos.Text = "Lín: 0  Car: 0"

        ' Guardar los esquemas de colores
        ' en: MyDocuments/ColorDictionary.csv
        Compilar.WriteColorDictionaryToFile()
        ' Si se cambia y se quiere leer
        'Compilar.UpdateColorDictionaryFromFile()

        Form1Activo.splitContainer1.Panel2.Visible = False
        Form1Activo.splitContainer1_Resize(Nothing, Nothing)

        ' Esto se usará para cargar                                 (16/Oct/20)
        ' si se indica en la línea de comandos,
        ' ya que se habrán cargado en LeerConfig
        Dim tmpCargarUltimo = False ' cargarUltimo
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
            ' Si hay más de un formulario abierto                   (16/Oct/20)
            ' abrirlo en una nueva ventana
            If MdiChildren.Count > 1 Then
                Nuevo()
                Form1Activo.nombreFichero = nombreUltimoFichero
            End If
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
    ''' Para saber si se deben o no habilitar los botones de los toolStrip
    ''' Antes estaba en MenuEdiDropDownOpening
    ''' </summary>
    Public Sub HabilitarBotones()
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
    ' Los menús de Ventana
    '

    Private m_ChildFormNumber As Integer = 0

    ''' <summary>
    ''' Crear una nueva ventana.
    ''' </summary>
    ''' <remarks>16/Oct/2020</remarks>
    Friend Sub NuevaVentana()
        ' Create a new instance of the child form.
        Dim ChildForm As New Form1()
        ' Make it a child of this MDI form before showing it.
        ChildForm.MdiParent = Me

        m_ChildFormNumber += 1
        ChildForm.Text = "Nueva Ventana #" & m_ChildFormNumber
        Me.Text = $"{MDIPrincipal.TituloMDI} [{ChildForm.Text}]"

        Form1Activo = ChildForm

        ChildForm.Width = Me.ClientSize.Width - 4
        ChildForm.Height = ClientSize.Height - 4 - menuStrip1.Height - panelHerramientas.Height
        ChildForm.Top = 0
        ChildForm.Left = 0

        ChildForm.Show()

        labelInfo.Text = ChildForm.Text


        ' Algunos de estos valores deben ser a nivel de Form1

        inicializando = True

        richTextBoxCodigo.Text = ""
        richTextBoxLineas.Text = ""
        richTextBoxCodigo.Modified = False

        labelInfo.Text = ""
        labelPos.Text = "Lín: 1  Car: 1"
        labelTamaño.Text = $"{0:#,##0} car. ({0:#,##0} palab.)"
        nombreUltimoFichero = ""

        ' A nive de Form1
        Form1Activo.lstSyntax.Items.Clear()
        Form1Activo.richTextBoxLineas.Clear()
        Form1Activo.nombreFichero = ""
        Form1Activo.TextoModificado = False
        Form1Activo.codigoNuevo = ""
        Form1Activo.codigoAnterior = ""

        inicializando = False
    End Sub

    Private Sub ShowNewForm(sender As Object,
                            e As EventArgs) Handles menuVentanaNueva.Click
        NuevaVentana()
    End Sub

    Private Sub CascadeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles menuVentanaCascade.Click
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub TileVerticalToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles menuVentanaTileVertical.Click
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub TileHorizontalToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles menuVentanaTileHorizontal.Click
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub ArrangeIconsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles menuVentanaArrangeAll.Click
        Me.LayoutMdi(MdiLayout.ArrangeIcons)
    End Sub

    Private Sub CloseAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles menuVentanaCloseAll.Click
        ' Close all child forms of the parent.
        For Each ChildForm As Form1 In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub

    Private Sub MaximizarTodasMenuVentana_Click(sender As Object, e As EventArgs) Handles menuVentanaMaximizarTodas.Click
        Me.Text = $"{MDIPrincipal.TituloMDI}"
        For Each ChildForm As Form1 In Me.MdiChildren
            ChildForm.WindowState = FormWindowState.Maximized
        Next
    End Sub

    Private Sub RestaurarMenuVentana_Click(sender As Object, e As EventArgs) Handles menuVentanaRestaurarTodas.Click
        Me.Text = $"{MDIPrincipal.TituloMDI}"
        For Each ChildForm As Form1 In Me.MdiChildren
            ChildForm.WindowState = FormWindowState.Normal
        Next
    End Sub

    Private Sub timerClipBoard_Tick(sender As Object,
                                    e As EventArgs) Handles timerClipBoard.Tick
        ' Se ejecutará cada 30 segundos (1000 ms = 1 s, 30000 ms = 30 s)
        Try
            If Clipboard.ContainsText Then
                Dim sActual = Clipboard.GetText()
                If Not String.IsNullOrEmpty(sActual) Then
                    If sActual <> sClipUltimo Then
                        ' Añadirlo a la colección 
                        AñadirRecorte(sActual)
                        sClipUltimo = sActual
                    End If
                End If
            End If
        Catch ex As Exception
            Debug.Assert(ex.Message = "")
        End Try
    End Sub

    Private Sub menuVentana_DropDownOpening(sender As Object,
                                            e As EventArgs) Handles menuVentana.DropDownOpening
        ' Forzar a mostrar el nombre en el menú de ventanas         (16/Oct/20)
        ' De una respuesta en stackoverflow
        ' https://stackoverflow.com/a/1348453/14338047

        If Me.ActiveMdiChild IsNot Nothing Then
            Dim ventanaActiva As Form = Me.ActiveMdiChild
            Me.ActivateMdiChild(Nothing)
            Me.ActivateMdiChild(ventanaActiva)
        End If
    End Sub

    Private Sub GuardarTodo_Click(sender As Object,
                                  e As EventArgs) Handles buttonGuardarTodo.Click, menuFileGuardarTodo.Click
        MostrarProcesando("Guardar todo", "Guardando todos los ficheros abiertos...", Me.MdiChildren.Count * 2)
        Dim t = 0
        For Each frm As Form1 In Me.MdiChildren
            nombreUltimoFichero = frm.nombreFichero
            labelInfo.Text = $"Guardando {nombreUltimoFichero}..."
            OnProgreso(labelInfo.Text)
            Application.DoEvents()
            Guardar()
        Next
        labelInfo.Text = $"Guardado {If(t = 1, "un fichero", $"{t} ficheros")}."
        OnProgreso(labelInfo.Text)
    End Sub

    Private Sub menuFileRecientes_DropDownOpening(sender As Object,
                                                  e As EventArgs) Handles menuFileRecientes.DropDownOpening
        Dim yaSeleccionado = False
        For i = 0 To menuFileRecientes.DropDownItems.Count - 1
            TryCast(menuFileRecientes.DropDownItems(i), ToolStripMenuItem).Checked = False
            If menuFileRecientes.DropDownItems(i).Text.IndexOf(nombreUltimoFichero) > 3 Then
                If yaSeleccionado Then Continue For
                menuFileRecientes.DropDownItems(i).Select()
                TryCast(menuFileRecientes.DropDownItems(i), ToolStripMenuItem).Checked = True
                yaSeleccionado = True
            End If
        Next
    End Sub

    '
    ' Para poner en gris el texto de Buscar y Reemplazar            (17/Oct/20)
    ' cuando están vacios
    ' Truco copiado del proyecto CSharpToVB
    '

    Private Sub comboBoxBuscar_Enter(sender As Object, e As EventArgs) Handles comboBoxBuscar.Enter
        If comboBoxBuscar.Text <> BuscarVacio Then
            'comboBoxBuscar.Text = ""
            comboBoxBuscar.ForeColor = SystemColors.ControlText
        End If
    End Sub

    Private Sub comboBoxBuscar_Leave(sender As Object, e As EventArgs) Handles comboBoxBuscar.Leave
        ' Esto hace lo mismo que comprobar si está en blanco... :-P
        'If Not Me.comboBoxBuscar.Text.Any Then
        If String.IsNullOrEmpty(Me.comboBoxBuscar.Text) Then '
            Me.comboBoxBuscar.ForeColor = SystemColors.GrayText
            Me.comboBoxBuscar.Text = BuscarVacio
        End If
    End Sub

    Private Sub comboBoxReemplazar_Enter(sender As Object, e As EventArgs) Handles comboBoxReemplazar.Enter
        If comboBoxReemplazar.Text <> ReemplazarVacio Then
            'comboBoxReemplazar.Text = ""
            comboBoxReemplazar.ForeColor = SystemColors.ControlText
        End If
    End Sub

    Private Sub comboBoxReemplazar_Leave(sender As Object, e As EventArgs) Handles comboBoxReemplazar.Leave
        'If Not Me.comboBoxReemplazar.Text.Any Then
        If String.IsNullOrEmpty(Me.comboBoxReemplazar.Text) Then
            Me.comboBoxReemplazar.ForeColor = SystemColors.GrayText
            Me.comboBoxReemplazar.Text = ReemplazarVacio
        End If
    End Sub

    Private buscarBoxAnt As String = BuscarVacio
    Private reemplazarBoxAnt As String = ReemplazarVacio

    Private Sub comboBoxBuscar_TextChanged(sender As Object,
                                           e As EventArgs) Handles comboBoxBuscar.TextChanged
        If inicializando Then Return
        If comboBoxBuscar.Text = "" OrElse comboBoxBuscar.Text = BuscarVacio Then
            comboBoxBuscar.ForeColor = SystemColors.GrayText
            inicializando = True
            comboBoxBuscar.Text = BuscarVacio
            inicializando = False
        Else
            If buscarBoxAnt = BuscarVacio Then
                inicializando = True
                comboBoxBuscar.Text = comboBoxBuscar.Text.QuitarPredeterminado(BuscarVacio)
                inicializando = False

                comboBoxBuscar.SelectionStart = comboBoxBuscar.Text.Length

            End If
            comboBoxBuscar.ForeColor = SystemColors.ControlText
        End If
        buscarBoxAnt = comboBoxBuscar.Text
    End Sub

    Private Sub comboBoxReemplazar_TextChanged(sender As Object,
                                               e As EventArgs) Handles comboBoxReemplazar.TextChanged
        If inicializando Then Return
        If comboBoxReemplazar.Text = "" OrElse comboBoxReemplazar.Text = ReemplazarVacio Then
            comboBoxReemplazar.ForeColor = SystemColors.GrayText
            inicializando = True
            comboBoxReemplazar.Text = ReemplazarVacio
            inicializando = False
        Else
            If reemplazarBoxAnt = ReemplazarVacio Then
                inicializando = True
                comboBoxReemplazar.Text = comboBoxReemplazar.Text.QuitarPredeterminado(ReemplazarVacio)
                inicializando = False

                comboBoxReemplazar.SelectionStart = comboBoxReemplazar.Text.Length

            End If
            comboBoxReemplazar.ForeColor = SystemColors.ControlText
        End If
        reemplazarBoxAnt = comboBoxReemplazar.Text
    End Sub

    'Public Sub comboBoxFileName_KeyDown(sender As Object,
    '                                    e As KeyEventArgs)
    '    If e.KeyCode = Keys.Enter Then
    '        e.Handled = True
    '        ' Abrir en nueva ventana                                (19/Oct/20)
    '        Nuevo()
    '        Form1Activo.nombreFichero = comboBoxFileName.Text
    '        Abrir(comboBoxFileName.Text)
    '    End If
    'End Sub

    'Public Sub comboBoxFileName_Validating(sender As Object,
    '                                       e As ComponentModel.CancelEventArgs)
    '    If inicializando Then Return

    '    Dim fic = comboBoxFileName.Text
    '    If String.IsNullOrWhiteSpace(fic) Then Return

    '    AñadirAlComboBoxFileName(fic)
    'End Sub

    Private Sub menuEditor_DropDownOpening(sender As Object,
                                           e As EventArgs) Handles menuEditor.DropDownOpening
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

    Private Sub menuEditorCambiarMayúsculas_DropDownOpening(sender As Object,
                                                            e As EventArgs) Handles menuEditorCambiarMayúsculas.DropDownOpening
        Dim b = richTextBoxCodigo.SelectionLength > 0
        menuMayúsculas.Enabled = b
        menuMinúsculas.Enabled = b
        menuTitulo.Enabled = b
        menuPrimeraMinúsculas.Enabled = b
    End Sub

    Private Sub menuEditorQuitarEspacios_DropDownOpening(sender As Object,
                                                         e As EventArgs) Handles menuEditorQuitarEspacios.DropDownOpening
        Dim b = richTextBoxCodigo.SelectionLength > 0
        menuQuitarEspaciosTrim.Enabled = b
        menuQuitarEspaciosTrimStart.Enabled = b
        menuQuitarEspaciosTrimEnd.Enabled = b
        menuQuitarEspaciosTodos.Enabled = b
    End Sub
End Class
