﻿'------------------------------------------------------------------------------
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

        ' Asignar el texto de la información a todas las ventanas
        For Each frm As Form1 In MdiChildren
            frm.labelInfo.Text = sCopyR
        Next
    End Sub


    Private Sub MDIPrincipal_FormClosing(sender As Object,
                                         e As FormClosingEventArgs) Handles Me.FormClosing
        ' Este evento se produce después de los Form1_Closing       (30/Oct/20)

        timerClipBoard.Stop()

        ' Para guardar la posición y tamaño actual de la ventana    (26/Oct/20)
        MDIPrincipal_Resize()

        ' Antes estaba en el Form1                                  (18/Oct/20)
        GuardarConfig()
    End Sub

    Private Sub MDIPrincipal_Resize() Handles Me.Resize
        If inicializando Then Return

        If Me.WindowState = FormWindowState.Normal Then
            tamForm = (Me.Left, Me.Top, Me.Height, Me.Width)
        Else
            tamForm = (Me.RestoreBounds.Left, Me.RestoreBounds.Top,
                       Me.RestoreBounds.Height, Me.RestoreBounds.Width)
        End If

    End Sub

    Private Sub MDIPrincipal_Move() Handles Me.Move
        If inicializando Then Return

        If Me.WindowState = FormWindowState.Normal Then
            tamForm = (Me.Left, Me.Top, Me.Height, Me.Width)
        End If

    End Sub


    ''' <summary>
    ''' Asignar los métodos de evento manualmente en vez de con Handles
    ''' ya que así todos los métodos estarán en la lista desplegable al seleccionar Form1
    ''' en la lista desplegable de las clases
    ''' </summary>
    ''' <remarks>Aparte de que hay que hacerlo manualmente hasta que Visual Studio 2019 no los conecte</remarks>
    Private Sub AsignaMetodosDeEventos()

        ' Ficheros: los menús de archivo y botones de la barra ficheros
        AddHandler menuFileAcercaDe.Click, Sub() AcercaDe()
        AddHandler menuFileSalir.Click, Sub() CurrentMDI.Close()
        AddHandler menuFileSeleccionarAbrir.Click, Sub() Abrir()
        AddHandler buttonSeleccionar.Click, Sub() Abrir()
        AddHandler menuFileGuardar.Click, Sub() Form1Activo.Guardar()
        AddHandler buttonGuardar.Click, Sub() Form1Activo.Guardar()
        AddHandler menuFileGuardarComo.Click, Sub() Form1Activo.GuardarComo()
        AddHandler buttonGuardarComo.Click, Sub() Form1Activo.GuardarComo()
        AddHandler menuFileNuevo.Click, Sub() Nuevo()
        AddHandler buttonNuevo.Click, Sub() Nuevo()

        AddHandler menuFileRecargar.Click, Sub() Form1Activo.Recargar()
        AddHandler buttonRecargar.Click, Sub() Form1Activo.Recargar()

        AddHandler menuRecargarFichero.Click, Sub() Form1Activo.Recargar()
        AddHandler menuCopiarPath.Click, Sub()
                                             Try
                                                 Clipboard.SetText(Form1Activo.nombreFichero)
                                             Catch ex As Exception
                                             End Try
                                         End Sub

        ' Edición: menús y botones de la barra de edición
        'AddHandler menuEdit.DropDownOpening, Sub() menuEditDropDownOpening()
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
        AddHandler menuCompilar.Click, Sub() Build()
        AddHandler buttonCompilar.Click, Sub() Build()
        AddHandler menuEjecutar.Click, Sub() CompilarEjecutar()
        AddHandler buttonEjecutar.Click, Sub() CompilarEjecutar()

        AddHandler buttonEvaluar.Click, Sub() Evaluar()
        AddHandler menuEvaluar.Click, Sub() Evaluar()
        AddHandler buttonColorearAlEvaluar.Click, Sub() colorearAlEvaluar = buttonColorearAlEvaluar.Checked
        AddHandler buttonCompilarAlEvaluar.Click, Sub() compilarAlEvaluar = buttonCompilarAlEvaluar.Checked

        AddHandler menuColorear.Click, Sub()
                                           cargando = True
                                           ColorearCodigo()
                                           cargando = False
                                       End Sub
        AddHandler buttonColorear.Click, Sub()
                                             cargando = True
                                             ColorearCodigo()
                                             cargando = False
                                         End Sub
        AddHandler menuNoColorear.Click, Sub() NoColorear()
        AddHandler buttonNoColorear.Click, Sub() NoColorear()

        ' Ocultar el panel de evaluación                            (05/Oct/20)
        ' en evento normal con el del botón                         (20/Oct/20)

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
    Private Sub Inicializar()

        DirDocumentos = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        DirConfiguracion = Path.Combine(DirDocumentos, Application.ProductName)
        If Not Directory.Exists(DirConfiguracion) Then
            Directory.CreateDirectory(DirConfiguracion)
        End If
        FicheroConfiguracion = Path.Combine(DirConfiguracion, Application.ProductName & ExtensionConfiguracion)

        ' Crear el de configuración "local"
        DirConfigLocal = Path.Combine(DirConfiguracion, "Config-Ficheros")
        If Not Directory.Exists(DirConfigLocal) Then
            Directory.CreateDirectory(DirConfigLocal)
        End If


        ' No mostrar el panel al iniciar
        ' Ponerlo después de inicializando = false                  (30/Sep/20)
        ' para que se ajuste el tamaño de FlowPanel
        MostrarPanelBuscar(False, esReemplazar:=False)

        variosLineaComandos = Environment.GetCommandLineArgs.Length

        LeerConfig()

        ' Mostrar los 15 primeros en el menú Recientes
        AsignarRecientes()

        inicializando = False

        esCtrlF = True

        ' Guardar los esquemas de colores
        ' en: MyDocuments/ColorDictionary.csv
        Compilar.WriteColorDictionaryToFile()
        ' Si se cambia y se quiere leer
        'Compilar.UpdateColorDictionaryFromFile()

        ' posicionarlo al principio
        richTextBoxCodigo.SelectionStart = 0
        richTextBoxCodigo.SelectionLength = 0

        ' Esto se usará para cargar                                 (16/Oct/20)
        ' si se indica en la línea de comandos,
        ' ya que se habrán cargado en LeerConfig
        ' Comprobar si se indica un fichero en la línea de comandos (28/Sep/20)
        ' Si hay varios, abrirlos todos                             (30/Oct/20)
        ' De ser así, se obvian los indicados en la configuración
        ' Más arriba (antes de LeerConfig) se asigna variosLineaComandos
        If variosLineaComandos > 1 Then
            Dim n = 0
            ' El primer fichero es el ejecutable                    (30/Oct/20)
            For i = 1 To variosLineaComandos - 1
                Dim sFicArg = Environment.GetCommandLineArgs(i)
                If String.IsNullOrWhiteSpace(sFicArg) Then Continue For

                ' Abrir siempre en ventanas nuevas                  (30/Oct/20)
                ' aunque puede que ya haya un formulario asignado
                ' Si hay más de un formulario abierto               (16/Oct/20)
                ' abrirlo en una nueva ventana
                n += 1
                If n > 0 Then
                    Nuevo()
                End If
                Form1Activo.nombreFichero = sFicArg
                Form1Activo.Abrir(sFicArg)

                ' Mostrar los números de línea
                If Form1Activo.nombreFichero <> "" Then Form1Activo.AñadirNumerosDeLinea()

                ' Iniciar la posición al principio
                'MostrarPosicion(New KeyEventArgs(Keys.Home))
            Next
            m_fProcesando.Close()
        End If

        buttonLenguaje.DropDownItems(0).Text = Compilar.LenguajeVisualBasic
        buttonLenguaje.DropDownItems(1).Text = Compilar.LenguajeCSharp

    End Sub

    ''' <summary>
    ''' Para saber si se deben o no habilitar los botones de los toolStrip
    ''' Antes estaba en MenuEdiDropDownOpening
    ''' </summary>
    Friend Sub HabilitarBotones()
        If inicializando Then Return
        If cargando Then Return

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
        ' el botón de mostrar/ocultar el panel de evaluación        (20/Oct/20)
        buttonMostrarEvaluacion.Enabled = Form1Activo.lstSyntax.Items.Count > 0
        If Form1Activo.splitContainer1.Panel2.Visible = True Then
            buttonMostrarEvaluacion.Text = "Ocultar panel de evaluación/error"
        Else
            buttonMostrarEvaluacion.Text = "Mostrar panel de evaluación/error"
        End If
        buttonMostrarEvaluacion.ToolTipText = buttonMostrarEvaluacion.Text

        buttonEditorMarcador.Enabled = b
        buttonEditorClasificarSeleccion.Enabled = richTextBoxCodigo.SelectionLength > 0

        b = Form1Activo.Bookmarks.Count > 0
        buttonEditorMarcadorAnteriorLocal.Enabled = b
        buttonEditorMarcadorSiguienteLocal.Enabled = b

        b = totalBookmarks > 0
        buttonEditorMarcadorAnterior.Enabled = b
        buttonEditorMarcadorSiguiente.Enabled = b
        buttonEditorMarcadorQuitarTodos.Enabled = b

        b = ColRecortes.Count > 0
        buttonEdicionRecortes.Enabled = b

        HabilitarBotonesNavegar()

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

        ' desde aquí se llama al evento Load de Form1
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

        ' A nive de Form1
        Form1Activo.lstSyntax.Items.Clear()
        Form1Activo.richTextBoxLineas.Clear()
        Form1Activo.nombreFichero = ""
        Form1Activo.TextoModificado = False
        Form1Activo.codigoNuevo = ""
        Form1Activo.codigoAnterior = ""

        inicializando = False
    End Sub

    Private Sub ShowNewForm() Handles menuVentanaNueva.Click
        NuevaVentana()
    End Sub

    Private Sub CascadeToolStripMenuItem_Click() Handles menuVentanaCascade.Click
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub TileVerticalToolStripMenuItem_Click() Handles menuVentanaTileVertical.Click
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub TileHorizontalToolStripMenuItem_Click() Handles menuVentanaTileHorizontal.Click
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub ArrangeIconsToolStripMenuItem_Click() Handles menuVentanaArrangeAll.Click
        Me.LayoutMdi(MdiLayout.ArrangeIcons)
    End Sub

    Private Sub CloseAllToolStripMenuItem_Click() Handles menuVentanaCloseAll.Click
        ' Close all child forms of the parent.
        For Each ChildForm As Form1 In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub

    Private Sub MaximizarTodasMenuVentana_Click() Handles menuVentanaMaximizarTodas.Click
        Me.Text = $"{MDIPrincipal.TituloMDI}"
        For Each ChildForm As Form1 In Me.MdiChildren
            ChildForm.WindowState = FormWindowState.Maximized
        Next
    End Sub

    Private Sub RestaurarMenuVentana_Click() Handles menuVentanaRestaurarTodas.Click
        Me.Text = $"{MDIPrincipal.TituloMDI}"
        For Each ChildForm As Form1 In Me.MdiChildren
            ChildForm.WindowState = FormWindowState.Normal
        Next
    End Sub

    Private Sub timerClipBoard_Tick() Handles timerClipBoard.Tick
        ' Se ejecutará cada 15 segundos (1000 ms = 1 s, 15000 ms = 15 s)
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

    Private Sub menuVentana_DropDownOpening() Handles menuVentana.DropDownOpening
        ' Forzar a mostrar el nombre en el menú de ventanas         (16/Oct/20)
        ' De una respuesta en stackoverflow
        ' https://stackoverflow.com/a/1348453/14338047

        '' asignar a Text el nombre del fichero sin path             (20/Oct/20)
        '' No es necesario, se asignaba el path completo en Form1_Activate!!!

        ' Finalmente lo pongo en TextoModificado                    (30/Oct/20)

        '' Si está modificado, añadir un asterisco                   (30/Oct/20)
        '' Lo quito del Activate, ya que si la ventana no se ha activado
        '' no tendrá el asterisco
        ''
        'For Each frm As Form1 In Me.MdiChildren
        '    'frm.Text = Path.GetFileName(frm.nombreFichero)
        '    ' solo añadirlo si no lo tiene
        '    If frm.TextoModificado AndAlso Not frm.Text.EndsWith(" *") Then
        '        frm.Text &= " *"
        '    End If
        'Next

        If ActiveMdiChild IsNot Nothing Then
            Dim ventanaActiva As Form = Me.ActiveMdiChild
            ActivateMdiChild(Nothing)
            ActivateMdiChild(ventanaActiva)
        End If
    End Sub

    Private Sub GuardarTodo_Click() Handles buttonGuardarTodo.Click, menuFileGuardarTodo.Click
        MostrarProcesando("Guardar todo", "Guardando todos los ficheros abiertos...", MdiChildren.Count * 2)
        Dim t = 0
        For Each frm As Form1 In Me.MdiChildren
            labelInfo.Text = $"Guardando {frm.nombreFichero}..."
            OnProgreso(labelInfo.Text)
            Application.DoEvents()
            frm.Guardar()
        Next
        labelInfo.Text = $"Guardado {If(t = 1, "un fichero", $"{t} ficheros")}."
        OnProgreso(labelInfo.Text)

        m_fProcesando.Close()
    End Sub

    Private Sub menuFileRecientes_DropDownOpening() Handles menuFileRecientes.DropDownOpening
        ' No hacer nada si no hay nombre de fichero                 (20/Oct/20)
        If String.IsNullOrWhiteSpace(Form1Activo.nombreFichero) Then Return

        Dim yaSeleccionado = False
        For i = 0 To menuFileRecientes.DropDownItems.Count - 1
            TryCast(menuFileRecientes.DropDownItems(i), ToolStripMenuItem).Checked = False
            If menuFileRecientes.DropDownItems(i).Text.IndexOf(Form1Activo.nombreFichero) > 3 Then
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
    ' Truco adaptado del proyecto CSharpToVB de Paul1956
    '

    Private Sub comboBoxBuscar_Enter() Handles comboBoxBuscar.Enter
        If comboBoxBuscar.Text <> BuscarVacio Then
            comboBoxBuscar.ForeColor = SystemColors.ControlText
        End If
    End Sub

    Private Sub comboBoxBuscar_Leave() Handles comboBoxBuscar.Leave
        ' Esto hace lo mismo que comprobar si está en blanco... :-P
        'If Not Me.comboBoxBuscar.Text.Any Then
        If String.IsNullOrEmpty(comboBoxBuscar.Text) Then '
            comboBoxBuscar.ForeColor = SystemColors.GrayText
            comboBoxBuscar.Text = BuscarVacio
        End If
    End Sub

    Private Sub comboBoxReemplazar_Enter() Handles comboBoxReemplazar.Enter
        If comboBoxReemplazar.Text <> ReemplazarVacio Then
            comboBoxReemplazar.ForeColor = SystemColors.ControlText
        End If
    End Sub

    Private Sub comboBoxReemplazar_Leave() Handles comboBoxReemplazar.Leave
        If String.IsNullOrEmpty(comboBoxReemplazar.Text) Then
            comboBoxReemplazar.ForeColor = SystemColors.GrayText
            comboBoxReemplazar.Text = ReemplazarVacio
        End If
    End Sub

    Private buscarBoxAnt As String = BuscarVacio
    Private reemplazarBoxAnt As String = ReemplazarVacio

    Private Sub comboBoxBuscar_TextChanged() Handles comboBoxBuscar.TextChanged
        If inicializando Then Return
        If String.IsNullOrEmpty(comboBoxBuscar.Text) OrElse comboBoxBuscar.Text = BuscarVacio Then
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

    Private Sub comboBoxReemplazar_TextChanged() Handles comboBoxReemplazar.TextChanged
        If inicializando Then Return

        If String.IsNullOrEmpty(comboBoxReemplazar.Text) OrElse comboBoxReemplazar.Text = ReemplazarVacio Then
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

    Private Sub menuEditor_DropDownOpening() Handles menuEditor.DropDownOpening
        Dim b = richTextBoxCodigo.SelectionLength > 0
        menuEditorQuitarIndentacion.Enabled = b
        menuEditorPonerIndentacion.Enabled = b
        menuEditorQuitarComentarios.Enabled = b
        menuEditorPonerComentarios.Enabled = b

        menuEditorClasificarSeleccion.Enabled = b

        menuEditorPonerTexto.Enabled = richTextBoxCodigo.TextLength > 0

        b = Form1Activo.Bookmarks.Count > 0
        'menuEditorMarcador.Enabled = b
        menuEditorMarcadorAnteriorLocal.Enabled = b
        menuEditorMarcadorSiguienteLocal.Enabled = b

        b = totalBookmarks > 0
        menuEditorMarcadorAnterior.Enabled = b
        menuEditorMarcadorSiguiente.Enabled = b
        menuEditorMarcadorQuitarTodos.Enabled = b
    End Sub

    Private Sub menuEditorCambiarMayúsculas_DropDownOpening() Handles menuEditorCambiarMayúsculas.DropDownOpening
        Dim b = richTextBoxCodigo.SelectionLength > 0
        menuMayúsculas.Enabled = b
        menuMinúsculas.Enabled = b
        menuTitulo.Enabled = b
        menuPrimeraMinúsculas.Enabled = b
    End Sub

    Private Sub menuEditorQuitarEspacios_DropDownOpening() Handles menuEditorQuitarEspacios.DropDownOpening
        Dim b = richTextBoxCodigo.SelectionLength > 0
        menuQuitarEspaciosTrim.Enabled = b
        menuQuitarEspaciosTrimStart.Enabled = b
        menuQuitarEspaciosTrimEnd.Enabled = b
        menuQuitarEspaciosTodos.Enabled = b
    End Sub

    ''' <summary>
    ''' Cuando se abre el menú de edición o se cambia la selección del código
    ''' asignar si están o no habilitados el menú en sí, el contextual y las barras de herramientas.
    ''' </summary>
    Public Sub menuEditDropDownOpening() Handles menuEdit.DropDownOpening, rtbCodigoContext.Opening
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

        ' También he puesto como menú pegar recorte                 (11/Oct/20)
        menuEditPegarRecorte.Enabled = ColRecortes.Count > 0

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

    Private Sub buttonMostrarEvaluacion_Click() Handles buttonMostrarEvaluacion.Click, menuOcultarEvaluar.Click
        If menuOcultarEvaluar.Text = "Ocultar panel de evaluación/error" Then
            Form1Activo.splitContainer1.Panel2.Visible = False
            menuOcultarEvaluar.Text = "Mostrar panel de evaluación/error"
        Else
            menuOcultarEvaluar.Text = "Ocultar panel de evaluación/error"
            Form1Activo.splitContainer1.Panel2.Visible = True
        End If
        buttonMostrarEvaluacion.Text = menuOcultarEvaluar.Text
        buttonMostrarEvaluacion.ToolTipText = buttonMostrarEvaluacion.Text
        Form1Activo.splitContainer1_Resize()
    End Sub

    Private Sub menuTools_DropDownOpening() Handles menuTools.DropDownOpening
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
        Else
            menuOcultarEvaluar.Text = "Mostrar panel de evaluación/error"
        End If
        buttonMostrarEvaluacion.Text = menuOcultarEvaluar.Text
        buttonMostrarEvaluacion.ToolTipText = buttonMostrarEvaluacion.Text
    End Sub

    '
    ' Para navegar en las últimas posiciones de edición.
    ' Se buscará la ventana indicada por el fichero, se mostrará el formulario
    ' y se posicionará en la última posición de esa ventana.
    '

    ''' <summary>
    ''' Las últimas posiciones de donde se ha estado editando.
    ''' Se guarda el fichero y la posición.
    ''' </summary>
    Friend UltimasPosiciones As New Dictionary(Of Integer, Marcadores)

    ''' <summary>
    ''' La posición actual en la navegación.
    ''' </summary>
    Friend PosNavegarActual As (Actual As Integer, Fichero As String, Posicion As Integer, Texto As String) = (-1, "", -1, "")

    Private Sub buttonNavegarAnterior_Click() Handles buttonNavegarAnterior.Click
        Navegar(anterior:=True)
    End Sub

    Private Sub buttonNavegarSiguiente_Click() Handles buttonNavegarSiguiente.Click
        Navegar(anterior:=False)
    End Sub

    ''' <summary>
    ''' Posicionarse en la posición anterior o siguiente de navegación.
    ''' </summary>
    ''' <param name="anterior">True si se ha pulsado en el botón anterior, 
    ''' False si se ha pulsado en siguiente</param>
    Private Sub Navegar(anterior As Boolean)
        If UltimasPosiciones.Count = 0 Then Return

        If anterior Then
            PosNavegarActual.Actual -= 1
            If PosNavegarActual.Actual < 0 Then
                ' Repetir las posiciones si se llega a la última
                PosNavegarActual.Actual = UltimasPosiciones.Count - 1
            End If
        Else
            PosNavegarActual.Actual += 1
            If PosNavegarActual.Actual >= UltimasPosiciones.Count Then
                PosNavegarActual.Actual = UltimasPosiciones.Count - 1
            End If
        End If
        ' Este caso no debería darse nunca, pero para comprobar
        If PosNavegarActual.Actual < 0 OrElse PosNavegarActual.Actual > UltimasPosiciones.Count - 1 Then
            PosNavegarActual.Actual = UltimasPosiciones.Count - 1
            Debug.Assert(False)
        End If

        HabilitarBotonesNavegar()

        PosNavegarActual.Fichero = UltimasPosiciones(PosNavegarActual.Actual).Fichero
        PosNavegarActual.Posicion = UltimasPosiciones(PosNavegarActual.Actual).Posicion
        'NavegarA(PosNavegarActual)
        ' actualizar el menú con la posición actual
        ' desde ahí se navega
        buttonNavegarMenu.DropDownItems(PosNavegarActual.Actual).PerformClick()
    End Sub

    ''' <summary>
    ''' Navegar al fichero y posición indicadas.
    ''' </summary>
    ''' <param name="fic"></param>
    ''' <param name="pos"></param>
    Private Sub NavegarA(posNav As (Actual As Integer, Fichero As String, Posicion As Integer, Texto As String))
        cargando = True
        ' Buscar el fichero en las ventanas abiertas
        For Each frm As Form1 In Me.MdiChildren
            If frm.nombreFichero = posNav.Fichero Then
                ' posicionarse
                frm.richTextBoxCodigo.SelectionStart = posNav.Posicion
                frm.richTextBoxCodigo.SelectionLength = 0
                frm.BringToFront()
                Exit For
            End If
        Next
        cargando = False
    End Sub

    ''' <summary>
    ''' Habilitar adecuadamente los botones de navegar.
    ''' </summary>
    Private Sub HabilitarBotonesNavegar()
        buttonNavegarAnterior.Enabled = True ' (PosNavegarActual.Actual > 0)
        buttonNavegarSiguiente.Enabled = (PosNavegarActual.Actual < UltimasPosiciones.Count - 1)
    End Sub

    ''' <summary>
    ''' Asignar la posición actual en el formulario indicado.
    ''' </summary>
    ''' <param name="elForm1"></param>
    Friend Sub AsignarNavegar(elForm1 As Form1)
        If cargando Then Return

        ' No asignar la misma ubicación que la anterior
        If PosNavegarActual.Posicion = elForm1.richTextBoxCodigo.SelectionStart Then
            If PosNavegarActual.Fichero = elForm1.nombreFichero Then
                Return
            End If
        End If
        PosNavegarActual.Posicion = elForm1.richTextBoxCodigo.SelectionStart
        PosNavegarActual.Fichero = elForm1.nombreFichero

        ' Asignar la línea completa                                 (22/Oct/20)
        Dim pos = PosicionActual0()
        PosNavegarActual.Texto = elForm1.richTextBoxCodigo.Lines(pos.Linea)
        Dim marcador = New Marcadores(elForm1)
        ' El índice será el número de posiciones,
        ' al empezar, Count es cero y ese es el primer índice.
        Dim index = UltimasPosiciones.Count
        UltimasPosiciones.Add(index, marcador)
        PosNavegarActual.Actual = index

        ' Asignar al menú del botón anterior las posiciones
        Dim s = PosNavegarActual.Texto.TrimStart
        If s.Length > 60 Then
            s = s.Substring(0, 60)
        End If
        s = $"{Path.GetFileName(PosNavegarActual.Fichero)}: {s}"
        For Each m As ToolStripMenuItem In buttonNavegarMenu.DropDownItems
            m.Checked = False
        Next
        Dim mnu As New ToolStripMenuItem(s)
        mnu.Tag = PosNavegarActual
        AddHandler mnu.Click, Sub(sender As Object, e As EventArgs)
                                  For Each m As ToolStripMenuItem In buttonNavegarMenu.DropDownItems
                                      m.Checked = False
                                  Next
                                  Dim m2 = TryCast(sender, ToolStripMenuItem)
                                  Dim pna = CType(m2.Tag, (Actual As Integer, Fichero As String, Posicion As Integer, Texto As String))
                                  NavegarA(pna)
                                  m2.Checked = True
                                  PosNavegarActual = pna
                                  HabilitarBotonesNavegar()
                              End Sub
        mnu.Checked = True
        buttonNavegarMenu.DropDownItems.Add(mnu)

    End Sub

    Private Sub buttonEditorMarcador_Click() Handles buttonEditorMarcador.Click, menuEditorMarcador.Click
        Form1Activo.MarcadorPonerQuitar()
    End Sub

    Private Sub buttonEditorMarcadorAnterior_Click() Handles buttonEditorMarcadorAnterior.Click, menuEditorMarcadorAnterior.Click, menuEditorMarcadorAnteriorLocal.Click
        ' Ir al marcador anterior, en todos los ficheros abiertos   (24/Oct/20)
        Dim b = Form1Activo.MarcadorAnterior
        If b = False Then
            ' Buscar en otro fichero
            ' Buscar cuál es el form1activo en la lista de ventanas
            ' e ir a la anterior si es que hay más
            Dim index = 0
            For i = Me.MdiChildren.Count - 1 To 0 Step -1
                Dim frm = Me.MdiChildren(i)
                If Form1Activo.Text = frm.Text Then
                    Do While i > 0
                        frm = Me.MdiChildren(i - 1)
                        Form1Activo = TryCast(frm, Form1)
                        If Form1Activo.Bookmarks.Count > 0 Then
                            b = Form1Activo.MarcadorAnterior()
                            If b Then
                                Form1Activo.BringToFront()
                                Exit For
                            End If
                        End If
                        i -= 1
                    Loop
                End If
            Next
        End If
        If b = False Then
            MessageBox.Show("Se ha llegado al principio de los ficheros con marcadores/bookmarks.",
                            "Marcador global anterior",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub buttonEditorMarcadorAnteriorLocal_Click() Handles buttonEditorMarcadorAnteriorLocal.Click
        Form1Activo.MarcadorAnterior()
    End Sub

    Private Sub EditorQuitarTodosLosMarcadores() Handles buttonEditorMarcadorQuitarTodos.Click, menuEditorMarcadorQuitarTodos.Click
        If MessageBox.Show("¿Seguro que quieres quitar todos los marcadores.",
                           "Quitar todos los marcadores",
                           MessageBoxButtons.YesNo,
                           MessageBoxIcon.Question) = DialogResult.Yes Then

            For Each frm As Form1 In Me.MdiChildren
                frm.MarcadorQuitarTodos()
            Next
        End If
    End Sub

    Private Sub buttonEditorMarcadorSiguiente_Click() Handles buttonEditorMarcadorSiguiente.Click, menuEditorMarcadorSiguiente.Click, menuEditorMarcadorSiguienteLocal.Click
        ' Ir al marcador siguiente, en todos los ficheros abiertos
        Dim b = Form1Activo.MarcadorSiguiente
        If b = False Then
            ' Buscar en otro fichero
            ' Buscar cuál es el form1activo en la lista de ventanas
            ' e ir a la siguiente si es que hay más
            ' aunque fallará si la siguiente ventana no tiene marcadores
            For i = 0 To Me.MdiChildren.Count - 1
                Dim frm = Me.MdiChildren(i)
                If Form1Activo.Text = frm.Text Then
                    Do While i < Me.MdiChildren.Count - 1
                        frm = Me.MdiChildren(i + 1)
                        Form1Activo = TryCast(frm, Form1)
                        If Form1Activo.Bookmarks.Count > 0 Then
                            b = Form1Activo.MarcadorSiguiente()
                            If b Then
                                Form1Activo.BringToFront()
                                Exit For
                            End If
                        End If
                        i += 1
                    Loop
                End If
            Next
        End If
        If b = False Then
            MessageBox.Show("Se ha llegado al final de los ficheros con marcadores/bookmarks.",
                            "Marcador global siguiente",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub buttonEditorMarcadorSiguienteLocal_Click() Handles buttonEditorMarcadorSiguienteLocal.Click
        Form1Activo.MarcadorSiguiente()
    End Sub

    Private Sub MDIPrincipal_DragEnter(sender As Object, e As DragEventArgs) Handles MyBase.DragEnter
        Form1Activo.Form1_DragEnter(sender, e)
    End Sub

    Private Sub MDIPrincipal_DragDrop(sender As Object, e As DragEventArgs) Handles MyBase.DragDrop
        Form1Activo.Form1_DragDrop(sender, e)
    End Sub
End Class
