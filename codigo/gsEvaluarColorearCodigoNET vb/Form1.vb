'------------------------------------------------------------------------------
' Evaluar, Compilar y colorear código                               (26/Sep/20)
'
' Editor de múltiples ficheros con opciones de editor de código
'
' Con opciones de compilación:
'   Compilar y ejecutar
'   Compilar (build)
'   Evaluar el código.
' Colorear el código (en el editor) y en formato HTML.
'
'   Al colorear y compilar/evaluar se muestra:
'       las palabras clave o los errores producidos al compilar
'
' Funciones generales de edición de texto:
'   Buscar y reemplazar
'   Poner y quitar comentarios
'   Poner y quitar indentación
'   Convertir a mayúsculas, minúsculas y título
'   Clasificar el texto
'   Quitar espacios de delante, detrás y todos los espacios
'   Poner y quitar palabras delante y detrás de cada línea seleccionada
' Navegación y marcadores (bookmarks)
'   Opción de navegar adelante y atrás en los sitios en que se ha estado de cada fichero
'   Opción de poner/quitar marcadores
'       ir al siguiente y anterior,
'       tanto en el fichero actual como en todos los ficheros
'
'------------------------------------------------------------------------------
'
' Para ver las diferentes revisiones, ver el fichero Revisiones.txt
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
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Public Class Form1

    ''' <summary>
    ''' El inicio de la selección al cerrar
    ''' </summary>
    Friend selectionStartAnt As Integer

    ''' <summary>
    ''' El final de la selección al cerrar
    ''' </summary>
    Friend selectionEndAnt As Integer

    Private _nombreFichero As String

    ''' <summary>
    ''' Nombre del fichero asignado al formulario actual.
    ''' Al asignarlo, si no está en la colección <see cref="UltimasVentanasAbiertas"/>
    ''' se añadirá.
    ''' </summary>
    Friend Property nombreFichero As String
        Get
            Return _nombreFichero
        End Get
        Set(value As String)
            _nombreFichero = value
            ' No guardarla si es una cadena vacía                   (26/Oct/20)
            If String.IsNullOrWhiteSpace(value) Then
                Return
            End If
            CompararString.IgnoreCase = True
            If Not UltimasVentanasAbiertas.Contains(value, New CompararString) Then
                UltimasVentanasAbiertas.Add(value)
            End If
        End Set
    End Property

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
            ' Asignar el nombre del fichero sin asterisco           (30/Oct/20)
            Me.Text = Path.GetFileName(nombreFichero)
            If value Then
                labelModificado.Text = "*"
                Me.Text &= " *"
            Else
                labelModificado.Text = " "
            End If
            richTextBoxCodigo.Modified = value
        End Set
    End Property

    '
    ' Los métodos de evento del formulario
    '

    Private Sub Form1_Load() Handles Me.Load
        ' Asignaciones al richTextBox                               (28/Sep/20)
        ' (por si las cambio por error en el diseñador)
        richTextBoxCodigo.EnableAutoDragDrop = False
        richTextBoxCodigo.AcceptsTab = False

        ' Asignaciones al Form
        Me.KeyPreview = True

        buttonLenguaje.DropDownItems(0).Text = Compilar.LenguajeVisualBasic
        buttonLenguaje.DropDownItems(1).Text = Compilar.LenguajeCSharp
        buttonLenguaje.Image = buttonLenguaje.DropDownItems(2).Image
        buttonLenguaje.Text = ExtensionTexto ' buttonLenguaje.DropDownItems(2).Text

        richTextBoxCodigo.ContextMenuStrip = CurrentMDI.rtbCodigoContext

        LeerConfigLocal()

        If Me.MdiParent Is CurrentMDI Then
            Me.BringToFront()
        End If

    End Sub

    ''' <summary>
    ''' Si se pregunta guardar, y quedan más ficheros sin guardar,
    ''' usar la misma respuesta (si se dice que Sí)
    ''' </summary>
    ''' <remarks>30/Oct/2020</remarks>
    Private Shared guardarSinPreguntar As Boolean

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
            If TextoModificado OrElse String.IsNullOrEmpty(nombreFichero) Then
                If CerrandoVarios AndAlso guardarSinPreguntar Then
                    Guardar()
                Else
                    ' Si está modificado, y se indicó guardar               (30/Oct/20)
                    ' guardar el resto de ficheros sin preguntar
                    ' esto es para el caso que se cierre el MDI y haya ficheros sin guardar
                    Dim msg = "¿Quieres guardarlo?"
                    If e.CloseReason = CloseReason.MdiFormClosing Then
                        CerrandoVarios = True
                        msg &= $"{vbCrLf}Se está cerrando la aplicación, si guardas este, también se guardarán los restantes ficheros que estén modificados."
                    Else
                        msg &= $"{vbCrLf}Si estás cerrando varias ventanas, también se guardarán los restantes ficheros que estén modificados."
                    End If

                    ' No preguntaba si quería guardar,                      (01/Oct/20)
                    ' porque usaba GuadarComo, y si el nombre es el mismo, no mostraba el diálogo

                    ' Dar la oportunidad de cancelar para seguir editando   (02/Oct/20)
                    Dim res As DialogResult = DialogResult.No
                    res = MessageBox.Show($"El texto está modificado,{vbCrLf}{msg}{vbCrLf}{vbCrLf}" &
                                          "Pulsa en Cancelar para seguir editando.",
                                          "Texto modificado y no guardado",
                                          MessageBoxButtons.YesNoCancel,
                                          MessageBoxIcon.Question)
                    If res = DialogResult.Yes Then
                        ' GuardarComo debe mostrar siempre el cuadro dediálogo.
                        ' Si no tiene nombre, preguntar                 (02/Oct/20)
                        ' Guardar se encarga de llamar a GuardarComo si no tiene nombre
                        Guardar()
                        If CerrandoVarios Then
                            guardarSinPreguntar = True
                        End If
                    ElseIf res = DialogResult.Cancel Then
                        guardarSinPreguntar = False
                        e.Cancel = True
                        Return
                    Else
                        guardarSinPreguntar = False
                    End If
                End If
            End If
            selectionStartAnt = -1
            selectionEndAnt = -1
            If richTextBoxCodigo.SelectionLength > 0 Then
                selectionStartAnt = richTextBoxCodigo.SelectionStart
                selectionEndAnt = richTextBoxCodigo.SelectionLength
            End If
        End If

        GuardarConfigLocal()

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

    Friend Sub Form1_Activated() Handles Me.Activated
        If cargando Then Return

        ' Asignar cuál es el Form1 activo
        Form1Activo = Me

        ' No asignar el nombre del fichero al Text                  (30/Oct/20)
        ' ya que se hará en TextoModificado
        'If Not String.IsNullOrWhiteSpace(nombreFichero) Then
        '    ' Aquí se asignaba el path completo!!!                  (20/Oct/20)
        '    Me.Text = Path.GetFileName(nombreFichero)
        'End If

        ' esto está en el evento menuVentana.DropDownOpening del MDI
        ' ya que si no, solo se pondrá si se activa la ventana
        '' si está modificado, añadir un asterisco                   (30/Oct/20)
        'If TextoModificado Then
        '    Me.Text &= " *"
        'End If

        If Me.WindowState = FormWindowState.Normal Then
            CurrentMDI.Text = $"{MDIPrincipal.TituloMDI} [{Me.Text}]"
        End If

        ' Activar los botones que correspondan
        CurrentMDI.HabilitarBotones()

        ' Referescar la lista de sintaxis, porque no se repinta     (23/Oct/20)
        lstSyntax.Refresh()
    End Sub

    Private Sub Form1_Resize() Handles Me.Resize
        If inicializando Then Return

        'If Me.WindowState = FormWindowState.Normal Then
        '    tamForm = (Me.Left, Me.Top, Me.Height, Me.Width)
        'Else
        '    tamForm = (Me.RestoreBounds.Left, Me.RestoreBounds.Top,
        '               Me.RestoreBounds.Height, Me.RestoreBounds.Width)
        'End If

        UtilFormEditor.labelInfo.Text = $"MDIPrincipal: Ancho: {CurrentMDI.Width}, Alto: {CurrentMDI.Height} / ClientSize: Ancho: {CurrentMDI.ClientSize.Width}, Alto: {CurrentMDI.ClientSize.Height} - Form1: Ancho: {Me.Width}, Alto: {Me.Height}"
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
                e.SuppressKeyPress = True

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

    Public Sub Form1_DragEnter(sender As Object, e As DragEventArgs) Handles Me.DragEnter, richTextBoxCodigo.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) OrElse
                e.Data.GetDataPresent(DataFormats.Text) OrElse
                e.Data.GetDataPresent("System.String") Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Public Sub Form1_DragDrop(sender As Object, e As DragEventArgs) Handles Me.DragDrop, richTextBoxCodigo.DragDrop
        If e.Data.GetDataPresent("System.String") Then
            Dim fic As String = CType(e.Data.GetData("System.String"), String)
            ' Comprobar que sea una URL o un fichero
            If fic.IndexOf("http") > -1 OrElse fic.IndexOf(":\") > -1 Then
                Abrir(fic)
                m_fProcesando.Close()
            End If
        ElseIf e.Data.GetDataPresent(DataFormats.FileDrop) Then
            ' Si se sueltan varios ficheros, abrirlos todos         (30/Oct/20)
            Dim files = CType(e.Data.GetData(DataFormats.FileDrop, True), String())
            ' Si se suelta 1, abrirlo en la misma ventana
            ' si son más de 1, abrirlos en ventanas nuevas
            If files.Length = 1 Then
                Dim fic = files(0)
                Abrir(fic)
            Else
                MostrarProcesando("Cargando ficheros con Drag & Drop", $"Cargando {files.Length} ficheros...{vbCrLf}", files.Length * 2)
                Dim fic As String
                ' Al tener una ventana abierta y hacer DragDrop
                ' se quedaba modificado el primer fichero
                ' así que... si hay una ventana abierta, se deja como estaba
                ' y los ficheros soltados se abren en neuvas ventanas
                ' así también vale por si ya tenía texto

                'Dim origen = TryCast(sender, Form)
                'Dim j As Integer = 0
                '' si el dragdrop no se manda desde el MDI
                'If origen IsNot MDIPrincipal Then
                '    j = 1
                '    fic = files(0)
                '    Abrir(fic)
                '    TextoModificado = False
                'End If
                For i = 0 To files.Length - 1
                    fic = files(i)
                    Nuevo()
                    Form1Activo.Abrir(fic)
                    Form1Activo.TextoModificado = False
                Next
                'If origen IsNot MDIPrincipal Then
                '    TextoModificado = False
                '    'Me.Close()
                '    If Me.Text = "" Then
                '        Me.Close()
                '    End If
                'End If
            End If
            m_fProcesando.Close()
        ElseIf e.Data.GetDataPresent(DataFormats.Text) Then
            richTextBoxCodigo.SelectedText = e.Data.GetData("System.String", True).ToString()
        End If
    End Sub

    '
    ' Guardar y Leer configuración para este formulario
    '

    ''' <summary>
    ''' Guardar los datos de configuración relacionados con el fichero del formulario actual.
    ''' </summary>
    Private Sub GuardarConfigLocal()
        Dim ficConfig = Path.Combine(DirConfigLocal, Bookmarks.Fichero & ExtensionConfiguracion)
        Dim cfg = New Config(ficConfig)

        Dim seccion = $"Marcadores {Bookmarks.Fichero}"
        cfg.SetValue(seccion, "Count", Bookmarks.Count)
        For i = 0 To Bookmarks.Count - 1
            cfg.SetKeyValue(seccion, $"Inicio-{i}", Bookmarks.Item(i).Inicio)
            cfg.SetKeyValue(seccion, $"SelStart-{i}", Bookmarks.Item(i).SelStart)
        Next

        ' Para la selección al cerrar                               (05/Oct/20)
        seccion = $"Selección {Bookmarks.Fichero}"
        cfg.SetValue(seccion, "selectionStartAnt", selectionStartAnt)
        cfg.SetValue(seccion, "selectionEndAnt", selectionEndAnt)

        cfg.Save()
    End Sub

    ''' <summary>
    ''' Leer los datos de configuración relacionados con el fichero del formulario actual.
    ''' </summary>
    Private Sub LeerConfigLocal()
        Bookmarks = New Marcadores(Me)

        Dim ficConfig = Path.Combine(DirConfigLocal, Bookmarks.Fichero & ExtensionConfiguracion)
        Dim cfg = New Config(ficConfig)

        Dim cuantos As Integer

        Bookmarks.Clear()
        Dim seccion = $"Marcadores {Bookmarks.Fichero}"
        cuantos = cfg.GetValue(seccion, "Count", 0)
        For j = 0 To cuantos - 1
            Dim inicio = cfg.GetValue(seccion, $"Inicio-{j}", 0)
            Dim selStart = cfg.GetValue(seccion, $"SelStart-{j}", 0)
            Bookmarks.Add(inicio, selStart)
        Next

        ' Para la selección al cerrar                               (05/Oct/20)
        seccion = $"Selección {Bookmarks.Fichero}"
        selectionStartAnt = cfg.GetValue(seccion, "selectionStartAnt", -1)
        selectionEndAnt = cfg.GetValue(seccion, "selectionEndAnt", -1)

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

    Public Sub splitContainer1_Resize() Handles splitContainer1.Resize
        If CurrentMDI.WindowState = FormWindowState.Minimized Then Return

        If splitContainer1.Panel2.Visible Then
            If splitPanel2 < 100 Then
                splitPanel2 = CInt(splitContainer1.Width * 0.25)
            End If
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
    Public Sub richTextBoxCodigo_VScroll() Handles richTextBoxCodigo.VScroll
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
                e.SuppressKeyPress = True
                QuitarIndentacion(richTextBoxCodigo)
                MostrarPosicion(e)
            End If
        ElseIf e.Alt = False AndAlso e.Control = False Then
            If e.KeyCode = Keys.Tab Then
                ' Adelante
                e.Handled = True
                e.SuppressKeyPress = True
                PonerIndentacion(richTextBoxCodigo)
                MostrarPosicion(e)
            ElseIf e.KeyCode = Keys.Enter Then
                e.Handled = True
                e.SuppressKeyPress = True

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
                    ' Ahora tiene vbLf en vez de vbCr               (23/Oct/20)
                    ' Esto es lo que añadía una línea de más        (30/Oct/20)
                    ' cuando se pulsaba ENTER desde el final de la línea anterior
                    'richTextBoxCodigo.SelectedText = vbLf & New String(" "c, col)
                    richTextBoxCodigo.SelectedText = New String(" "c, col)
                    inicializando = False

                    '' Si se da esta condición, (creo que siempre se da),
                    '' ir al inicio, borrar e ir al final
                    '' Esta condición se da cuando se pulsa INTRO    (30/Oct/20)
                    '' cuando no está al principio de la línea
                    'If richTextBoxCodigo.GetLineFromCharIndex(richTextBoxCodigo.SelectionStart) > ln Then
                    '    'Debug.Assert(rtEditor.Lines(ln + 1) = "")
                    '    SendKeys.Send("{HOME}{BS}{END}")
                    'End If
                End If
                'End If
                MostrarPosicion(e)
            End If
        End If

    End Sub

    Public Sub richTextBoxCodigo_TextChanged() Handles richTextBoxCodigo.TextChanged
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

    Private Sub richTextBoxCodigo_SelectionChanged() Handles richTextBoxCodigo.SelectionChanged, richTextBoxCodigo.MouseClick
        If cargando Then Return

        ' Guardar la posición para usar en Navegar...               (21/Oct/20)
        CurrentMDI.AsignarNavegar(Me)

        CurrentMDI.HabilitarBotones()
        MostrarPosicion(Nothing)
    End Sub

    Private Sub richTextBoxCodigo_FontChanged() Handles richTextBoxCodigo.FontChanged
        richTextBoxLineas.Font = New Font(richTextBoxCodigo.Font.FontFamily, richTextBoxCodigo.Font.Size)
    End Sub

    '
    ' Para los marcadores / Bookmarks                               (28/Sep/20)
    ' Usando la clase Marcadores                                    (23/Oct/20)
    '

    ''' <summary>
    ''' Los marcadores para este formulario.
    ''' </summary>
    ''' <returns></returns>
    Friend Property Bookmarks As Marcadores

    ''' <summary>
    ''' Poner los marcadores, si hay... (30/Sep/20)
    ''' </summary>
    Friend Sub PonerLosMarcadores()
        If Bookmarks.Count = 0 Then Return

        inicializando = True

        ' Recordar la posición                                      (30/Sep/20)
        Dim selStart = richTextBoxCodigo.SelectionStart

        Bookmarks.Sort()
        Dim colMarcadorTemp = Bookmarks.ToList
        Bookmarks.Clear()
        For i = 0 To colMarcadorTemp.Count - 1
            richTextBoxCodigo.SelectionStart = colMarcadorTemp(i)
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
    Friend Sub MarcadorPonerQuitar()
        Dim posActual = PosicionActual0()
        If Bookmarks.Contains(posActual.Inicio) Then
            ' quitarlo
            Dim posAnt = Bookmarks.GetSelectionStart(posActual.Inicio)
            Bookmarks.Remove(posActual.Inicio)

            totalBookmarks -= 1

            richTextBoxCodigo.SelectionStart = posAnt
            richTextBoxCodigo.SelectionLength = 0
            Dim fcol = richTextBoxLineas.GetFirstCharIndexFromLine(posActual.Linea)
            richTextBoxLineas.SelectionStart = fcol
            richTextBoxLineas.SelectionLength = 5
            'richTextBoxLineas.SelectionBullet = False
            ' así es como se pone en AñadirNumerosDeLinea
            richTextBoxLineas.SelectedText = $" {(posActual.Linea + 1).ToString("0").PadLeft(4)}"
        Else
            Bookmarks.Add(posActual.Inicio, posActual.SelStart)
            totalBookmarks += 1

            ' Poner los marcadores en richTextBoxLineas
            Dim fcol = richTextBoxLineas.GetFirstCharIndexFromLine(posActual.Linea)
            richTextBoxLineas.SelectionStart = fcol
            richTextBoxLineas.SelectionLength = 5

            ' Poner delante la imagen del marcador
            ' Usando la imagen bookmark_003_8x10.png
            richTextBoxLineas.SelectedRtf = $"{picBookmark}{(posActual.Linea + 1).ToString("0").PadLeft(4)}" & "}"

            richTextBoxCodigo.SelectionStart = posActual.SelStart
            richTextBoxCodigo.SelectionLength = 0
        End If
        Bookmarks.Sort()
    End Sub

    ''' <summary>
    ''' La imagen a usar cuando se muestra un marcador en richTextBoxLineas.
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
    Friend Function MarcadorAnterior() As Boolean
        If Bookmarks.Count = 0 Then Return False

        Dim posActual = PosicionActual0()
        Dim res = Bookmarks.Where(Function(x) x < posActual.Inicio)
        If res.Count > 0 Then
            Dim pos = Bookmarks.GetSelectionStart(res.Last)
            richTextBoxCodigo.SelectionStart = pos
            Return True
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

            Return False
        End If
    End Function

    ''' <summary>
    ''' Ir al marcador siguiente.
    ''' Si está después del último, ir al anterior.
    ''' </summary>
    Friend Function MarcadorSiguiente() As Boolean
        If Bookmarks.Count = 0 Then Return False

        Dim posActual = PosicionActual0()
        Dim res = Bookmarks.Where(Function(x) x > posActual.Inicio)
        If res.Count > 0 Then
            Dim pos = Bookmarks.GetSelectionStart(res.First)
            richTextBoxCodigo.SelectionStart = pos

            Return True
        Else
            ' si no hay más marcadores, ir al anterior
            richTextBoxCodigo.SelectionStart = 0
            MarcadorSiguiente()

            Return False
        End If
    End Function

    ''' <summary>
    ''' Quitar todos los marcadores.
    ''' </summary>
    Friend Sub MarcadorQuitarTodos()
        If Bookmarks.Count = 0 Then Return

        Bookmarks.Clear()
        AñadirNumerosDeLinea()
    End Sub

    ''' <summary>
    ''' Añadir los números de línea
    ''' </summary>
    ''' <remarks>Como método separado 18/Sep/20</remarks>
    Friend Sub AñadirNumerosDeLinea()
        If inicializando Then Return
        If String.IsNullOrEmpty(richTextBoxCodigo.Text) Then Return

        Dim finlinea = richTextBoxCodigo.Text.ComprobarFinLinea
        Dim lineas = richTextBoxCodigo.Lines.Length
        richTextBoxLineas.Text = ""
        For i = 1 To lineas
            richTextBoxLineas.Text &= $" {i.ToString("0").PadLeft(4)}{finlinea}"
        Next
        ' Sincronizar los scrolls
        Form1Activo.richTextBoxCodigo_VScroll()

        PonerLosMarcadores()
    End Sub

    '
    ' Los métodos para abrir, guardar y recargar en el Form1        (23/Oct/20)
    '

    ''' <summary>
    ''' Abre nuevamente el último fichero
    ''' desechando los datos realizados
    ''' </summary>
    Friend Sub Recargar()
        Abrir(nombreFichero)
    End Sub

    ''' <summary>
    ''' Abre el fichero indicado en el parámetro, 
    ''' si no está en el combo de ficheros, añadirlo al principio.
    ''' De añadirlo al princpio se encarga <see cref="AñadirAUltimosFicherosAbiertos"/>.
    ''' </summary>
    ''' <param name="fic">El path completo del fichero a abrir</param>
    ''' <remarks>En el combo se muestra solo el nombre sin el path si el path es el directorio de documentos
    ''' (o el que se haya asignado como predeterminado) en otro caso se muestra el path completo</remarks>
    Friend Sub Abrir(fic As String)
        If String.IsNullOrWhiteSpace(fic) Then Return

        Dim sDirFic = Path.GetDirectoryName(fic)
        If Not File.Exists(fic) Then
            If String.IsNullOrWhiteSpace(sDirFic) Then
                fic = Path.Combine(DirDocumentos, fic)
            End If
            If File.Exists(fic) Then
                Abrir(fic)
            End If
            Return
        End If

        If String.IsNullOrWhiteSpace(sDirFic) Then
            fic = Path.Combine(DirDocumentos, fic)
        End If

        ' Leer la configuración para este fichero                   (23/Oct/20)
        ' Se asignará el valor a Bookmarks y el texto seleccionado
        ' Leerlo antes de abrir el fichero
        LeerConfigLocal()

        labelInfo.Text = $"Abriendo {fic}..."
        If m_fProcesando Is Nothing Then
            MostrarProcesando(labelInfo.Text, labelInfo.Text, 2)
        End If
        OnProgreso(labelInfo.Text)

        Dim sCodigo = ""
        Using sr As New StreamReader(fic, detectEncodingFromByteOrderMarks:=True, encoding:=Encoding.UTF8)
            sCodigo = sr.ReadToEnd()
        End Using

        ' Si se deben cambiar los TAB por 8 espacios                (05/Oct/20)
        ' cambiarlos por los indicados en EspaciosIndentar          (23/Oct/20)
        If CambiarTabs Then
            'sCodigo = sCodigo.Replace(vbTab, "        ")
            Dim sTabs = New String(" "c, EspaciosIndentar)
            sCodigo = sCodigo.Replace(vbTab, sTabs)
        End If

        codigoAnterior = sCodigo

        richTextBoxCodigo.Text = sCodigo

        ' El nombre del fichero con el path completo                (17/Oct/20)
        ' Ya no se usa NombreUltimoFichero a nivel global           (23/Oct/20)
        ' Siempre se hará referencia a Form1Activo.nombreFichero
        nombreFichero = fic

        ' En la ventana mostrar solo el nombre del fichero          (19/Oct/20)
        ' independientemente del path
        Me.Text = Path.GetFileName(fic)

        AñadirAUltimosFicherosAbiertos(fic)

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

        ' Mostrar información del fichero
        labelInfo.Text = $"{Path.GetFileName(fic)} ({sDirFic})"
        CurrentMDI.Text = $"{MDIPrincipal.TituloMDI} [{Me.Text}]"
        Application.DoEvents()

        ' Limpiar el contenido de la sintaxis                       (23/Oct/20)
        lstSyntax.Items.Clear()

        ' Si hay que colorear el fichero cargado
        If colorearAlCargar Then
            ' Para que tarde menos en colorear                      (23/Oct/20)
            ' pero tener en cuenta si cargando ya estaba asignado
            Dim cargandoTmp = cargando
            cargando = True
            ColorearCodigo()
            cargando = cargandoTmp
            ' Mostrar el panel de sintaxis
            ' Si es texto, como no se colorea,                      (30/Oct/20)
            ' no mostrar el panel de sintaxis
            If buttonLenguaje.Text = ExtensionTexto Then
                splitContainer1.Panel2.Visible = False
            Else
                splitContainer1.Panel2.Visible = True
            End If
        Else
            splitContainer1.Panel2.Visible = False
        End If
        splitContainer1_Resize()

        '
        ' Todo esto debe estar después de LeerConfigLocal:
        ' PonerLosMarcadores y la selección del texto anterior
        '

        ' Poner los marcadores de este fichero                      (23/Oct/20)
        ' Antes estaba en Inicializar de MDIPrincipal
        'PonerLosMarcadores()

        ' En AñadirNumerosDeLinea se llama a PonerLosMarcadores
        AñadirNumerosDeLinea()


        ' Marcar el texto que antes estaba seleccionado             (23/Oct/20)
        ' Antes estaba en Inicializar de MDIPrincipal
        If selectionStartAnt > -1 Then
            richTextBoxCodigo.SelectionStart = selectionStartAnt
            richTextBoxCodigo.SelectionLength = selectionEndAnt
        End If

        MostrarPosicion(Nothing)

        labelTamaño.Text = $"{richTextBoxCodigo.Text.Length:#,##0} car. ({richTextBoxCodigo.Text.CuantasPalabras:#,##0} palab.)"

        ' Si es texto, deshabilitar los botones que correspondan    (08/Oct/20)
        'If sLenguaje = ExtensionTexto Then
        '    CurrentMDI.HabilitarBotones()
        'End If
        ' Llamar siempre a este método para habilitar los botones   (26/Oct/20)
        CurrentMDI.HabilitarBotones()

        Me.TextoModificado = False

    End Sub

    ''' <summary>
    ''' Guarda el fichero indicado en el parámetro
    ''' </summary>
    ''' <param name="fic">El path completo del fichero a guardar</param>
    Public Sub Guardar(fic As String)
        labelInfo.Text = $"Guardando {fic}..."
        If m_fProcesando Is Nothing Then
            MostrarProcesando(labelInfo.Text, labelInfo.Text, 2)
        End If
        OnProgreso(labelInfo.Text)

        Dim sCodigo = richTextBoxCodigo.Text

        Dim sDirFic = Path.GetDirectoryName(fic)
        If String.IsNullOrWhiteSpace(sDirFic) Then
            fic = Path.Combine(DirDocumentos, fic)
        End If

        ' Si se deben cambiar los TAB por 8 espacios                (05/Oct/20)
        ' cambiarlos por los indicados en EspaciosIndentar          (23/Oct/20)
        If CambiarTabs Then
            Dim sTabs = New String(" "c, EspaciosIndentar)
            sCodigo = sCodigo.Replace(vbTab, sTabs)
        End If

        Using sw As New StreamWriter(fic, append:=False, encoding:=Encoding.UTF8)
            sw.WriteLine(sCodigo)
        End Using
        codigoAnterior = sCodigo

        labelInfo.Text = $"{Path.GetFileName(fic)} ({Path.GetDirectoryName(fic)})"
        nombreFichero = fic

        ' En la ventana mostrar solo el nombre del fichero          (19/Oct/20)
        ' independientemente del path
        Me.Text = Path.GetFileName(fic)
        CurrentMDI.Text = $"{MDIPrincipal.TituloMDI} [{Me.Text}]"
        Application.DoEvents()

        labelTamaño.Text = $"{richTextBoxCodigo.Text.Length:#,##0} car. ({richTextBoxCodigo.Text.CuantasPalabras:#,##0} palab.)"

        Me.TextoModificado = False

        AñadirAUltimosFicherosAbiertos(fic)

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
            CurrentMDI.HabilitarBotones()
        End If

        Me.TextoModificado = False
        labelInfo.Text = "Fichero guardado con " & labelTamaño.Text
        OnProgreso(labelInfo.Text)

    End Sub

    ''' <summary>
    ''' Muestra el cuadro de diálogo de Guardar como.
    ''' </summary>
    Friend Sub GuardarComo()
        Dim fichero = nombreFichero

        Using sFD As New SaveFileDialog
            sFD.Title = "Seleccionar fichero para guardar el código"
            sFD.FileName = fichero
            sFD.InitialDirectory = DirDocumentos
            sFD.RestoreDirectory = True
            sFD.Filter = "Código de Visual Basic y CSharp (*.vb;*.cs)|*.vb;*.cs|Textos (*.txt;*.log;*.md)|*.txt;*.log;*.md|Todos los ficheros (*.*)|*.*"
            If sFD.ShowDialog = DialogResult.Cancel Then
                Return
            End If
            fichero = sFD.FileName
            nombreFichero = sFD.FileName
            ' Guardarlo
            Guardar(fichero)
        End Using
    End Sub

    ''' <summary>
    ''' Guarda el fichero actual (<see cref="nombreFichero"/>).
    ''' Si no tiene nombre muestra el diálogo de guardar como
    ''' </summary>
    Friend Sub Guardar()
        If String.IsNullOrEmpty(nombreFichero) Then
            GuardarComo()
            Return
        End If
        Guardar(nombreFichero)
    End Sub

    Private Sub buttonLenguaje_DropDownItemClicked(sender As Object,
                                                   e As ToolStripItemClickedEventArgs) Handles buttonLenguaje.DropDownItemClicked
        Dim it = e.ClickedItem
        buttonLenguaje.Text = it.Text
        buttonLenguaje.Image = it.Image
    End Sub
End Class
