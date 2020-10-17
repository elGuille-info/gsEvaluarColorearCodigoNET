'------------------------------------------------------------------------------
' Módulo con las opciones compartidas para manipular el editor      (09/Oct/20)
'
' Aquí irán los métodos para manipular las opciones del editor
' se asignará el richTextBoxCodigo que esté activo y el resto de opciones
' usarán ese editor, con idea de tener las opciones de los menús y toolStrip
' en esete módulo en vez de en el Form1.
'
'
' (c) Guillermo (elGuille) Som, 2020
'------------------------------------------------------------------------------
Option Strict On
Option Infer On

Imports Microsoft.VisualBasic
Imports vb = Microsoft.VisualBasic
Imports System
'Imports System.Data
Imports System.Collections.Generic
Imports System.Text
Imports System.Linq
Imports System.IO
Imports System.Reflection
'
Imports Microsoft.CodeAnalysis
'Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.Emit
Imports Microsoft.CodeAnalysis.CSharp
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports Microsoft.CodeAnalysis.Text


Friend Module UtilFormEditor

#Region " Las propiedades para acceder a MDIPrincipal y el Form1 activo "

    ''' <summary>
    ''' La instancia del formulario MDI.
    ''' </summary>
    ''' <returns></returns>
    Public Property CurrentMDI As MDIPrincipal

    ''' <summary>
    ''' La instancia del Form1 activo
    ''' </summary>
    ''' <returns></returns>
    Public Property Form1Activo As Form1

#End Region

#Region " Los RichTextBox y los eventos relacionados "

    ''' <summary>
    ''' El richTextBoxCodigo del Form1 activo
    ''' </summary>
    Public ReadOnly Property richTextBoxCodigo As RichTextBox
        Get
            Return Form1Activo.richTextBoxCodigo
        End Get
    End Property

    ''' <summary>
    ''' El richTextBoxLineas del Form1 activo
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property richTextBoxLineas As RichTextBox
        Get
            Return Form1Activo.richTextBoxLineas
        End Get
    End Property

#End Region

#Region " Definiciones de API para sincronizar los scrollbar de los richtextbox "

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
    Public Function GetScrollPos(hWnd As IntPtr, nBar As Integer) As Integer
    End Function

    <System.Runtime.InteropServices.DllImport("User32.dll")>
    Public Function SendMessage(hWnd As IntPtr, msg As Long, wParam As Long, lParam As Long) As Integer
    End Function

#End Region

#Region " statusBar: Los controles y menú contextual "

    ''' <summary>
    ''' El LabelInfo del MDIPrincipal
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property labelInfo As ToolStripStatusLabel
        Get
            Return Form1Activo.labelInfo
        End Get
    End Property

    ''' <summary>
    ''' El LabelFuente del MDIPrincipal
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property labelFuente As ToolStripStatusLabel
        Get
            Return Form1Activo.labelFuente
        End Get
    End Property

    ''' <summary>
    ''' El LabelModificado del MDIPrincipal
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property labelModificado As ToolStripStatusLabel
        Get
            Return Form1Activo.labelModificado
        End Get
    End Property

    ''' <summary>
    ''' El LabelPos del MDIPrincipal
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property labelPos As ToolStripStatusLabel
        Get
            Return Form1Activo.labelPos
        End Get
    End Property

    ''' <summary>
    ''' El LabelTamaño del MDIPrincipal
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property labelTamaño As ToolStripStatusLabel
        Get
            Return Form1Activo.labelTamaño
        End Get
    End Property

    ''' <summary>
    ''' El ButtonLenguaje del MDIPrincipal
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property buttonLenguaje As ToolStripSplitButton
        Get
            Return Form1Activo.buttonLenguaje
        End Get
    End Property

#End Region

#Region " Los 'campos' generales "

    ''' <summary>
    ''' Si se está inicializando.
    ''' Usado para que no se provoquen eventos en cadena
    ''' </summary>
    Public inicializando As Boolean = True

    ''' <summary>
    ''' Si se deben cambiar los TAB por 8 espacios
    ''' </summary>
    Public CambiarTabs As Boolean = True

    ''' <summary>
    ''' El inicio de la selección al cerrar
    ''' </summary>
    Public selectionStartAnt As Integer

    ''' <summary>
    ''' El final de la selección al cerrar
    ''' </summary>
    Public selectionEndAnt As Integer

    ''' <summary>
    ''' El número de caracteres para indentar
    ''' </summary>
    Public EspaciosIndentar As Integer = 4

    ''' <summary>
    ''' Si se deben mostrar los números de línea en el código coloreado como HTML
    ''' </summary>
    Public Property MostrarLineasHTML As Boolean
        Get
            Return CurrentMDI.chkMostrarLineasHTML.Checked
        End Get
        Set(value As Boolean)
            CurrentMDI.chkMostrarLineasHTML.Checked = value
        End Set
    End Property

    ''' <summary>
    ''' El nombre del fichero de configuración
    ''' </summary>
    Public FicheroConfiguracion As String

    ''' <summary>
    ''' El directorio de documentos
    ''' Valor predeterminado para guardar los ficheros
    ''' </summary>
    Public DirDocumentos As String

    ''' <summary>
    ''' Si se pulsa Ctrl+F o Ctrl+H para indicar 
    ''' los datos a buscar o reemplazar
    ''' </summary>
    Public esCtrlF As Boolean = True

    ''' <summary>
    ''' El nombre de la fuente a usar en el editor
    ''' </summary>
    Public fuenteNombre As String = "Consolas"

    ''' <summary>
    ''' El tamaño de la fuente a usar en el editor
    ''' </summary>
    Public fuenteTamaño As String = "11" 'gsCol.FuenteTamPre

    ''' <summary>
    ''' Si se debe cargar el último fichero abierto 
    ''' cuando se cerró la aplicación
    ''' </summary>
    Public cargarUltimo As Boolean

    ''' <summary>
    ''' Nombre del último fichero asignado al código.
    ''' Se usará el NombreFichero del formulario activo.
    ''' </summary>
    Public Property nombreUltimoFichero As String
        Get
            If Form1Activo IsNot Nothing Then
                Return Form1Activo.nombreFichero
            Else
                Return ""
            End If
        End Get
        Set(value As String)
            If Form1Activo IsNot Nothing Then
                Form1Activo.nombreFichero = value
                CompararString.IgnoreCase = True
                If Not nombresUltimos.Contains(value, New CompararString) Then
                    nombresUltimos.Add(value)
                End If
            End If
        End Set
    End Property

    ''' <summary>
    ''' Los nombres de los ficheros abiertos en cada sesión.
    ''' </summary>
    ''' <remarks>16/Oct/2020</remarks>
    Public nombresUltimos As New List(Of String)

    ''' <summary>
    ''' Si se debe colorear el código al abrir el fichero
    ''' </summary>
    Public colorearAlCargar As Boolean

    ''' <summary>
    ''' Si se debe colorear el código al evaluar
    ''' </summary>
    Public Property colorearAlEvaluar As Boolean
        Get
            Return CurrentMDI.buttonColorearAlEvaluar.Checked
        End Get
        Set(value As Boolean)
            CurrentMDI.buttonColorearAlEvaluar.Checked = value
        End Set
    End Property

    ''' <summary>
    ''' Si se deben comprobar los errores al evaluar.
    ''' Si es <see cref="True"/> se pre-compila el código.
    ''' </summary>
    Public Property compilarAlEvaluar As Boolean
        Get
            Return CurrentMDI.buttonCompilarAlEvaluar.Checked
        End Get
        Set(value As Boolean)
            CurrentMDI.buttonCompilarAlEvaluar.Checked = value
        End Set
    End Property

    '''' <summary>
    '''' El nuevo código del editor
    '''' </summary>
    'Public codigoNuevo As String

    '''' <summary>
    '''' El código anterior del editor.
    '''' Usado para comparar si ha habido cambios
    '''' </summary>
    'Public codigoAnterior As String

    '''' <summary>
    '''' Indica si se ha modificado el texto.
    '''' Cuando cambia el texto actual (<see cref="codigoNuevo"/>)
    '''' se comprueba con <see cref="codigoAnterior"/> por si hay cambios.
    '''' </summary>
    'Public Property TextoModificado As Boolean
    '    Get
    '        Return richTextBoxCodigo.Modified
    '    End Get
    '    Set(value As Boolean)
    '        If value Then
    '            labelModificado.Text = "*"
    '        Else
    '            labelModificado.Text = " "
    '        End If
    '        richTextBoxCodigo.Modified = value
    '    End Set
    'End Property

#End Region

#Region " Guardar y Leer la configuración "

    ''' <summary>
    ''' Guardar los datos de configuración
    ''' </summary>
    Public Sub GuardarConfig()
        Dim cfg = New Config(FicheroConfiguracion)
        Dim cuantos = 0

        ' Si cargarUltimo es falso no guardar el último fichero     (16/Sep/20)
        cfg.SetValue("Ficheros", "CargarUltimo", cargarUltimo)
        If cargarUltimo Then
            cfg.SetValue("Ficheros", "Ultimo", nombreUltimoFichero)
        Else
            cfg.SetValue("Ficheros", "Ultimo", "")
        End If

        ' Leer los últimos ficheros abiertos                        (16/Oct/20)
        cuantos = nombresUltimos.Count
        cfg.SetKeyValue("Ficheros", "Ultimos-cuantos", cuantos)
        For i = 0 To cuantos - 1
            Dim s = nombresUltimos(i)
            cfg.SetKeyValue("Ficheros", $"Ultimos-Ficheros{i}", s)
        Next

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
        cuantos = comboBoxFileName.Items.Count
        cfg.SetValue("Ficheros", "Count", cuantos)
        For i = 0 To comboBoxFileName.Items.Count - 1 ' To 0 Step -1
            ' No añadir el path si es el directorio de Documentos
            Dim s = comboBoxFileName.Items(i).ToString
            If Path.GetDirectoryName(s) = DirDocumentos Then
                s = Path.GetFileName(s)
            End If
            cfg.SetKeyValue("Ficheros", $"Fichero{j}", s)
            j += 1
            If j = MaxFicsConfig Then Exit For
        Next

        ' El nombre y tamaño de la fuente                           (11/Sep/20)
        cfg.SetValue("Fuente", "Nombre", fuenteNombre)
        cfg.SetValue("Fuente", "Tamaño", fuenteTamaño)

        '' El tamaño y la posición de la ventana
        'cfg.SetValue("Ventana", "Left", tamForm.L)
        'cfg.SetValue("Ventana", "Top", tamForm.T)
        'cfg.SetValue("Ventana", "Height", tamForm.H)
        'cfg.SetValue("Ventana", "Width", tamForm.W)

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
    Public Sub LeerConfig()
        Dim cfg = New Config(FicheroConfiguracion)
        Dim cuantos = 0

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

        cuantos = cfg.GetValue("Ficheros", "Count", 0)
        comboBoxFileName.Items.Clear()

        For i = 0 To MaxFicsConfig - 1
            If i >= cuantos Then Exit For
            Dim s = cfg.GetValue("Ficheros", $"Fichero{i}", "-1")
            If s = "-1" Then Exit For
            ' No añadir el path si es el directorio de Documentos
            If Path.GetDirectoryName(s) = DirDocumentos Then
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
        richTextBoxLineas.Font = New Font(fuenteNombre, CSng(fuenteTamaño))
        ' No cambiar la fuente del panel de sintaxis                (03/Oct/20)
        'richTextBoxSyntax.Font = New Font(fuenteNombre, CSng(fuenteTamaño))

        'If Not MDIPrincipal.PrimerForm1 Is Me Then
        '    ' El tamaño y la posición de la ventana
        '    tamForm.L = cfg.GetValue("Ventana", "Left", -1)
        '    tamForm.T = cfg.GetValue("Ventana", "Top", -1)
        '    tamForm.H = cfg.GetValue("Ventana", "Height", -1)
        '    tamForm.W = cfg.GetValue("Ventana", "Width", -1)

        '    If tamForm.L <> -1 Then Me.Left = tamForm.L
        '    If tamForm.T <> -1 Then Me.Top = tamForm.T
        '    If tamForm.H > -1 Then Me.Height = tamForm.H
        '    If tamForm.W > -1 Then Me.Width = tamForm.W
        'End If

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

        ' Leer los últimos ficheros abiertos                        (16/Oct/20)
        cuantos = cfg.GetValue("Ficheros", "Ultimos-cuantos", 0)
        nombresUltimos.Clear()
        For i = 0 To cuantos - 1
            Dim s = cfg.GetValue("Ficheros", $"Ultimos-Ficheros{i}", "-1")
            If s = "-1" Then Exit For
            nombresUltimos.Add(s)
        Next
        ' Si se indica cargarUltimo, abrirlos todos                 (16/Oct/20)
        If cargarUltimo Then
            cuantos = nombresUltimos.Count ' + 1
            ' En Abrir se llama a OnProgreso una vez
            ' en ColorearCodigo 2 veces
            If colorearAlCargar Then cuantos = nombresUltimos.Count * 3
            MostrarProcesando($"Cargando {nombresUltimos.Count} ficheros", "Cargando los ficheros...", cuantos)
            For i = 0 To nombresUltimos.Count - 1
                m_fProcesando.Text = $"Cargando {i + 1} de {nombresUltimos.Count} ficheros"
                If String.IsNullOrWhiteSpace(nombresUltimos(i)) Then Continue For
                m_fProcesando.Mensaje1 = $"Cargando {nombresUltimos(i)}...{vbCrLf}"
                Application.DoEvents()
                If m_fProcesando.Cancelar Then Exit For
                If i = nombresUltimos.Count - 1 Then
                    m_fProcesando.ValorActual = m_fProcesando.Maximo - 1
                End If
                Nuevo()
                Form1Activo.nombreFichero = nombresUltimos(i)
                Abrir(Form1Activo.nombreFichero)
                Application.DoEvents()
            Next
            m_fProcesando.Close()
        End If

    End Sub

#End Region

#Region " Para la ventana de progreso "

    Friend m_fProcesando As New fProcesando

    ''' <summary>
    ''' Mostrar la ventana de procesando (mientras cargan los ficheros, etc.)
    ''' </summary>
    ''' <param name="titulo">El título de la ventana.</param>
    ''' <param name="msg">El mensaje a mostrar.</param>
    ''' <param name="max">El valor máximo de la barra de progreso.</param>
    ''' <remarks>16/Oct/2020</remarks>
    Public Sub MostrarProcesando(titulo As String, msg As String, max As Integer)
        'm_fProcesando = New fProcesando
        m_fProcesando.Texto = titulo
        m_fProcesando.Mensaje = msg
        m_fProcesando.Mensaje1 = m_fProcesando.Mensaje
        m_fProcesando.Maximo = max
        m_fProcesando.ValorActual = 0
        m_fProcesando.ShowInTaskbar = False
        m_fProcesando.Show()
        m_fProcesando.BringToFront()
        ' Que siempre esté encima del resto.
        m_fProcesando.TopMost = True
        Application.DoEvents()

    End Sub

    ''' <summary>
    ''' Asigna a la ventana de progreso el texto a mostrar en la segunda línea,
    ''' el de la primera línea se habrá asignado antes.
    ''' </summary>
    ''' <param name="msg">El mensaje a añadir a la ventana de progreso.</param>
    ''' <returns>El porcentaje que queda</returns>
    ''' <remarks>16/Oct/2020</remarks>
    Public Function OnProgreso(msg As String) As Integer
        m_fProcesando.Mensaje = m_fProcesando.Mensaje1.Replace("...", ".") & msg
        m_fProcesando.ValorActual += 1
        Application.DoEvents()
        ' Devolver el porcentaje que resta                          (17/Oct/20)
        Return m_fProcesando.PorcentajeRestante
    End Function

#End Region

#Region " Para los marcadores y bookmarks "

    '
    ' Para los marcadores / Bookmarks                               (28/Sep/20)
    '

    ''' <summary>
    ''' Colección con los marcadores del código que se está editando.
    ''' </summary>
    Public ColMarcadores As New List(Of Integer)

    ''' <summary>
    ''' El fichero usado con los marcadores.
    ''' </summary>
    Public MarcadorFichero As String

    ''' <summary>
    ''' Poner los marcadores, si hay... (30/Sep/20)
    ''' </summary>
    Public Sub PonerLosMarcadores()
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
    Public Sub MarcadorPonerQuitar()
        Dim posActual = PosicionActual()
        If ColMarcadores.Contains(posActual.Inicio) Then
            ' quitarlo
            ColMarcadores.Remove(posActual.Inicio)

            richTextBoxCodigo.SelectionStart = If(posActual.Inicio - 1 < 0, 0, posActual.Inicio - 1) ' posActual.Inicio - 1
            richTextBoxCodigo.SelectionLength = 0
            'richTextBoxCodigo.SelectionBullet = False
            Dim fcol = richTextBoxLineas.GetFirstCharIndexFromLine(posActual.Linea - 1)
            richTextBoxLineas.SelectionStart = fcol
            richTextBoxLineas.SelectionLength = 5
            richTextBoxLineas.SelectionBullet = False
            ' así es como se pone en AñadirNumerosDeLinea
            richTextBoxLineas.SelectedText = $" {(posActual.Linea).ToString("0").PadLeft(4)}"
        Else
            ColMarcadores.Add(posActual.Inicio)
            ' Poner los marcadores en richTextBoxLineas
            'richTextBoxLineas.SelectionBullet = True
            Dim fcol = richTextBoxLineas.GetFirstCharIndexFromLine(posActual.Linea - 1)
            richTextBoxLineas.SelectionStart = fcol
            richTextBoxLineas.SelectionLength = 5
            'richTextBoxLineas.SelectionBullet = True

            ' Poner delante la imagen del marcador
            ' Usando la imagen bookmark_003_8x10.png
            richTextBoxLineas.SelectedRtf = $"{picBookmark}{(posActual.Linea).ToString("0").PadLeft(4)}" & "}"
            'richTextBoxLineas.SelectedText = $"*{(posActual.Linea).ToString("0").PadLeft(4)}"

            'richTextBoxCodigo.SelectionBullet = True
            richTextBoxCodigo.SelectionStart = If(posActual.Inicio - 1 < 0, 0, posActual.Inicio - 1)
            richTextBoxCodigo.SelectionLength = 0
        End If
        ColMarcadores.Sort()
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
    Public Sub MarcadorAnterior()
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
    Public Sub MarcadorSiguiente()
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
    Public Sub MarcadorQuitarTodos()
        If ColMarcadores.Count = 0 Then Return

        ColMarcadores.Clear()
        AñadirNumerosDeLinea()

        MarcadorFichero = ""
    End Sub

    ''' <summary>
    ''' Añadir los números de línea
    ''' </summary>
    ''' <remarks>Como método separado 18/Sep/20</remarks>
    Public Sub AñadirNumerosDeLinea()
        If inicializando Then Return
        If String.IsNullOrEmpty(richTextBoxCodigo.Text) Then Return

        Dim lineas = richTextBoxCodigo.Lines.Length
        richTextBoxLineas.Text = ""
        For i = 1 To lineas
            richTextBoxLineas.Text &= $" {i.ToString("0").PadLeft(4)}{vbCr}"
        Next
        ' Sincronizar los scrolls
        Form1Activo.richTextBoxCodigo_VScroll(Nothing, Nothing)

        PonerLosMarcadores()
    End Sub

#End Region

#Region " Métodos para la posición actual "

    ''' <summary>
    ''' Averigua la línea, columna (y primer caracter de la línea) de la posición actual en richTextBoxCodigo.
    ''' </summary>
    ''' <returns>Una tupla con la Fila, Columna y la posición del primer caracter de la línea (Inicio)</returns>
    Public Function PosicionActual() As (Linea As Integer, Columna As Integer, Inicio As Integer)
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
    Public Sub MostrarPosicion(e As KeyEventArgs)
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


#End Region

#Region " Para la indentación "

    '
    ' Para la indentación                                           (28/Sep/20)
    '

    ''' <summary>
    ''' Poner indentación
    ''' </summary>
    ''' <param name="rtEditor"></param>
    Public Sub PonerIndentacion(rtEditor As RichTextBox)
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
    Public Sub QuitarIndentacion(rtEditor As RichTextBox)
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

#End Region

#Region " Para los comentarios "

    '
    ' Para los comentarios                                          (28/Sep/20)
    '

    ''' <summary>
    ''' Poner comentarios a las líneas seleccionadas,
    ''' si no hay seleccionada, se pone donde esté el cursor
    ''' </summary>
    ''' <param name="rtEditor"></param>
    Public Sub PonerComentarios(rtEditor As RichTextBox)
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
    Public Sub QuitarComentarios(rtEditor As RichTextBox)
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

#End Region

#Region " Colorear: Los botones, menús y métodos "

    'Friend buttonColorear As ToolStripButton = CurrentMDI.buttonColorear
    'Friend buttonColorearHTML As ToolStripButton = CurrentMDI.buttonColorearHTML
    'Friend buttonNoColorear As ToolStripButton = CurrentMDI.buttonNoColorear

    'Friend menuColorear As ToolStripMenuItem = CurrentMDI.menuColorear
    'Friend menuColorearHTML As ToolStripMenuItem = CurrentMDI.menuColorearHTML
    'Friend menuNoColorear As ToolStripMenuItem = CurrentMDI.menuNoColorear


    ''' <summary>
    ''' Colorear la selección actual del código
    ''' </summary>
    Public Sub ColorearSeleccion()
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

    ''' <summary>
    ''' Mostrar el código sin colorear
    ''' </summary>
    Public Sub NoColorear()
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
    Public Sub ColorearCodigo()
        If richTextBoxCodigo.TextLength = 0 Then Return

        ' Si buttonLenguaje.Text es Texto, no colorear              (08/Oct/20)
        If buttonLenguaje.Text = ExtensionTexto Then Return

        labelInfo.Text = $"Coloreando el código de {buttonLenguaje.Text}..."
        CurrentMDI.Refresh()
        OnProgreso(labelInfo.Text)

        Dim modif = richTextBoxCodigo.Modified
        inicializando = True

        Dim colCodigo = Compilar.ColoreaRichTextBox(richTextBoxCodigo, buttonLenguaje.Text)

        MostrarResultadoEvaluar(colCodigo, False)

        inicializando = False
        richTextBoxCodigo.SelectionStart = 0
        richTextBoxCodigo.SelectionLength = 0
        richTextBoxCodigo.Modified = modif

        labelInfo.Text = $"Código coloreado para {buttonLenguaje.Text}."
        OnProgreso(labelInfo.Text)
        CurrentMDI.Refresh()
    End Sub

    ''' <summary>
    ''' Colorear el código en formato HTML y mostrarlo en ventana aparte
    ''' </summary>
    Public Sub ColorearHTML()
        If richTextBoxCodigo.TextLength = 0 Then Return

        labelInfo.Text = $"Coloreando en HTML para {buttonLenguaje.Text}..."
        CurrentMDI.Refresh()

        Dim sHTML = Compilar.ColoreaHTML(richTextBoxCodigo.Text,
                                         buttonLenguaje.Text,
                                         mostrarLineas:=chkMostrarLineasHTML.Checked)

        labelInfo.Text = "Mostrando el visor de HTML..."
        CurrentMDI.Refresh()

        Dim fHTML As New FormVisorHTML(comboBoxFileName.Text, sHTML)
        fHTML.Show()

        labelInfo.Text = ""
    End Sub

#End Region

#Region " Compilar, evaluar: Los botones, menús y métodos "

    'Friend buttonCompilar As ToolStripButton = CurrentMDI.buttonCompilar
    'Friend buttonEjecutar As ToolStripButton = CurrentMDI.buttonEjecutar
    'Friend buttonEvaluar As ToolStripButton = CurrentMDI.buttonEvaluar

    'Friend menuEvaluar As ToolStripMenuItem = CurrentMDI.menuEvaluar
    'Friend menuEjecutar As ToolStripMenuItem = CurrentMDI.menuEjecutar
    'Friend menuCompilar As ToolStripMenuItem = CurrentMDI.menuCompilar

    'Friend menuOcultarEvaluar As ToolStripMenuItem = CurrentMDI.menuOcultarEvaluar


    '
    ' Los métodos de compilar, ejecutar, evaluar
    '

    ''' <summary>
    ''' Compila y ejecuta el código actual.
    ''' Si se producen errores muestra la información con los errores.
    ''' </summary>
    Public Sub CompilarEjecutar()
        If richTextBoxCodigo.TextLength = 0 Then Return

        ' guardar el código antes de compilar y ejecutar
        labelInfo.Text = "Compilando el código..."
        CurrentMDI.Refresh()

        ' Guardar (con el nombre del formulario activo)             (16/Oct/20)
        If Form1Activo.TextoModificado Then
            Guardar()
        End If

        Dim fichero = nombreUltimoFichero

        ' Si no tiene el path, añadirlo                             (27/Sep/20)
        If String.IsNullOrEmpty(Path.GetDirectoryName(fichero)) Then
            fichero = Path.Combine(DirDocumentos, fichero)
        End If

        Dim res = Compilar.CompileRun(fichero, run:=True)
        If res.Result Is Nothing Then
            Form1Activo.lstSyntax.Items.Clear()
            Dim s = "Error al compilar, seguramente el fichero está siendo usado por otro proceso."
            Form1Activo.lstSyntax.Items.Add(s)
            labelInfo.Text = s
            Form1Activo.splitContainer1.Panel2.Visible = True
            Form1Activo.splitContainer1_Resize(Nothing, Nothing)

            Return
        End If
        If res.Result.Success Then
            labelInfo.Text = $"Se ha compilado y ejecutado satisfactoriamente."

            Form1Activo.lstSyntax.Items.Clear()
            Form1Activo.splitContainer1.Panel2.Visible = False
            Form1Activo.splitContainer1_Resize(Nothing, Nothing)
        Else
            ' Mostrar los errores
            MostrarErrores(res.Result)
        End If

    End Sub

    ''' <summary>
    ''' Compilar el código sin ejecutar
    ''' </summary>
    Public Sub Build()
        If richTextBoxCodigo.TextLength = 0 Then Return

        If Form1Activo.TextoModificado Then
            Guardar()
        End If

        Dim filepath = nombreUltimoFichero
        labelInfo.Text = $"Compilando para {buttonLenguaje.Text}..."
        CurrentMDI.Refresh()

        ' Si no tiene el path, añadirlo                             (27/Sep/20)
        If String.IsNullOrEmpty(Path.GetDirectoryName(filepath)) Then
            filepath = Path.Combine(DirDocumentos, filepath)
        End If

        Dim res = Compilar.CompileRun(filepath, run:=False)

        Form1Activo.lstSyntax.Items.Clear()

        If res.Result Is Nothing Then
            Dim s = "Error al compilar, seguramente el fichero está siendo usado por otro proceso."
            Form1Activo.lstSyntax.Items.Add(s)
            labelInfo.Text = s
            Form1Activo.splitContainer1.Panel2.Visible = True
            Form1Activo.splitContainer1_Resize(Nothing, Nothing)

            Return
        End If

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
    Public Sub Evaluar()
        If richTextBoxCodigo.TextLength = 0 Then Return

        If Form1Activo.TextoModificado Then
            Guardar()
        End If

        Dim res As EmitResult
        Form1Activo.lstSyntax.Items.Clear()

        If compilarAlEvaluar Then
            Form1Activo.splitContainer1.Panel2.Visible = False
            Form1Activo.splitContainer1_Resize(Nothing, Nothing)
            Form1Activo.lstSyntax.Items.Clear()

            labelInfo.Text = $"Compilando para {buttonLenguaje.Text}..."
            CurrentMDI.Refresh()

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
            CurrentMDI.Refresh()

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
    Public Sub MostrarErrores(res As EmitResult)
        Dim errors = 0
        Dim warnings = 0

        Form1Activo.lstSyntax.Items.Clear()
        ' Ponerlo en ownerdraw para que coloree
        Form1Activo.lstSyntax.DrawMode = DrawMode.OwnerDrawFixed

        For Each diag As Diagnostic In res.Diagnostics
            Dim dcsI As New DiagInfo(diag)
            Form1Activo.lstSyntax.Items.Add(dcsI)

            If diag.Severity = DiagnosticSeverity.Error Then
                errors += 1
            ElseIf diag.Severity = DiagnosticSeverity.Warning Then
                warnings += 1
            End If
        Next

        labelInfo.Text = $"Compilado con {errors} error{If(errors = 1, "", "es")} y {warnings} advertencia{If(warnings = 1, "", "s")}."

        Form1Activo.splitContainer1.Panel2.Visible = True
        Form1Activo.splitContainer1_Resize(Nothing, Nothing)
    End Sub

    ''' <summary>
    ''' Mostrar el resultado de evaluar el código.
    ''' Las clases, métodos, palabras clave, etc. definidos en el código.
    ''' </summary>
    ''' <param name="colCodigo">Colección del tipo <see cref="Dictionary(Of String, List(Of DiagClassifSpanInfo))"/>
    ''' con la lista con los valores sacados de <see cref="ClassifiedSpan"/> </param>
    Public Sub MostrarResultadoEvaluar(colCodigo As Dictionary(Of String, Dictionary(Of String, ClassifSpanInfo)),
                                       mostrarSyntax As Boolean)
        Dim codTiposCount As (Clases As Integer, Metodos As Integer, Keywords As Integer)
        codTiposCount.Clases = If(colCodigo.Keys.Contains("class name"), colCodigo("class name").Count, 0) + If(colCodigo.Keys.Contains("module name"), colCodigo("module name").Count, 0)
        codTiposCount.Metodos = If(colCodigo.Keys.Contains("method name"), colCodigo("method name").Count, 0)
        codTiposCount.Keywords = If(colCodigo.Keys.Contains("keyword"), colCodigo("keyword").Count, 0) + If(colCodigo.Keys.Contains("keyword - control"), colCodigo("keyword - control").Count, 0)

        labelInfo.Text = $"Código con {codTiposCount.Clases} clase{If(codTiposCount.Clases = 1, "", "s")}, {codTiposCount.Metodos} método{If(codTiposCount.Metodos = 1, "", "s")} y {codTiposCount.Keywords} palabra{If(codTiposCount.Keywords = 1, "", "s")} clave."

        Form1Activo.lstSyntax.Items.Clear()
        ' Ponerlo en ownerdraw para que coloree
        Form1Activo.lstSyntax.DrawMode = DrawMode.OwnerDrawFixed

        ' Clasificar las claves                                     (25/Sep/20)
        Dim claves = From kv In colCodigo Order By kv.Key Ascending Select kv

        For Each kv In claves
            Dim s1 = kv.Key

            ' No mostrar el contenido de estos símbolos
            If s1 = "comment" OrElse s1.StartsWith("string") OrElse
                s1 = "punctuation" OrElse s1.StartsWith("operator") OrElse
                s1 = "number" OrElse s1.StartsWith("xml") Then
            Else
                Form1Activo.lstSyntax.Items.Add(s1)
                Dim claves2 = From kv0 In kv.Value Order By kv0.Key Ascending Select kv0
                For Each kv1 In claves2
                    Form1Activo.lstSyntax.Items.Add(kv1.Value)
                Next
            End If
        Next

        If mostrarSyntax Then
            Form1Activo.splitContainer1.Panel2.Visible = True
            Form1Activo.splitContainer1_Resize(Nothing, Nothing)
        End If
    End Sub

#End Region

#Region " Clasificar "

    '
    ' Clasificar
    '

    ''' <summary>
    ''' Clasificar el texto seleccionado
    ''' </summary>
    Public Sub ClasificarSeleccion()
        If richTextBoxCodigo.SelectionLength = 0 Then Return

        Dim selStart = richTextBoxCodigo.SelectionStart
        Dim lineas() As String = richTextBoxCodigo.SelectedText.Split(vbCr.ToCharArray)
        Dim sb As New System.Text.StringBuilder
        Dim j = lineas.Length - 1
        If String.IsNullOrWhiteSpace(lineas(j)) Then j -= 1
        Dim k = 0

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

#End Region

#Region " Cambiar las mayúsculas y minúsculas "

    'Friend menuMayúsculas As ToolStripMenuItem = CurrentMDI.menuMayúsculas
    'Friend menuMinúsculas As ToolStripMenuItem = CurrentMDI.menuMinúsculas
    'Friend menuTitulo As ToolStripMenuItem = CurrentMDI.menuTitulo

    'Friend menuPrimeraMinúsculas As ToolStripMenuItem = CurrentMDI.menuPrimeraMinúsculas

    '
    ' Cambiar las mayúsculas y minúsculas
    '

    ''' <summary>
    ''' Cambiar las mayúsculas y minúsculas según el valor del tipo <see cref="CasingValues"/>.
    ''' </summary>
    ''' <param name="queCase"></param>
    Public Sub CambiarMayúsculasMinúsculas(queCase As CasingValues)
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

#End Region

#Region " Quitar espacios de delante y destrás "

    'Friend menuQuitarEspaciosTodos As ToolStripMenuItem = CurrentMDI.menuQuitarEspaciosTodos
    'Friend menuQuitarEspaciosTrim As ToolStripMenuItem = CurrentMDI.menuQuitarEspaciosTrim
    'Friend menuQuitarEspaciosTrimEnd As ToolStripMenuItem = CurrentMDI.menuQuitarEspaciosTrimEnd
    'Friend menuQuitarEspaciosTrimStart As ToolStripMenuItem = CurrentMDI.menuQuitarEspaciosTrimStart


    '
    ' Quitar los espacios
    '

    ''' <summary>
    ''' Quitar los espacios de delante, detrás o ambos
    ''' </summary>
    ''' <param name="delante"></param>
    ''' <param name="detras"></param>
    Public Sub QuitarEspacios(Optional delante As Boolean = True,
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

#End Region

#Region " Poner y quitar texto del final "

    ''' <summary>
    ''' Poner el texto indicado en txtPonerTexto al final de cada línea seleccionada.
    ''' </summary>
    Public Sub PonerTextoAlFinal()
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

    ''' <summary>
    ''' Quitar el texto indicado en txtPonerTexto del final de cada línea seleccionada.
    ''' </summary>
    Public Sub QuitarTextoDelFinal()
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

#End Region

#Region " Para los recortes (lo copiado en edición) "

    '
    ' Para los recortes
    '

    ''' <summary>
    ''' El número máximo de recortes.
    ''' </summary>
    Public Const MaxRecortes As Integer = 10

    ''' <summary>
    ''' Colección para los últimos textos copiados o cortados en la aplicación.
    ''' Guardar solo <see cref="MaxRecortes"/> recortes.
    ''' </summary>
    Public Property ColRecortes As New List(Of String)

    ''' <summary>
    ''' Añadir la cadena indicada a la colección de recortes.
    ''' </summary>
    ''' <param name="str">La cadena a añadir.</param>
    Public Sub AñadirRecorte(str As String)
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
    Public Sub MostrarRecortes()
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

#End Region

#Region " Ficheros: métodos, menús y botones "

    '
    ' Los menús de ficheros
    '

    'Friend menuFile As ToolStripMenuItem = CurrentMDI.menuFile
    'Friend menuFileAcercaDe As ToolStripMenuItem = CurrentMDI.menuFileAcercaDe
    'Friend menuFileGuardar As ToolStripMenuItem = CurrentMDI.menuFileGuardar
    'Friend menuFileGuardarComo As ToolStripMenuItem = CurrentMDI.menuFileGuardarComo
    'Friend menuFileNuevo As ToolStripMenuItem = CurrentMDI.menuFileNuevo
    'Friend menuFileRecargar As ToolStripMenuItem = CurrentMDI.menuFileRecargar
    'Friend menuFileRecientes As ToolStripMenuItem = CurrentMDI.menuFileRecientes
    'Friend menuFileSalir As ToolStripMenuItem = CurrentMDI.menuFileSalir
    'Friend menuFileSeleccionarAbrir As ToolStripMenuItem = CurrentMDI.menuFileSeleccionarAbrir

    '
    ' Los botones de la barra de herramientas
    '

    'Friend buttonAbrir As ToolStripButton = CurrentMDI.buttonAbrir
    'Friend buttonGuardar As ToolStripButton = CurrentMDI.buttonGuardar
    'Friend buttonGuardarComo As ToolStripButton = CurrentMDI.buttonGuardarComo
    'Friend buttonNuevo As ToolStripButton = CurrentMDI.buttonNuevo
    'Friend buttonRecargar As ToolStripButton = CurrentMDI.buttonRecargar

    Public ReadOnly Property comboBoxFileName As ToolStripComboBox
        Get
            Return CurrentMDI.comboBoxFileName
        End Get
    End Property


    ''' <summary>
    ''' Para saber que no se debe colorear, ni compilar, etc.
    ''' </summary>
    Public Const ExtensionTexto As String = "Texto"

    ''' <summary>
    ''' El número máximo de ficheros en el menú recientes
    ''' </summary>
    Public Const MaxFicsMenu As Integer = 9

    ''' <summary>
    ''' El número máximo de ficheros a guardar en la configuración
    ''' y a mostrar en el combo de los nombres de los ficheros.
    ''' </summary>
    Public Const MaxFicsConfig As Integer = 50

    ''' <summary>
    ''' Muestra la ventana informativa sobre esta utilidad
    ''' y las DLL que utiliza
    ''' </summary>
    Public Sub AcercaDe()
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
    ' Los métodos de ficheros
    '

    ''' <summary>
    ''' Asigna los ficheros al menú recientes
    ''' </summary>
    Public Sub AsignarRecientes()
        CurrentMDI.menuFileRecientes.DropDownItems.Clear()
        For i = 0 To MaxFicsMenu - 1
            ' Salir si en el combo hay menos ficheros que el contador actual
            ' el bucle lo hace hasta un máximo de MaxFicsMenu (10)
            If (i + 1) > comboBoxFileName.Items.Count Then Exit For

            Dim s = $"{i + 1} - {comboBoxFileName.Items(i).ToString}"
            ' Crear el menú y asignar el método de evento para Click
            Dim m = New ToolStripMenuItem(s)
            AddHandler m.Click, Sub() AbrirReciente(m.Text)
            CurrentMDI.menuFileRecientes.DropDownItems.Add(m)
        Next
    End Sub

    ''' <summary>
    ''' Abre el fichero seleccionado del menú recientes
    ''' </summary>
    ''' <param name="fic">Path completo del fichero a abrir</param>
    Public Sub AbrirReciente(fic As String)
        ' El nombre está después del guión -
        ' posición máxima será 4: 1 - Nombre
        ' pero si hay 2 cifras, será 5: 10 - Nombre
        ' Tomando a partir del la 4ª posición está bien
        If String.IsNullOrEmpty(fic) Then Return

        Dim ficT = fic.Substring(4).Trim()

        ' Comprobar si ya está abierto,                             (17/Oct/20)
        ' si es así, mostrar la ventana
        ' si no está abierto, abrirlo en una ventana nueva
        For Each frm As Form1 In CurrentMDI.MdiChildren
            If frm.nombreFichero = ficT Then
                frm.BringToFront()
                Return
            End If
        Next
        ' Si llega aquí es que no estaba abierto
        Nuevo()
        Form1Activo.nombreFichero = ficT
        Abrir(Form1Activo.nombreFichero)


        '' Si es el que está abierto, salir
        '' Solo si el texto no está modificado                       (14/Sep/20)
        '' por si se quiere re-abrir
        'If ficT = nombreUltimoFichero AndAlso Form1Activo.TextoModificado = False Then Return

        '' Si está modificado, preguntar si se quiere guardar
        'If Form1Activo.TextoModificado Then
        '    GuardarComo()
        'End If

        '' pero no abrir ese...
        '' en nombreUltimoFichero se asigna el seleccionado del menú recientes
        'nombreUltimoFichero = ficT
        'Abrir(nombreUltimoFichero)
    End Sub


    ''' <summary>
    ''' Muestra el cuadro de diálogos de abrir ficheros
    ''' </summary>
    Public Sub Abrir()
        Using oFD As New OpenFileDialog
            oFD.Title = "Seleccionar fichero de código a abrir"
            Dim sFN = comboBoxFileName.Text
            sFN = If(String.IsNullOrEmpty(sFN),
                Path.Combine(DirDocumentos, $"prueba.{If(buttonLenguaje.Text = Compilar.LenguajeVisualBasic, ".vb", ".cs")}"),
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
    Public Sub Recargar()
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
    Public Sub Abrir(fic As String)
        If String.IsNullOrWhiteSpace(fic) Then
            Return
        End If

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

        labelInfo.Text = $"Abriendo {fic}..."
        OnProgreso(labelInfo.Text)

        ' solo ocultarlo si no es el mismo fichero                  (25/Sep/20)
        If fic <> Form1Activo.nombreFichero Then
            Form1Activo.splitContainer1.Panel2.Visible = False
            Form1Activo.splitContainer1_Resize(Nothing, Nothing)
            'richTextBoxSyntax.Text = ""
            Form1Activo.lstSyntax.Items.Clear()
        End If

        Dim sCodigo = ""
        Using sr As New StreamReader(fic, detectEncodingFromByteOrderMarks:=True, encoding:=Encoding.UTF8)
            sCodigo = sr.ReadToEnd()
        End Using

        ' Si se deben cambiar los TAB por 8 espacios                (05/Oct/20)
        If CambiarTabs Then
            sCodigo = sCodigo.Replace(vbTab, "        ")
        End If

        Form1Activo.codigoAnterior = sCodigo

        richTextBoxCodigo.Text = sCodigo
        'richTextBoxCodigo.Modified = False
        richTextBoxCodigo.SelectionStart = 0
        richTextBoxCodigo.SelectionLength = 0
        MostrarPosicion(Nothing)

        ' El nombre del fichero con el path completo                (17/Oct/20)
        Form1Activo.nombreFichero = fic
        Form1Activo.Text = fic

        ' Comprobar si hay que añadir el fichero a la lista de recientes
        sDirFic = Path.GetDirectoryName(fic)
        If sDirFic = DirDocumentos Then
            fic = Path.GetFileName(fic)
        End If
        nombreUltimoFichero = fic
        comboBoxFileName.Text = fic

        AñadirAlComboBoxFileName(fic)

        If MarcadorFichero <> fic Then
            MarcadorQuitarTodos()
            MarcadorFichero = fic
        End If

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
        'Form1Activo.Text = fic ' labelInfo.Text
        Application.DoEvents()
        CurrentMDI.Text = $"{MDIPrincipal.TituloMDI} [{Form1Activo.Text}]"

        'labelTamaño.Text = $"{richTextBoxCodigo.Text.Length:#,##0} car. ({richTextBoxCodigo.Text.CuantasPalabras:#,##0} palab.)"

        ' Si hay que colorear el fichero cargado
        If colorearAlCargar Then
            ColorearCodigo()
        End If

        labelTamaño.Text = $"{richTextBoxCodigo.Text.Length:#,##0} car. ({richTextBoxCodigo.Text.CuantasPalabras:#,##0} palab.)"

        ' Si es texto, deshabilitar los botones que correspondan    (08/Oct/20)
        If sLenguaje = ExtensionTexto Then
            CurrentMDI.HabilitarBotones()
        End If

        Form1Activo.TextoModificado = False

    End Sub

    ''' <summary>
    ''' Guarda el fichero indicado en el parámetro
    ''' </summary>
    ''' <param name="fic">El path completo del fichero a guardar</param>
    Public Sub Guardar(fic As String)
        labelInfo.Text = $"Guardando {fic}..."
        OnProgreso(labelInfo.Text)

        Dim sCodigo = richTextBoxCodigo.Text

        Dim sDirFic = Path.GetDirectoryName(fic)
        If String.IsNullOrWhiteSpace(sDirFic) Then
            fic = Path.Combine(DirDocumentos, fic)
        End If

        ' Si se deben cambiar los TAB por 8 espacios                (05/Oct/20)
        If CambiarTabs Then
            sCodigo = sCodigo.Replace(vbTab, "        ")
        End If

        Using sw As New StreamWriter(fic, append:=False, encoding:=Encoding.UTF8)
            sw.WriteLine(sCodigo)
        End Using
        Form1Activo.codigoAnterior = sCodigo

        labelInfo.Text = $"{Path.GetFileName(fic)} ({Path.GetDirectoryName(fic)})"
        Form1Activo.nombreFichero = fic ' labelInfo.Text

        Form1Activo.Text = fic ' labelInfo.Text
        Application.DoEvents()
        CurrentMDI.Text = $"{MDIPrincipal.TituloMDI} [{Form1Activo.Text}]"

        'labelTamaño.Text = $"{richTextBoxCodigo.Text.Length:#,##0} car."
        labelTamaño.Text = $"{richTextBoxCodigo.Text.Length:#,##0} car. ({richTextBoxCodigo.Text.CuantasPalabras:#,##0} palab.)"

        Form1Activo.TextoModificado = False

        'richTextBoxCodigo.Modified = False
        'labelModificado.Text = " "
        Dim fic2 = fic
        If Path.GetDirectoryName(fic) = DirDocumentos Then
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
            CurrentMDI.HabilitarBotones()
        End If

        Form1Activo.TextoModificado = False
        labelInfo.Text = "Fichero guardado con " & labelTamaño.Text
        OnProgreso(labelInfo.Text)

    End Sub

    ''' <summary>
    ''' Borra la ventana de código y deja en blanco el <see cref="nombreUltimoFichero"/>.
    ''' Si se ha modificado el que había, pregunta si lo quieres guardar
    ''' </summary>
    Public Sub Nuevo()
        CurrentMDI.NuevaVentana()

        'If TextoModificado Then GuardarComo()

        'richTextBoxCodigo.Text = ""
        'richTextBoxLineas.Text = ""
        ''richTextBoxSyntax.Text = ""
        'Form1Activo.lstSyntax.Items.Clear()
        'nombreUltimoFichero = ""
        'TextoModificado = False
        'richTextBoxCodigo.Modified = False
        'labelInfo.Text = ""
        'labelPos.Text = "Lín: 1  Car: 1"
        ''labelTamaño.Text = "0 car."
        'labelTamaño.Text = $"{0:#,##0} car. ({0:#,##0} palab.)"
        'codigoAnterior = ""
    End Sub

    ''' <summary>
    ''' Muestra el cuadro de diálogo de Guardar como.
    ''' </summary>
    Public Sub GuardarComo()

        Dim fichero = comboBoxFileName.Text

        'If nombreUltimoFichero <> comboBoxFileName.Text Then
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
            nombreUltimoFichero = sFD.FileName
            ' Guardarlo
            Guardar(fichero)
        End Using
        'End If
    End Sub

    ''' <summary>
    ''' Guarda el fichero actual (<see cref="nombreUltimoFichero"/>).
    ''' Si no tiene nombre muestra el diálogo de guardar como
    ''' </summary>
    Public Sub Guardar()
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
    Public Sub AñadirAlComboBoxFileName(fic As String)
        If String.IsNullOrWhiteSpace(fic) Then Return

        ' No añadir el path si es el directorio de Documentos       (24/Sep/20)
        If Path.GetDirectoryName(fic) = DirDocumentos Then
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


    Public Sub comboBoxFileName_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            Abrir(comboBoxFileName.Text)
        End If
    End Sub

    Public Sub comboBoxFileName_Validating(sender As Object, e As ComponentModel.CancelEventArgs)
        If inicializando Then Return

        Dim fic = comboBoxFileName.Text
        If String.IsNullOrWhiteSpace(fic) Then Return

        AñadirAlComboBoxFileName(fic)
    End Sub


#End Region

#Region " Edición: Los menús, botones y métodos "

    'Friend menuEdit As ToolStripMenuItem = CurrentMDI.menuEdit
    'Friend menuEditBuscar As ToolStripMenuItem = CurrentMDI.menuEditBuscar
    'Friend menuEditBuscarSiguiente As ToolStripMenuItem = CurrentMDI.menuEditBuscarSiguiente
    'Friend menuEditCopiar As ToolStripMenuItem = CurrentMDI.menuEditCopiar
    'Friend menuEditCortar As ToolStripMenuItem = CurrentMDI.menuEditCortar
    'Friend menuEditDeshacer As ToolStripMenuItem = CurrentMDI.menuEditDeshacer
    'Friend menuEditPegar As ToolStripMenuItem = CurrentMDI.menuEditPegar
    'Friend menuEditReemplazar As ToolStripMenuItem = CurrentMDI.menuEditReemplazar
    'Friend menuEditReemplazarSiguiente As ToolStripMenuItem = CurrentMDI.menuEditReemplazarSiguiente
    'Friend menuEditReemplazarTodos As ToolStripMenuItem = CurrentMDI.menuEditReemplazarTodos
    'Friend menuEditRehacer As ToolStripMenuItem = CurrentMDI.menuEditRehacer
    'Friend menuEditSeleccionarTodo As ToolStripMenuItem = CurrentMDI.menuEditSeleccionarTodo

    'Friend buttonCopiar As ToolStripButton = CurrentMDI.buttonCopiar
    'Friend buttonCortar As ToolStripButton = CurrentMDI.buttonCortar
    'Friend buttonDeshacer As ToolStripButton = CurrentMDI.buttonDeshacer
    'Friend buttonEdicionRecortes As ToolStripButton = CurrentMDI.buttonEdicionRecortes
    'Friend buttonPegar As ToolStripButton = CurrentMDI.buttonPegar
    'Friend buttonRehacer As ToolStripButton = CurrentMDI.buttonRehacer
    'Friend buttonSeleccionar As ToolStripButton = CurrentMDI.buttonSeleccionar
    'Friend buttonSeleccionarTodo As ToolStripButton = CurrentMDI.buttonSeleccionarTodo

    'Friend rtbCodigoContext As ContextMenuStrip = CurrentMDI.rtbCodigoContext

    ' Las expresiones lambda para asignar con AddHandler en el menú editar
    Public lambdaUndo As EventHandler = Sub(sender As Object, e As EventArgs)
                                            If richTextBoxCodigo.CanUndo Then richTextBoxCodigo.Undo()
                                        End Sub
    Public lambdaRedo As EventHandler = Sub(sender As Object, e As EventArgs)
                                            If richTextBoxCodigo.CanRedo Then richTextBoxCodigo.Redo()
                                        End Sub
    Public lambdaPaste As EventHandler = Sub(sender As Object, e As EventArgs) Pegar()
    Public lambdaCopy As EventHandler = Sub(sender As Object, e As EventArgs)
                                            AñadirRecorte(richTextBoxCodigo.SelectedText)
                                            richTextBoxCodigo.Copy()
                                        End Sub
    Public lambdaCut As EventHandler = Sub(sender As Object, e As EventArgs)
                                           AñadirRecorte(richTextBoxCodigo.SelectedText)
                                           richTextBoxCodigo.Cut()
                                       End Sub
    Public lambdaSelectAll As EventHandler = Sub(sender As Object, e As EventArgs) richTextBoxCodigo.SelectAll()


    ''' <summary>
    ''' Crear un menú contextual para richTextBoxCodigo
    ''' para los comandos de edición
    ''' </summary>
    ''' <remarks>Usando la extensión Clonar: 27/Sep/20</remarks>
    Public Sub CrearContextMenuCodigo()
        CurrentMDI.rtbCodigoContext.Items.Clear()
        CurrentMDI.rtbCodigoContext.Items.AddRange({CurrentMDI.menuEditDeshacer.Clonar(lambdaUndo),
                                        CurrentMDI.menuEditRehacer.Clonar(lambdaRedo), CurrentMDI.toolEditSeparator1,
                                        CurrentMDI.menuEditCortar.Clonar(lambdaCut), CurrentMDI.menuEditCopiar.Clonar(lambdaCopy),
                                        CurrentMDI.menuEditPegar.Clonar(lambdaPaste), CurrentMDI.toolEditSeparator2,
                                        CurrentMDI.menuEditSeleccionarTodo.Clonar(lambdaSelectAll)})
        'richTextBoxCodigo.ContextMenuStrip = rtbCodigoContext
    End Sub


    '
    ' Los métodos de edición
    '

    ''' <summary>
    ''' Pega el texto del portapapeles en el editor
    ''' </summary>
    Public Sub Pegar()
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
    Public Sub menuEditDropDownOpening()
        If inicializando Then Return

        inicializando = True

        ' para saber si hay texto en el control
        Dim b = richTextBoxCodigo.TextLength > 0

        CurrentMDI.menuEditDeshacer.Enabled = richTextBoxCodigo.CanUndo
        CurrentMDI.menuEditRehacer.Enabled = richTextBoxCodigo.CanRedo
        CurrentMDI.menuEditCopiar.Enabled = richTextBoxCodigo.SelectionLength > 0
        CurrentMDI.menuEditCortar.Enabled = CurrentMDI.menuEditCopiar.Enabled
        CurrentMDI.menuEditSeleccionarTodo.Enabled = b
        CurrentMDI.menuEditPegar.Enabled = richTextBoxCodigo.CanPaste(DataFormats.GetFormat("Text"))

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
        CurrentMDI.menuEditPegarRecorte.Enabled = ColRecortes.Count > 0

        CurrentMDI.menuEditBuscar.Enabled = b
        CurrentMDI.menuEditReemplazar.Enabled = b

        ' Actualizar también los del menú contextual
        For j = 0 To CurrentMDI.rtbCodigoContext.Items.Count - 1
            For i = 0 To CurrentMDI.menuEdit.DropDownItems.Count - 1
                If CurrentMDI.rtbCodigoContext.Items(j).Name = CurrentMDI.menuEdit.DropDownItems(i).Name Then
                    CurrentMDI.rtbCodigoContext.Items(j).Enabled = CurrentMDI.menuEdit.DropDownItems(i).Enabled
                    Exit For
                End If
            Next
        Next

        inicializando = False
    End Sub


#End Region

#Region " Buscar y reemplazar: Métodos, menús, botones... "

    '
    ' Buscar y reemplazar
    '

    'Friend buttonBuscarSiguiente As ToolStripButton = CurrentMDI.buttonBuscarSiguiente
    'Friend buttonReemplazarSiguiente As ToolStripButton = CurrentMDI.buttonReemplazarSiguiente
    'Friend buttonReemplazarTodo As ToolStripButton = CurrentMDI.buttonReemplazarTodo


    ''' <summary>
    ''' El número máximo de items en buscar y reemplazar
    ''' </summary>
    Public Const BuscarMaxItems As Integer = 20

    ''' <summary>
    ''' La posición desde la que se está buscando
    ''' </summary>
    Public buscarPos As Integer

    ''' <summary>
    ''' La posición en que está la primera coincidencia
    ''' al buscar
    ''' </summary>
    Public buscarPosIni As Integer

    ''' <summary>
    ''' Para indicar que se está buscando
    ''' la primera coincidencia de la búsqueda
    ''' </summary>
    Public buscarPrimeraCoincidencia As Boolean = True

    ''' <summary>
    ''' La cadena a buscar
    ''' </summary>
    Public buscarQueBuscar As String

    ''' <summary>
    ''' La cadena a poner cuando se reemplaza
    ''' </summary>
    Public buscarQueReemplazar As String

    ''' <summary>
    ''' Si se busca teniendo en cuenta las mayúsculas y minúsculas
    ''' </summary>
    Public buscarMatchCase As Boolean

    ''' <summary>
    ''' Si se busca la palabra completa
    ''' </summary>
    Public buscarWholeWord As Boolean

    ''' <summary>
    ''' Para que al buscar para reemplazar no se meta en unbucle sin fin.
    ''' </summary>
    Public paraReemplazarSiguiente As Boolean 'Integer

    ''' <summary>
    ''' Muestra el panel de buscar/reemplazar
    ''' y reinicia la posición de búsqueda
    ''' </summary>
    ''' <param name="esBuscar">True si es Buscar, false si es Reemplazar</param>
    Public Sub BuscarReemplazar(esBuscar As Boolean)
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
    Public Sub MostrarPanelBuscar(mostrar As Boolean)
        CurrentMDI.panelBuscar.Visible = mostrar
        CurrentMDI.menuMostrar_Buscar.Checked = mostrar
        If mostrar Then
            esCtrlF = True
        End If
    End Sub

    ''' <summary>
    ''' Buscar siguiente coincidencia.
    ''' Devuelve False si no hay más
    ''' </summary>
    Public Function BuscarSiguiente(esReemplazar As Boolean) As Boolean
        ' Buscar en el texto lo indicado
        ' Se empieza desde la posición actual del texto

        buscarQueBuscar = comboBoxBuscar.Text
        ' Cambiar los \t por tabs y los \n por vbCr                 (29/Sep/20)
        buscarQueBuscar = buscarQueBuscar.Replace("\n", vbCr).Replace("\t", vbTab)

        If String.IsNullOrEmpty(buscarQueBuscar) Then
            If CurrentMDI.panelBuscar.Visible = False Then
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
                If paraReemplazarSiguiente = False Then
                    paraReemplazarSiguiente = True ' += 1
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

    ''' <summary>
    ''' Reemplaza el texto si la selección es lo buscado y
    ''' sigue buscando
    ''' </summary>
    Public Sub ReemplazarSiguiente()
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
        paraReemplazarSiguiente = False '0
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
    Public Sub ReemplazarTodo()
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
    Public Sub comboBoxReemplazar_Validating()
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
    Public Sub comboBoxBuscar_Validating()
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

#End Region

#Region " Editor: Los menús y botones "

    'Friend menuEditor As ToolStripMenuItem = CurrentMDI.menuEditor
    'Friend menuEditorCambiarMayúsculas As ToolStripMenuItem = CurrentMDI.menuEditorCambiarMayúsculas
    'Friend menuEditorClasificarSeleccion As ToolStripMenuItem = CurrentMDI.menuEditorClasificarSeleccion
    'Friend menuEditorMarcador As ToolStripMenuItem = CurrentMDI.menuEditorMarcador
    'Friend menuEditorMarcadorAnterior As ToolStripMenuItem = CurrentMDI.menuEditorMarcadorAnterior
    'Friend menuEditorMarcadorQuitarTodos As ToolStripMenuItem = CurrentMDI.menuEditorMarcadorQuitarTodos
    'Friend menuEditorMarcadorSiguiente As ToolStripMenuItem = CurrentMDI.menuEditorMarcadorSiguiente
    'Friend menuEditorPonerComentarios As ToolStripMenuItem = CurrentMDI.menuEditorPonerComentarios
    'Friend menuEditorPonerIndentacion As ToolStripMenuItem = CurrentMDI.menuEditorPonerIndentacion
    'Friend menuEditorPonerTexto As ToolStripMenuItem = CurrentMDI.menuEditorPonerTexto
    'Friend menuEditorPonerTextoAlFinal As ToolStripMenuItem = CurrentMDI.menuEditorPonerTextoAlFinal
    'Friend menuEditorQuitarComentarios As ToolStripMenuItem = CurrentMDI.menuEditorQuitarComentarios
    'Friend menuEditorQuitarEspacios As ToolStripMenuItem = CurrentMDI.menuEditorQuitarEspacios
    'Friend menuEditorQuitarIndentacion As ToolStripMenuItem = CurrentMDI.menuEditorQuitarIndentacion
    'Friend menuEditorQuitarTextoDelfinal As ToolStripMenuItem = CurrentMDI.menuEditorQuitarTextoDelfinal

    'Friend buttonEditorClasificarSeleccion As ToolStripButton = CurrentMDI.buttonEditorClasificarSeleccion
    'Friend buttonEditorMarcador As ToolStripButton = CurrentMDI.buttonEditorMarcador
    'Friend buttonEditorMarcadorAnterior As ToolStripButton = CurrentMDI.buttonEditorMarcadorAnterior
    'Friend buttonEditorMarcadorQuitarTodos As ToolStripButton = CurrentMDI.buttonEditorMarcadorQuitarTodos
    'Friend buttonEditorMarcadorSiguiente As ToolStripButton = CurrentMDI.buttonEditorMarcadorSiguiente
    'Friend buttonEditorPonerComentarios As ToolStripButton = CurrentMDI.buttonEditorPonerComentarios
    'Friend buttonEditorPonerIndentacion As ToolStripButton = CurrentMDI.buttonEditorPonerIndentacion
    'Friend buttonEditorQuitarComentarios As ToolStripButton = CurrentMDI.buttonEditorQuitarComentarios
    'Friend buttonEditorQuitarIndentacion As ToolStripButton = CurrentMDI.buttonEditorQuitarIndentacion

#End Region

#Region " Herramientas: Los paneles, menús y contextual del panel herramientas "

    'Friend menuTools As ToolStripMenuItem = CurrentMDI.menuTools

    'Friend menuMostrar_Buscar As ToolStripMenuItem = CurrentMDI.menuMostrar_Buscar
    'Friend menuMostrar_Compilar As ToolStripMenuItem = CurrentMDI.menuMostrar_Compilar
    'Friend menuMostrar_Edicion As ToolStripMenuItem = CurrentMDI.menuMostrar_Edicion
    'Friend menuMostrar_Editor As ToolStripMenuItem = CurrentMDI.menuMostrar_Editor
    'Friend menuMostrar_Ficheros As ToolStripMenuItem = CurrentMDI.menuMostrar_Ficheros

    'Friend panelBuscar As Panel = CurrentMDI.panelBuscar
    'Friend panelHerramientas As FlowLayoutPanel = CurrentMDI.panelHerramientas

    'Friend barraHerramientasContext As ContextMenuStrip = CurrentMDI.barraHerramientasContext

    'Friend menuOpciones As ToolStripMenuItem = CurrentMDI.menuOpciones


    Public ReadOnly Property comboBoxBuscar As ToolStripComboBox
        Get
            Return CurrentMDI.comboBoxBuscar
        End Get
    End Property

    Public ReadOnly Property comboBoxReemplazar As ToolStripComboBox
        Get
            Return CurrentMDI.comboBoxReemplazar
        End Get
    End Property

    Public ReadOnly Property buttonColorearAlEvaluar As ToolStripButton
        Get
            Return CurrentMDI.buttonColorearAlEvaluar
        End Get
    End Property

    Public ReadOnly Property buttonCompilarAlEvaluar As ToolStripButton
        Get
            Return CurrentMDI.buttonCompilarAlEvaluar
        End Get
    End Property

    Public ReadOnly Property chkMostrarLineasHTML As ToolStripButton
        Get
            Return CurrentMDI.chkMostrarLineasHTML
        End Get
    End Property

    Public ReadOnly Property buttonMatchCase As ToolStripButton
        Get
            Return CurrentMDI.buttonMatchCase
        End Get
    End Property

    Public ReadOnly Property buttonWholeWord As ToolStripButton
        Get
            Return CurrentMDI.buttonWholeWord
        End Get
    End Property

    Public ReadOnly Property toolStripFicheros As ToolStrip
        Get
            Return CurrentMDI.toolStripFicheros
        End Get
    End Property

    Public ReadOnly Property toolStripEdicion As ToolStrip
        Get
            Return CurrentMDI.toolStripEdicion
        End Get
    End Property

    Public ReadOnly Property toolStripCompilar As ToolStrip
        Get
            Return CurrentMDI.toolStripCompilar
        End Get
    End Property

    Public ReadOnly Property toolStripEditor As ToolStrip
        Get
            Return CurrentMDI.toolStripEditor
        End Get
    End Property

    Public ReadOnly Property txtPonerTexto As ToolStripTextBox
        Get
            Return CurrentMDI.txtPonerTexto
        End Get
    End Property


    '
    ' Cambio de tamaño del panel de herramientas
    '

    Public Sub panelHerramientas_SizeChanged(sender As Object, e As EventArgs)
        Dim iant = inicializando

        inicializando = True
        'Dim tAnt = splitContainer1.Top
        'splitContainer1.Top = panelHerramientas.Top + panelHerramientas.Height + 10
        'splitContainer1.Height += (tAnt - splitContainer1.Top)

        'Dim tAnt = Form1Activo.Top

        AjustarAltoPanelHerramientas()

        ' asignar el tamaño a todos los form1
        For Each frm As Form In Application.OpenForms
            If TypeOf frm Is Form1 Then
                With frm
                    If .WindowState = FormWindowState.Normal Then
                        .Top = 0
                        .Left = 0
                        .Width = CurrentMDI.ClientSize.Width - 4
                        .Height = CurrentMDI.ClientSize.Height - 4 - CurrentMDI.menuStrip1.Height - CurrentMDI.panelHerramientas.Height
                    End If
                End With
            End If
        Next


        inicializando = iant
    End Sub

    ''' <summary>
    ''' Ajustar el alto del panel de herramientas
    ''' según estén visibles las barras de herramientas que están
    ''' en la segunda línea.
    ''' </summary>
    Public Sub AjustarAltoPanelHerramientas()
        'If inicializando Then Return

        ' Ajustar el alto del FlowPanel (01/Oct/20)
        ' Comprobar si las barras de herramientas el Top es mayor de 10
        ' en ese caso, el FlowPanel ponerlo a 63, si no será 28
        Dim esVisible = False
        For Each c As Control In CurrentMDI.panelHerramientas.Controls
            If c.Visible AndAlso c.Top > 10 Then
                esVisible = True
                Exit For
            End If
        Next
        If esVisible Then
            CurrentMDI.HabilitarBotones()
            CurrentMDI.panelHerramientas.Height = 63
        Else
            CurrentMDI.panelHerramientas.Height = 31
        End If
    End Sub




#End Region


End Module
