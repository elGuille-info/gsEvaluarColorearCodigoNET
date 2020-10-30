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

    ' SendeMessage está declarada en el módulo WindowsAPI           (26/Oct/20)
    ' comentar esto si se añade ese módulo
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
    ''' Tupla para guardar el tamaño y posición del formulario
    ''' </summary>
    Public tamForm As (Left As Integer, Top As Integer, Height As Integer, Width As Integer)

    ''' <summary>
    ''' Llevar la cuenta de los marcadores que hay de forma global.
    ''' </summary>
    Public totalBookmarks As Integer

    ''' <summary>
    ''' Si se está inicializando.
    ''' Usado para que no se provoquen eventos en cadena
    ''' </summary>
    Public inicializando As Boolean = True

    ''' <summary>
    ''' Mientras están cargándose los ficheros será True.
    ''' </summary>
    Public cargando As Boolean

    ''' <summary>
    ''' Si se deben cambiar los TAB por 8 espacios
    ''' </summary>
    Public CambiarTabs As Boolean = True

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
    ''' El nombre del fichero de configuración global
    ''' </summary>
    Public FicheroConfiguracion As String

    ''' <summary>
    ''' La extensión a usar en los ficheros de configuración
    ''' </summary>
    ''' <remarks>23/Oct/2020</remarks>
    Public Const ExtensionConfiguracion As String = ".appconfig.txt"

    ''' <summary>
    ''' El Directorio donde se guarda la configuración
    ''' </summary>
    ''' <remarks>23/Oct/2020</remarks>
    Public DirConfiguracion As String

    ''' <summary>
    ''' El directorio para guardar las configuraciones de los ficheros.
    ''' </summary>
    Public DirConfigLocal As String

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
    ''' si se indican ficheros en la línea de comandos.
    ''' </summary>
    ''' <remarks>30/Oct/2020</remarks>
    Public variosLineaComandos As Integer

    ''' <summary>
    ''' Los nombres de los ficheros abiertos en cada sesión.
    ''' </summary>
    ''' <remarks>16/Oct/2020</remarks>
    Public UltimasVentanasAbiertas As New List(Of String)

    ''' <summary>
    ''' Lista de los ficheros que se han usado en la aplicación.
    ''' Antes estaban en ComboBoxfileName.
    ''' </summary>
    ''' <remarks>19/Oct/2020</remarks>
    Public UltimosFicherosAbiertos As New List(Of String)

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
    ''' Si es True se pre-compila el código.
    ''' </summary>
    Public Property compilarAlEvaluar As Boolean
        Get
            Return CurrentMDI.buttonCompilarAlEvaluar.Checked
        End Get
        Set(value As Boolean)
            CurrentMDI.buttonCompilarAlEvaluar.Checked = value
        End Set
    End Property

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

        ' Guardar los últimos ficheros abiertos                     (16/Oct/20)
        ' No guardar si está en blanco el nombre del fichero

        ' No guardar las últimas ventanas abiertas                  (30/Oct/20)
        ' si se han indicado ficheros desde la línea de comandos.
        ' Esto lo hago así lo puedo usar para abrir varios ficheros
        ' que usaré para modificar sin que sean de los que quiero "recordar"
        If variosLineaComandos = 0 Then
            cuantos = -1
            For i = 0 To UltimasVentanasAbiertas.Count - 1
                Dim s = UltimasVentanasAbiertas(i)
                If String.IsNullOrWhiteSpace(s) Then Continue For
                cuantos += 1
                cfg.SetKeyValue("Ventanas", $"Fichero{cuantos}", s)
            Next
            cfg.SetValue("Ventanas", "Cuantas", cuantos + 1)
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
        cuantos = UltimosFicherosAbiertos.Count
        cfg.SetValue("Ficheros", "Count", cuantos)
        For i = 0 To UltimosFicherosAbiertos.Count - 1 ' To 0 Step -1
            ' No añadir el path si es el directorio de Documentos
            Dim s = UltimosFicherosAbiertos(i).ToString
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

        ' para el texto a quitar/poner                              (08/Oct/20)
        cfg.SetValue("Quitar Poner Texto", "PonerTexto", txtPonerTexto.Text)

        ' El tamaño y la posición de la ventana                     (26/Oct/20)
        cfg.SetValue("Ventana", "Left", tamForm.Left)
        cfg.SetValue("Ventana", "Top", tamForm.Top)
        cfg.SetValue("Ventana", "Height", tamForm.Height)
        cfg.SetValue("Ventana", "Width", tamForm.Width)

        cfg.Save()
    End Sub

    ''' <summary>
    ''' Leer el fichero de configuración
    ''' y asignar los valores usados anteriormente.
    ''' </summary>
    Public Sub LeerConfig()
        Dim cfg = New Config(FicheroConfiguracion)

        ' El tamaño y la posición de la ventana                     (26/Oct/20)
        tamForm.Left = cfg.GetValue("Ventana", "Left", -1)
        tamForm.Top = cfg.GetValue("Ventana", "Top", -1)
        tamForm.Height = cfg.GetValue("Ventana", "Height", -1)
        tamForm.Width = cfg.GetValue("Ventana", "Width", -1)

        ' Comprobar que no esté fuera de la pantalla
        If tamForm.Left <> -1 Then
            ' No se mostraba al cambiar del monitor externo         (30/Oct/20)
            ' a la pantalla del portátil.
            ' Invierto la comparación... aunque no sé...
            ' cuando vuelva a ser negativo no se verá en su sitio...
            If Screen.PrimaryScreen.WorkingArea.Left < tamForm.Left Then
                CurrentMDI.Left = tamForm.Left
            End If
        End If
        If tamForm.Top <> -1 Then
            If tamForm.Top < 0 Then
                CurrentMDI.Top = 0
            Else
                CurrentMDI.Top = tamForm.Top
            End If
            'If Screen.PrimaryScreen.WorkingArea.Top < tamForm.Top Then
            '    CurrentMDI.Top = tamForm.Top
            'Else
            '    CurrentMDI.Top = 0
            'End If
        End If
        If tamForm.Height > -1 Then CurrentMDI.Height = tamForm.Height
        If tamForm.Width > -1 Then CurrentMDI.Width = tamForm.Width

        ' Si cargarUltimo es falso no asignar el último fichero     (16/Sep/20)
        cargarUltimo = cfg.GetValue("Ficheros", "CargarUltimo", True)

        buttonLenguaje.Text = cfg.GetValue("Herramientas", "Lenguaje", Compilar.LenguajeVisualBasic)
        If buttonLenguaje.Text = Compilar.LenguajeVisualBasic Then
            buttonLenguaje.Image = buttonLenguaje.DropDownItems(0).Image
        ElseIf buttonLenguaje.Text = Compilar.LenguajeCSharp Then
            buttonLenguaje.Image = buttonLenguaje.DropDownItems(1).Image
        Else
            buttonLenguaje.Image = buttonLenguaje.DropDownItems(2).Image
        End If
        colorearAlCargar = cfg.GetValue("Herramientas", "Colorear", True)
        chkMostrarLineasHTML.Checked = cfg.GetValue("Herramientas", "Mostrar Líneas HTML", True)

        colorearAlEvaluar = cfg.GetValue("Herramientas", "ColorearEvaluar", False)
        compilarAlEvaluar = cfg.GetValue("Herramientas", "CompilarEvaluar", True)

        ' Para la forma de clasificar                               (02/Oct/20)
        CompararString.IgnoreCase = cfg.GetValue("Clasificar", "IgnoreCase", False)
        CompararString.CompareOrdinal = cfg.GetValue("Clasificar", "CompareOrdinal", False)

        ' Leer la lista de los últimos ficheros que se han abierto
        ' (aunque no estén abiertos ahora)
        Dim cuantos = cfg.GetValue("Ficheros", "Count", 0)
        UltimosFicherosAbiertos.Clear()
        For i = 0 To MaxFicsConfig - 1
            If i >= cuantos Then Exit For
            Dim s = cfg.GetValue("Ficheros", $"Fichero{i}", "-1")
            If s = "-1" Then Exit For
            ' No añadir el path si es el directorio de Documentos
            If Path.GetDirectoryName(s) = DirDocumentos Then
                s = Path.GetFileName(s)
            End If
            UltimosFicherosAbiertos.Add(s)
        Next
        '' No clasificar los elementos

        ' El nombre y tamaño de la fuente                           (11/Sep/20)
        fuenteNombre = cfg.GetValue("Fuente", "Nombre", "Consolas")
        fuenteTamaño = cfg.GetValue("Fuente", "Tamaño", "11")
        labelFuente.Text = $"{fuenteNombre}; {fuenteTamaño}"

        ' Solo asignar si son diferentes                            (02/Oct/20)

        richTextBoxCodigo.Font = New Font(fuenteNombre, CSng(fuenteTamaño))
        richTextBoxLineas.Font = New Font(fuenteNombre, CSng(fuenteTamaño))

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
        comboBoxBuscar.Text = If(String.IsNullOrWhiteSpace(buscarQueBuscar), BuscarVacio, buscarQueBuscar)
        If comboBoxBuscar.Text = BuscarVacio Then
            comboBoxBuscar.ForeColor = SystemColors.GrayText
        Else
            comboBoxBuscar.ForeColor = SystemColors.ControlText
        End If
        comboBoxReemplazar.Text = If(String.IsNullOrWhiteSpace(buscarQueReemplazar), ReemplazarVacio, buscarQueReemplazar)
        If comboBoxReemplazar.Text = ReemplazarVacio Then
            comboBoxReemplazar.ForeColor = SystemColors.GrayText
        Else
            comboBoxReemplazar.ForeColor = SystemColors.ControlText
        End If

        Dim j = 0

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

        ' para el texto a quitar/poner                              (08/Oct/20)
        txtPonerTexto.Text = cfg.GetValue("Quitar Poner Texto", "PonerTexto", "")

        ' Leer los últimos ficheros abiertos                        (16/Oct/20)
        cuantos = cfg.GetValue("Ventanas", "Cuantas", 0)
        UltimasVentanasAbiertas.Clear()
        For i = 0 To cuantos - 1
            Dim s = cfg.GetValue("Ventanas", $"Fichero{i}", "")
            'If s = "-1" Then Exit For
            ' Si está el nombre en blanco, continuar                (17/Oct/20)
            If String.IsNullOrWhiteSpace(s) Then Continue For
            ' UltimasVentanasAbiertas debe tener el path completo   (20/Oct/20)
            's = Path.GetFileName(s)
            UltimasVentanasAbiertas.Add(s)
        Next
        ' Si se indica cargarUltimo, abrirlos todos                 (16/Oct/20)
        ' Si se indica algún fichero en la línea de comandos        (30/Oct/20)
        ' No abrir los anteriores (uso la variable variosLineaComandos)
        If cargarUltimo AndAlso variosLineaComandos < 2 Then
            cargando = True
            cuantos = UltimasVentanasAbiertas.Count
            j = cuantos
            ' En Abrir se llama a OnProgreso una vez, en ColorearCodigo 2 veces
            If colorearAlCargar Then cuantos *= 3
            MostrarProcesando($"Cargando {j} ficheros", "Cargando los ficheros...", cuantos)
            For i = 0 To j - 1
                m_fProcesando.Text = $"Cargando {i + 1} de {j} ficheros"

                ' Esto se supone que no se dará...                  (17/Oct/20)
                If String.IsNullOrWhiteSpace(UltimasVentanasAbiertas(i)) Then Continue For

                m_fProcesando.Mensaje1 = $"Cargando {UltimasVentanasAbiertas(i)}...{vbCrLf}"
                Application.DoEvents()
                If m_fProcesando.Cancelar Then Exit For
                If i = UltimasVentanasAbiertas.Count - 1 Then
                    m_fProcesando.ValorActual = m_fProcesando.Maximo - 1
                End If
                AbrirEnNuevaVentana(UltimasVentanasAbiertas(i))
                Application.DoEvents()
            Next
            m_fProcesando.Close()
            cargando = False

            '
            ' Esto ya no hace falta...                              (26/Oct/20)
            '
            ' Activar el último formulario cargado                  (24/Oct/20)
            ' para que se actualicen los botones, etc.
            ' Llamo directamente al evento Activated                (26/Oct/20)
            'Form1Activo.Form1_Activated() 'BringToFront()
        End If

    End Sub

#End Region

#Region " Para la ventana de progreso "

    Friend m_fProcesando As fProcesando

    ''' <summary>
    ''' Mostrar la ventana de procesando (mientras cargan los ficheros, etc.)
    ''' </summary>
    ''' <param name="titulo">El título de la ventana.</param>
    ''' <param name="msg">El mensaje a mostrar.</param>
    ''' <param name="max">El valor máximo de la barra de progreso.</param>
    ''' <remarks>16/Oct/2020</remarks>
    Public Sub MostrarProcesando(titulo As String, msg As String, max As Integer)
        m_fProcesando = New fProcesando
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

#Region " Métodos para la posición actual "

    ''' <summary>
    ''' Averigua la línea, columna (y primer caracter de la línea) de la posición actual en richTextBoxCodigo.
    ''' </summary>
    ''' <returns>Una tupla con la Fila, SelStart y la primera posición de la línea (Inicio)</returns>
    ''' <remarks>En esta implementación no se le añade +1 a la posición ni a la línea.</remarks>
    Public Function PosicionActual0() As (Linea As Integer, SelStart As Integer, Inicio As Integer)
        Dim pos As Integer = richTextBoxCodigo.SelectionStart
        Dim lin As Integer = richTextBoxCodigo.GetLineFromCharIndex(pos)
        Dim fcol = richTextBoxCodigo.GetFirstCharIndexFromLine(lin)

        Return (lin, pos, fcol)
    End Function

    ''' <summary>
    ''' Averigua la línea, columna (y primer caracter de la línea) de la posición actual en richTextBoxCodigo.
    ''' </summary>
    ''' <returns>Una tupla con la Fila, Columna y la posición del primer caracter de la línea (Inicio)</returns>
    ''' <remarks>Esta implementación le añade 1 a la posición y a la línea.
    ''' Usarla para mostrar la fila y columna</remarks>
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
            Dim finlinea = selAnt.ComprobarFinLinea
            Dim lineas() As String = selAnt.Split(finlinea.ToCharArray)
            sIndent = New String(" "c, EspaciosIndentar)
            Dim sb As New System.Text.StringBuilder
            Dim j As Integer = lineas.Length - 1
            Dim k As Integer

            For i = 0 To j
                If lineas(i).Length > 0 AndAlso i < j Then
                    k += EspaciosIndentar
                End If
                If i = j Then
                    'If lineas(i).Length = 0 Then
                    'End If
                    sb.AppendFormat("{0}{1}", sIndent, lineas(i))
                Else
                    sb.AppendFormat("{0}{1}{2}", sIndent, lineas(i), finlinea)
                End If
            Next

            Dim lenAnt = selAnt.Length

            ' Poner la selección y sustituir el texto
            rtEditor.SelectionStart = selStart
            rtEditor.SelectedText = sb.ToString.TrimEnd() & finlinea
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
            Dim finlinea = rtEditor.SelectedText.ComprobarFinLinea
            Dim lineas() As String = rtEditor.SelectedText.Split(finlinea.ToCharArray)
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
                            sb.AppendFormat("{0}{1}", lineas(i).Substring(EspaciosIndentar), finlinea)
                        End If
                    ElseIf lineas(i).StartsWith(vbTab) Then
                        ' Considerar los tab como indentación       (22/Nov/05)
                        If i = j Then
                            sb.AppendFormat("{0}", lineas(i).Substring(1))
                        Else
                            sb.AppendFormat("{0}{1}", lineas(i).Substring(1), finlinea)
                        End If
                    Else
                        If i = j Then
                            sb.AppendFormat("{0}", lineas(i))
                        Else
                            sb.AppendFormat("{0}{1}", lineas(i), finlinea)
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
            Dim finLinea = rtEditor.SelectedText.ComprobarFinLinea
            Dim lineas() As String = rtEditor.SelectedText.Split(finLinea.ToCharArray)
            Dim sb As New System.Text.StringBuilder
            Dim j As Integer = lineas.Length - 1
            If String.IsNullOrEmpty(lineas(j)) Then j -= 1
            Dim k As Integer = 0

            For i As Integer = 0 To j
                ' Poner el comentario después de los espacios inciales  (05/Oct/20)
                Dim p = lineas(i).TrimStart().Length
                Dim esp = lineas(i).Length - p
                'Debug.Assert(esp <> lineas(i).Length)
                If esp = lineas(i).Length Then
                    sb.AppendFormat("{0}{1}", lineas(i), finLinea)
                Else
                    sb.AppendFormat("{0}{1}{2}{3}", lineas(i).Substring(0, esp), sSep, lineas(i).Substring(esp), finLinea)
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
            Dim finLinea = rtEditor.SelectedText.ComprobarFinLinea
            Dim lineas() As String = rtEditor.SelectedText.Split(finLinea.ToCharArray)
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
                        sb.AppendFormat("{0}{1}", lineas(i), finLinea)
                    Else
                        sb.AppendFormat("{0}{1}{2}", lineas(i).Substring(0, p), lineas(i).Substring(p + sSep.Length), finLinea)
                    End If

                    k += 1
                Else
                    ' si no tiene comentario, dejarlo como está
                    sb.AppendFormat("{0}{1}", lineas(i), finLinea)
                    k += 1
                End If
            Next
            Dim sbT = sb.ToString
            If j = 1 Then
                sbT = sbT.TrimEnd(finLinea.ToCharArray)
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
        If m_fProcesando Is Nothing Then
            MostrarProcesando(labelInfo.Text, labelInfo.Text, 2)
        End If
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

        Dim fHTML As New FormVisorHTML(Form1Activo.nombreFichero, sHTML)
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
            Form1Activo.Guardar()
        End If

        Dim fichero = Form1Activo.nombreFichero

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
            Form1Activo.splitContainer1_Resize()

            Return
        End If
        If res.Result.Success Then
            labelInfo.Text = $"Se ha compilado y ejecutado satisfactoriamente."

            Form1Activo.lstSyntax.Items.Clear()
            Form1Activo.splitContainer1.Panel2.Visible = False
            Form1Activo.splitContainer1_Resize()
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
            Form1Activo.Guardar()
        End If

        Dim filepath = Form1Activo.nombreFichero
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
            Form1Activo.splitContainer1_Resize()

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
            Form1Activo.Guardar()
        End If

        Dim res As EmitResult
        Form1Activo.lstSyntax.Items.Clear()

        If compilarAlEvaluar Then
            Form1Activo.splitContainer1.Panel2.Visible = False
            Form1Activo.splitContainer1_Resize()
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
        Form1Activo.splitContainer1_Resize()
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
            Form1Activo.splitContainer1_Resize()
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
        Dim finlinea = richTextBoxCodigo.SelectedText.ComprobarFinLinea
        Dim lineas() As String = richTextBoxCodigo.SelectedText.Split(finlinea.ToCharArray)
        Dim sb As New System.Text.StringBuilder
        Dim j = lineas.Length - 1
        If String.IsNullOrWhiteSpace(lineas(j)) Then j -= 1
        Dim k = 0

        ' Clasificar el array usando el comparador de CompararString
        Array.Sort(lineas, 0, j + 1, New CompararString)

        For i As Integer = 0 To j
            sb.AppendFormat("{0}{1}", lineas(i), finlinea)
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
        Dim finlinea = richTextBoxCodigo.SelectedText.ComprobarFinLinea
        Dim lineas() As String = richTextBoxCodigo.SelectedText.Split(finlinea.ToCharArray)
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
                    sb.AppendFormat("{0}{1}", lineas(i).QuitarTodosLosEspacios, finlinea)
                ElseIf delante AndAlso detras Then
                    sb.AppendFormat("{0}{1}", lineas(i).Trim(), finlinea)
                ElseIf delante Then
                    sb.AppendFormat("{0}{1}", lineas(i).TrimStart(), finlinea)
                ElseIf detras Then
                    sb.AppendFormat("{0}{1}", lineas(i).TrimEnd(), finlinea)
                End If
                k += 1
            Else
                sb.AppendFormat("{0}{1}", lineas(i), finlinea)
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
            Dim finlinea = richTextBoxCodigo.SelectedText.ComprobarFinLinea
            Dim lineas() As String = richTextBoxCodigo.SelectedText.TrimEnd(ChrW(13)).Split(finlinea.ToCharArray)
            Dim sb As New System.Text.StringBuilder
            Dim j As Integer = lineas.Length - 1

            For i As Integer = 0 To j
                'If i = j AndAlso String.IsNullOrEmpty(lineas(i)) Then Exit For
                sb.AppendFormat("{0}{1}{2}", lineas(i), sPoner, finlinea)
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
            Dim finlinea = richTextBoxCodigo.SelectedText.ComprobarFinLinea
            Dim lineas() As String = richTextBoxCodigo.SelectedText.TrimEnd(ChrW(13)).Split(finlinea.ToCharArray)
            Dim sb As New System.Text.StringBuilder
            Dim j As Integer = lineas.Length - 1

            For i As Integer = 0 To j
                If lineas(i).TrimEnd.EndsWith(sQuitar, StringComparison.OrdinalIgnoreCase) Then
                    Dim k = lineas(i).LastIndexOf(sQuitar)
                    If k = 0 Then
                        sb.AppendFormat("{0}{1}", "", finlinea)
                    Else
                        sb.AppendFormat("{0}{1}", lineas(i).Substring(0, k), finlinea)
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
            Form1Activo.AñadirNumerosDeLinea()

            ' Seleccionar el texto después de pegar                 (04/Oct/20)
            ' Si es texto, no seleccionar                           (30/Oct/20)
            If buttonLenguaje.Text <> ExtensionTexto Then
                richTextBoxCodigo.SelectionStart = pos
                richTextBoxCodigo.SelectionLength = s.Length + (selStart - pos)
                ' y colorearlo si procede
                ColorearSeleccion()
            End If
            richTextBoxCodigo.SelectionStart = selStart 'pos
        End If
    End Sub

#End Region

#Region " Ficheros: métodos, menús y botones "

    ''' <summary>
    ''' Para saber que no se debe colorear, ni compilar, etc.
    ''' </summary>
    Public Const ExtensionTexto As String = "Texto"

    ''' <summary>
    ''' El número máximo de ficheros en el menú recientes
    ''' </summary>
    Public Const MaxFicsMenu As Integer = 25 '9

    ''' <summary>
    ''' El número máximo de ficheros a guardar en la configuración
    ''' y a mostrar en el combo de los nombres de los ficheros.
    ''' Estos ficheros solo se podrán ver en las opciones del programa,
    ''' ya que solo estarán accesible los 25 mostrados en Recientes
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

        MessageBox.Show($"{producto} v{vers} ({fvi.FileVersion})" & vbCrLf & vbCrLf &
                        $"{descL}" & vbCrLf &
                        $"{desc1}",
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
            If (i + 1) > UltimosFicherosAbiertos.Count Then Exit For

            Dim s = $"{i + 1} - {UltimosFicherosAbiertos(i)}"
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
            If frm.Text = ficT Then 'If Path.GetFileName(frm.nombreFichero) = ficT Then
                frm.BringToFront()
                frm.Activate()
                Return
            End If
        Next
        ' Si llega aquí es que no estaba abierto
        Nuevo()
        ' Aquí siempre debe tener el path                           (26/Oct/20)
        Dim sDir = Path.GetDirectoryName(ficT)
        If String.IsNullOrWhiteSpace(sDir) Then
            sDir = DirDocumentos
            ficT = Path.Combine(sDir, ficT)
        End If
        Form1Activo.nombreFichero = ficT
        Form1Activo.Abrir(Form1Activo.nombreFichero)
        Form1Activo.BringToFront()
        'Form1Activo.Activate()
    End Sub


    ''' <summary>
    ''' Muestra el cuadro de diálogos de abrir ficheros
    ''' </summary>
    Public Sub Abrir()
        Using oFD As New OpenFileDialog
            oFD.Title = "Seleccionar fichero de código a abrir"
            Dim sFN = Form1Activo.nombreFichero
            sFN = If(String.IsNullOrEmpty(sFN),
                Path.Combine(DirDocumentos, $"prueba.{If(buttonLenguaje.Text = Compilar.LenguajeVisualBasic, ".vb", ".cs")}"),
                sFN)
            oFD.FileName = sFN
            oFD.Filter = "Código de Visual Basic y CSharp (*.vb;*.cs)|*.vb;*.cs|Textos (*.txt;*.log;*.md)|*.txt;*.log;*.md|Todos los ficheros (*.*)|*.*"
            oFD.InitialDirectory = Path.GetDirectoryName(sFN)
            oFD.Multiselect = False
            oFD.RestoreDirectory = True
            If oFD.ShowDialog() = DialogResult.OK Then
                AbrirEnNuevaVentana(oFD.FileName)
            End If
        End Using
    End Sub

    ''' <summary>
    ''' Abrir el fichero indicado en una nueva ventana.
    ''' </summary>
    ''' <param name="fic">El path completo del fichero a abrir.</param>
    ''' <remarks>19/Oct/2020</remarks>
    Public Sub AbrirEnNuevaVentana(fic As String)
        Dim cargandoTmp = cargando
        cargando = True
        Nuevo()
        Form1Activo.nombreFichero = fic
        Form1Activo.Abrir(fic)
        cargando = cargandoTmp
    End Sub

    '
    ' Los métodos pasados a Form1                                   (23/Oct/20)
    ' Recargar, Abrir(fic) y Guardar(fic), Guardar y GuardarComo
    '

    ''' <summary>
    ''' Borra la ventana de código y deja en blanco el <see cref="nombreUltimoFichero"/>.
    ''' Si se ha modificado el que había, pregunta si lo quieres guardar
    ''' </summary>
    Public Sub Nuevo()
        CurrentMDI.NuevaVentana()
    End Sub


    ''' <summary>
    ''' Añade (si no está) el fichero indicado al principio de la lista de ficheros y
    ''' asigna la lista al menú de los ficheros recientes.
    ''' </summary>
    ''' <param name="fic">El fichero a añadir al <see cref="UltimosFicherosAbiertos"/></param>
    Public Sub AñadirAUltimosFicherosAbiertos(fic As String)
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
        col.AddRange(UltimosFicherosAbiertos)

        UltimosFicherosAbiertos.Clear()
        ' quitar los repetidos
        ' hacerlo así para que el último abiero/guardado            (28/Sep/20)
        ' siempre esté el primero de la lista
        UltimosFicherosAbiertos.AddRange(col.Distinct().ToArray())

        AsignarRecientes()

        inicializando = False
    End Sub


#End Region

#Region " Edición: Los menús, botones y métodos "

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
                                        CurrentMDI.menuEditRehacer.Clonar(lambdaRedo), CurrentMDI.tsSepEdit1,
                                        CurrentMDI.menuEditCortar.Clonar(lambdaCut), CurrentMDI.menuEditCopiar.Clonar(lambdaCopy),
                                        CurrentMDI.menuEditPegar.Clonar(lambdaPaste), CurrentMDI.tsSepEdit2,
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
            ' El texto pegado                                       (26/Oct/20)
            Dim sTexto = Clipboard.GetText()

            Dim selStart = richTextBoxCodigo.SelectionStart - sTexto.Length
            Dim lin = richTextBoxCodigo.GetLineFromCharIndex(selStart)
            Dim pos = richTextBoxCodigo.GetFirstCharIndexFromLine(lin + 1)
            ' Esta es la primera línea del texto pegado             (26/Oct/20)
            'Dim sLinea = richTextBoxCodigo.Lines(lin + 1)

            ' obligar a poner las líneas                            (18/Sep/20)
            Form1Activo.AñadirNumerosDeLinea()

            ' Seleccionar el texto después de pegar                 (04/Oct/20)
            ' A ver si al no colorear cambia la posición            (30/Oct/20)
            ' ¡Efectivamente!
            If buttonLenguaje.Text = ExtensionTexto Then
                richTextBoxCodigo.SelectionStart = selStart
                richTextBoxCodigo.SelectionLength = sTexto.Length
            Else
                richTextBoxCodigo.SelectionStart = pos
                richTextBoxCodigo.SelectionLength = sTexto.Length
                ' y colorearlo si procede
                ColorearSeleccion()
            End If
        End If
    End Sub

#End Region

#Region " Buscar y reemplazar: Métodos, menús, botones... "

    '
    ' Buscar y reemplazar
    '

    Public Const BuscarVacio As String = "Buscar..."
    Public Const ReemplazarVacio As String = "Reemplazar..."

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
        MostrarPanelBuscar(True, esReemplazar:=Not esBuscar)
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
    ''' <param name="esReemplazar">Si se debe mostrar el toolStrip de Reemplazar</param>
    Public Sub MostrarPanelBuscar(mostrar As Boolean, esReemplazar As Boolean)
        CurrentMDI.panelBuscar.Visible = mostrar
        CurrentMDI.menuMostrar_Buscar.Checked = mostrar
        If mostrar Then
            CurrentMDI.toolStripReemplazar.Visible = esReemplazar
        End If
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

        Dim buscarPosAnt = buscarPos
        If buscarPos >= richTextBoxCodigo.Text.Length Then
            buscarPos = 0
        ElseIf buscarPos < 0 Then
            buscarPos = 0
        End If

        ' Se ve que en la versión de .NET 5.0 RC2 han modificado 
        ' la forma de usar Find
        ' y ahora si se busca algo que está más de na vez, devuelve -1
        ' como si no estuviera... 
        ' NO ES ASÍ, fue fallo mío y buscaba algo que no estaba...  (30/Oct/20)
        ' lo que indica la documentación sobre FIND es que la cadena de búsqueda
        ' no puede contener varias líneas, si las tiene, no lo encuentra.
        buscarPos = richTextBoxCodigo.Find(buscarQueBuscar, buscarPos, rtbFinds)
        If buscarPos = -1 Then
            If esReemplazar Then
                If paraReemplazarSiguiente = False Then
                    paraReemplazarSiguiente = True
                    buscarPos = 0
                    If Not BuscarSiguiente(esReemplazar) Then
                        Dim msg As String
                        If buscarPosAnt = 0 Then
                            msg = $"No se ha encontrado el texto buscado: {buscarQueBuscar}."
                        Else
                            msg = $"No hay más coincidencias de: {buscarQueBuscar}."
                        End If
                        MessageBox.Show(msg,
                                        "Reemplazar siguiente",
                                        MessageBoxButtons.OK, MessageBoxIcon.Information)
                        esCtrlF = True
                        Return False
                    End If
                End If
                Return False
            Else
                If buscarPosAnt = 0 OrElse buscarPosAnt >= richTextBoxCodigo.Text.Length Then
                    MessageBox.Show($"No se ha encontrado el texto buscado: {buscarQueBuscar}.",
                                    "Buscar siguiente",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information)
                    esCtrlF = True
                    Return False
                End If
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
                    MessageBox.Show($"No hay más coincidencias de: {buscarQueBuscar}, se ha llegado al inicio de la búsqueda.",
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
            MessageBox.Show($"{t} cambio{plural} realizado{plural} de: {buscarQueBuscar} por: {buscarQueReemplazar}.",
                            "Reemplazar todos",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show($"No se encontró el texto especificado: {buscarQueBuscar}",
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

        If comboBoxReemplazar.Text <> ReemplazarVacio Then
            comboBoxReemplazar.ForeColor = SystemColors.ControlText
        Else
            comboBoxReemplazar.ForeColor = SystemColors.GrayText
        End If

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

        If comboBoxBuscar.Text <> BuscarVacio Then
            comboBoxBuscar.ForeColor = SystemColors.ControlText
        Else
            comboBoxBuscar.ForeColor = SystemColors.GrayText
        End If
    End Sub

#End Region

#Region " Herramientas: Los paneles, menús y contextual del panel herramientas "

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
