<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormOpciones
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormOpciones))
        Me.tabControl1 = New System.Windows.Forms.TabControl()
        Me.tabPage1 = New System.Windows.Forms.TabPage()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.chkCaseSensitive = New System.Windows.Forms.CheckBox()
        Me.chkCompareOrdinal = New System.Windows.Forms.CheckBox()
        Me.chkMostrarLineasHTML = New System.Windows.Forms.CheckBox()
        Me.grbAlEvaluar = New System.Windows.Forms.GroupBox()
        Me.chkColorearEvaluar = New System.Windows.Forms.CheckBox()
        Me.chkCompilarEvaluar = New System.Windows.Forms.CheckBox()
        Me.chkColorearCargar = New System.Windows.Forms.CheckBox()
        Me.chkCargarUltimo = New System.Windows.Forms.CheckBox()
        Me.tabPage2 = New System.Windows.Forms.TabPage()
        Me.groupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtFic = New System.Windows.Forms.TextBox()
        Me.btnOrdenar = New System.Windows.Forms.Button()
        Me.btnQuitar = New System.Windows.Forms.Button()
        Me.lstFics = New System.Windows.Forms.ListBox()
        Me.label1 = New System.Windows.Forms.Label()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.grbColores = New System.Windows.Forms.GroupBox()
        Me.txtIndentar = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.comboFuenteTamaño = New System.Windows.Forms.ComboBox()
        Me.comboFuenteNombre = New System.Windows.Forms.ComboBox()
        Me.labelTamaño = New System.Windows.Forms.Label()
        Me.labelFuente = New System.Windows.Forms.Label()
        Me.TabPage4 = New System.Windows.Forms.TabPage()
        Me.grbBuscarReemplazar = New System.Windows.Forms.GroupBox()
        Me.chkWholeWord = New System.Windows.Forms.CheckBox()
        Me.chkMatchCase = New System.Windows.Forms.CheckBox()
        Me.grbReemplazar = New System.Windows.Forms.GroupBox()
        Me.lstReemplazar = New System.Windows.Forms.ListBox()
        Me.btnQuitarReemplazar = New System.Windows.Forms.Button()
        Me.btnOrdenarReemplazar = New System.Windows.Forms.Button()
        Me.grbBuscar = New System.Windows.Forms.GroupBox()
        Me.lstBuscar = New System.Windows.Forms.ListBox()
        Me.btnQuitarBuscar = New System.Windows.Forms.Button()
        Me.btnOrdenarBuscar = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TabPage5 = New System.Windows.Forms.TabPage()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.lstRecortes = New System.Windows.Forms.ListBox()
        Me.btnQuitarRecortes = New System.Windows.Forms.Button()
        Me.btnOrdenarRecortes = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.btnAceptar = New System.Windows.Forms.Button()
        Me.btnCancelar = New System.Windows.Forms.Button()
        Me.toolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.chkCambiarTab = New System.Windows.Forms.CheckBox()
        Me.tabControl1.SuspendLayout()
        Me.tabPage1.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.grbAlEvaluar.SuspendLayout()
        Me.tabPage2.SuspendLayout()
        Me.groupBox1.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.grbColores.SuspendLayout()
        Me.TabPage4.SuspendLayout()
        Me.grbBuscarReemplazar.SuspendLayout()
        Me.grbReemplazar.SuspendLayout()
        Me.grbBuscar.SuspendLayout()
        Me.TabPage5.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'tabControl1
        '
        Me.tabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tabControl1.Controls.Add(Me.tabPage1)
        Me.tabControl1.Controls.Add(Me.tabPage2)
        Me.tabControl1.Controls.Add(Me.TabPage3)
        Me.tabControl1.Controls.Add(Me.TabPage4)
        Me.tabControl1.Controls.Add(Me.TabPage5)
        Me.tabControl1.Location = New System.Drawing.Point(12, 12)
        Me.tabControl1.Name = "tabControl1"
        Me.tabControl1.SelectedIndex = 0
        Me.tabControl1.Size = New System.Drawing.Size(549, 317)
        Me.tabControl1.TabIndex = 0
        '
        'tabPage1
        '
        Me.tabPage1.Controls.Add(Me.chkCambiarTab)
        Me.tabPage1.Controls.Add(Me.GroupBox4)
        Me.tabPage1.Controls.Add(Me.chkMostrarLineasHTML)
        Me.tabPage1.Controls.Add(Me.grbAlEvaluar)
        Me.tabPage1.Controls.Add(Me.chkColorearCargar)
        Me.tabPage1.Controls.Add(Me.chkCargarUltimo)
        Me.tabPage1.Location = New System.Drawing.Point(4, 24)
        Me.tabPage1.Name = "tabPage1"
        Me.tabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPage1.Size = New System.Drawing.Size(541, 289)
        Me.tabPage1.TabIndex = 0
        Me.tabPage1.Text = "Opciones generales"
        Me.tabPage1.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.chkCaseSensitive)
        Me.GroupBox4.Controls.Add(Me.chkCompareOrdinal)
        Me.GroupBox4.Location = New System.Drawing.Point(12, 195)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(511, 77)
        Me.GroupBox4.TabIndex = 5
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Clasificar:"
        '
        'chkCaseSensitive
        '
        Me.chkCaseSensitive.AutoSize = True
        Me.chkCaseSensitive.Location = New System.Drawing.Point(6, 22)
        Me.chkCaseSensitive.Name = "chkCaseSensitive"
        Me.chkCaseSensitive.Size = New System.Drawing.Size(221, 19)
        Me.chkCaseSensitive.TabIndex = 0
        Me.chkCaseSensitive.Text = "Distinguir &mayúsculas de minúsculas"
        Me.toolTip1.SetToolTip(Me.chkCaseSensitive, "Si al clasificar se distinguen las mayúsculas de las minúsculas.")
        Me.chkCaseSensitive.UseVisualStyleBackColor = True
        '
        'chkCompareOrdinal
        '
        Me.chkCompareOrdinal.AutoSize = True
        Me.chkCompareOrdinal.Location = New System.Drawing.Point(6, 47)
        Me.chkCompareOrdinal.Name = "chkCompareOrdinal"
        Me.chkCompareOrdinal.Size = New System.Drawing.Size(422, 19)
        Me.chkCompareOrdinal.TabIndex = 0
        Me.chkCompareOrdinal.Text = "&Ordinal: Las mayúsculas antes de las minúsculas (debe distinguir may/min)"
        Me.toolTip1.SetToolTip(Me.chkCompareOrdinal, "Si al clasificar se ponen las mayúsculas antes de las minúsculas," & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "debe distingui" &
        "r mayúsculas de minúsculas.")
        Me.chkCompareOrdinal.UseVisualStyleBackColor = True
        '
        'chkMostrarLineasHTML
        '
        Me.chkMostrarLineasHTML.AutoSize = True
        Me.chkMostrarLineasHTML.Location = New System.Drawing.Point(18, 87)
        Me.chkMostrarLineasHTML.Name = "chkMostrarLineasHTML"
        Me.chkMostrarLineasHTML.Size = New System.Drawing.Size(270, 19)
        Me.chkMostrarLineasHTML.TabIndex = 3
        Me.chkMostrarLineasHTML.Text = "Mostrar números de línea al colorear en HTML"
        Me.toolTip1.SetToolTip(Me.chkMostrarLineasHTML, "Si al colorear en HTML se deben mostrar los números de línea")
        Me.chkMostrarLineasHTML.UseVisualStyleBackColor = True
        '
        'grbAlEvaluar
        '
        Me.grbAlEvaluar.Controls.Add(Me.chkColorearEvaluar)
        Me.grbAlEvaluar.Controls.Add(Me.chkCompilarEvaluar)
        Me.grbAlEvaluar.Location = New System.Drawing.Point(12, 112)
        Me.grbAlEvaluar.Name = "grbAlEvaluar"
        Me.grbAlEvaluar.Size = New System.Drawing.Size(511, 77)
        Me.grbAlEvaluar.TabIndex = 4
        Me.grbAlEvaluar.TabStop = False
        Me.grbAlEvaluar.Text = "Al evaluar:"
        '
        'chkColorearEvaluar
        '
        Me.chkColorearEvaluar.AutoSize = True
        Me.chkColorearEvaluar.Location = New System.Drawing.Point(6, 22)
        Me.chkColorearEvaluar.Name = "chkColorearEvaluar"
        Me.chkColorearEvaluar.Size = New System.Drawing.Size(123, 19)
        Me.chkColorearEvaluar.TabIndex = 0
        Me.chkColorearEvaluar.Text = "Colorear el código"
        Me.toolTip1.SetToolTip(Me.chkColorearEvaluar, "Si se colorea el código al Evaluar")
        Me.chkColorearEvaluar.UseVisualStyleBackColor = True
        '
        'chkCompilarEvaluar
        '
        Me.chkCompilarEvaluar.AutoSize = True
        Me.chkCompilarEvaluar.Location = New System.Drawing.Point(6, 47)
        Me.chkCompilarEvaluar.Name = "chkCompilarEvaluar"
        Me.chkCompilarEvaluar.Size = New System.Drawing.Size(126, 19)
        Me.chkCompilarEvaluar.TabIndex = 0
        Me.chkCompilarEvaluar.Text = "Comprobar errores"
        Me.toolTip1.SetToolTip(Me.chkCompilarEvaluar, "Si se comprueban errores al Evaluar" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(se pre-compilará el código para saber los e" &
        "rrores producidos)")
        Me.chkCompilarEvaluar.UseVisualStyleBackColor = True
        '
        'chkColorearCargar
        '
        Me.chkColorearCargar.AutoSize = True
        Me.chkColorearCargar.Location = New System.Drawing.Point(18, 37)
        Me.chkColorearCargar.Name = "chkColorearCargar"
        Me.chkColorearCargar.Size = New System.Drawing.Size(171, 19)
        Me.chkColorearCargar.TabIndex = 1
        Me.chkColorearCargar.Text = "&Colorear al cargar el fichero"
        Me.chkColorearCargar.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        Me.toolTip1.SetToolTip(Me.chkColorearCargar, "Si al abrir un fichero se debe colorear" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10))
        Me.chkColorearCargar.UseVisualStyleBackColor = True
        '
        'chkCargarUltimo
        '
        Me.chkCargarUltimo.AutoSize = True
        Me.chkCargarUltimo.Location = New System.Drawing.Point(18, 12)
        Me.chkCargarUltimo.Name = "chkCargarUltimo"
        Me.chkCargarUltimo.Size = New System.Drawing.Size(233, 19)
        Me.chkCargarUltimo.TabIndex = 0
        Me.chkCargarUltimo.Text = "Al iniciar cargar el último fichero &usado"
        Me.toolTip1.SetToolTip(Me.chkCargarUltimo, "Si al iniciar la aplicación se carga el último fichero " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "abierto en la sesión ant" &
        "erior" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10))
        Me.chkCargarUltimo.UseVisualStyleBackColor = True
        '
        'tabPage2
        '
        Me.tabPage2.Controls.Add(Me.groupBox1)
        Me.tabPage2.Location = New System.Drawing.Point(4, 24)
        Me.tabPage2.Name = "tabPage2"
        Me.tabPage2.Padding = New System.Windows.Forms.Padding(6)
        Me.tabPage2.Size = New System.Drawing.Size(541, 289)
        Me.tabPage2.TabIndex = 1
        Me.tabPage2.Text = "Ficheros recientes"
        Me.tabPage2.UseVisualStyleBackColor = True
        '
        'groupBox1
        '
        Me.groupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.groupBox1.Controls.Add(Me.txtFic)
        Me.groupBox1.Controls.Add(Me.btnOrdenar)
        Me.groupBox1.Controls.Add(Me.btnQuitar)
        Me.groupBox1.Controls.Add(Me.lstFics)
        Me.groupBox1.Controls.Add(Me.label1)
        Me.groupBox1.Location = New System.Drawing.Point(12, 12)
        Me.groupBox1.Margin = New System.Windows.Forms.Padding(12)
        Me.groupBox1.Name = "groupBox1"
        Me.groupBox1.Size = New System.Drawing.Size(517, 265)
        Me.groupBox1.TabIndex = 0
        Me.groupBox1.TabStop = False
        Me.groupBox1.Text = "Lista de ficheros abiertos recientemente"
        '
        'txtFic
        '
        Me.txtFic.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtFic.Location = New System.Drawing.Point(6, 236)
        Me.txtFic.Name = "txtFic"
        Me.txtFic.Size = New System.Drawing.Size(266, 23)
        Me.txtFic.TabIndex = 2
        '
        'btnOrdenar
        '
        Me.btnOrdenar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOrdenar.Image = CType(resources.GetObject("btnOrdenar.Image"), System.Drawing.Image)
        Me.btnOrdenar.Location = New System.Drawing.Point(303, 235)
        Me.btnOrdenar.Name = "btnOrdenar"
        Me.btnOrdenar.Size = New System.Drawing.Size(75, 23)
        Me.btnOrdenar.TabIndex = 3
        Me.btnOrdenar.Text = "Ordenar"
        Me.btnOrdenar.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnOrdenar.UseVisualStyleBackColor = True
        '
        'btnQuitar
        '
        Me.btnQuitar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnQuitar.AutoSize = True
        Me.btnQuitar.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnQuitar.Image = CType(resources.GetObject("btnQuitar.Image"), System.Drawing.Image)
        Me.btnQuitar.Location = New System.Drawing.Point(393, 234)
        Me.btnQuitar.Name = "btnQuitar"
        Me.btnQuitar.Size = New System.Drawing.Size(118, 25)
        Me.btnQuitar.TabIndex = 4
        Me.btnQuitar.Text = "Quitar selección"
        Me.btnQuitar.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnQuitar.UseVisualStyleBackColor = True
        '
        'lstFics
        '
        Me.lstFics.Font = New System.Drawing.Font("Consolas", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.lstFics.FormattingEnabled = True
        Me.lstFics.ItemHeight = 14
        Me.lstFics.Location = New System.Drawing.Point(6, 45)
        Me.lstFics.Name = "lstFics"
        Me.lstFics.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.lstFics.Size = New System.Drawing.Size(505, 158)
        Me.lstFics.TabIndex = 1
        '
        'label1
        '
        Me.label1.Location = New System.Drawing.Point(6, 19)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(505, 23)
        Me.label1.TabIndex = 0
        Me.label1.Text = "Selecciona los que quieras borrar o clasifícalos usando los botones"
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.grbColores)
        Me.TabPage3.Location = New System.Drawing.Point(4, 24)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(541, 289)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Colores y fuente"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'grbColores
        '
        Me.grbColores.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grbColores.Controls.Add(Me.txtIndentar)
        Me.grbColores.Controls.Add(Me.Label3)
        Me.grbColores.Controls.Add(Me.comboFuenteTamaño)
        Me.grbColores.Controls.Add(Me.comboFuenteNombre)
        Me.grbColores.Controls.Add(Me.labelTamaño)
        Me.grbColores.Controls.Add(Me.labelFuente)
        Me.grbColores.Location = New System.Drawing.Point(12, 12)
        Me.grbColores.Name = "grbColores"
        Me.grbColores.Size = New System.Drawing.Size(517, 265)
        Me.grbColores.TabIndex = 0
        Me.grbColores.TabStop = False
        Me.grbColores.Text = "Colorear y fuentes"
        '
        'txtIndentar
        '
        Me.txtIndentar.Location = New System.Drawing.Point(135, 80)
        Me.txtIndentar.Name = "txtIndentar"
        Me.txtIndentar.Size = New System.Drawing.Size(30, 23)
        Me.txtIndentar.TabIndex = 8
        Me.txtIndentar.Text = "4"
        Me.txtIndentar.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(6, 83)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(123, 20)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Espacios indentar:"
        '
        'comboFuenteTamaño
        '
        Me.comboFuenteTamaño.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.comboFuenteTamaño.Items.AddRange(New Object() {"8", "9", "10", "11", "12", "13", "14", "16", "18"})
        Me.comboFuenteTamaño.Location = New System.Drawing.Point(79, 51)
        Me.comboFuenteTamaño.Name = "comboFuenteTamaño"
        Me.comboFuenteTamaño.Size = New System.Drawing.Size(121, 23)
        Me.comboFuenteTamaño.TabIndex = 3
        '
        'comboFuenteNombre
        '
        Me.comboFuenteNombre.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.comboFuenteNombre.Items.AddRange(New Object() {"Courier New", "Consolas", "Segoe UI", "Lucida Console", "Arial", "Verdana", "Comic Sans MS"})
        Me.comboFuenteNombre.Location = New System.Drawing.Point(79, 22)
        Me.comboFuenteNombre.Name = "comboFuenteNombre"
        Me.comboFuenteNombre.Size = New System.Drawing.Size(121, 23)
        Me.comboFuenteNombre.TabIndex = 1
        '
        'labelTamaño
        '
        Me.labelTamaño.Location = New System.Drawing.Point(6, 54)
        Me.labelTamaño.Margin = New System.Windows.Forms.Padding(3)
        Me.labelTamaño.Name = "labelTamaño"
        Me.labelTamaño.Size = New System.Drawing.Size(67, 20)
        Me.labelTamaño.TabIndex = 2
        Me.labelTamaño.Text = "Tamaño:"
        '
        'labelFuente
        '
        Me.labelFuente.Location = New System.Drawing.Point(6, 25)
        Me.labelFuente.Margin = New System.Windows.Forms.Padding(3)
        Me.labelFuente.Name = "labelFuente"
        Me.labelFuente.Size = New System.Drawing.Size(67, 20)
        Me.labelFuente.TabIndex = 0
        Me.labelFuente.Text = "Fuente:"
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.grbBuscarReemplazar)
        Me.TabPage4.Location = New System.Drawing.Point(4, 24)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage4.Size = New System.Drawing.Size(541, 289)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Buscar / Reemplazar"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'grbBuscarReemplazar
        '
        Me.grbBuscarReemplazar.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grbBuscarReemplazar.Controls.Add(Me.chkWholeWord)
        Me.grbBuscarReemplazar.Controls.Add(Me.chkMatchCase)
        Me.grbBuscarReemplazar.Controls.Add(Me.grbReemplazar)
        Me.grbBuscarReemplazar.Controls.Add(Me.grbBuscar)
        Me.grbBuscarReemplazar.Controls.Add(Me.Label2)
        Me.grbBuscarReemplazar.Location = New System.Drawing.Point(12, 12)
        Me.grbBuscarReemplazar.Margin = New System.Windows.Forms.Padding(12)
        Me.grbBuscarReemplazar.Name = "grbBuscarReemplazar"
        Me.grbBuscarReemplazar.Size = New System.Drawing.Size(517, 265)
        Me.grbBuscarReemplazar.TabIndex = 0
        Me.grbBuscarReemplazar.TabStop = False
        Me.grbBuscarReemplazar.Text = "Opciones de  buscar y reemplazar"
        '
        'chkWholeWord
        '
        Me.chkWholeWord.AutoSize = True
        Me.chkWholeWord.Image = CType(resources.GetObject("chkWholeWord.Image"), System.Drawing.Image)
        Me.chkWholeWord.Location = New System.Drawing.Point(6, 225)
        Me.chkWholeWord.Name = "chkWholeWord"
        Me.chkWholeWord.Size = New System.Drawing.Size(198, 19)
        Me.chkWholeWord.TabIndex = 4
        Me.chkWholeWord.Text = "Comprobar palabra completa"
        Me.chkWholeWord.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.chkWholeWord.UseVisualStyleBackColor = True
        '
        'chkMatchCase
        '
        Me.chkMatchCase.AutoSize = True
        Me.chkMatchCase.Image = CType(resources.GetObject("chkMatchCase.Image"), System.Drawing.Image)
        Me.chkMatchCase.Location = New System.Drawing.Point(6, 200)
        Me.chkMatchCase.Name = "chkMatchCase"
        Me.chkMatchCase.Size = New System.Drawing.Size(240, 19)
        Me.chkMatchCase.TabIndex = 3
        Me.chkMatchCase.Text = "Comprobar mayúsculas y minúsculas"
        Me.chkMatchCase.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.chkMatchCase.UseVisualStyleBackColor = True
        '
        'grbReemplazar
        '
        Me.grbReemplazar.Controls.Add(Me.lstReemplazar)
        Me.grbReemplazar.Controls.Add(Me.btnQuitarReemplazar)
        Me.grbReemplazar.Controls.Add(Me.btnOrdenarReemplazar)
        Me.grbReemplazar.Location = New System.Drawing.Point(262, 45)
        Me.grbReemplazar.Name = "grbReemplazar"
        Me.grbReemplazar.Size = New System.Drawing.Size(250, 149)
        Me.grbReemplazar.TabIndex = 2
        Me.grbReemplazar.TabStop = False
        Me.grbReemplazar.Text = "Reemplazar"
        '
        'lstReemplazar
        '
        Me.lstReemplazar.Font = New System.Drawing.Font("Consolas", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.lstReemplazar.FormattingEnabled = True
        Me.lstReemplazar.ItemHeight = 14
        Me.lstReemplazar.Location = New System.Drawing.Point(6, 20)
        Me.lstReemplazar.Name = "lstReemplazar"
        Me.lstReemplazar.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.lstReemplazar.Size = New System.Drawing.Size(238, 88)
        Me.lstReemplazar.TabIndex = 0
        '
        'btnQuitarReemplazar
        '
        Me.btnQuitarReemplazar.AutoSize = True
        Me.btnQuitarReemplazar.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnQuitarReemplazar.Image = CType(resources.GetObject("btnQuitarReemplazar.Image"), System.Drawing.Image)
        Me.btnQuitarReemplazar.Location = New System.Drawing.Point(126, 118)
        Me.btnQuitarReemplazar.Name = "btnQuitarReemplazar"
        Me.btnQuitarReemplazar.Size = New System.Drawing.Size(118, 25)
        Me.btnQuitarReemplazar.TabIndex = 2
        Me.btnQuitarReemplazar.Text = "Quitar selección"
        Me.btnQuitarReemplazar.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnQuitarReemplazar.UseVisualStyleBackColor = True
        '
        'btnOrdenarReemplazar
        '
        Me.btnOrdenarReemplazar.AutoSize = True
        Me.btnOrdenarReemplazar.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnOrdenarReemplazar.Image = CType(resources.GetObject("btnOrdenarReemplazar.Image"), System.Drawing.Image)
        Me.btnOrdenarReemplazar.Location = New System.Drawing.Point(6, 118)
        Me.btnOrdenarReemplazar.Name = "btnOrdenarReemplazar"
        Me.btnOrdenarReemplazar.Size = New System.Drawing.Size(76, 25)
        Me.btnOrdenarReemplazar.TabIndex = 1
        Me.btnOrdenarReemplazar.Text = "Ordenar"
        Me.btnOrdenarReemplazar.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnOrdenarReemplazar.UseVisualStyleBackColor = True
        '
        'grbBuscar
        '
        Me.grbBuscar.Controls.Add(Me.lstBuscar)
        Me.grbBuscar.Controls.Add(Me.btnQuitarBuscar)
        Me.grbBuscar.Controls.Add(Me.btnOrdenarBuscar)
        Me.grbBuscar.Location = New System.Drawing.Point(6, 45)
        Me.grbBuscar.Name = "grbBuscar"
        Me.grbBuscar.Size = New System.Drawing.Size(250, 149)
        Me.grbBuscar.TabIndex = 1
        Me.grbBuscar.TabStop = False
        Me.grbBuscar.Text = "Buscar"
        '
        'lstBuscar
        '
        Me.lstBuscar.Font = New System.Drawing.Font("Consolas", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.lstBuscar.FormattingEnabled = True
        Me.lstBuscar.ItemHeight = 14
        Me.lstBuscar.Location = New System.Drawing.Point(6, 20)
        Me.lstBuscar.Name = "lstBuscar"
        Me.lstBuscar.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.lstBuscar.Size = New System.Drawing.Size(238, 88)
        Me.lstBuscar.TabIndex = 0
        '
        'btnQuitarBuscar
        '
        Me.btnQuitarBuscar.AutoSize = True
        Me.btnQuitarBuscar.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnQuitarBuscar.Image = CType(resources.GetObject("btnQuitarBuscar.Image"), System.Drawing.Image)
        Me.btnQuitarBuscar.Location = New System.Drawing.Point(126, 118)
        Me.btnQuitarBuscar.Name = "btnQuitarBuscar"
        Me.btnQuitarBuscar.Size = New System.Drawing.Size(118, 25)
        Me.btnQuitarBuscar.TabIndex = 2
        Me.btnQuitarBuscar.Text = "Quitar selección"
        Me.btnQuitarBuscar.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnQuitarBuscar.UseVisualStyleBackColor = True
        '
        'btnOrdenarBuscar
        '
        Me.btnOrdenarBuscar.AutoSize = True
        Me.btnOrdenarBuscar.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnOrdenarBuscar.Image = CType(resources.GetObject("btnOrdenarBuscar.Image"), System.Drawing.Image)
        Me.btnOrdenarBuscar.Location = New System.Drawing.Point(6, 118)
        Me.btnOrdenarBuscar.Name = "btnOrdenarBuscar"
        Me.btnOrdenarBuscar.Size = New System.Drawing.Size(76, 25)
        Me.btnOrdenarBuscar.TabIndex = 1
        Me.btnOrdenarBuscar.Text = "Ordenar"
        Me.btnOrdenarBuscar.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnOrdenarBuscar.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.Location = New System.Drawing.Point(6, 19)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(505, 23)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Selecciona lo que quieras borrar o clasifícalos usando los botones"
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.GroupBox2)
        Me.TabPage5.Location = New System.Drawing.Point(4, 24)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage5.Size = New System.Drawing.Size(541, 289)
        Me.TabPage5.TabIndex = 4
        Me.TabPage5.Text = "Edición"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.GroupBox3)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(517, 265)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Recortes de edición"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.lstRecortes)
        Me.GroupBox3.Controls.Add(Me.btnQuitarRecortes)
        Me.GroupBox3.Controls.Add(Me.btnOrdenarRecortes)
        Me.GroupBox3.Location = New System.Drawing.Point(6, 45)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(505, 149)
        Me.GroupBox3.TabIndex = 3
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Recortes"
        '
        'lstRecortes
        '
        Me.lstRecortes.Font = New System.Drawing.Font("Consolas", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.lstRecortes.FormattingEnabled = True
        Me.lstRecortes.ItemHeight = 14
        Me.lstRecortes.Location = New System.Drawing.Point(6, 20)
        Me.lstRecortes.Name = "lstRecortes"
        Me.lstRecortes.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.lstRecortes.Size = New System.Drawing.Size(493, 88)
        Me.lstRecortes.TabIndex = 0
        '
        'btnQuitarRecortes
        '
        Me.btnQuitarRecortes.AutoSize = True
        Me.btnQuitarRecortes.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnQuitarRecortes.Image = CType(resources.GetObject("btnQuitarRecortes.Image"), System.Drawing.Image)
        Me.btnQuitarRecortes.Location = New System.Drawing.Point(381, 118)
        Me.btnQuitarRecortes.Name = "btnQuitarRecortes"
        Me.btnQuitarRecortes.Size = New System.Drawing.Size(118, 25)
        Me.btnQuitarRecortes.TabIndex = 2
        Me.btnQuitarRecortes.Text = "Quitar selección"
        Me.btnQuitarRecortes.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnQuitarRecortes.UseVisualStyleBackColor = True
        '
        'btnOrdenarRecortes
        '
        Me.btnOrdenarRecortes.AutoSize = True
        Me.btnOrdenarRecortes.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnOrdenarRecortes.Image = CType(resources.GetObject("btnOrdenarRecortes.Image"), System.Drawing.Image)
        Me.btnOrdenarRecortes.Location = New System.Drawing.Point(6, 118)
        Me.btnOrdenarRecortes.Name = "btnOrdenarRecortes"
        Me.btnOrdenarRecortes.Size = New System.Drawing.Size(76, 25)
        Me.btnOrdenarRecortes.TabIndex = 1
        Me.btnOrdenarRecortes.Text = "Ordenar"
        Me.btnOrdenarRecortes.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnOrdenarRecortes.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.Location = New System.Drawing.Point(6, 19)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(505, 23)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Selecciona lo que quieras borrar o clasifícalos usando los botones"
        '
        'btnAceptar
        '
        Me.btnAceptar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAceptar.Location = New System.Drawing.Point(405, 335)
        Me.btnAceptar.Name = "btnAceptar"
        Me.btnAceptar.Size = New System.Drawing.Size(75, 23)
        Me.btnAceptar.TabIndex = 1
        Me.btnAceptar.Text = "Aceptar"
        Me.btnAceptar.UseVisualStyleBackColor = True
        '
        'btnCancelar
        '
        Me.btnCancelar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancelar.Location = New System.Drawing.Point(486, 335)
        Me.btnCancelar.Name = "btnCancelar"
        Me.btnCancelar.Size = New System.Drawing.Size(75, 23)
        Me.btnCancelar.TabIndex = 2
        Me.btnCancelar.Text = "Cancelar"
        Me.btnCancelar.UseVisualStyleBackColor = True
        '
        'chkCambiarTab
        '
        Me.chkCambiarTab.AutoSize = True
        Me.chkCambiarTab.Location = New System.Drawing.Point(18, 62)
        Me.chkCambiarTab.Name = "chkCambiarTab"
        Me.chkCambiarTab.Size = New System.Drawing.Size(277, 19)
        Me.chkCambiarTab.TabIndex = 2
        Me.chkCambiarTab.Text = "Al cargar o guardar, cambiar TAB por 8 espacios"
        Me.chkCambiarTab.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        Me.toolTip1.SetToolTip(Me.chkCambiarTab, "Si al abrir un fichero se deben cambiar los TAB por espacios." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Si está selecciona" &
        "da, también se quitarán al guardar." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10))
        Me.chkCambiarTab.UseVisualStyleBackColor = True
        '
        'FormOpciones
        '
        Me.AcceptButton = Me.btnAceptar
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancelar
        Me.ClientSize = New System.Drawing.Size(573, 370)
        Me.Controls.Add(Me.btnAceptar)
        Me.Controls.Add(Me.btnCancelar)
        Me.Controls.Add(Me.tabControl1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimumSize = New System.Drawing.Size(550, 400)
        Me.Name = "FormOpciones"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Opciones"
        Me.tabControl1.ResumeLayout(False)
        Me.tabPage1.ResumeLayout(False)
        Me.tabPage1.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.grbAlEvaluar.ResumeLayout(False)
        Me.grbAlEvaluar.PerformLayout()
        Me.tabPage2.ResumeLayout(False)
        Me.groupBox1.ResumeLayout(False)
        Me.groupBox1.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        Me.grbColores.ResumeLayout(False)
        Me.grbColores.PerformLayout()
        Me.TabPage4.ResumeLayout(False)
        Me.grbBuscarReemplazar.ResumeLayout(False)
        Me.grbBuscarReemplazar.PerformLayout()
        Me.grbReemplazar.ResumeLayout(False)
        Me.grbReemplazar.PerformLayout()
        Me.grbBuscar.ResumeLayout(False)
        Me.grbBuscar.PerformLayout()
        Me.TabPage5.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend chkCargarUltimo As System.Windows.Forms.CheckBox
    Private btnAceptar As System.Windows.Forms.Button
    Private btnCancelar As System.Windows.Forms.Button
    Private btnOrdenar As System.Windows.Forms.Button
    Private btnOrdenarBuscar As Button
    Private btnOrdenarReemplazar As Button
    Private btnQuitar As System.Windows.Forms.Button
    Private btnQuitarBuscar As Button
    Private btnQuitarReemplazar As Button
    Private grbAlEvaluar As GroupBox
    Private grbBuscar As GroupBox
    Private grbBuscarReemplazar As GroupBox
    Private grbReemplazar As GroupBox
    Private groupBox1 As System.Windows.Forms.GroupBox
    Private label1 As System.Windows.Forms.Label
    Private Label2 As Label
    Private tabControl1 As System.Windows.Forms.TabControl
    Private tabPage1 As System.Windows.Forms.TabPage
    Private tabPage2 As System.Windows.Forms.TabPage
    Private TabPage3 As TabPage
    Private TabPage4 As TabPage
    Private toolTip1 As ToolTip
    Private txtFic As System.Windows.Forms.TextBox
    Private grbColores As GroupBox
    Private labelTamaño As Label
    Private labelFuente As Label
    Private WithEvents TabPage5 As TabPage
    Private WithEvents GroupBox2 As GroupBox
    Private WithEvents GroupBox3 As GroupBox
    Private WithEvents btnQuitarRecortes As Button
    Private WithEvents btnOrdenarRecortes As Button
    Private WithEvents Label4 As Label
    Private WithEvents txtIndentar As TextBox
    Private WithEvents Label3 As Label
    Private WithEvents lstBuscar As ListBox
    Private WithEvents lstFics As ListBox
    Private WithEvents lstReemplazar As ListBox
    Private WithEvents comboFuenteNombre As ComboBox
    Private WithEvents comboFuenteTamaño As ComboBox
    Private WithEvents lstRecortes As ListBox
    Private WithEvents GroupBox4 As GroupBox
    Friend WithEvents chkIgnoreCase As CheckBox
    Private WithEvents chkCaseSensitive As CheckBox
    Private WithEvents chkCompareOrdinal As CheckBox
    Private WithEvents chkColorearEvaluar As CheckBox
    Private WithEvents chkCompilarEvaluar As CheckBox
    Private WithEvents chkColorearCargar As CheckBox
    Private WithEvents chkMatchCase As CheckBox
    Private WithEvents chkWholeWord As CheckBox
    Private WithEvents chkMostrarLineasHTML As CheckBox
    Private chkCambiarTab As CheckBox
End Class
