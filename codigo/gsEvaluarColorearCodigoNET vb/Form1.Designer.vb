<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Friend Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.statusContextMenu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.menuCopiarPath = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuRecargarFichero = New System.Windows.Forms.ToolStripMenuItem()
        Me.statusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.labelModificado = New System.Windows.Forms.ToolStripStatusLabel()
        Me.labelInfo = New System.Windows.Forms.ToolStripStatusLabel()
        Me.labelFuente = New System.Windows.Forms.ToolStripStatusLabel()
        Me.buttonLenguaje = New System.Windows.Forms.ToolStripSplitButton()
        Me.menuVB = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuCS = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuTxt = New System.Windows.Forms.ToolStripMenuItem()
        Me.labelTamaño = New System.Windows.Forms.ToolStripStatusLabel()
        Me.labelPos = New System.Windows.Forms.ToolStripStatusLabel()
        Me.richTextBoxCodigo = New System.Windows.Forms.RichTextBox()
        Me.richTextBoxLineas = New System.Windows.Forms.RichTextBox()
        Me.splitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.lstSyntax = New System.Windows.Forms.ListBox()
        Me.statusContextMenu.SuspendLayout()
        Me.statusStrip1.SuspendLayout()
        CType(Me.splitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitContainer1.Panel1.SuspendLayout()
        Me.splitContainer1.Panel2.SuspendLayout()
        Me.splitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'statusContextMenu
        '
        Me.statusContextMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuCopiarPath, Me.menuRecargarFichero})
        Me.statusContextMenu.Name = "StatusContextMenu"
        Me.statusContextMenu.Size = New System.Drawing.Size(309, 48)
        '
        'menuCopiarPath
        '
        Me.menuCopiarPath.Image = CType(resources.GetObject("menuCopiarPath.Image"), System.Drawing.Image)
        Me.menuCopiarPath.Name = "menuCopiarPath"
        Me.menuCopiarPath.Size = New System.Drawing.Size(308, 22)
        Me.menuCopiarPath.Text = "Copiar PATH completo"
        Me.menuCopiarPath.ToolTipText = "Copia en el portapapeles la ruta completa del fichero actual"
        '
        'menuRecargarFichero
        '
        Me.menuRecargarFichero.Image = CType(resources.GetObject("menuRecargarFichero.Image"), System.Drawing.Image)
        Me.menuRecargarFichero.Name = "menuRecargarFichero"
        Me.menuRecargarFichero.Size = New System.Drawing.Size(308, 22)
        Me.menuRecargarFichero.Text = "Recargar el fichero (sin guardar los cambios)"
        Me.menuRecargarFichero.ToolTipText = "Abre nuevamente el fichero actual desechando los cambios realizados"
        '
        'statusStrip1
        '
        Me.statusStrip1.ContextMenuStrip = Me.statusContextMenu
        Me.statusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.labelModificado, Me.labelInfo, Me.labelFuente, Me.buttonLenguaje, Me.labelTamaño, Me.labelPos})
        Me.statusStrip1.Location = New System.Drawing.Point(0, 553)
        Me.statusStrip1.Name = "statusStrip1"
        Me.statusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 16, 0)
        Me.statusStrip1.Size = New System.Drawing.Size(1382, 24)
        Me.statusStrip1.TabIndex = 2
        Me.statusStrip1.Text = "StatusStrip1"
        '
        'labelModificado
        '
        Me.labelModificado.BorderSides = System.Windows.Forms.ToolStripStatusLabelBorderSides.Right
        Me.labelModificado.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.labelModificado.Name = "labelModificado"
        Me.labelModificado.Size = New System.Drawing.Size(16, 19)
        Me.labelModificado.Text = "*"
        Me.labelModificado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'labelInfo
        '
        Me.labelInfo.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.labelInfo.Font = New System.Drawing.Font("Consolas", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.labelInfo.Name = "labelInfo"
        Me.labelInfo.Size = New System.Drawing.Size(1087, 19)
        Me.labelInfo.Spring = True
        Me.labelInfo.Text = "LabelInfo"
        Me.labelInfo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'labelFuente
        '
        Me.labelFuente.BorderSides = CType((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me.labelFuente.Name = "labelFuente"
        Me.labelFuente.Size = New System.Drawing.Size(77, 19)
        Me.labelFuente.Text = "Consolas; 11"
        '
        'buttonLenguaje
        '
        Me.buttonLenguaje.AutoToolTip = False
        Me.buttonLenguaje.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.buttonLenguaje.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuVB, Me.menuCS, Me.menuTxt})
        Me.buttonLenguaje.Image = CType(resources.GetObject("buttonLenguaje.Image"), System.Drawing.Image)
        Me.buttonLenguaje.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.buttonLenguaje.Name = "buttonLenguaje"
        Me.buttonLenguaje.Size = New System.Drawing.Size(32, 22)
        Me.buttonLenguaje.Text = "Lenguajes"
        Me.buttonLenguaje.ToolTipText = "Lenguajes"
        '
        'menuVB
        '
        Me.menuVB.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.menuVB.Image = CType(resources.GetObject("menuVB.Image"), System.Drawing.Image)
        Me.menuVB.Name = "menuVB"
        Me.menuVB.Size = New System.Drawing.Size(89, 22)
        Me.menuVB.Text = "VB"
        '
        'menuCS
        '
        Me.menuCS.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.menuCS.Image = CType(resources.GetObject("menuCS.Image"), System.Drawing.Image)
        Me.menuCS.Name = "menuCS"
        Me.menuCS.Size = New System.Drawing.Size(89, 22)
        Me.menuCS.Text = "C#"
        '
        'menuTxt
        '
        Me.menuTxt.Image = CType(resources.GetObject("menuTxt.Image"), System.Drawing.Image)
        Me.menuTxt.Name = "menuTxt"
        Me.menuTxt.Size = New System.Drawing.Size(89, 22)
        Me.menuTxt.Text = "txt"
        '
        'labelTamaño
        '
        Me.labelTamaño.BorderSides = System.Windows.Forms.ToolStripStatusLabelBorderSides.Left
        Me.labelTamaño.Name = "labelTamaño"
        Me.labelTamaño.Size = New System.Drawing.Size(60, 19)
        Me.labelTamaño.Text = "4,466 car."
        '
        'labelPos
        '
        Me.labelPos.BorderSides = System.Windows.Forms.ToolStripStatusLabelBorderSides.Left
        Me.labelPos.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.labelPos.Name = "labelPos"
        Me.labelPos.Size = New System.Drawing.Size(93, 19)
        Me.labelPos.Text = "Lín: 399  Car: 20"
        '
        'richTextBoxCodigo
        '
        Me.richTextBoxCodigo.AllowDrop = True
        Me.richTextBoxCodigo.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.richTextBoxCodigo.DetectUrls = False
        Me.richTextBoxCodigo.Font = New System.Drawing.Font("Consolas", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.richTextBoxCodigo.HideSelection = False
        Me.richTextBoxCodigo.Location = New System.Drawing.Point(70, 0)
        Me.richTextBoxCodigo.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.richTextBoxCodigo.Name = "richTextBoxCodigo"
        Me.richTextBoxCodigo.Size = New System.Drawing.Size(982, 553)
        Me.richTextBoxCodigo.TabIndex = 1
        Me.richTextBoxCodigo.Text = ""
        Me.richTextBoxCodigo.WordWrap = False
        '
        'richTextBoxLineas
        '
        Me.richTextBoxLineas.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.richTextBoxLineas.BackColor = System.Drawing.Color.FromArgb(CType(CType(253, Byte), Integer), CType(CType(253, Byte), Integer), CType(CType(253, Byte), Integer))
        Me.richTextBoxLineas.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.richTextBoxLineas.CausesValidation = False
        Me.richTextBoxLineas.Font = New System.Drawing.Font("Consolas", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.richTextBoxLineas.ForeColor = System.Drawing.Color.Teal
        Me.richTextBoxLineas.Location = New System.Drawing.Point(4, 0)
        Me.richTextBoxLineas.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.richTextBoxLineas.Name = "richTextBoxLineas"
        Me.richTextBoxLineas.ReadOnly = True
        Me.richTextBoxLineas.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.None
        Me.richTextBoxLineas.ShortcutsEnabled = False
        Me.richTextBoxLineas.Size = New System.Drawing.Size(59, 553)
        Me.richTextBoxLineas.TabIndex = 0
        Me.richTextBoxLineas.TabStop = False
        Me.richTextBoxLineas.Text = "   1" & Global.Microsoft.VisualBasic.ChrW(10) & "   2" & Global.Microsoft.VisualBasic.ChrW(10) & "   3" & Global.Microsoft.VisualBasic.ChrW(10) & "   4" & Global.Microsoft.VisualBasic.ChrW(10) & "   5" & Global.Microsoft.VisualBasic.ChrW(10) & "   6" & Global.Microsoft.VisualBasic.ChrW(10) & "   7" & Global.Microsoft.VisualBasic.ChrW(10) & "   8" & Global.Microsoft.VisualBasic.ChrW(10) & "   9" & Global.Microsoft.VisualBasic.ChrW(10) & "  10" & Global.Microsoft.VisualBasic.ChrW(10) & "  11" & Global.Microsoft.VisualBasic.ChrW(10) & "  12" & Global.Microsoft.VisualBasic.ChrW(10) & "  13"
        '
        'splitContainer1
        '
        Me.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.splitContainer1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.splitContainer1.Name = "splitContainer1"
        '
        'splitContainer1.Panel1
        '
        Me.splitContainer1.Panel1.Controls.Add(Me.richTextBoxLineas)
        Me.splitContainer1.Panel1.Controls.Add(Me.richTextBoxCodigo)
        Me.splitContainer1.Panel1MinSize = 300
        '
        'splitContainer1.Panel2
        '
        Me.splitContainer1.Panel2.Controls.Add(Me.lstSyntax)
        Me.splitContainer1.Panel2MinSize = 0
        Me.splitContainer1.Size = New System.Drawing.Size(1382, 553)
        Me.splitContainer1.SplitterDistance = 1049
        Me.splitContainer1.SplitterWidth = 5
        Me.splitContainer1.TabIndex = 8
        '
        'lstSyntax
        '
        Me.lstSyntax.BackColor = System.Drawing.Color.FromArgb(CType(CType(253, Byte), Integer), CType(CType(253, Byte), Integer), CType(CType(253, Byte), Integer))
        Me.lstSyntax.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lstSyntax.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstSyntax.Font = New System.Drawing.Font("Consolas", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.lstSyntax.HorizontalScrollbar = True
        Me.lstSyntax.ItemHeight = 14
        Me.lstSyntax.Location = New System.Drawing.Point(0, 0)
        Me.lstSyntax.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.lstSyntax.Name = "lstSyntax"
        Me.lstSyntax.ScrollAlwaysVisible = True
        Me.lstSyntax.Size = New System.Drawing.Size(328, 553)
        Me.lstSyntax.TabIndex = 3
        Me.lstSyntax.TabStop = False
        '
        'Form1
        '
        Me.AllowDrop = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1382, 577)
        Me.Controls.Add(Me.splitContainer1)
        Me.Controls.Add(Me.statusStrip1)
        Me.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Evaluar el código de C# o VB"
        Me.statusContextMenu.ResumeLayout(False)
        Me.statusStrip1.ResumeLayout(False)
        Me.statusStrip1.PerformLayout()
        Me.splitContainer1.Panel1.ResumeLayout(False)
        Me.splitContainer1.Panel2.ResumeLayout(False)
        CType(Me.splitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    '
    ' Declaraciones de los controles
    '

    '
    ' Los controles del statusBar
    '
    Private statusContextMenu As ContextMenuStrip
    Friend menuCopiarPath As ToolStripMenuItem
    Friend menuRecargarFichero As ToolStripMenuItem
    Friend statusStrip1 As System.Windows.Forms.StatusStrip
    Friend labelFuente As ToolStripStatusLabel
    Friend labelInfo As ToolStripStatusLabel
    Friend labelModificado As ToolStripStatusLabel
    Friend labelPos As ToolStripStatusLabel
    Friend labelTamaño As ToolStripStatusLabel
    Friend buttonLenguaje As ToolStripSplitButton
    Private menuCS As ToolStripMenuItem
    Private menuVB As ToolStripMenuItem
    Private menuTxt As ToolStripMenuItem

    Friend WithEvents richTextBoxCodigo As RichTextBox
    Friend WithEvents richTextBoxLineas As RichTextBox
    Friend WithEvents lstSyntax As ListBox
    Friend WithEvents splitContainer1 As SplitContainer

End Class
