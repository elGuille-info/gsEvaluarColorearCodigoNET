<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FormRecortes
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
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormRecortes))
        Me.labelPortapapeles = New System.Windows.Forms.Label()
        Me.lstRecortes = New System.Windows.Forms.ListBox()
        Me.toolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnCerrar = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'labelPortapapeles
        '
        Me.labelPortapapeles.BackColor = System.Drawing.Color.AliceBlue
        Me.labelPortapapeles.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.labelPortapapeles.Font = New System.Drawing.Font("Consolas", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.labelPortapapeles.Location = New System.Drawing.Point(0, 0)
        Me.labelPortapapeles.Margin = New System.Windows.Forms.Padding(0)
        Me.labelPortapapeles.Name = "labelPortapapeles"
        Me.labelPortapapeles.Size = New System.Drawing.Size(116, 27)
        Me.labelPortapapeles.TabIndex = 0
        Me.labelPortapapeles.Text = " Recortes"
        Me.labelPortapapeles.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lstRecortes
        '
        Me.lstRecortes.BackColor = System.Drawing.Color.AliceBlue
        Me.lstRecortes.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstRecortes.Font = New System.Drawing.Font("Consolas", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.lstRecortes.FormattingEnabled = True
        Me.lstRecortes.ItemHeight = 18
        Me.lstRecortes.Items.AddRange(New Object() {"", "     1: 123456789.123456789.123456789.1234567895", "     2- .", "     3- .", "     4- .", "     5- .", "     6- .", "     7- .", "     8- .", "     9- .", "    10- ."})
        Me.lstRecortes.Location = New System.Drawing.Point(0, 27)
        Me.lstRecortes.Margin = New System.Windows.Forms.Padding(0)
        Me.lstRecortes.Name = "lstRecortes"
        Me.lstRecortes.Size = New System.Drawing.Size(449, 236)
        Me.lstRecortes.TabIndex = 1
        '
        'toolTip1
        '
        Me.toolTip1.AutoPopDelay = 9000
        Me.toolTip1.BackColor = System.Drawing.Color.AliceBlue
        Me.toolTip1.InitialDelay = 500
        Me.toolTip1.ReshowDelay = 100
        '
        'btnCerrar
        '
        Me.btnCerrar.FlatAppearance.BorderSize = 0
        Me.btnCerrar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCerrar.Image = CType(resources.GetObject("btnCerrar.Image"), System.Drawing.Image)
        Me.btnCerrar.Location = New System.Drawing.Point(428, -1)
        Me.btnCerrar.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.btnCerrar.Name = "btnCerrar"
        Me.btnCerrar.Size = New System.Drawing.Size(23, 18)
        Me.btnCerrar.TabIndex = 3
        Me.btnCerrar.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        '
        'FormRecortes
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Window
        Me.ClientSize = New System.Drawing.Size(449, 278)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnCerrar)
        Me.Controls.Add(Me.lstRecortes)
        Me.Controls.Add(Me.labelPortapapeles)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormRecortes"
        Me.Opacity = 0.85R
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "FormRecortes"
        Me.ResumeLayout(False)

    End Sub

    Private WithEvents labelPortapapeles As Label
    Private WithEvents lstRecortes As ListBox
    Private toolTip1 As ToolTip
    Private WithEvents btnCerrar As Button
End Class
