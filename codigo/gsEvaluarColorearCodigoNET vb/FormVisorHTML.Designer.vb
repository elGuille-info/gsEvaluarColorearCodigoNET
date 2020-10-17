<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FormVisorHTML
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
        Me.WebBrowser1 = New System.Windows.Forms.WebBrowser()
        Me.panelBrowser = New System.Windows.Forms.Panel()
        Me.SuspendLayout()
        '
        'WebBrowser1
        '
        Me.WebBrowser1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.WebBrowser1.Location = New System.Drawing.Point(6, 6)
        Me.WebBrowser1.Name = "WebBrowser1"
        Me.WebBrowser1.Size = New System.Drawing.Size(787, 438)
        Me.WebBrowser1.TabIndex = 1
        '
        'panelBrowser
        '
        Me.panelBrowser.BackColor = System.Drawing.Color.DeepSkyBlue
        Me.panelBrowser.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.panelBrowser.Dock = System.Windows.Forms.DockStyle.Fill
        Me.panelBrowser.Location = New System.Drawing.Point(0, 0)
        Me.panelBrowser.Name = "panelBrowser"
        Me.panelBrowser.Size = New System.Drawing.Size(800, 450)
        Me.panelBrowser.TabIndex = 1
        '
        'FormVisorHTML
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.WebBrowser1)
        Me.Controls.Add(Me.panelBrowser)
        Me.Name = "FormVisorHTML"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FormVisorHTML"
        Me.ResumeLayout(False)

    End Sub

    Private WebBrowser1 As System.Windows.Forms.WebBrowser
    Private panelBrowser As Panel
End Class
