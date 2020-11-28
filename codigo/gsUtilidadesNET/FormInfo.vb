'------------------------------------------------------------------------------
' FormInfo                                                          (27/Nov/20)
'
' Muestra la información de este programa y las clases que expone
'
'
' (c) Guillermo (elGuille) Som, 2020
'------------------------------------------------------------------------------
Option Strict On
Option Infer On

Imports System
'Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.Text
Imports System.Linq
Imports Microsoft.VisualBasic
'Imports vb = Microsoft.VisualBasic

''' <summary>
''' Formulario para mostrar la información de esta utilidad y las clases que la componen.
''' </summary>
Public Class FormInfo
    Private iniciando As Boolean = True

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.CenterToScreen()

        iniciando = False
        FormInfo_Resize(Nothing, Nothing)

        Dim sb As New StringBuilder
        sb.AppendLine("Información:")
        sb.AppendLine("============")
        sb.AppendLine(UtilidadesCompilarColorear.Version)
        sb.AppendLine()
        sb.AppendLine()

        Dim col = UtilidadesCompilarColorear.ClasesExpuestas
        sb.AppendLine("Clases expuestas:")
        sb.AppendLine("=================")
        For i = 0 To col.Count - 1
            sb.AppendLine(col(i))
        Next

        Me.TxtVersion.Text = sb.ToString

    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        Debug.WriteLine($"W: {Me.Width}, H: {Me.Height}")
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub FormInfo_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        If iniciando Then Return

        If Me.Width > Screen.PrimaryScreen.WorkingArea.Width Then
            Me.Width = Screen.PrimaryScreen.WorkingArea.Width - 24
        End If
        If Me.Height > Screen.PrimaryScreen.WorkingArea.Height Then
            Me.Height = Screen.PrimaryScreen.WorkingArea.Height - 24
        End If
    End Sub
End Class
