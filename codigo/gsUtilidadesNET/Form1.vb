'------------------------------------------------------------------------------
' Form1                                                             (27/Nov/20)
' Formulario sin funcionalidad para crear esta biblioteca como aplicación de Windows
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
Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Me.Width > Screen.PrimaryScreen.WorkingArea.Width Then
            Me.Width = Screen.PrimaryScreen.WorkingArea.Width - 24
        End If
        If Me.Height > Screen.PrimaryScreen.WorkingArea.Height Then
            Me.Height = Screen.PrimaryScreen.WorkingArea.Height - 24
        End If

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
End Class
