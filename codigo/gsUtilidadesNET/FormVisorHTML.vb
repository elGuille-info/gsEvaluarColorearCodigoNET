'------------------------------------------------------------------------------
' Formulario para ver el código HTML                                (24/Sep/20)
'
' El control WebBrowser no se puede manejar en el diseñador de formularios.
' Todo hay que hacerlo en el .Designer
'
' (c) Guillermo (elGuille) Som, 2020
'------------------------------------------------------------------------------
Option Strict On
Option Infer On

Imports Microsoft.VisualBasic
Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports System.Linq

''' <summary>
''' Formulario visor de páginas HTML
''' </summary>
Public Class FormVisorHTML

    ''' <summary>
    ''' El constructor al que se le pasará el título de la ventana y el código HTML a mostrar
    ''' </summary>
    ''' <param name="titulo">Título a poner en la ventana del navegador</param>
    ''' <param name="codigoHTML">El código HTML a mostrar en el navegador</param>
    ''' <remarks>Se crea un fichero temporal llamado HTMLTemp.html.</remarks>
    Sub New(titulo As String, codigoHTML As String)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.Text = titulo

        Dim ficTmp As String = Path.Combine(Path.GetTempPath(), "HTMLTemp.html")

        ' Para el formato de lectura y escritura de los ficheros    (03/Nov/20)
        Dim enc As Encoding
        If FormatoEncoding = FormatosEncoding.UTF8 Then
            enc = Encoding.UTF8
        ElseIf FormatoEncoding = FormatosEncoding.Default Then
            enc = Encoding.Default
        Else
            enc = Encoding.Latin1
        End If
        Using sw As New StreamWriter(ficTmp, False, enc)
            sw.WriteLine(codigoHTML)
            sw.Flush()
            sw.Close()
        End Using
        Me.WebBrowser1.Navigate(New Uri(ficTmp))

    End Sub

    Private Sub FormVisorHTML_Load(sender As Object, e As EventArgs) Handles Me.Load
        Width = CInt(Screen.PrimaryScreen.Bounds.Width * 0.45)
        Height = CInt(Screen.PrimaryScreen.Bounds.Height * 0.65)

        Me.CenterToScreen()

    End Sub

End Class