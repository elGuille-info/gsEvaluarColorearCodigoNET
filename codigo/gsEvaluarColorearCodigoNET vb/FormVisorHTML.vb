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
Imports System.Collections.Generic
Imports System.IO

Public Class FormVisorHTML

    ''' <summary>
    ''' El constructor al que se le pasará el título de la ventana y el código HTML a mostrar
    ''' </summary>
    ''' <param name="titulo">Título a poner en la ventana del navegador</param>
    ''' <param name="codigoHTML">El código HTML a mostrar en el navegador</param>
    ''' <remarks>Se crea un fichero temporal llamado HTMLTemp.html.
    ''' Si no se indica en el &lt;pre> la fuente o el tamaño, se usará Consolas a 11 puntos</remarks>
    Sub New(titulo As String, codigoHTML As String)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.Text = titulo

        Dim ficTmp As String = Path.Combine(Path.GetTempPath(), "HTMLTemp.html")

        Using sw As New StreamWriter(ficTmp, False, System.Text.Encoding.UTF8)
            ' Aquí si incluir el style
            sw.WriteLine("<style>pre{{font-family:{0}; font-size:{1}.0pt;}}</style>", "Consolas", "11")
            sw.WriteLine(codigoHTML)
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