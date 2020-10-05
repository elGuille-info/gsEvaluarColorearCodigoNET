'------------------------------------------------------------------------------
' Ventana con los recortes                                          (30/Sep/20)
'
'
' (c) Guillermo (elGuille) Som, 2020
'------------------------------------------------------------------------------
Option Strict On
Option Infer On

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic

Public Class FormRecortes

    ''' <summary>
    ''' El texto seleccionado de la lista.
    ''' </summary>
    ''' <returns></returns>
    Public Property TextoSeleccionado As String

    Private colRecortes As New List(Of String)

    ''' <summary>
    ''' Crear una ventana en la posición indicada y con los recortes.
    ''' </summary>
    ''' <param name="posicion">La posición donde se mostrará la ventana.</param>
    ''' <param name="recortes">Colección con los recortes.</param>
    ''' <param name="maxRecortes">El número máximo de recortes.</param>
    Public Sub New(posicion As Point, recortes As List(Of String), maxRecortes As Integer)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        Me.Left = posicion.X
        Me.Top = posicion.Y
        Me.colRecortes.Clear()
        Me.colRecortes.AddRange(recortes)

        Dim _48Espacios = New String(" "c, 48)
        Dim _40Espacios = New String(" "c, 40)

        lstRecortes.Items.Clear()
        ' El primer y último item no se usan,
        ' son solo para separar los recortes de los bordes
        lstRecortes.Items.Add($"{_48Espacios} ")
        For i = 0 To maxRecortes - 1
            If i < colRecortes.Count Then
                lstRecortes.Items.Add($"{(i + 1).ToString("0").PadLeft(6)}: {colRecortes(i).PadRight(40)} ")
            Else
                lstRecortes.Items.Add($"{(i + 1).ToString("0").PadLeft(6)}: {_40Espacios} ")
            End If
        Next
        lstRecortes.Items.Add($"{_48Espacios} ")
        If lstRecortes.Items.Count > 0 Then
            lstRecortes.SelectedIndex = 1
        End If
    End Sub

    Private Sub FormRecortes_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.KeyPreview = True

    End Sub

    Private Sub FormRecortes_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Me.KeyPress
        ' Si se pulsa ESC cerrar la ventana                         (04/Oct/20)
        If e.KeyChar = ChrW(27) Then
            e.Handled = True
            Me.DialogResult = DialogResult.Cancel
        End If
    End Sub

    Private Sub FormRecortes_Click(sender As Object, e As EventArgs) Handles Me.Click
        Me.DialogResult = DialogResult.Cancel
    End Sub

    Private Sub lstRecortes_MouseMove(sender As Object, e As MouseEventArgs) Handles lstRecortes.MouseMove
        If lstRecortes.Items.Count = 0 Then Return

        Dim point = New Point(e.X, e.Y)
        Dim i = lstRecortes.IndexFromPoint(point)

        ' No aceptar el primero ni el último
        If i < 1 OrElse i = lstRecortes.Items.Count - 1 Then
            Return
        End If
        If i < lstRecortes.Items.Count Then
            lstRecortes.SelectedIndex = i
        End If

        If i < colRecortes.Count Then
            Try
                toolTip1.SetToolTip(lstRecortes, colRecortes(i - 1))
            Catch ex As Exception
                'toolTip1.SetToolTip(lstRecortes, $"<índice Recortes({i}) fuera de rango>")
            End Try
        Else
            toolTip1.SetToolTip(lstRecortes, "")
        End If

    End Sub

    Private Sub lstRecortes_MouseClick(sender As Object, e As MouseEventArgs) Handles lstRecortes.MouseClick
        Dim i = lstRecortes.SelectedIndex
        ' No aceptar el primero ni el último
        If i < 1 OrElse i = lstRecortes.Items.Count - 1 Then
            Return
        End If

        ' Hay que usar el anterior,                                 (04/Oct/20)
        ' ya que el primero (0) está en blanco
        TextoSeleccionado = colRecortes(i - 1)

        Me.DialogResult = DialogResult.OK
    End Sub

    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        Me.DialogResult = DialogResult.Cancel
    End Sub

End Class