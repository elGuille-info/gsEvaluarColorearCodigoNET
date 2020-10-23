'------------------------------------------------------------------------------
' Marcador, clase para los marcadores (bookmark) y última posición  (20/Oct/20)
'
' Usarla para los marcadores (Bookmarks)
' Posiciones en las ventanas abiertas
'
'
' (c) Guillermo (elGuille) Som, 2020
'------------------------------------------------------------------------------
Option Strict On
Option Infer On

Imports System
'Imports System.Data
Imports System.Collections.Generic
Imports System.Text
Imports System.Linq
Imports Microsoft.VisualBasic
'Imports vb = Microsoft.VisualBasic

Public Class Marcadores

    Public Sub New(frm As Form1)
        Me.Formulario = frm
        Me.Posicion = frm.richTextBoxCodigo.SelectionStart
        Posiciones.Clear()
    End Sub

    ''' <summary>
    ''' El fichero al que hace referencia la lista de Posiciones.
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Fichero As String
        Get
            Return Formulario.nombreFichero
        End Get
    End Property

    ''' <summary>
    ''' La posición actual del texto en el richTextBoxCodigo
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Posicion As Integer

    ''' <summary>
    ''' El formulario (<see cref="Form1"/>) al que hacen referencia las posiciones.
    ''' </summary>
    ''' <returns></returns>
    Public Property Formulario As Form1

    ''' <summary>
    ''' Las posiciones para el formulario indicado.
    ''' </summary>
    ''' <returns></returns>
    Public Property Posiciones As New List(Of Integer)

    'Private colPosiciones As New Dictionary(Of String, List(Of Integer))

    'Default Public ReadOnly Property Item(fichero As String) As List(Of Integer)
    '    Get
    '        If Not colPosiciones.Keys.Contains(fichero) Then
    '            colPosiciones.Add(fichero, New List(Of Integer))
    '        End If
    '        Return colPosiciones(fichero)
    '    End Get
    'End Property

End Class
