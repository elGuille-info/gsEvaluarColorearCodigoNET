'------------------------------------------------------------------------------
' Para realizar clasificaciones sensibles a mayúsculas/minúsculas   (01/Oct/20)
'
' Esta clase la usaba en gsEditor(/ 2008) en el 2005
' Se usará de esta forma:
'   CompararString.IgnoreCase = clasif_caseSensitive
'   CompararString.UsarCompareOrdinal = clasif_compareOrdinal
'   Array.Sort(lineas, 0, j + 1, New CompararString)
'
'
' (c) Guillermo (elGuille) Som, 2005, 2020
'------------------------------------------------------------------------------
Option Strict On
Option Infer On

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic

''' <summary>
''' Clase para hacer las comparaciones de cadenas.
''' Se puede usar cuando se necesita un comparador basado en IComparer(Of String).Compare.
''' </summary>
''' <example>
''' Por ejemplo:
''' Array.Sort(lineas, 0, j + 1, New CompararString)
''' </example>
Public Class CompararString
    Implements System.Collections.Generic.IComparer(Of String)

    Public Shared IgnoreCase As Boolean

    Public Shared UsarCompareOrdinal As Boolean

    Public Function Compare(x As String, y As String) As Integer _
                        Implements System.Collections.Generic.IComparer(Of String).Compare

        If UsarCompareOrdinal AndAlso IgnoreCase = False Then
            Return String.CompareOrdinal(x, y)
        End If
        Return String.Compare(x, y, IgnoreCase)
    End Function

End Class

