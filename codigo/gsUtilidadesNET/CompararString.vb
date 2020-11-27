'------------------------------------------------------------------------------
' Para realizar clasificaciones sensibles a mayúsculas/minúsculas   (01/Oct/20)
'
' Esta clase la usaba en gsEditor(/ 2008) en el 2005
' Se usará de esta forma:
'   CompararString.IgnoreCase = clasif_caseSensitive
'   CompararString.UsarCompareOrdinal = clasif_compareOrdinal
'   Array.Sort(lineas, 0, j + 1, New CompararString)
'
' Añado la implementación de IEqualityComparer(Of String)           (05/Oct/20)
'
' (c) Guillermo (elGuille) Som, 2005, 2020
'------------------------------------------------------------------------------
Option Strict On
Option Infer On

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
'Imports System.Diagnostics.CodeAnalysis
'Imports System.Reflection.Metadata

''' <summary>
''' Clase para hacer las comparaciones de cadenas que implementa IComparer(Of String) y IEqualityComparer(Of String).
''' Se puede usar cuando se necesita un comparador basado en IComparer(Of String).Compare.
''' </summary>
''' <remarks>
''' Por ejemplo:
''' Array.Sort(lineas, 0, j + 1, New CompararString)
''' 
''' CompararString.IgnoreCase = True
''' Dim compStr As New CompararString
''' 
'''     If Not colCodigo.Keys.Contains(classifSpan.ClassificationType, compStr) Then
'''         colCodigo.Add(classifSpan.ClassificationType, New Dictionary(Of String, ClassifSpanInfo))
'''     End If
'''
''' </remarks>
Public Class CompararString
    Implements IComparer(Of String)
    Implements IEqualityComparer(Of String)

    ''' <summary>
    ''' Distinguir mayúsculas de minúsculas.
    ''' True si no se hace distinción entre mayúsculas y minúsculas.
    ''' </summary>
    Public Shared Property IgnoreCase As Boolean

    ''' <summary>
    ''' Si se ponen las mayúsculas antes de las minúsculas, 
    ''' debe distinguir mayúsculas de minúsculas (<see cref="IgnoreCase"/> debe ser True).
    ''' True si en la comparación se ponen antes las minúsculas que las mayúsculas.
    ''' </summary>
    Public Shared Property CompareOrdinal As Boolean

    ''' <summary>
    ''' Función que devuelve un valor según el resultado de la comparación.
    ''' Una valor menor de cero si x es menor que y.
    ''' Un valor cero si ambos son iguales.
    ''' Un valor mayor de cero si x es mayor que y.
    ''' </summary>
    ''' <param name="x"></param>
    ''' <param name="y"></param>
    ''' <returns></returns>
    Public Function Compare(x As String, y As String) As Integer Implements IComparer(Of String).Compare

        If CompareOrdinal AndAlso IgnoreCase = False Then
            Return String.CompareOrdinal(x, y)
        End If

        Return String.Compare(x, y, IgnoreCase)
    End Function

    Public Overloads Function Equals(x As String, y As String) As Boolean Implements IEqualityComparer(Of String).Equals
        Return Me.Compare(x, y) = 0
    End Function

    ' Usa DisallowNull de System.Diagnostics.CodeAnalysis
    'Public Overloads Function GetHashCode(<DisallowNull> obj As String) As Integer Implements IEqualityComparer(Of String).GetHashCode
    Public Overloads Function GetHashCode(obj As String) As Integer Implements IEqualityComparer(Of String).GetHashCode
        Return MyBase.GetHashCode
    End Function
End Class

