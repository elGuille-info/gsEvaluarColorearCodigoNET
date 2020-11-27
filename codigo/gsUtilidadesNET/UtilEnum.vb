'------------------------------------------------------------------------------
' Utilidades                                                        (27/Nov/20)
' Clase con la definición de la enumeración FormatosEncoding y
' utilidades para manejar enumeraciones
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
'Imports Microsoft.VisualBasic
'Imports vb = Microsoft.VisualBasic

''' <summary>
''' Clase con la definición de la enumeración FormatosEncoding y utilidades para manejar enumeraciones.
''' </summary>
Public Module UtilEnum

    ''' <summary>
    ''' Para los formatos de los ficheros.
    ''' </summary>
    ''' <remarks>03/Nov/2020</remarks>
    Public Enum FormatosEncoding
        ''' <summary>
        ''' Latin1 character set (ISO-8859-1)
        ''' </summary>
        ''' <remarks>En .NET framework 4.8 no se admite Latin1</remarks>
        Latin1
        ''' <summary>
        ''' UTF-8 format
        ''' </summary>
        UTF8
        ''' <summary>
        ''' The default encoding for this .NET implementation.
        ''' No es recomendable usar Default en .NET 5.0.
        ''' En .NET Framework 4.8 sí usar Default.
        ''' </summary>
        ''' <remarks>On .NET Core, the Default property always returns the UTF8Encoding. 
        ''' UTF-8 is supported on all the operating systems (Windows, Linux, and macOS) 
        ''' on which .NET Core applications run.</remarks>
        [Default]
    End Enum

    ''' <summary>
    ''' El formato a usar al leer/guardar los ficheros.
    ''' </summary>
    ''' <remarks>03/Nov/2020</remarks>
    Public ReadOnly Property FormatoEncoding As FormatosEncoding = FormatosEncoding.Latin1

    ''' <summary>
    ''' Comprueba si el nombre corresponde con un valor de la enumeración.
    ''' </summary>
    ''' <param name="arg">El nombre de la enumeración</param>
    ''' <param name="val">El valor de la enumeración</param>
    ''' <param name="class">La enumeración que define el valor</param>
    ''' <param name="exception">True si no se lanza una excepción y se devuelve False</param>
    ''' <returns>
    ''' Un valor True si el valor está definido en la enumeración.
    ''' Si no está definido y exception es True lanza una excepción, en otro caso devuelve False.
    ''' </returns>
    Public Function CheckValidEnumValue(arg As String, val As Object, [class] As Type,
                                        Optional [exception] As Boolean = False) As Boolean
        ' Si el valor no está definido en la enumeración lanzar una excepción
        If Not [Enum].IsDefined([class], val) Then
            If [exception] Then
                Throw New ComponentModel.InvalidEnumArgumentException(arg, CInt(val), [class])
            Else
                Return False
            End If
        End If

        Return True
    End Function
End Module
