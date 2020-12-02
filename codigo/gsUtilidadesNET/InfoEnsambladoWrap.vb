'------------------------------------------------------------------------------
' InfoEnsambladoWrap                                                (02/Dic/20)
' Clase que llama a los métodos de InfoEnsamblado en Mostrar contenido ensamblado
'
'
' (c) Guillermo (elGuille) Som, 2020
'------------------------------------------------------------------------------
Option Strict On
Option Infer On

Imports System
'Imports System.Data
'Imports System.Collections.Generic
'Imports System.Text
'Imports System.Linq
'Imports Microsoft.VisualBasic
'Imports vb = Microsoft.VisualBasic

''' <summary>
''' Esta clase accede a <see cref="InfoEnsamblado"/> de Mostrar contenido ensamblado.
''' Expone los métodos y propiedades compartidas de esa clase.
''' </summary>
Public Class InfoEnsambladoWrap

    ''' <summary>
    ''' Acceso a <see cref="InfoEnsamblado.ReturnValue"/>.
    ''' </summary>
    ''' <returns></returns>
    Public Shared Property ReturnValue As Integer
        Get
            Return InfoEnsamblado.ReturnValue
        End Get
        Set(value As Integer)
            InfoEnsamblado.ReturnValue = value
        End Set
    End Property

    ''' <summary>
    ''' Acceso a <see cref="InfoEnsamblado.GuardarInfo"/>.
    ''' </summary>
    ''' <param name="args"></param>
    ''' <param name="fic"></param>
    ''' <returns></returns>
    ''' <remarks>El formato usado para guardar es Latin1</remarks>
    Public Shared Function GuardarInfo(args As String(), fic As String) As Boolean
        Return InfoEnsamblado.GuardarInfo(args, fic)
    End Function

    ''' <summary>
    ''' Acceso a <see cref="InfoEnsamblado.InfoTipo"/>.
    ''' </summary>
    ''' <param name="args"></param>
    ''' <param name="mostrarComandos"></param>
    ''' <returns></returns>
    Public Shared Function InfoTipo(args As String(), Optional mostrarComandos As Boolean = False) As String
        Return InfoEnsamblado.InfoTipo(args, mostrarComandos)
    End Function

    ''' <summary>
    ''' Acceso a <see cref="InfoEnsamblado.MostrarAyuda"/>.
    ''' </summary>
    ''' <param name="mostrarEnConsola"></param>
    ''' <param name="esperar"></param>
    ''' <returns></returns>
    Public Shared Function MostrarAyuda(mostrarEnConsola As Boolean, esperar As Boolean) As String
        Return InfoEnsamblado.MostrarAyuda(mostrarEnConsola, esperar)
    End Function

End Class
