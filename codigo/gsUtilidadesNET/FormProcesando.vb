'------------------------------------------------------------------------------
' Formulario para cuando se está procesando los datos               (25/Jul/08)
'
' Ahora está en gsGestNETUtil                                       (27/Jul/08)
'
' ©Guillermo 'guille' Som, 2008, 2020
'------------------------------------------------------------------------------
Option Strict On

Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports System
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Linq

''' <summary>
''' Formulario para mostrar un diálogo cuando se está procesando (acciones largas)
''' </summary>
Public Class FormProcesando

    Private ReadOnly inicializando As Boolean = True

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        inicializando = False
    End Sub

    ''' <summary>
    ''' El estilo de la barra de progreso.
    ''' </summary>
    ''' <returns></returns>
    Public Property EstiloProgreso() As ProgressBarStyle
        Get
            Return progressBar1.Style
        End Get
        Set(ByVal value As ProgressBarStyle)
            progressBar1.Style = value
        End Set
    End Property

    ''' <summary>
    ''' El título de la ventana.
    ''' </summary>
    ''' <returns></returns>
    Public Property Texto() As String
        Get
            Return Me.Text
        End Get
        Set(ByVal value As String)
            Me.Text = value
        End Set
    End Property

    'Private m_Mensaje1 As String
    ''' <summary>
    ''' Por si se necesita hacer una copia de la propiedad Mensaje.
    ''' </summary>
    Public Property Mensaje1() As String

    ''' <summary>
    ''' El mensaje a mostrar (máximo dos líneas).
    ''' </summary>
    ''' <returns></returns>
    Public Property Mensaje() As String
        Get
            Return Me.txtInfo.Text
        End Get
        Set(ByVal value As String)
            txtInfo.Text = value
            Me.Refresh()
            Application.DoEvents()
        End Set
    End Property

    ''' <summary>
    ''' El valor máximo de la barra de progreso.
    ''' </summary>
    ''' <returns></returns>
    Public Property Maximo() As Integer '= If(Me.progressBar1 Is Nothing, 100, Me.progressBar1.Maximum)
        Get
            If inicializando Then Return 100
            Return Me.progressBar1.Maximum
        End Get
        Set(value As Integer)
            Me.progressBar1.Maximum = value
        End Set
    End Property

    ''' <summary>
    ''' El valor actual de la barra de progreso.
    ''' Se ignora si se añade un valor mayor que el máximo.
    ''' </summary>
    ''' <returns></returns>
    Public Property ValorActual() As Integer
        Get
            If inicializando Then Return 0
            'Return If(Me.progressBar1 Is Nothing, 0, Me.progressBar1.Value)
            Return Me.progressBar1.Value
        End Get
        Set(ByVal value As Integer)
            If value <= Me.progressBar1.Maximum Then
                Me.progressBar1.Value = value
            End If
            Me.Refresh()
            Application.DoEvents()
        End Set
    End Property

    ''' <summary>
    ''' El porcentaje restante.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>17/Oct/2020</remarks>
    Public ReadOnly Property PorcentajeRestante As Integer
        Get
            If inicializando Then Return 1
            Return If(Maximo < 1, 1, ValorActual * 100 \ Maximo)
        End Get
    End Property

    ''' <summary>
    ''' Si se ha pulsado en el botón Cancelar.
    ''' </summary>
    ''' <returns></returns>
    Public Property Cancelar() As Boolean '= False

    Private Sub BtnCancelar_Click(ByVal sender As Object,
                                  ByVal e As EventArgs) _
                                  Handles btnCancelar.Click
        Cancelar = True
        Application.DoEvents()
    End Sub
End Class