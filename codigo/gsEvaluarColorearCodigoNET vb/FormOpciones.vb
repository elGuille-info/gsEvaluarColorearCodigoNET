'------------------------------------------------------------------------------
' Opciones de la aplicación de compilar y ejecutar
'
'
' (c) Guillermo (elGuille) Som, 2020
'------------------------------------------------------------------------------
Option Strict On
Option Infer On

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic

Public Class FormOpciones


    '''' <summary>
    '''' Si se deben cambiar de línea si el texto sobrepasa el borde derecho
    '''' </summary>
    '''' <returns></returns>
    '''' <remarks>Por ahora siempre será false (17/Sep/20)</remarks>
    'Private Property WordWrap As Boolean
    '    Get
    '        Return False 'chkWordWrap.Checked
    '    End Get
    '    Set(value As Boolean)
    '        chkWordWrap.Checked = False ' value
    '    End Set
    'End Property

    'Private elForm As Form1

    ''' <summary>
    ''' Crear una nueva instancia usando los valores de <see cref="Form1"/>
    ''' </summary>
    ''' <param name="elForm">Un objeto del tipo <see cref="Form1"/> para editar los cambios</param>
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        'Me.elForm = elForm

        ' Add any initialization after the InitializeComponent() call.
        lstFics.Items.Clear()
        lstFics.Items.AddRange(UltimosFicherosAbiertos.ToArray)
        lstBuscar.Items.Clear()
        lstBuscar.Items.AddRange(comboBoxBuscar.ToList().ToArray)
        lstReemplazar.Items.Clear()
        lstReemplazar.Items.AddRange(comboBoxReemplazar.ToList().ToArray)

        chkCargarUltimo.Checked = cargarUltimo
        chkColorearCargar.Checked = colorearAlCargar
        chkColorearEvaluar.Checked = buttonColorearAlEvaluar.Checked
        chkCompilarEvaluar.Checked = buttonCompilarAlEvaluar.Checked
        chkMatchCase.Checked = buscarMatchCase
        chkWholeWord.Checked = buscarWholeWord
        chkMostrarLineasHTML.Checked = MostrarLineasHTML
        txtIndentar.Text = EspaciosIndentar.ToString
        comboFuenteNombre.Text = fuenteNombre
        comboFuenteTamaño.Text = fuenteTamaño
        lstRecortes.Items.Clear()
        lstRecortes.Items.AddRange(ColRecortes.ToArray)
        chkCaseSensitive.Checked = CompararString.IgnoreCase
        chkCompareOrdinal.Checked = CompararString.CompareOrdinal
        chkCambiarTab.Checked = CambiarTabs

    End Sub


    Private Sub FormOpciones_Load(sender As Object, e As EventArgs) Handles Me.Load
        'chkCargarUltimo.Checked = CargarUltimo
        'chkColorearCargar.Checked = ColorearAlCargar
        'chkWordWrap.Checked = WordWrap

        AñadirEventos()
        'AsignarItems()

    End Sub

    Private Sub FormOpciones_FormClosing(sender As Object,
                                         e As FormClosingEventArgs) Handles Me.FormClosing

        ' Si el usuario cierra el formulario desde la "x"
        ' se considera que se cancela
        If e.CloseReason = CloseReason.UserClosing Then
            Me.DialogResult = DialogResult.Cancel
        End If
    End Sub

    ''' <summary>
    ''' Crear los métodos de evento del formulario
    ''' </summary>
    Private Sub AñadirEventos()
        Dim buttonHandlerQuitar = Sub(sender As Object, e As EventArgs) BorrarLosSeleccionados(sender)
        Dim buttonHandlerOrdenar = Sub(sender As Object, e As EventArgs) OrdenarLosSeleccionados(sender)

        AddHandler lstFics.KeyDown, AddressOf lstFics_KeyDown
        AddHandler lstBuscar.KeyDown, AddressOf lstFics_KeyDown
        AddHandler lstReemplazar.KeyDown, AddressOf lstFics_KeyDown
        AddHandler lstRecortes.KeyDown, AddressOf lstFics_KeyDown

        AddHandler btnCancelar.Click, Sub() Me.DialogResult = DialogResult.Cancel
        AddHandler btnAceptar.Click, Sub()
                                         cargarUltimo = chkCargarUltimo.Checked
                                         colorearAlCargar = chkColorearCargar.Checked
                                         buscarMatchCase = chkMatchCase.Checked
                                         buscarWholeWord = chkWholeWord.Checked
                                         UltimosFicherosAbiertos.Clear()
                                         UltimosFicherosAbiertos.AddRange(lstFics.ToList().ToArray)
                                         comboBoxBuscar.Items.Clear()
                                         comboBoxBuscar.Items.AddRange(lstBuscar.ToList().ToArray)
                                         comboBoxReemplazar.Items.Clear()
                                         comboBoxReemplazar.Items.AddRange(lstReemplazar.ToList().ToArray)
                                         MostrarLineasHTML = chkMostrarLineasHTML.Checked
                                         EspaciosIndentar = txtIndentar.AsInteger
                                         fuenteNombre = comboFuenteNombre.Text
                                         fuenteTamaño = comboFuenteTamaño.Text
                                         ColRecortes.Clear()
                                         ColRecortes.AddRange(lstRecortes.ToList)
                                         CompararString.IgnoreCase = chkCaseSensitive.Checked
                                         CompararString.CompareOrdinal = chkCompareOrdinal.Checked
                                         CambiarTabs = chkCambiarTab.Checked

                                         Me.DialogResult = DialogResult.OK
                                     End Sub

        AddHandler btnQuitar.Click, buttonHandlerQuitar
        AddHandler btnQuitarBuscar.Click, buttonHandlerQuitar
        AddHandler btnQuitarReemplazar.Click, buttonHandlerQuitar
        AddHandler btnQuitarRecortes.Click, buttonHandlerQuitar

        AddHandler btnOrdenar.Click, buttonHandlerOrdenar
        AddHandler btnOrdenarBuscar.Click, buttonHandlerOrdenar
        AddHandler btnOrdenarReemplazar.Click, buttonHandlerOrdenar
        AddHandler btnOrdenarRecortes.Click, buttonHandlerOrdenar

        AddHandler lstFics.SelectedIndexChanged,
                    Sub() txtFic.Text = If(lstFics.SelectedItem Is Nothing,
                                            txtFic.Text,
                                            lstFics.SelectedItem.ToString
                                            )

    End Sub

    Private Sub lstFics_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Delete Then
            e.Handled = True
            BorrarLosSeleccionados(TryCast(sender, ListBox))
        End If
    End Sub

    ''' <summary>
    ''' Borrar los elementos seleccionados de la lista indicada.
    ''' En realidad el que hace el envío es el botón que se ha pulsado.
    ''' </summary>
    Private Sub BorrarLosSeleccionados(sender As Object)
        Dim lista As ListBox = Nothing
        If sender Is btnOrdenar Then
            lista = lstFics
        ElseIf sender Is btnOrdenarBuscar Then
            lista = lstBuscar
        ElseIf sender Is btnOrdenarRecortes Then
            lista = lstRecortes
        ElseIf sender Is btnOrdenarReemplazar Then
            lista = lstReemplazar
        End If

        If lista Is Nothing Then Return
        If lista.SelectedIndices.Count = 0 Then Return

        For i = lista.SelectedIndices.Count - 1 To 0 Step -1
            lista.Items.RemoveAt(lista.SelectedIndices(i))
        Next
    End Sub

    ''' <summary>
    ''' Ordenar los elementos de la lista indicada.
    ''' En realidad el parámetro es del botón que se ha pulsado
    ''' </summary>
    Private Sub OrdenarLosSeleccionados(sender As Object)
        Dim lista As ListBox = Nothing
        If sender Is btnOrdenar Then
            lista = lstFics
        ElseIf sender Is btnOrdenarBuscar Then
            lista = lstBuscar
        ElseIf sender Is btnOrdenarRecortes Then
            lista = lstRecortes
        ElseIf sender Is btnOrdenarReemplazar Then
            lista = lstReemplazar
        End If

        If lista Is Nothing Then Return
        If lista.SelectedIndices.Count = 0 Then Return

        Dim col = lista.ToList() '.Sort
        col.Sort()
        lista.Items.AddRange(col.ToArray)
    End Sub

End Class