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


    ''' <summary>
    ''' Si se deben cambiar de línea si el texto sobrepasa el borde derecho
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>Por ahora siempre será false (17/Sep/20)</remarks>
    Private Property WordWrap As Boolean
        Get
            Return False 'chkWordWrap.Checked
        End Get
        Set(value As Boolean)
            chkWordWrap.Checked = False ' value
        End Set
    End Property

    Private elForm1 As Form1

    ''' <summary>
    ''' Crear una nueva instancia usando los valores de <see cref="Form1"/>
    ''' </summary>
    ''' <param name="elForm">Un objeto del tipo <see cref="Form1"/> para editar los cambios</param>
    Public Sub New(elForm As Form1)

        ' This call is required by the designer.
        InitializeComponent()

        elForm1 = elForm

        ' Add any initialization after the InitializeComponent() call.
        lstFics.Items.Clear()
        lstFics.Items.AddRange(elForm.comboBoxFileName.ToList().ToArray)
        lstBuscar.Items.Clear()
        lstBuscar.Items.AddRange(elForm.comboBoxBuscar.ToList().ToArray)
        lstReemplazar.Items.Clear()
        lstReemplazar.Items.AddRange(elForm.comboBoxReemplazar.ToList().ToArray)

        chkCargarUltimo.Checked = elForm.cargarUltimo
        chkColorearCargar.Checked = elForm.colorearAlCargar
        chkColorearEvaluar.Checked = elForm.buttonColorearAlEvaluar.Checked
        chkCompilarEvaluar.Checked = elForm.buttonCompilarAlEvaluar.Checked
        chkMatchCase.Checked = elForm.buscarMatchCase
        chkWholeWord.Checked = elForm.buscarWholeWord
        chkMostrarLineasHTML.Checked = elForm.mostrarLineasHTML
        txtIndentar.Text = elForm.EspaciosIndentar.ToString
        comboFuenteNombre.Text = elForm.fuenteNombre
        comboFuenteTamaño.Text = elForm.fuenteTamaño

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
        Dim buttonHandlerQuitar = Sub(sender As Object, e As EventArgs) BorrarLosSeleccionados(TryCast(sender, ListBox))
        Dim buttonHandlerOrdenar = Sub(sender As Object, e As EventArgs) OrdenarLosSeleccionados(TryCast(sender, ListBox))

        AddHandler lstFics.KeyDown, AddressOf lstFics_KeyDown
        AddHandler lstBuscar.KeyDown, AddressOf lstFics_KeyDown
        AddHandler lstReemplazar.KeyDown, AddressOf lstFics_KeyDown

        AddHandler btnCancelar.Click, Sub() Me.DialogResult = DialogResult.Cancel
        AddHandler btnAceptar.Click, Sub()
                                         elForm1.cargarUltimo = chkCargarUltimo.Checked
                                         elForm1.colorearAlCargar = chkColorearCargar.Checked
                                         elForm1.buscarMatchCase = chkMatchCase.Checked
                                         elForm1.buscarWholeWord = chkWholeWord.Checked
                                         elForm1.comboBoxFileName.Items.Clear()
                                         elForm1.comboBoxFileName.Items.AddRange(lstFics.ToList().ToArray)
                                         elForm1.comboBoxBuscar.Items.Clear()
                                         elForm1.comboBoxBuscar.Items.AddRange(lstBuscar.ToList().ToArray)
                                         elForm1.comboBoxReemplazar.Items.Clear()
                                         elForm1.comboBoxReemplazar.Items.AddRange(lstReemplazar.ToList().ToArray)
                                         elForm1.mostrarLineasHTML = chkMostrarLineasHTML.Checked
                                         elForm1.EspaciosIndentar = txtIndentar.AsInteger
                                         elForm1.fuenteNombre = comboFuenteNombre.Text
                                         elForm1.fuenteTamaño = comboFuenteTamaño.Text

                                         Me.DialogResult = DialogResult.OK
                                     End Sub

        AddHandler btnQuitar.Click, buttonHandlerQuitar ' Sub(sender As Object, e As EventArgs) BorrarLosSeleccionados(TryCast(sender, ListBox))
        AddHandler btnQuitarBuscar.Click, buttonHandlerQuitar
        AddHandler btnQuitarReemplazar.Click, buttonHandlerQuitar

        AddHandler btnOrdenar.Click, buttonHandlerOrdenar
        AddHandler btnOrdenarBuscar.Click, buttonHandlerOrdenar
        AddHandler btnOrdenarReemplazar.Click, buttonHandlerOrdenar

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
    ''' </summary>
    Private Sub BorrarLosSeleccionados(lista As ListBox)
        If lista Is Nothing Then Return
        If lista.SelectedIndices.Count = 0 Then Return

        For i = lista.SelectedIndices.Count - 1 To 0 Step -1
            lista.Items.RemoveAt(lista.SelectedIndices(i))
        Next
    End Sub

    ''' <summary>
    ''' Ordenar los elementos de la lista indicada.
    ''' </summary>
    ''' <param name="lista"></param>
    Private Sub OrdenarLosSeleccionados(lista As ListBox)
        If lista Is Nothing Then Return
        If lista.SelectedIndices.Count = 0 Then Return

        Dim col = lista.ToList() '.Sort
        col.Sort()
        lista.Items.AddRange(col.ToArray)
    End Sub

End Class