'------------------------------------------------------------------------------
' Opciones de la aplicación de compilar y ejecutar
'
'
' (c) Guillermo (elGuille) Som, 2020
'------------------------------------------------------------------------------
Option Strict On
Option Infer On
#If ESX86 Then
Imports gsUtilidadesNETx86
#Else
Imports gsUtilidadesNET
#End If

Imports Microsoft.VisualBasic

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics.Tracing

Public Class FormOpciones

    ''' <summary>
    ''' Crear una nueva instancia usando los valores de <see cref="Form1"/>
    ''' </summary>
    ''' <param name="elForm">Un objeto del tipo <see cref="Form1"/> para editar los cambios</param>
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        lstFics.Items.Clear()
        lstFics.Items.AddRange(UltimosFicherosAbiertos.ToArray)
        lstBuscar.Items.Clear()
        lstBuscar.Items.AddRange(ComboBoxBuscar.ToList().ToArray)
        lstReemplazar.Items.Clear()
        lstReemplazar.Items.AddRange(ComboBoxReemplazar.ToList().ToArray)

        chkCargarUltimo.Checked = cargarUltimo
        chkColorearCargar.Checked = colorearAlCargar
        chkColorearEvaluar.Checked = CurrentMDI.buttonColorearAlEvaluar.Checked
        chkCompilarEvaluar.Checked = CurrentMDI.buttonCompilarAlEvaluar.Checked
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
        chkAvisarBuscar.Checked = avisarFinBusqueda

        cboMacros.Items.Clear()
        cboMacros.Items.AddRange(CurrentMDI.colMacros.ToArray)
        cboMacros.Text = CurrentMDI.MacroActual
        chkHabilitarEjecutarMacro.Checked = CurrentMDI.HabilitarEjecutarMacro

    End Sub


    Private Sub FormOpciones_Load(sender As Object, e As EventArgs) Handles Me.Load

        Me.CenterToScreen()

    End Sub

    Private Sub FormOpciones_FormClosing(sender As Object,
                                         e As FormClosingEventArgs) Handles Me.FormClosing

        ' Si el usuario cierra el formulario desde la "x"
        ' se considera que se cancela
        If e.CloseReason = CloseReason.UserClosing Then
            Me.DialogResult = DialogResult.Cancel
        End If
    End Sub

    Private Sub LstFics_KeyDown(sender As Object,
                                e As KeyEventArgs) Handles lstFics.KeyDown, lstBuscar.KeyDown, lstReemplazar.KeyDown, lstRecortes.KeyDown
        If e.KeyCode = Keys.Delete Then
            e.Handled = True
            BorrarLosSeleccionados(TryCast(sender, ListBox), Nothing)
        End If
    End Sub

    Private Sub LstFics_SelectedIndexChanged(sender As Object,
                                             e As EventArgs) Handles lstFics.SelectedIndexChanged
        txtFic.Text = If(lstFics.SelectedItem Is Nothing,
                            txtFic.Text,
                            lstFics.SelectedItem.ToString
                            )
    End Sub

    ''' <summary>
    ''' Borrar los elementos seleccionados de la lista indicada.
    ''' En realidad el que hace el envío es el botón que se ha pulsado.
    ''' </summary>
    Private Sub BorrarLosSeleccionados(sender As Object,
                                       e As EventArgs) Handles btnQuitar.Click, btnQuitarBuscar.Click, btnQuitarReemplazar.Click, btnQuitarRecortes.Click
        Dim lista As ListBox '= Nothing
        If sender Is btnQuitar Then
            lista = lstFics
        ElseIf sender Is btnQuitarBuscar Then
            lista = lstBuscar
        ElseIf sender Is btnQuitarRecortes Then
            lista = lstRecortes
        ElseIf sender Is btnQuitarReemplazar Then
            lista = lstReemplazar
        Else
            lista = TryCast(sender, ListBox)
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
    Private Sub OrdenarLosSeleccionados(sender As Object, e As EventArgs) Handles btnOrdenar.Click, btnOrdenarBuscar.Click, btnOrdenarReemplazar.Click, btnOrdenarRecortes.Click
        Dim lista As ListBox '= Nothing
        If sender Is btnOrdenar Then
            lista = lstFics
        ElseIf sender Is btnOrdenarBuscar Then
            lista = lstBuscar
        ElseIf sender Is btnOrdenarRecortes Then
            lista = lstRecortes
        ElseIf sender Is btnOrdenarReemplazar Then
            lista = lstReemplazar
        Else
            lista = TryCast(sender, ListBox)
        End If

        If lista Is Nothing Then Return
        If lista.SelectedIndices.Count = 0 Then Return

        Dim col = lista.ToList()
        col.Sort()
        lista.Items.AddRange(col.ToArray)
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles btnCancelar.Click
        Me.DialogResult = DialogResult.Cancel
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        cargarUltimo = chkCargarUltimo.Checked
        colorearAlCargar = chkColorearCargar.Checked
        buscarMatchCase = chkMatchCase.Checked
        buscarWholeWord = chkWholeWord.Checked
        UltimosFicherosAbiertos.Clear()
        UltimosFicherosAbiertos.AddRange(lstFics.ToList().ToArray)
        ComboBoxBuscar.Items.Clear()
        ComboBoxBuscar.Items.AddRange(lstBuscar.ToList().ToArray)
        ComboBoxReemplazar.Items.Clear()
        ComboBoxReemplazar.Items.AddRange(lstReemplazar.ToList().ToArray)
        MostrarLineasHTML = chkMostrarLineasHTML.Checked
        EspaciosIndentar = txtIndentar.AsInteger
        fuenteNombre = comboFuenteNombre.Text
        fuenteTamaño = comboFuenteTamaño.Text
        ColRecortes.Clear()
        ColRecortes.AddRange(lstRecortes.ToList)
        CompararString.IgnoreCase = chkCaseSensitive.Checked
        CompararString.CompareOrdinal = chkCompareOrdinal.Checked
        CambiarTabs = chkCambiarTab.Checked
        avisarFinBusqueda = chkAvisarBuscar.Checked

        CurrentMDI.MacroActual = cboMacros.Text
        CurrentMDI.colMacros.AddRange(cboMacros.ToList)

        CurrentMDI.HabilitarEjecutarMacro = chkHabilitarEjecutarMacro.Checked

        Me.DialogResult = DialogResult.OK

    End Sub

    Private Sub CboMacros_Validating(sender As Object, e As ComponentModel.CancelEventArgs) Handles cboMacros.Validating
        Dim s = cboMacros.Text
        ' Comprobar si tiene CTRL o SHIFT o ALT
        Dim j = s.IndexOf("CTRL", StringComparison.OrdinalIgnoreCase)
        If j > -1 Then
            s = s.Replace("CTRL", "^", StringComparison.OrdinalIgnoreCase)
        End If
        j = s.IndexOf("ALT", StringComparison.OrdinalIgnoreCase)
        If j > -1 Then
            s = s.Replace("ALT", "%", StringComparison.OrdinalIgnoreCase)
        End If
        j = s.IndexOf("SHIFT", StringComparison.OrdinalIgnoreCase)
        If j > -1 Then
            s = s.Replace("SHIFT", "+", StringComparison.OrdinalIgnoreCase)
        End If
        If Not cboMacros.Items.Contains(s) Then
            cboMacros.Items.Add(s)
        End If
    End Sub
End Class