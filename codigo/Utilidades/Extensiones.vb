'------------------------------------------------------------------------------
' Módulo con extensiones para String (y otras clases)               (25/Sep/20)
'
'
' (c) Guillermo (elGuille) Som, 2020
'------------------------------------------------------------------------------
Option Strict On
Option Infer On

Imports Microsoft.VisualBasic
Imports System
Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions
Imports System.Text

Public Module Extensiones

    ' Extensión AsInteger                                           (28/Sep/20)

    ''' <summary>
    ''' Devuelve un valor Integer de la propiedad Text del TextBox indicado.
    ''' </summary>
    ''' <param name="txt">El TextBox a extender</param>
    ''' <remarks>28/sep/2020</remarks>
    <Extension>
    Public Function AsInteger(txt As TextBox) As Integer
        Dim i As Integer = 0

        Integer.TryParse(txt.Text, i)

        Return i
    End Function

    ' Extensión Clonar                                              (27/Sep/20)

    ''' <summary>
    ''' Clonar un menú item del tipo <see cref="ToolStripMenuItem"/>.
    ''' </summary>
    ''' <param name="mnu">El <see cref="ToolStripMenuItem"/> al que se va a asignar el nuevo menú.</param>
    ''' <param name="eClick">Manejador del evento Click.</param>
    ''' <param name="eSelect">Manuejador del evento DropDownOpening.</param>
    ''' <returns>Una nueva copia del tipo <see cref="ToolStripMenuItem"/>.</returns>
    <Extension>
    Public Function Clonar(mnu As ToolStripMenuItem,
                           eClick As EventHandler,
                           Optional eSelect As EventHandler = Nothing) As ToolStripMenuItem
        Dim mnuC As New ToolStripMenuItem
        AddHandler mnuC.Click, eClick
        If eSelect IsNot Nothing Then
            AddHandler mnuC.DropDownOpening, eSelect
        End If
        mnuC.Checked = mnu.Checked
        mnuC.Enabled = mnu.Enabled
        mnuC.Font = mnu.Font
        mnuC.Image = mnu.Image
        mnuC.Name = mnu.Name
        mnuC.ShortcutKeys = mnu.ShortcutKeys
        mnuC.ShowShortcutKeys = mnu.ShowShortcutKeys
        mnuC.Tag = mnu.Tag
        mnuC.Text = mnu.Text
        mnuC.ToolTipText = mnu.ToolTipText

        Return mnuC
    End Function


    ' Extensión ToList                                              (26/Sep/20)

    ''' <summary>
    ''' Convierte una colección de items de un <see cref="ToolStripComboBox"/> en un <see cref="List(Of String)"/>.
    ''' </summary>
    ''' <param name="elCombo">El comboBox con los elementos a convertir en un <see cref="List(Of String)"/></param>
    ''' <returns></returns>
    <Extension>
    Public Function ToList(elCombo As ToolStripComboBox) As List(Of String)
        Dim col As New List(Of String)

        For i = 0 To elCombo.Items.Count - 1
            col.Add(elCombo.Items(i).ToString)
        Next

        Return col
    End Function

    ''' <summary>
    ''' Convierte una colección de items de un <see cref="ComboBox"/> en un <see cref="List(Of String)"/>.
    ''' </summary>
    ''' <param name="elCombo">El comboBox con los elementos a convertir en un <see cref="List(Of String)"/></param>
    ''' <returns></returns>
    <Extension>
    Public Function ToList(elCombo As ComboBox) As List(Of String)
        Dim col As New List(Of String)

        For i = 0 To elCombo.Items.Count - 1
            col.Add(elCombo.Items(i).ToString)
        Next

        Return col
    End Function

    ''' <summary>
    ''' Convierte una colección de items de un <see cref="ListBox"/> en un <see cref="List(Of String)"/>.
    ''' </summary>
    ''' <param name="elListBox">El listBox con los elementos a convertir en un <see cref="List(Of String)"/></param>
    ''' <returns></returns>
    <Extension>
    Public Function ToList(elListBox As ListBox) As List(Of String)
        Dim col As New List(Of String)

        For i = 0 To elListBox.Items.Count - 1
            col.Add(elListBox.Items(i).ToString)
        Next

        Return col
    End Function

    ' Extensión para reemplazar palabras completas                  (25/Sep/20)

    '''' <summary>
    '''' Reemplaza palabras completas usando RegEx
    '''' </summary>
    '''' <param name="str">La cadena a la que se aplica la extensión</param>
    '''' <param name="oldValue">El valor a buscar</param>
    '''' <param name="newValue">El nuevo valor a poner</param>
    '''' <returns>Una cadena con los cambios</returns>
    '''' <remarks>Ver <seealso cref="ReplaceWholeWord(String, String, String)"/> 
    '''' que según el autor del código original (palota) es más rápida que esta
    '''' que está adaptado del punto 6 de la respuesta anterior a esta (que es parecida):
    '''' https://stackoverflow.com/a/10151013</remarks>
    '<Extension>
    'Public Function ReplaceFullWord(str As String, oldValue As String, newValue As String) As String
    '    If str Is Nothing Then Return Nothing

    '    Return Regex.Replace(str, $"\b{oldValue}\b", newValue, RegexOptions.IgnoreCase)
    'End Function

    ''' <summary>
    ''' Devuelve una nueva cadena en la que todas las apariciones de oldValue
    ''' en la instancia actual se reemplazan por newValue, teniendo en cuenta
    ''' que se buscarán palabras completas.
    ''' </summary>
    ''' <param name="str">La cadena a la que se aplica la extensión</param>
    ''' <param name="oldValue">El valor a buscar</param>
    ''' <param name="newValue">El nuevo valor a poner</param>
    ''' <returns>Una cadena con los cambios</returns>
    ''' <remarks>Código convertido del original en C# de palota:
    ''' https://stackoverflow.com/a/62782791/14338047</remarks>
    <Extension>
    Public Function ReplaceWord(str As String, oldValue As String, newValue As String) As String
        Dim IsWordChar = Function(c As Char) Char.IsLetterOrDigit(c) OrElse c = "_"c

        Dim sb As StringBuilder = Nothing
        Dim p As Integer = 0, j As Integer = 0

        Do While j < str.Length AndAlso __Assign(j, str.IndexOf(oldValue, j, StringComparison.Ordinal)) >= 0
            If (j = 0 OrElse Not IsWordChar(str(j - 1))) AndAlso
                (j + oldValue.Length = str.Length OrElse Not IsWordChar(str(j + oldValue.Length))) Then

                sb = If(sb, New StringBuilder())
                sb.Append(str, p, j - p)
                sb.Append(newValue)
                j += oldValue.Length
                p = j
            Else
                j += 1
            End If
        Loop

        If sb Is Nothing Then Return str
        sb.Append(str, p, str.Length - p)
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Función para la equivalencia en C# de:
    ''' while (j &lt; text.Length && (j = unvalor)>=0 )
    ''' </summary>
    ''' <typeparam name="T">El tipo de datos</typeparam>
    ''' <param name="target">La variable a la que se le asignará el valor de la expresión de value</param>
    ''' <param name="value">Expresión con el valor a asignar a target</param>
    ''' <returns>Devuelve el valor de value</returns>
    Private Function __Assign(Of T)(ByRef target As T, value As T) As T
        target = value
        Return value
    End Function
End Module


