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

    '
    ' Extensiones reemplazar si no está lo que se va a reemplazar   (04/Oct/20)
    '

    ''' <summary>
    ''' Reemplazar buscar por poner si no está poner.
    ''' </summary>
    ''' <param name="texto">La cadena en la que se hará la búsqueda y el reemplazo.</param>
    ''' <param name="buscar">La cadena a buscar sin distinguir entre mayúsculas y minúsculas.</param>
    ''' <param name="poner">La cadena a poner si previamente no está.</param>
    ''' <returns>Una cadena con los cambios realizados.</returns>
    <Extension>
    Public Function ReplaceSiNoEstaPoner(texto As String, buscar As String, poner As String) As String

        Dim j = texto.IndexOf(poner)
        ' si está lo que se quiere poner, devolver la cadena actual sin cambios
        If j > -1 Then Return texto

        Return texto.Replace(buscar, poner)
    End Function

    ''' <summary>
    ''' Reemplazar buscar por poner si no está poner.
    ''' </summary>
    ''' <param name="texto">La cadena en la que se hará la búsqueda y el reemplazo.</param>
    ''' <param name="buscar">La cadena a buscar usando la compración indicada.</param>
    ''' <param name="poner">La cadena a poner si previamente no está.</param>
    ''' <param name="comparar">El tipo de comparación a relizar: Ordinal o OrdinalIgnoreCase.</param>
    ''' <returns>Una cadena con los cambios realizados.</returns>
    <Extension>
    Public Function ReplaceSiNoEstaPoner(texto As String, buscar As String, poner As String,
                                         comparar As StringComparison) As String

        Dim j = texto.IndexOf(poner)
        ' si está lo que se quiere poner, devolver la cadena actual sin cambios
        If j > -1 Then Return texto

        Return texto.Replace(buscar, poner, comparar)
    End Function


    ''' <summary>
    ''' Reemplazar buscar por poner si no está poner.
    ''' </summary>
    ''' <param name="texto">La cadena en la que se hará la búsqueda y el reemplazo.</param>
    ''' <param name="buscar">La cadena a buscar (palabra completa) usando la comparación indicada.</param>
    ''' <param name="poner">La cadena a poner si previamente no está.</param>
    ''' <param name="comparar">El tipo de comparación a relizar: Ordinal o OrdinalIgnoreCase.</param>
    ''' <returns>Una cadena con los cambios realizados.</returns>
    <Extension>
    Public Function ReplaceWordSiNoEstaPoner(texto As String, buscar As String, poner As String,
                                             comparar As StringComparison) As String
        Dim j = texto.IndexOf(poner)
        ' si está lo que se quiere poner, devolver la cadena actual sin cambios
        If j > -1 Then Return texto

        Return ReplaceWord(texto, buscar, poner, comparar)
    End Function


    '
    ' Extensión quitar todos los espacios
    '

    ''' <summary>
    ''' Quitar todos los espacios que tenga la cadena
    ''' </summary>
    ''' <param name="texto">Cadena a la que se quitarán los espacios.</param>
    ''' <returns>Una nueva cadena con TODOS los espacios quitados.</returns>
    <Extension>
    Public Function QuitarTodosLosEspacios(texto As String) As String
        Dim col As MatchCollection = Regex.Matches(texto, "\S+")
        Dim sb As New StringBuilder
        For Each m As Match In col
            sb.Append(m.Value)
        Next

        Return sb.ToString
    End Function

    '
    ' Extensión contar palabras y saber las palabras usando Regex.
    '

    ''' <summary>
    ''' Contar las palabras de una cadena de texto usando <see cref="Regex"/>.
    ''' </summary>
    ''' <param name="texto">El texto con las palabras a contar.</param>
    ''' <returns>Un valor entero con el número de palabras</returns>
    ''' <example>
    ''' Adaptado usando una cadena en vez del Text del RichTextBox
    ''' (sería del RichTextBox para WinForms)
    ''' El código lo he adaptado de:
    ''' https://social.msdn.microsoft.com/Forums/en-US/
    '''     81e438ed-9d35-47d7-a800-1fabab0f3d52/
    '''     c-how-to-add-a-word-counter-to-a-richtextbox
    '''     ?forum=csharplanguage
    ''' </example>
    <Extension>
    Public Function CuantasPalabras(texto As String) As Integer
        Dim col As MatchCollection = Regex.Matches(texto, "[\W]+")
        Return col.Count
    End Function

    '
    ' Extensiones de cadena y cambiar a mayúsculas/minúsculas       (01/Oct/20)
    '

    Public Enum CasingValues As Integer
        ''' <summary>
        ''' No se hacen cambios
        ''' </summary>
        Normal
        ''' <summary>
        ''' Todas las letras a mayúsculas
        ''' </summary>
        Upper
        ''' <summary>
        ''' Todas las letras a minúsculas.
        ''' </summary>
        Lower
        ''' <summary>
        ''' La primera letra de cada palabra a mayúsculas.
        ''' </summary>
        Title
        ''' <summary>
        ''' La primera letra de cada palabra en mayúsculas.
        ''' Equivalente a <see cref="Title"/>.
        ''' </summary>
        FirstToUpper
        ''' <summary>
        ''' La primera letra de cada palabra en minúsculas
        ''' </summary>
        FirstToLower

    End Enum

    ''' <summary>
    ''' Cambia el texto a Upper, Lower, TitleCase/FirstToUpper o FirstToLower.
    ''' Se devuelve una nueva cadena con los cambios.
    ''' Valores posibles:
    ''' Normal
    ''' Upper
    ''' Lower
    ''' Title o FirstToLower
    ''' FirstToLower
    ''' </summary>
    ''' <param name="text">La cadena a la que se aplicará</param>
    ''' <param name="queCase">Un valor </param>
    ''' <returns>Una cadena con los cambios</returns>
    <Extension>
    Public Function CambiarCase(text As String, queCase As CasingValues) As String
        Select Case queCase
            Case CasingValues.Lower
                text = text.ToLower
            Case CasingValues.Upper
                text = text.ToUpper
            'Case CasingValues.Normal
            Case CasingValues.Title, CasingValues.FirstToUpper ' Title
                text = ToTitle(text)
            Case CasingValues.FirstToLower 'camelCase
                text = ToLowerFirst(text)
            Case Else ' Normal
                '
        End Select

        Return text
    End Function

    ''' <summary>
    ''' Devuelve una cadena en tipo Título
    ''' la primera letra de cada palabra en mayúsculas.
    ''' Usando System.Globalization.CultureInfo.CurrentCulture
    ''' que es más eficaz que
    ''' System.Threading.Thread.CurrentThread.CurrentCulture
    ''' </summary>
    <Extension>
    Public Function ToTitle(text As String) As String
        ' según la documentación usar CultureInfo.CurrentCulture es más eficaz
        ' que CurrentThread.CurrentCulture
        Dim cultureInfo = System.Globalization.CultureInfo.CurrentCulture
        Dim txtInfo = cultureInfo.TextInfo
        If text Is Nothing Then
            Return ""
        End If
        Return txtInfo.ToTitleCase(text)
    End Function

    ''' <summary>
    ''' Devuelve la cadena indicada con el primer carácter en minúsculas.
    ''' Si tiene espacios delante, pone en minúscula el primer carácter que no sea espacio.
    ''' </summary>
    ''' <param name="text"></param>
    ''' <returns></returns>
    <Extension>
    Public Function ToLowerFirstChar(text As String) As String
        If String.IsNullOrWhiteSpace(text) Then
            Return text
        End If

        Dim sb As New StringBuilder
        Dim b = False
        For i = 0 To text.Length - 1
            If Not b AndAlso Not Char.IsWhiteSpace(text(i)) Then
                sb.Append(text(i).ToString.ToLower)
                b = True
            Else
                sb.Append(text(i))
            End If
        Next

        Return sb.ToString
    End Function

    ''' <summary>
    ''' Convierte en minúsculas el primer carácter de cada palabra en la cadena indicada.
    ''' </summary>
    <Extension>
    Public Function ToLowerFirst(text As String) As String
        If String.IsNullOrWhiteSpace(text) Then
            Return text
        End If

        Dim col = Palabras(text)
        Dim sb As New StringBuilder
        For i = 0 To col.Count - 1
            sb.AppendFormat("{0}", col(i).ToLowerFirstChar)
        Next

        Return sb.ToString
    End Function

    ''' <summary>
    ''' Devuelve una lista con las palabras del texto indicado.
    ''' </summary>
    ''' <param name="text">La cadena de la que se extraerán las palabras.</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' En realidad no devuelve solo las palabras,
    ''' ya que cada elemento contendrá los espacios y otros símbolos que estén con esa palabra:
    ''' Si la palabra tiene espacios delante también los añade, si tiene un espacio o un símbolo detrás
    ''' también lo añade.
    ''' Si al final hay espacios en blanco, los elimina.
    ''' </remarks>
    ''' <example>    Private Sub Hola(str As String) 
    ''' Devolverá: "    Private ", "Sub ", "Hola(", "str ", "As ", "String)"
    ''' </example>
    <Extension>
    Public Function Palabras(text As String) As List(Of String)
        ' busca palabra con (o sin) espacios delante (\s*),
        ' cualquier cosa (.),
        ' una o más palabras (\w+) y
        ' cualquier cosa (.)
        Dim s = "\s*.\w+."
        Dim res = Regex.Matches(text, s)
        Dim col As New List(Of String)
        For Each m As Match In res
            col.Add(m.Value)
        Next

        Return col
    End Function

    '
    ' Extensión AsInteger                                           (28/Sep/20)
    '

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

    '
    ' Extensión Clonar                                              (27/Sep/20)
    '

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

    '
    ' Extensión ToList                                              (26/Sep/20)
    '

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

    '
    ' Extensión para reemplazar palabras completas                  (25/Sep/20)
    '

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


    '
    ' Versiones si se comprueban mayúsculas y minúsculas            (04/Oct/20)
    '

    ''' <summary>
    ''' Reemplaza todas las ocurrencias de 'buscar' por 'poner' en el texto,
    ''' teniendo en cuenta mayúsculas y minúsculas en la cadena a buscar.
    ''' </summary>
    ''' <param name="texto">La cadena en la que se hará la búsqueda y el reemplazo.</param>
    ''' <param name="buscar">El valor a buscar (palabra completa) distingue mayúsculas y minúsculas.</param>
    ''' <param name="poner">El nuevo valor a poner.</param>
    ''' <returns>Una cadena con los cambios.</returns>
    <Extension>
    Public Function ReplaceWordOrdinal(texto As String, buscar As String, poner As String) As String
        Return ReplaceWord(texto, buscar, poner, StringComparison.Ordinal)
    End Function


    ''' <summary>
    ''' Reemplaza todas las ocurrencias de 'buscar' por 'poner' en el texto,
    ''' ignorando mayúsculas y minúsculas en la cadena a buscar.
    ''' </summary>
    ''' <param name="texto">La cadena en la que se hará la búsqueda y el reemplazo.</param>
    ''' <param name="buscar">El valor a buscar (palabra completa) sin distinguir mayúsculas y minúsculas.</param>
    ''' <param name="poner">El nuevo valor a poner.</param>
    ''' <returns>Una cadena con los cambios.</returns>
    <Extension>
    Public Function ReplaceWordIgnoreCase(texto As String, buscar As String, poner As String) As String
        Return ReplaceWord(texto, buscar, poner, StringComparison.OrdinalIgnoreCase)
    End Function


    ''' <summary>
    ''' Devuelve una nueva cadena en la que todas las apariciones de oldValue
    ''' en la instancia actual se reemplazan por newValue, teniendo en cuenta
    ''' que se buscarán palabras completas.
    ''' </summary>
    ''' <param name="texto">La cadena en la que se hará la búsqueda y el reemplazo.</param>
    ''' <param name="buscar">El valor a buscar (palabra completa).</param>
    ''' <param name="poner">El nuevo valor a poner.</param>
    ''' <param name="comparar">El tipo de comparación Ordinal / OrdinalIgnoreCase.</param>
    ''' <returns>Una cadena con los cambios.</returns>
    ''' <remarks>Código convertido del original en C# de palota:
    ''' https://stackoverflow.com/a/62782791/14338047</remarks>
    <Extension>
    Public Function ReplaceWord(texto As String, buscar As String, poner As String,
                                comparar As StringComparison) As String
        Dim IsWordChar = Function(c As Char) Char.IsLetterOrDigit(c) OrElse c = "_"c

        Dim sb As StringBuilder = Nothing
        Dim p As Integer = 0, j As Integer = 0

        ' Comprueba sin distinguir mayúsculas y minúsculas          (04/Oct/20)
        'Dim ordinal = StringComparison.Ordinal
        'Dim noOrdinal = StringComparison.OrdinalIgnoreCase

        Do While j < texto.Length AndAlso __Assign(j, texto.IndexOf(buscar, j, comparar)) >= 0
            If (j = 0 OrElse Not IsWordChar(texto(j - 1))) AndAlso
                (j + buscar.Length = texto.Length OrElse Not IsWordChar(texto(j + buscar.Length))) Then

                sb = If(sb, New StringBuilder())
                sb.Append(texto, p, j - p)
                sb.Append(poner)
                j += buscar.Length
                p = j
            Else
                j += 1
            End If
        Loop

        If sb Is Nothing Then Return texto
        sb.Append(texto, p, texto.Length - p)
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Función para la equivalencia en C# de:
    ''' while (j &lt; text.Length &amp;& (j = unvalor)>=0 )
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


