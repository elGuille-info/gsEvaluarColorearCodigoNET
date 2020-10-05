'------------------------------------------------------------------------------
' M�dulo con extensiones para String (y otras clases)               (25/Sep/20)
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
    ' Extensiones reemplazar si no est� lo que se va a reemplazar   (04/Oct/20)
    '

    ''' <summary>
    ''' Reemplazar buscar por poner si no est� poner.
    ''' </summary>
    ''' <param name="texto">La cadena en la que se har� la b�squeda y el reemplazo.</param>
    ''' <param name="buscar">La cadena a buscar sin distinguir entre may�sculas y min�sculas.</param>
    ''' <param name="poner">La cadena a poner si previamente no est�.</param>
    ''' <returns>Una cadena con los cambios realizados.</returns>
    <Extension>
    Public Function ReplaceSiNoEstaPoner(texto As String, buscar As String, poner As String) As String

        Dim j = texto.IndexOf(poner)
        ' si est� lo que se quiere poner, devolver la cadena actual sin cambios
        If j > -1 Then Return texto

        Return texto.Replace(buscar, poner)
    End Function

    ''' <summary>
    ''' Reemplazar buscar por poner si no est� poner.
    ''' </summary>
    ''' <param name="texto">La cadena en la que se har� la b�squeda y el reemplazo.</param>
    ''' <param name="buscar">La cadena a buscar usando la compraci�n indicada.</param>
    ''' <param name="poner">La cadena a poner si previamente no est�.</param>
    ''' <param name="comparar">El tipo de comparaci�n a relizar: Ordinal o OrdinalIgnoreCase.</param>
    ''' <returns>Una cadena con los cambios realizados.</returns>
    <Extension>
    Public Function ReplaceSiNoEstaPoner(texto As String, buscar As String, poner As String,
                                         comparar As StringComparison) As String

        Dim j = texto.IndexOf(poner)
        ' si est� lo que se quiere poner, devolver la cadena actual sin cambios
        If j > -1 Then Return texto

        Return texto.Replace(buscar, poner, comparar)
    End Function


    ''' <summary>
    ''' Reemplazar buscar por poner si no est� poner.
    ''' </summary>
    ''' <param name="texto">La cadena en la que se har� la b�squeda y el reemplazo.</param>
    ''' <param name="buscar">La cadena a buscar (palabra completa) usando la comparaci�n indicada.</param>
    ''' <param name="poner">La cadena a poner si previamente no est�.</param>
    ''' <param name="comparar">El tipo de comparaci�n a relizar: Ordinal o OrdinalIgnoreCase.</param>
    ''' <returns>Una cadena con los cambios realizados.</returns>
    <Extension>
    Public Function ReplaceWordSiNoEstaPoner(texto As String, buscar As String, poner As String,
                                             comparar As StringComparison) As String
        Dim j = texto.IndexOf(poner)
        ' si est� lo que se quiere poner, devolver la cadena actual sin cambios
        If j > -1 Then Return texto

        Return ReplaceWord(texto, buscar, poner, comparar)
    End Function


    '
    ' Extensi�n quitar todos los espacios
    '

    ''' <summary>
    ''' Quitar todos los espacios que tenga la cadena
    ''' </summary>
    ''' <param name="texto">Cadena a la que se quitar�n los espacios.</param>
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
    ' Extensi�n contar palabras y saber las palabras usando Regex.
    '

    ''' <summary>
    ''' Contar las palabras de una cadena de texto usando <see cref="Regex"/>.
    ''' </summary>
    ''' <param name="texto">El texto con las palabras a contar.</param>
    ''' <returns>Un valor entero con el n�mero de palabras</returns>
    ''' <example>
    ''' Adaptado usando una cadena en vez del Text del RichTextBox
    ''' (ser�a del RichTextBox para WinForms)
    ''' El c�digo lo he adaptado de:
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
    ' Extensiones de cadena y cambiar a may�sculas/min�sculas       (01/Oct/20)
    '

    Public Enum CasingValues As Integer
        ''' <summary>
        ''' No se hacen cambios
        ''' </summary>
        Normal
        ''' <summary>
        ''' Todas las letras a may�sculas
        ''' </summary>
        Upper
        ''' <summary>
        ''' Todas las letras a min�sculas.
        ''' </summary>
        Lower
        ''' <summary>
        ''' La primera letra de cada palabra a may�sculas.
        ''' </summary>
        Title
        ''' <summary>
        ''' La primera letra de cada palabra en may�sculas.
        ''' Equivalente a <see cref="Title"/>.
        ''' </summary>
        FirstToUpper
        ''' <summary>
        ''' La primera letra de cada palabra en min�sculas
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
    ''' <param name="text">La cadena a la que se aplicar�</param>
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
    ''' Devuelve una cadena en tipo T�tulo
    ''' la primera letra de cada palabra en may�sculas.
    ''' Usando System.Globalization.CultureInfo.CurrentCulture
    ''' que es m�s eficaz que
    ''' System.Threading.Thread.CurrentThread.CurrentCulture
    ''' </summary>
    <Extension>
    Public Function ToTitle(text As String) As String
        ' seg�n la documentaci�n usar CultureInfo.CurrentCulture es m�s eficaz
        ' que CurrentThread.CurrentCulture
        Dim cultureInfo = System.Globalization.CultureInfo.CurrentCulture
        Dim txtInfo = cultureInfo.TextInfo
        If text Is Nothing Then
            Return ""
        End If
        Return txtInfo.ToTitleCase(text)
    End Function

    ''' <summary>
    ''' Devuelve la cadena indicada con el primer car�cter en min�sculas.
    ''' Si tiene espacios delante, pone en min�scula el primer car�cter que no sea espacio.
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
    ''' Convierte en min�sculas el primer car�cter de cada palabra en la cadena indicada.
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
    ''' <param name="text">La cadena de la que se extraer�n las palabras.</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' En realidad no devuelve solo las palabras,
    ''' ya que cada elemento contendr� los espacios y otros s�mbolos que est�n con esa palabra:
    ''' Si la palabra tiene espacios delante tambi�n los a�ade, si tiene un espacio o un s�mbolo detr�s
    ''' tambi�n lo a�ade.
    ''' Si al final hay espacios en blanco, los elimina.
    ''' </remarks>
    ''' <example>    Private Sub Hola(str As String) 
    ''' Devolver�: "    Private ", "Sub ", "Hola(", "str ", "As ", "String)"
    ''' </example>
    <Extension>
    Public Function Palabras(text As String) As List(Of String)
        ' busca palabra con (o sin) espacios delante (\s*),
        ' cualquier cosa (.),
        ' una o m�s palabras (\w+) y
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
    ' Extensi�n AsInteger                                           (28/Sep/20)
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
    ' Extensi�n Clonar                                              (27/Sep/20)
    '

    ''' <summary>
    ''' Clonar un men� item del tipo <see cref="ToolStripMenuItem"/>.
    ''' </summary>
    ''' <param name="mnu">El <see cref="ToolStripMenuItem"/> al que se va a asignar el nuevo men�.</param>
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
    ' Extensi�n ToList                                              (26/Sep/20)
    '

    ''' <summary>
    ''' Convierte una colecci�n de items de un <see cref="ToolStripComboBox"/> en un <see cref="List(Of String)"/>.
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
    ''' Convierte una colecci�n de items de un <see cref="ComboBox"/> en un <see cref="List(Of String)"/>.
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
    ''' Convierte una colecci�n de items de un <see cref="ListBox"/> en un <see cref="List(Of String)"/>.
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
    ' Extensi�n para reemplazar palabras completas                  (25/Sep/20)
    '

    '''' <summary>
    '''' Reemplaza palabras completas usando RegEx
    '''' </summary>
    '''' <param name="str">La cadena a la que se aplica la extensi�n</param>
    '''' <param name="oldValue">El valor a buscar</param>
    '''' <param name="newValue">El nuevo valor a poner</param>
    '''' <returns>Una cadena con los cambios</returns>
    '''' <remarks>Ver <seealso cref="ReplaceWholeWord(String, String, String)"/> 
    '''' que seg�n el autor del c�digo original (palota) es m�s r�pida que esta
    '''' que est� adaptado del punto 6 de la respuesta anterior a esta (que es parecida):
    '''' https://stackoverflow.com/a/10151013</remarks>
    '<Extension>
    'Public Function ReplaceFullWord(str As String, oldValue As String, newValue As String) As String
    '    If str Is Nothing Then Return Nothing

    '    Return Regex.Replace(str, $"\b{oldValue}\b", newValue, RegexOptions.IgnoreCase)
    'End Function


    '
    ' Versiones si se comprueban may�sculas y min�sculas            (04/Oct/20)
    '

    ''' <summary>
    ''' Reemplaza todas las ocurrencias de 'buscar' por 'poner' en el texto,
    ''' teniendo en cuenta may�sculas y min�sculas en la cadena a buscar.
    ''' </summary>
    ''' <param name="texto">La cadena en la que se har� la b�squeda y el reemplazo.</param>
    ''' <param name="buscar">El valor a buscar (palabra completa) distingue may�sculas y min�sculas.</param>
    ''' <param name="poner">El nuevo valor a poner.</param>
    ''' <returns>Una cadena con los cambios.</returns>
    <Extension>
    Public Function ReplaceWordOrdinal(texto As String, buscar As String, poner As String) As String
        Return ReplaceWord(texto, buscar, poner, StringComparison.Ordinal)
    End Function


    ''' <summary>
    ''' Reemplaza todas las ocurrencias de 'buscar' por 'poner' en el texto,
    ''' ignorando may�sculas y min�sculas en la cadena a buscar.
    ''' </summary>
    ''' <param name="texto">La cadena en la que se har� la b�squeda y el reemplazo.</param>
    ''' <param name="buscar">El valor a buscar (palabra completa) sin distinguir may�sculas y min�sculas.</param>
    ''' <param name="poner">El nuevo valor a poner.</param>
    ''' <returns>Una cadena con los cambios.</returns>
    <Extension>
    Public Function ReplaceWordIgnoreCase(texto As String, buscar As String, poner As String) As String
        Return ReplaceWord(texto, buscar, poner, StringComparison.OrdinalIgnoreCase)
    End Function


    ''' <summary>
    ''' Devuelve una nueva cadena en la que todas las apariciones de oldValue
    ''' en la instancia actual se reemplazan por newValue, teniendo en cuenta
    ''' que se buscar�n palabras completas.
    ''' </summary>
    ''' <param name="texto">La cadena en la que se har� la b�squeda y el reemplazo.</param>
    ''' <param name="buscar">El valor a buscar (palabra completa).</param>
    ''' <param name="poner">El nuevo valor a poner.</param>
    ''' <param name="comparar">El tipo de comparaci�n Ordinal / OrdinalIgnoreCase.</param>
    ''' <returns>Una cadena con los cambios.</returns>
    ''' <remarks>C�digo convertido del original en C# de palota:
    ''' https://stackoverflow.com/a/62782791/14338047</remarks>
    <Extension>
    Public Function ReplaceWord(texto As String, buscar As String, poner As String,
                                comparar As StringComparison) As String
        Dim IsWordChar = Function(c As Char) Char.IsLetterOrDigit(c) OrElse c = "_"c

        Dim sb As StringBuilder = Nothing
        Dim p As Integer = 0, j As Integer = 0

        ' Comprueba sin distinguir may�sculas y min�sculas          (04/Oct/20)
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
    ''' Funci�n para la equivalencia en C# de:
    ''' while (j &lt; text.Length &amp;& (j = unvalor)>=0 )
    ''' </summary>
    ''' <typeparam name="T">El tipo de datos</typeparam>
    ''' <param name="target">La variable a la que se le asignar� el valor de la expresi�n de value</param>
    ''' <param name="value">Expresi�n con el valor a asignar a target</param>
    ''' <returns>Devuelve el valor de value</returns>
    Private Function __Assign(Of T)(ByRef target As T, value As T) As T
        target = value
        Return value
    End Function
End Module


