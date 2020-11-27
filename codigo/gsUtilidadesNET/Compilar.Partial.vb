'------------------------------------------------------------------------------
' Compilar.Partial                                                  (25/Sep/20)
' Código usado por Compilar para colorear el código
'
'--------------------------------------------------------------------------
' Estas definiciones están adaptadas, y/o literalmente copiadas,
' de la clase ColorSelector usada en CSharpToVB (Visual Basic) de Paul1956:
'   https://github.com/paul1956/CSharpToVB
'--------------------------------------------------------------------------
'
'
' (c) Guillermo (elGuille) Som, 2020
'------------------------------------------------------------------------------
Option Strict On
Option Infer On

Imports System
Imports System.Drawing
Imports System.Collections.Generic

Imports System.IO
Imports System.Text
'Imports Microsoft.CodeAnalysis.Classification

Partial Public Class Compilar

    '--------------------------------------------------------------------------
    ' Estas definiciones están adaptadas, y/o literalmente copiadas,
    ' de la clase ColorSelector usada en CSharpToVB (Visual Basic) de Paul1956:
    '   https://github.com/paul1956/CSharpToVB
    '--------------------------------------------------------------------------
    ' 103, 204, 241
    ' 58; 155;205
    Private Shared ReadOnly colorMapping As New Dictionary(Of String, Color)(StringComparer.OrdinalIgnoreCase) From {
         {"class name", Color.FromArgb(0, 128, 128)},
         {"comment", Color.FromArgb(0, 100, 0)},
         {"constant name", Color.Black},
         {"default", Color.FromArgb(65, 110, 178)}, ' Color.Black}, 25/Sep: (65, 110, 178)
         {"delegate name", Color.FromArgb(0, 128, 128)},
         {"enum name", Color.FromArgb(0, 128, 128)},
         {"enum member name", Color.FromArgb(0, 128, 128)},
         {"error", Color.Red},
         {"excluded code", Color.FromArgb(128, 128, 128)},
         {"event name", Color.Black},
         {"extension method name", Color.FromArgb(160, 80, 40)}, '.Black}, 03/Oct
         {"field name", Color.Black},
         {"identifier", Color.FromArgb(64, 48, 160)}, 'Color.Black}, ' 24/Sep/20: (64, 48, 160) 25/Sep: (73, 172, 205)
         {"interface name", Color.FromArgb(0, 128, 128)},
         {"keyword", Color.FromArgb(0, 0, 255)},
         {"keyword - control", Color.FromArgb(143, 8, 196)},
         {"label name", Color.Black},
         {"local name", Color.FromArgb(65, 110, 178)}, ' Color.Black}, ' 25/Sep: (65, 110, 178)
         {"method name", Color.FromArgb(118, 85, 40)}, 'Color.Black},
         {"module name", Color.FromArgb(0, 128, 128)},
         {"namespace name", Color.Black},
         {"number", Color.Black},
         {"operator", Color.Black},
         {"operator - overloaded", Color.Black},
         {"parameter name", Color.FromArgb(65, 110, 178)}, ' Color.Black}, ' 25/Sep: (65, 110, 178)
         {"preprocessor keyword", Color.Gray},
         {"preprocessor text", Color.Black},
         {"property name", Color.Black},
         {"punctuation", Color.Black},
         {"static symbol", Color.FromArgb(118, 85, 40)}, ' Color.Black}, ' 23/Sep/20
         {"string - escape character", Color.FromArgb(143, 8, 196)},' Color.Yellow}, ' 01/Oct/20
         {"string - verbatim", Color.FromArgb(128, 0, 0)}, ' texto entre comillas de múltiples líneas en C#
         {"string", Color.FromArgb(163, 21, 21)}, ' texto entre comillas
         {"struct name", Color.FromArgb(43, 145, 175)},
         {"text", Color.Black},
         {"type parameter name", Color.DarkGray},
         {"xml doc comment - attribute name", Color.FromArgb(128, 128, 128)},
         {"xml doc comment - attribute quotes", Color.FromArgb(128, 128, 128)},
         {"xml doc comment - attribute value", Color.FromArgb(128, 128, 128)},
         {"xml doc comment - cdata section", Color.FromArgb(128, 128, 128)},
         {"xml doc comment - comment", Color.FromArgb(128, 128, 128)},
         {"xml doc comment - delimiter", Color.FromArgb(128, 128, 128)},
         {"xml doc comment - entity reference", Color.FromArgb(0, 128, 0)},
         {"xml doc comment - name", Color.FromArgb(128, 128, 128)},
         {"xml doc comment - processing instruction", Color.FromArgb(128, 128, 128)},
         {"xml doc comment - text", Color.FromArgb(0, 128, 0)},
         {"xml literal - attribute name", Color.FromArgb(1, 115, 254)}, ' 03/Oct antes 128, 128, 128 
         {"xml literal - attribute quotes", Color.FromArgb(128, 128, 128)},
         {"xml literal - attribute value", Color.FromArgb(1, 115, 254)}, ' 03/Oct antes 128, 128, 128 
         {"xml literal - cdata section", Color.FromArgb(128, 128, 128)},
         {"xml literal - comment", Color.FromArgb(128, 128, 128)},
         {"xml literal - delimiter", Color.FromArgb(1, 115, 254)}, ' 03/Oct antes 100, 100, 185
         {"xml literal - embedded expression", Color.FromArgb(128, 128, 128)},
         {"xml literal - entity reference", Color.FromArgb(185, 100, 100)},
         {"xml literal - name", Color.FromArgb(132, 70, 70)},
         {"xml literal - processing instruction", Color.FromArgb(128, 128, 128)},
         {"xml literal - text", Color.FromArgb(85, 85, 85)}
     }

    ''' <summary>
    ''' Obtiene el color a partir del nombre indicado en <see cref="ClassifiedSpan.ClassificationType"/>
    ''' de un objeto del tipo <see cref="ClassifiedSpan"/>
    ''' </summary>
    ''' <param name="classificationTypeName">Texto del ClassificationType</param>
    ''' <returns>Un objeto de tipo <see cref="System.Drawing.Color"/> con el color correspondiente</returns>
    Public Shared Function GetColorFromName(classificationTypeName As String) As Color
        If String.IsNullOrWhiteSpace(classificationTypeName) Then
            Return colorMapping("default")
        End If
        Dim returnValue As Color = Nothing
        If colorMapping.TryGetValue(classificationTypeName, returnValue) Then
            Return returnValue
        End If
        'Debug.Print($"GetColorFromName missing({classificationTypeName})")
        Return colorMapping("error")
    End Function

    ''' <summary>
    ''' Devuelve una cadena con el color asociado con el nombre indicado.
    ''' </summary>
    ''' <param name="classificationTypeName">Texto del ClassificationType</param>
    ''' <returns>El formato será #RGB en hexadecimal</returns>
    Public Shared Function GetStringColorFromName(classificationTypeName As String) As String
        If String.IsNullOrWhiteSpace(classificationTypeName) Then
            Return stringForColor(colorMapping("default"))
        End If
        Dim ReturnValue As Color = Nothing
        If colorMapping.TryGetValue(classificationTypeName, ReturnValue) Then
            Return stringForColor(ReturnValue)
        End If
        'Debug.Print($"GetColorFromName missing({classificationTypeName})")
        Return stringForColor(colorMapping("error"))
    End Function

    ''' <summary>
    ''' Devuelve una cadena en formato #RRGGBB (en hexadecimal) del color indicado
    ''' </summary>
    ''' <param name="co">El <see cref="Color"/> a convertir en cadena</param>
    ''' <returns>Una cadena (en formato hexadecimal) con los colores al estilo #RRGGBB</returns>
    Private Shared Function stringForColor(co As Color) As String
        Return $"#{co.R.ToString("X2")}{co.G.ToString("X2")}{co.B.ToString("X2")}"
    End Function

    Private Shared ReadOnly colorDictionaryPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ColorDictionary.csv")

    ''' <summary>
    ''' Guarda las definiciones de los colores en el fichero
    ''' indicado por <see cref="colorDictionaryPath"/>
    ''' </summary>
    Public Shared Sub WriteColorDictionaryToFile()
        WriteColorDictionaryToFile(colorDictionaryPath)
    End Sub

    ''' <summary>
    ''' Guarda las definiciones de los colores en el fichero indicado
    ''' </summary>
    ''' <param name="ficPath">El fichero (con formato CSV) con las definiciones de los colores</param>
    Public Shared Sub WriteColorDictionaryToFile(ficPath As String)
        Using FileStream As FileStream = File.OpenWrite(ficPath)
            Using sw As New StreamWriter(FileStream)
                sw.WriteLine($"Key,R,G,B")
                For Each kvp As KeyValuePair(Of String, Color) In colorMapping
                    sw.WriteLine($"{kvp.Key},{kvp.Value.R},{kvp.Value.G},{kvp.Value.B}")
                Next
                sw.Flush()
                sw.Close()
            End Using
            FileStream.Close()
        End Using
    End Sub

    ''' <summary>
    ''' Lee las definiciones de los colores del fichero
    ''' indicado por <see cref="colorDictionaryPath"/>
    ''' </summary>
    Public Shared Sub UpdateColorDictionaryFromFile()
        UpdateColorDictionaryFromFile(colorDictionaryPath)
    End Sub

    ''' <summary>
    ''' Actualiza los colores de las definiciones desde el fichero indicado
    ''' </summary>
    ''' <param name="ficPath">El fichero (con formato CSV) con las definiciones de los colores</param>
    Public Shared Sub UpdateColorDictionaryFromFile(ficPath As String)
        If Not File.Exists(ficPath) Then
            WriteColorDictionaryToFile(ficPath)
            Exit Sub
        End If
        Using fs As FileStream = File.OpenRead(ficPath)
            Using sr As New StreamReader(fs)
                sr.ReadLine()
                Do While sr.Peek() <> -1
                    Dim line As String = sr.ReadLine()
                    Dim Split() As String = line.Split(","c)
                    Dim key As String = Split(0)
                    Dim R As Integer = Convert.ToInt32(Split(1), Globalization.CultureInfo.InvariantCulture)
                    Dim G As Integer = Convert.ToInt32(Split(2), Globalization.CultureInfo.InvariantCulture)
                    Dim B As Integer = Convert.ToInt32(Split(3), Globalization.CultureInfo.InvariantCulture)
                    colorMapping(key) = Color.FromArgb(R, G, B)
                Loop
                sr.Close()
            End Using
            fs.Close()
        End Using
    End Sub

End Class