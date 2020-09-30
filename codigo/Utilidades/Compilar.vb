'------------------------------------------------------------------------------
' Compilar                                                          (20/Sep/20)
' Con los métodos para compilar el código de C# o VB
' También colorea el código (para usarlo directamente en un RichTextBox)
' Evalúa el código y muestra los errores y advertencias producidos.
'
'
' (c) Guillermo (elGuille) Som, 2020
'------------------------------------------------------------------------------
Option Strict On
Option Infer On

Imports System
Imports System.Linq
Imports System.Collections.Generic

Imports System.IO

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CSharp
Imports Microsoft.CodeAnalysis.Emit
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.Text
Imports System.Text
Imports Microsoft.CodeAnalysis.Host.Mef
Imports Microsoft.CodeAnalysis.Classification
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports System.Drawing.Text
Imports System.ComponentModel
Imports System.Security.Cryptography

Public Class Compilar

    ''' <summary>
    ''' El nombre del lenguaje VB (como está definido en Compilation.Language)
    ''' </summary>
    Private Const LenguajeVisualBasic As String = "Visual Basic"
    ''' <summary>
    ''' El nombre del lenguaje C# (como está definido en Compilation.Language)
    ''' </summary>
    Private Const LenguajeCSharp As String = "C#"


    ''' <summary>
    ''' Devuelve true si el código a compilar contiene InitializeComponent()
    ''' y se considera que es una aplicación de Windows (WinForm)
    ''' </summary>
    Private Shared Property EsWinForm As Boolean

    ''' <summary>
    ''' Compila el código indicado para saber si tiene errores o advertencias
    ''' </summary>
    ''' <param name="sourceCode">El código a compilar</param>
    ''' <param name="lenguaje">El lenguaje a usar (Visual Basic o C#)</param>
    ''' <returns>Un objeto <see cref="EmitResult"/> con la información de la compilación.
    ''' Si hay errores se devuelve en EmitResult.Diagnostics y EmitResult.Success será false</returns>
    ''' <remarks>El código compilado no se puede ejecutar, solo es para evaluarlo</remarks>
    Public Shared Function ComprobarCodigo(sourceCode As String, lenguaje As String) As EmitResult
        Dim outputDLL As String

        If lenguaje = LenguajeVisualBasic Then
            outputDLL = "tempVB.dll"
        Else
            outputDLL = "tempCS.dll"
        End If

        Return GenerarCompilacion(sourceCode, outputDLL, lenguaje).Emit(outputDLL)
    End Function

    ''' <summary>
    ''' Compila el código del fichero indicado.
    ''' Crea el ensamblado y lo guarda en el directorio del código fuente.
    ''' Crea el fichero runtimeconfig.json
    ''' Devuelve el nombre del ejecutable para usar con "dotnet OutputPath".
    ''' Ejecuta (si así se indica) el ensamblado generado.
    ''' </summary>
    ''' <param name="file">El path del fichero con el código a compilar</param>
    ''' <param name="run">True si se debe ejecutar después de compilar</param>
    ''' <returns>Devuelve una tupla con el resultado de la compilación (los errores en Result.Diagnostics)
    ''' y el path de la DLL generada en OutputPath</returns>
    ''' <remarks>Al compilar se crea el fichero runtimeconfig.json para que se pueda ejecutar con "dotnet OutputPath"</remarks>
    Public Shared Function CompileRun(file As String,
                                      Optional run As Boolean = True) As (Result As EmitResult, OutputPath As String)

        Dim res = CompileFile(file)
        If res.Result.Success = False Then Return (res.Result, "")

        If run Then
            Try
                ' Algunas veces no se ejecuta,                      (17/Sep/20)
                ' porque el path contiene espacios.
                Process.Start("dotnet", $"{ChrW(34)}{res.OutputPath}{ChrW(34)}")
            Catch
            End Try
        End If

        Return (res.Result, res.OutputPath)
    End Function

    ''' <summary>
    ''' Compila el fichero indicado y devuelve el path de la DLL generada
    ''' e información del resultado de la compilación y si es o no una aplicación de Windows (WinForm)
    ''' </summary>
    ''' <param name="filepath">El path al fichero a compilar</param>
    ''' <returns>Una Tupla con el resultado de la compilación:
    ''' Result tendrá el resultado de la compilación en Success (True si fue bien)
    ''' y los errores y advertencias en Diagnostics.
    ''' OutputPath tendrá el path completo de la DLL generada.
    ''' EsWinForm será true si se generó una aplicación de WindowsForms, si no será de Consola.
    ''' </returns>
    ''' <remarks>Al compilar se crea el fichero runtimeconfig.json para que se pueda ejecutar con "dotnet OutputPath"</remarks>
    Public Shared Function CompileFile(filepath As String) As (Result As EmitResult, OutputPath As String, EsWinForm As Boolean)
        Dim sourceCode As String

        Using sr = New StreamReader(filepath, System.Text.Encoding.UTF8, True)
            sourceCode = sr.ReadToEnd()
        End Using

        EsWinForm = sourceCode.IndexOf("InitializeComponent()") > -1

        Dim outputDll = Path.GetFileNameWithoutExtension(filepath) & ".dll"
        Dim outputPath = Path.Combine(Path.GetDirectoryName(filepath), outputDll)
        Dim extension = Path.GetExtension(filepath).ToLowerInvariant()

        ' Si existe la DLL de salida, eliminarla
        If File.Exists(outputPath) Then
            File.Delete(outputPath)
        End If

        Dim lenguaje = If(extension = ".vb", LenguajeVisualBasic, LenguajeCSharp)
        Dim result = GenerarCompilacion(sourceCode, outputDll, lenguaje).Emit(outputPath)

        ' para ejecutar una DLL usando dotnet, necesitamos un fichero de configuración
        CrearJson((result, outputPath, EsWinForm))

        Return (result, outputPath, EsWinForm)
    End Function

    ''' <summary>
    ''' Crea el fichero runtimeconfig.json para la aplicación.
    ''' </summary>
    ''' <param name="res">Tupla con el valor de EsWinForm y el path de la DLL
    ''' (de Result solo se utiliza Result.Success)</param>
    ''' <remarks>Si EsWinForm es true se crea el json para una aplicación de tipo WindowsDesktop.App
    ''' si no, será de tipo NETCore.App</remarks>
    Public Shared Sub CrearJson(res As (Result As EmitResult, OutputPath As String, EsWinForm As Boolean))
        ' para ejecutar una DLL usando dotnet, necesitamos un fichero de configuración

        ' No crearlo si hay error al compilar
        If res.Result.Success = False Then Return

        Dim jsonFile = Path.ChangeExtension(res.OutputPath, "runtimeconfig.json")

        Dim jsonText = ""

        If res.EsWinForm Then
            Dim version = WindowsDesktopApp().Version
            ' Aplicación de escritorio (Windows Forms)
            ' Microsoft.WindowsDesktop.App
            ' 5.0.0-preview.8.20411.6
            ' 5.0.0-rc.1.20452.2
            jsonText = "
{
    ""runtimeOptions"": {
    ""tfm"": ""net5.0-windows"",
    ""framework"": {
        ""name"": ""Microsoft.WindowsDesktop.App"",
        ""version"": """ & version & """
    }
    }
}"
        Else
            Dim version = NETCoreApp().Version
            ' Tipo consola
            ' Microsoft.NETCore.App
            ' 5.0.0-preview.8.20407.11
            ' 5.0.0-rc.1.20451.14
            jsonText = "
{
    ""runtimeOptions"": {
    ""tfm"": ""net5.0"",
    ""framework"": {
        ""name"": ""Microsoft.NETCore.App"",
        ""version"": """ & version & """
    }
    }
}"
        End If

        Using sw = New StreamWriter(jsonFile, False, Encoding.UTF8)
            sw.WriteLine(jsonText)
        End Using
    End Sub

    ''' <summary>
    ''' Compila y genera una DLL con el código de Visual Basic o C# indicado
    ''' </summary>
    ''' <param name="sourceCode">El código (de Visual Basic) a compilar</param>
    ''' <param name="outputDll">El fichero de salida para la compilación (solo el nombre, sin path)</param>
    ''' <param name="lenguaje">El lenguaje a usar (Visual Basic o C#)</param>
    ''' <returns>Devuelve un objeto Compilation (<see cref="Compilation"/>)</returns>
    ''' <remarks>
    ''' Si el código contiene InitializeComponent() se considera que es aplicación de Windows
    ''' si no, será aplicación de Consola.
    ''' </remarks>
    ''' <example>
    ''' Para crear una DLL (ejecutable) usar:
    ''' result = GeneraCompilacion(src, outDLL, lenguaje).Emit(outputPath)
    ''' CrearJson((result, outputPath, EsWinForm))
    ''' o llamar a <see cref="CompileFile"/>
    ''' </example>
    Private Shared Function GenerarCompilacion(sourceCode As String, outputDll As String, lenguaje As String) As Compilation
        Dim codeString = SourceText.From(sourceCode)
        Dim tree As SyntaxTree
        ' Añadir todas las referencias
        Dim references = Referencias()

        Dim outpKind = OutputKind.ConsoleApplication
        If EsWinForm Then _
            outpKind = OutputKind.WindowsApplication

        If lenguaje = LenguajeVisualBasic Then
            Dim options = VisualBasicParseOptions.Default.WithLanguageVersion(VisualBasic.LanguageVersion.Default)

            tree = VisualBasic.SyntaxFactory.ParseSyntaxTree(codeString, options)

            Return VisualBasicCompilation.Create(
                outputDll,
                {tree},
                references:=references,
                options:=New VisualBasicCompilationOptions(
                    outpKind,
                    optimizationLevel:=OptimizationLevel.Release,
                    assemblyIdentityComparer:=DesktopAssemblyIdentityComparer.Default))

        Else
            Dim options = CSharpParseOptions.Default.WithLanguageVersion(CSharp.LanguageVersion.Default)

            tree = CSharp.SyntaxFactory.ParseSyntaxTree(codeString, options)

            Return CSharpCompilation.Create(
                outputDll,
                {tree},
                references:=references,
                options:=New CSharpCompilationOptions(
                    outpKind,
                    optimizationLevel:=OptimizationLevel.Release,
                    assemblyIdentityComparer:=DesktopAssemblyIdentityComparer.Default))

        End If
    End Function

    ''' <summary>
    ''' Compila y genera una DLL con el código de Visual Basic indicado
    ''' </summary>
    ''' <param name="sourceCode">El código (de Visual Basic) a compilar</param>
    ''' <param name="outputDll">El fichero de salida para la compilación (solo el nombre, sin path)</param>
    ''' <returns>Devuelve un objeto Compilation (<see cref="VisualBasicCompilation"/>)</returns>
    ''' <remarks>
    ''' Si el código contiene InitializeComponent() se considera que es aplicación de Windows
    ''' si no, será aplicación de Consola.
    ''' </remarks>
    <Obsolete("Este método está obsoleto, usar GenerarCompilacion indicando VB como lenguaje")>
    Private Shared Function VBGenerateCode(sourceCode As String, outputDll As String) As VisualBasicCompilation
        Dim codeString = SourceText.From(sourceCode)
        Dim options = VisualBasicParseOptions.Default.WithLanguageVersion(VisualBasic.LanguageVersion.Default)

        Dim tree = VisualBasic.SyntaxFactory.ParseSyntaxTree(codeString, options)

        ' Añadir todas las referencias
        Dim references = Referencias()

        Dim outpKind = OutputKind.ConsoleApplication
        If EsWinForm Then _
            outpKind = OutputKind.WindowsApplication

        Return VisualBasicCompilation.Create(
                outputDll,
                {tree},
                references:=references,
                options:=New VisualBasicCompilationOptions(
                    outpKind,
                    optimizationLevel:=OptimizationLevel.Release,
                    assemblyIdentityComparer:=DesktopAssemblyIdentityComparer.Default))

    End Function

    ''' <summary>
    ''' Compila y genera una DLL con el código de CSharp indicado
    ''' </summary>
    ''' <param name="sourceCode">El código (de CSharp) a compilar</param>
    ''' <param name="outputDll">El fichero de salida para la compilación (solo el nombre, sin path)</param>
    ''' <returns>Devuelve un objeto Compilation (<see cref="CSharpCompilation"/>)</returns>
    ''' <remarks>
    ''' Si el código contiene InitializeComponent() se considera que es aplicación de Windows
    ''' si no, será aplicación de Consola.
    ''' </remarks>
    <Obsolete("Este método está obsoleto, usar GenerarCompilacion indicando C# como lenguaje")>
    Private Shared Function CSGenerateCode(sourceCode As String, outputDll As String) As CSharpCompilation
        Dim codeString = SourceText.From(sourceCode)
        Dim options = CSharpParseOptions.Default.WithLanguageVersion(CSharp.LanguageVersion.Default)

        Dim tree = CSharp.SyntaxFactory.ParseSyntaxTree(codeString, options)

        ' Añadir todas las referencias
        Dim references = Referencias()

        Dim outpKind = OutputKind.ConsoleApplication
        If EsWinForm Then _
            outpKind = OutputKind.WindowsApplication

        Return CSharpCompilation.Create(
                outputDll,
                {tree},
                references:=references,
                options:=New CSharpCompilationOptions(
                    outpKind,
                    optimizationLevel:=OptimizationLevel.Release,
                    assemblyIdentityComparer:=DesktopAssemblyIdentityComparer.Default))

    End Function

    ''' <summary>
    ''' Evalúa el código indicado para contar las clases/módulos y palabras clave.
    ''' </summary>
    ''' <param name="sourceCode">El código a evaluar</param>
    ''' <param name="lenguaje">El lenguaje a usar (Visual Basic o C#)</param>
    ''' <returns>Una colección del tipo <see cref="Dictionary(Of String, List(Of String))"/> con los elementos del código</returns>
    Public Shared Function EvaluaCodigo(sourceCode As String,
                                        lenguaje As String) As Dictionary(Of String, List(Of String))

        Dim colSpans = GetClasSpans(sourceCode, lenguaje)
        Return EvaluaCodigo(sourceCode, colSpans)
    End Function

    ''' <summary>
    ''' Evalúa el código indicado para contar las clases/módulos y palabras clave.
    ''' </summary>
    ''' <param name="sourceCode">El código a evaluar</param>
    ''' <param name="colSpans">Un enumerador del tipo IEnumerable(Of ClassifiedSpan)</param>
    ''' <returns>Una colección del tipo <see cref="Dictionary(Of String, List(Of String))"/> con los elementos del código</returns>
    Private Shared Function EvaluaCodigo(sourceCode As String,
                                         colSpans As IEnumerable(Of ClassifiedSpan)) As Dictionary(Of String, List(Of String))

        Dim colCodigo As New Dictionary(Of String, List(Of String))

        Dim source = SourceText.From(sourceCode)

        For Each classifSpan In colSpans
            Dim word = source.ToString(classifSpan.TextSpan)

            If Not colCodigo.Keys.Contains(classifSpan.ClassificationType) Then
                colCodigo.Add(classifSpan.ClassificationType, New List(Of String))
            End If
            If Not colCodigo(classifSpan.ClassificationType).Contains(word) Then
                colCodigo(classifSpan.ClassificationType).Add(word)
            End If
        Next

        Return colCodigo
    End Function

    ''' <summary>
    ''' Devuelve un enumerador del tipo <see cref="ClassifiedSpan"/> del código y el lenguaje indicados
    ''' </summary>
    ''' <param name="sourceCode">El código a evaluar</param>
    ''' <param name="lenguaje">El lenguaje a usar (Visual Basic o C#)</param>
    ''' <returns>Un enumerador del tipo <see cref="ClassifiedSpan"/></returns>
    ''' <remarks>Esto también lo uso para colorear en <see cref="ColoreaRichTextBox"/></remarks>
    Private Shared Function GetClasSpans(sourceCode As String,
                                         lenguaje As String) As IEnumerable(Of ClassifiedSpan)
        Dim host = MefHostServices.Create(MefHostServices.DefaultAssemblies)
        Dim workspace = New AdhocWorkspace(host)

        Dim source As SourceText = SourceText.From(sourceCode)
        Dim tree As SyntaxTree
        Dim comp As Compilation

        If lenguaje = LenguajeVisualBasic Then
            tree = VisualBasicSyntaxTree.ParseText(source)
            comp = VisualBasicCompilation.Create("temp.vb").AddReferences(Referencias).AddSyntaxTrees(tree)
        Else
            tree = CSharpSyntaxTree.ParseText(source)
            comp = CSharpCompilation.Create("temp.cs").AddReferences(Referencias).AddSyntaxTrees(tree)
        End If

        ' Para comprobar cómo se llaman internamente los lenguajes  (24/Sep/20)
        'Dim slang = comp.Language ' "C#" o "Visual Basic"

        Dim semantic = comp.GetSemanticModel(tree)
        Dim clasSpans = Classifier.GetClassifiedSpans(semantic, New TextSpan(0, source.Length), workspace)

        Return clasSpans
    End Function


    ''' <summary>
    ''' Colección con las referencias usadas para compilar.
    ''' Tendrá tanto las referencias para aplicaciones WindowsDesktop como de NETCore.
    ''' </summary>
    Private Shared Property ColReferencias As List(Of MetadataReference)

    ''' <summary>
    ''' Devuelve todas las referencias para aplicaciones de NETCore y WindowsDesktop.
    ''' </summary>
    ''' <returns>Una lista del tipo <see cref="MetadataReference"/> con las referencias</returns>
    Private Shared Function Referencias() As List(Of MetadataReference)
        If ColReferencias IsNot Nothing Then Return ColReferencias

        ColReferencias = New List(Of MetadataReference)

        Dim dirCore = System.Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory()
        ColReferencias = ReferenciasDir(dirCore)

        ' Para las aplicaciones de Windows Forms

        ' Buscar la versión mayor del directorio de aplicaciones de escritorio
        Dim dirWinDesk = WindowsDesktopApp().Dir
        ColReferencias.AddRange(ReferenciasDir(dirWinDesk))

        Return ColReferencias
    End Function

    ''' <summary>
    ''' Devuelve las referencias encontradas en el directorio indicado.
    ''' </summary>
    ''' <param name="dirCore">Directorio con las DLL a referenciar</param>
    ''' <returns>Una lista del tipo <see cref="MetadataReference"/> con las referencias encontradas</returns>
    Private Shared Function ReferenciasDir(dirCore As String) As List(Of MetadataReference)
        Dim col = New List(Of MetadataReference)()
        Dim dll = New List(Of String)()

        dll.AddRange(Directory.GetFiles(dirCore, "System*.dll"))
        dll.AddRange(Directory.GetFiles(dirCore, "Microsoft*.dll"))

        Dim noInc = Path.Combine(dirCore, "Microsoft.DiaSymReader.Native.amd64.dll")
        If dll.Contains(noInc) Then dll.Remove(noInc)

        For i = 0 To dll.Count - 1
            col.Add(MetadataReference.CreateFromFile(dll(i)))
        Next

        Return col
    End Function

    ''' <summary>
    ''' Devuelve el directorio y la versión mayor
    ''' del path con las DLL de Microsoft.WindowsDesktop.App.
    ''' </summary>
    Public Shared Function WindowsDesktopApp() As (Dir As String, Version As String)
        ' Buscar el directorio para las referencias de WindowsDesktop.App (08/Sep/20)

        Dim dirCore = System.Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory()
        ' Buscar la versión mayor del directorio de aplicaciones de escritorio
        Dim dirWinDesk As String
        Dim mayor As String
        Dim dirSep = Path.DirectorySeparatorChar
        Dim j = dirCore.IndexOf($"dotnet{dirSep}shared{dirSep}")

        If j = -1 Then
            mayor = "5.0.0-rc.1.20452.2"
            dirWinDesk = $"C:\Program Files\dotnet\shared\Microsoft.WindowsDesktop.App\{mayor}"
        Else
            j += ($"dotnet{dirSep}shared{dirSep}").Length
            dirWinDesk = Path.Combine(dirCore.Substring(0, j), "Microsoft.WindowsDesktop.App")
            Dim dirs = Directory.GetDirectories(dirWinDesk).ToList()
            dirs.Sort()
            mayor = Path.GetFileName(dirs.Last())
            dirWinDesk = Path.Combine(dirWinDesk, mayor)
        End If

        Return (dirWinDesk, mayor)
    End Function

    ''' <summary>
    ''' Devuelve el directorio y la versión mayor 
    ''' del path con las DLL para aplicaciones Microsoft.NETCore.App.
    ''' </summary>
    Public Shared Function NETCoreApp() As (Dir As String, Version As String)
        ' Buscar el directorio para las referencias de NETCore.App  (08/Sep/20)

        Dim dirCore = System.Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory()
        Dim mayor As String
        Dim j = dirCore.IndexOf("Microsoft.NETCore.App")

        If j = -1 Then
            mayor = "5.0.0-rc.1.20451.14"
        Else
            j += ("Microsoft.NETCore.App").Length
            Dim dirCoreApp = dirCore.Substring(0, j)
            Dim dirs = Directory.GetDirectories(dirCoreApp).ToList()
            dirs.Sort()
            mayor = Path.GetFileName(dirs.Last())
        End If

        Return (dirCore, mayor)
    End Function


#Region " Para colorear "

    ''' <summary>
    ''' Colorea el código indicado a formato HTML (para usar en sitios WEB)
    ''' </summary>
    ''' <param name="sourceCode">El código a colorear</param>
    ''' <param name="lenguaje">El lenguaje a usar (Visual Basic o C#)</param>
    ''' <param name="mostrarLineas">True si se mostrarán los números de línea en el código convertido</param>
    ''' <returns>Una cadena con el código HTML coloreado</returns>
    ''' <remarks></remarks>
    Public Shared Function ColoreaHTML(sourceCode As String, lenguaje As String, mostrarLineas As Boolean) As String
        ' Iniciado el 23/Sep/20 y finalizado el 25/Sep/20

        ' El color de fondo y de las letras de los números de línea
        Const cFondoNum = "#f5f5f5" ' "#eaeaea"
        Dim cNum = GetStringColorFromName("class name")

        Dim colSpans = GetClasSpans(sourceCode:=sourceCode, lenguaje)

        Dim codigoHTML = sourceCode

        Dim colL As New List(Of (lin As Integer, pos As Point, span As String))

        Dim source = SourceText.From(sourceCode)
        Dim selectionStart As Integer
        Dim selectionLength As Integer

        '#Const CAMBIARSTATICSYMBOL = True


        '#If CAMBIARSTATICSYMBOL Then
        Dim colCodigo = EvaluaCodigo(sourceCode, colSpans)
        '#End If

        For Each classifSpan In colSpans
            Dim word = source.ToString(classifSpan.TextSpan)
            Dim linea = source.Lines.GetLinePosition(classifSpan.TextSpan.Start)

            Dim colRGB = GetStringColorFromName(classifSpan.ClassificationType) ' "#------" '

            selectionStart = classifSpan.TextSpan.Start
            selectionLength = word.Length

            Dim span = word
            colL.Add((linea.Line, New Point(selectionStart, selectionLength), span))
            Dim n = colL.Count - 1
            If colRGB <> "#000000" AndAlso word.Length > 1 Then

                '#If CAMBIARSTATICSYMBOL Then

                ' Si la palabra está en static symbol,              (25/Sep/20)
                ' colorear como static symbol y poner en negrita
                ' Si está en class name, ídem, pero teniendo en cuenta que
                ' solo se pone en negrita si es C#
                ' Si está en property name cambiarlo al color de property name

                ' Comprobaciones por si no existe en la colección
                Dim esClassName = If(colCodigo.Keys.Contains("class name"), colCodigo("class name").Contains(word), False)
                Dim esStaticSymbol = If(colCodigo.Keys.Contains("static symbol"), colCodigo("static symbol").Contains(word), False)
                Dim esPropertyName = If(colCodigo.Keys.Contains("property name"), colCodigo("property name").Contains(word), False)
                If esClassName Then
                    ' En C# poner las clases en negrita
                    If lenguaje = "C#" Then
                        span = $"<b><span style='color:{GetStringColorFromName("class name")}'>{word}</span></b>"
                    Else
                        span = $"<span style='color:{GetStringColorFromName("class name")}'>{word}</span>"
                    End If
                ElseIf esStaticSymbol Then
                    span = $"<b><span style='color:{GetStringColorFromName("static symbol")}'>{word}</span></b>"
                ElseIf esPropertyName Then
                    span = $"<span style='color:{GetStringColorFromName("property name")}'>{word}</span>"
                ElseIf classifSpan.ClassificationType = "method name" Then
                    span = $"<b><span style='color:{colRGB}'>{word}</span></b>"
'#Else
'                If classifSpan.ClassificationType = "method name" OrElse classifSpan.ClassificationType = "static symbol" Then
'                    span = $"<b><span style='color:{colRGB}'>{word}</span></b>"
'#End If
                Else
                    span = $"<span style='color:{colRGB}'>{word}</span>"
                End If
            End If
            colL(n) = (colL(n).lin, colL(n).pos, span)

            Application.DoEvents()
        Next

        Dim htmlL = codigoHTML.Split(vbCr)

        ' Recorrer las colecciones para añadir los cambios de línea
        ' y corregir el coloreado y comentarios multilíneas
        For i = 0 To colL.Count - 1
            Dim l1 = colL(i).lin
            Dim s = htmlL(l1)
            Dim word = sourceCode.Substring(colL(i).pos.X, colL(i).pos.Y)
            If word.Length = 1 Then
                ' si son iguales no hacer nada
                If word <> colL(i).span Then
                    Dim j = s.LastIndexOf("</span>")
                    If j = -1 Then j = -6
                    Dim k = s.IndexOf(word, j + 7)
                    If k > -1 Then
                        s = s.Substring(0, k) & colL(i).span & s.Substring(k + 1)
                    End If
                End If
            Else
                ' Si es un comentario múltiple que no empieza al principio
                ' se supone que hay código delante de ese comentario: using System; /* comentario */
                If s.TrimStart().Contains(" /*") AndAlso s.Contains("*/") = True Then
                    s = s.Replace(word, colL(i).span)

                ElseIf s.TrimStart().Contains(" /*") AndAlso s.Contains("*/") = False Then
                    ' aunque puede continuar en líneas diferentes:
                    ' using System; /* comentario
                    '   que sigue en otra línea */
                    ' word contendrá el vbCr si es en dos líneas
                    Dim j1 = word.IndexOf(vbCr)
                    If j1 > -1 Then
                        Dim word1 = word.Substring(0, j1)
                        Dim word2 = word.Substring(j1 + vbCr.Length)
                        ' borrar las líneas de hmlL que tengan
                        ' ese código
                        For i1 = l1 To htmlL.Length - 1
                            If htmlL(i1).IndexOf(word1) = -1 Then
                                If htmlL(i1).IndexOf(word2) = -1 Then
                                    Exit For
                                End If
                            End If
                            ' Cambio de mayúculas y minúsculas
                            ' para que sea más complicado que sea casualida ;-)
                            htmlL(i1) = "<BoRrAr EstO>"
                        Next
                        If mostrarLineas Then
                            Dim s1 = colL(i).span.Replace(vbCr & word2, "").Replace("</span>", "")
                            colL(i) = (colL(i).lin, colL(i).pos, s1)
                            s = s.Replace(word1, colL(i).span) & $"{vbCr}<span style='color:{cNum}; background:{cFondoNum}'>{(l1 + 2).ToString("0").PadLeft(4)} </span>&nbsp;" & word2 & "</span>"
                        Else
                            s = s.Replace(word1, colL(i).span)
                        End If
                    Else
                        s = s.Replace(word, colL(i).span)
                    End If

                ElseIf s.Contains("/*") Then
                    ' Los comentarios de varias líneas se devuelven
                    ' directamente con los retornos de carro, etc.

                    ' Esto va bien si no hay texto delante
                    s = colL(i).span
                    Dim l2 = l1 'colL(i).lin
                    ' Si son distintas, es que word 
                    ' contiene todo el comentario
                    If word <> s Then
                        ' borrar las líneas de hmlL que tengan
                        ' ese código
                        For i1 = l1 To htmlL.Length - 1
                            If word.IndexOf(htmlL(i1)) = -1 Then
                                Exit For
                            End If
                            ' Cambio de mayúculas y minúsculas
                            ' para que sea más complicado que sea casualida ;-)
                            htmlL(i1) = "<BoRrAr EstO>"
                            If i1 > l2 Then l2 = i1
                        Next
                    End If
                    ' ponerle los números de línea
                    If mostrarLineas Then
                        Dim j1 = 0
                        For i1 = l1 To l2
                            j1 = s.IndexOf(vbCr, j1)
                            If j1 = -1 Then
                                s &= $"{vbCr}<span style='color:{cNum}; background:{cFondoNum}'>{(i1 + 2).ToString("0").PadLeft(4)} </span>&nbsp;"

                                Exit For
                            End If
                            s = s.Substring(0, j1) & $"{vbCr}<span style='color:{cNum}; background:{cFondoNum}'>{(i1 + 2).ToString("0").PadLeft(4)} </span>&nbsp;" & s.Substring(j1 + 1)
                        Next
                    Else
                        Dim j1 = 0
                        For i1 = l1 To l2
                            j1 = s.IndexOf(vbCr, j1)
                            If j1 = -1 Then
                                s &= $"{vbCr}"

                                Exit For
                            End If
                            s = s.Substring(0, j1) & $"{vbCr}&nbsp;" & s.Substring(j1 + 1)
                        Next
                    End If

                ElseIf colL(i).span.Contains($"<span style='color:{GetStringColorFromName("string - verbatim")}'>") OrElse
                        colL(i).span.Contains($"<span style='color:{GetStringColorFromName("string")}'>") Then
                    ' Comprobar si la línea siguiente tiene más código
                    ' de esta línea
                    Dim l2 = l1
                    Dim s1 = ""
                    For i1 = l1 To htmlL.Length - 1
                        s1 &= htmlL(i1) & vbCr
                        If s1.Contains(word) Then
                            For i2 = l1 To i1
                                htmlL(i2) = "<BoRrAr EstO>"
                                If i1 > l2 Then l2 = i1
                            Next
                            Exit For
                        End If
                    Next
                    If mostrarLineas Then
                        ' Comprobar si tiene vbCr antes del final
                        ' cuento lo que hay y si hay más de uno...
                        Dim cvbCr = s1.Count(Function(c) c = ChrW(13))
                        If cvbCr > 2 Then
                            s1 = s1.ReplaceWord(word, colL(i).span)

                            Dim j1 = 0
                            For i1 = l1 To l2
                                j1 = s1.IndexOf(vbCr, j1)
                                If j1 = -1 OrElse j1 = s1.Length - 1 Then
                                    Exit For
                                End If
                                s1 = s1.Substring(0, j1) & $"{vbCr}<span style='color:{cNum}; background:{cFondoNum}'>{(i1 + 2).ToString("0").PadLeft(4)} </span>&nbsp;" & s1.Substring(j1 + 1)
                            Next
                            s = s1

                        Else
                            s = s1.ReplaceWord(word, colL(i).span)
                        End If
                    Else
                        s = s1.ReplaceWord(word, colL(i).span)
                    End If

                Else
                    s = s.ReplaceWord(word, colL(i).span)
                End If

            End If

            htmlL(l1) = s
        Next

        codigoHTML = ""
        Dim primeraLinea = -1
        For i = 0 To htmlL.Length - 1

            If htmlL(i) = "<BoRrAr EstO>" Then Continue For
            If mostrarLineas Then
                ' Añadir números de líneas para el webbrowser
                If primeraLinea = -1 Then
                    primeraLinea = 0
                    codigoHTML &= $"<span style='color:{cNum}; background:{cFondoNum}'>{(i + 1).ToString("0").PadLeft(4)} </span>" & htmlL(i) & vbCr '"<br>&nbsp;"
                Else
                    codigoHTML &= $"<span style='color:{cNum}; background:{cFondoNum}'>{(i + 1).ToString("0").PadLeft(4)} </span>" & htmlL(i) & vbCr '"<br>&nbsp;"
                End If
            Else
                If primeraLinea = -1 Then
                    primeraLinea = 0
                    codigoHTML = htmlL(i) & vbCr '"<br>&nbsp;"
                Else
                    codigoHTML &= htmlL(i) & vbCr '"<br>&nbsp;"
                End If

            End If
        Next

        'codigoHTML = codigoHTML.Replace(vbCr & vbCr, "")

        Return "<pre style='font-family:Consolas; font-size: 11pt; font-weight:semi-bold'>" & codigoHTML & "</pre>"
    End Function


    ''' <summary>
    ''' Colorea el contenido del texto del richTextBox (<paramref name="richtb"/>) y 
    ''' asigna el resultado nuevamente en el mismo control
    ''' </summary>
    ''' <param name="richtb">Un RichTextBox al que se asignará el código coloreado</param>
    ''' <param name="lenguaje">El lenguaje a usar (Visual Basic o C#)</param>
    ''' <returns>Una colección del tipo <see cref="Dictionary(Of String, List(Of String))"/> con los elementos del código
    ''' sacados de la evaluación de <see cref="ClassifiedSpan"/>.</returns>
    Public Shared Function ColoreaRichTextBox(richtb As RichTextBox,
                                              lenguaje As String) As Dictionary(Of String, List(Of String))

        Dim colSpans = GetClasSpans(sourceCode:=richtb.Text, lenguaje)

        Dim colCodigo = EvaluaCodigo(richtb.Text, colSpans)

        Dim selectionStartAnt = 0

        With richtb
            Dim source = SourceText.From(.Text)
            For Each classifSpan In colSpans
                Dim word = source.ToString(classifSpan.TextSpan)

                ' No todas las clasificaciones están marcdas como "static symbol"
                ' por eso solo pone en negrita algunas
                ' Esta colección solo tiene "static symbol" (a día de hoy 23/Sep/2020)
                ' NOTA: No siempre marca correctamente los static symbol
                If ClassificationTypeNames.AdditiveTypeNames.Contains(classifSpan.ClassificationType) Then
                    .SelectionStart = selectionStartAnt
                    .SelectionLength = word.Length
                    .SelectionFont = New Font(.SelectionFont, FontStyle.Bold)
                Else

                    .SelectionStart = classifSpan.TextSpan.Start
                    selectionStartAnt = .SelectionStart
                    .SelectionLength = word.Length

                    ' Si la palabra está en static symbol,          (25/Sep/20)      
                    ' colorear como static symbol y poner en negrita
                    ' Si está en class name, ídem, pero teniendo en cuenta que
                    ' solo se pone en negrita si ses C#

                    ' Comprobaciones por si no existe en la colección
                    Dim esClassName = If(colCodigo.Keys.Contains("class name"), colCodigo("class name").Contains(word), False)
                    Dim esStaticSymbol = If(colCodigo.Keys.Contains("static symbol"), colCodigo("static symbol").Contains(word), False)
                    Dim esPropertyName = If(colCodigo.Keys.Contains("property name"), colCodigo("property name").Contains(word), False)
                    If esClassName Then
                        .SelectionColor = GetColorFromName("class name")
                        ' En C# poner las clases en negrita
                        If lenguaje = "C#" Then
                            .SelectionFont = New Font(.SelectionFont, FontStyle.Bold)
                        End If
                    ElseIf esStaticSymbol Then
                        .SelectionColor = GetColorFromName("static symbol")
                        .SelectionFont = New Font(.SelectionFont, FontStyle.Bold)
                    ElseIf esPropertyName Then
                        .SelectionColor = GetColorFromName("property name")
                    Else
                        .SelectionColor = GetColorFromName(classifSpan.ClassificationType)
                    End If

                    .SelectedText = word

                End If

                Application.DoEvents()

            Next
        End With

        Return colCodigo
    End Function

    '
    ' El resto de funciones está en Compilar.Partial
    '

#End Region

End Class
