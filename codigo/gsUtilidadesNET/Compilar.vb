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
Imports System.CodeDom.Compiler

''' <summary>
''' Clase para evaluar, compilar y colorear código de VB y C#
''' </summary>
Public Class Compilar

    ''' <summary>
    ''' Si está instalado dotnet (cualquier versión) devuelve True.
    ''' </summary>
    Public Shared ReadOnly Property TieneDotnet As Boolean

    ''' <summary>
    ''' El mensaje de error si no existe el path de dotnet.
    ''' Si existe dotnet, se asigna el path.
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property TieneDotNetMsg As String

    '''' <summary>
    '''' Evento para comprobar si tiene o no instalado dotnet
    '''' </summary>
    '''' <param name="msg">El mensaje que indica si existe el directorio de dotnet</param>
    '''' <param name="estaDotnet">True si el directorio de dotnet existe o False si no existe</param>
    'Public Shared Event TieneDotNetEvent(msg As String, estaDotnet As Boolean)

    ''Private Shared Sub OnTieneDotNetEvent(msg As String, estaDotnet As Boolean)
    ''    RaiseEvent TieneDotNetEvent(msg, estaDotnet)
    ''End Sub

    Shared Sub New()
        Try
            Dim s = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
            Dim dotnetPath = Path.Join(s, "dotnet")

            If Directory.Exists(dotnetPath) Then
                TieneDotnet = True
                TieneDotNetMsg = dotnetPath
            Else
                TieneDotnet = False
                TieneDotNetMsg = $"No tienes dotnet instalado o no se encuentra el path de dotnet en '{dotnetPath}'"
            End If
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
            TieneDotnet = False
            TieneDotNetMsg = ex.Message
        End Try
        'RaiseEvent TieneDotNetEvent(TieneDotNetMsg, TieneDotnet)
    End Sub

    ''' <summary>
    ''' El nombre del lenguaje VB (como está definido en Compilation.Language)
    ''' </summary>
    Public Const LenguajeVisualBasic As String = "Visual Basic"
    ''' <summary>
    ''' El nombre del lenguaje C# (como está definido en Compilation.Language)
    ''' </summary>
    Public Const LenguajeCSharp As String = "C#"


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
        If res.Result Is Nothing OrElse res.Result.Success = False Then Return (res.Result, "")

        If run Then
            Try
                ' Algunas veces no se ejecuta,                      (17/Sep/20)
                ' porque el path contiene espacios.
                Dim p = Process.Start("dotnet", $"{ChrW(34)}{res.OutputPath}{ChrW(34)}")

                ' No hay forma de moverla, el truco es:             (26/Oct/20)
                ' al poner el cursor sobre la ventana en la barra de tareas
                ' se muestra un popup de la ventana, pulsar en la barra y darle a mover...

            Catch ex As Exception
                Debug.WriteLine(ex.Message)
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
            Try
                File.Delete(outputPath)
            Catch ex As Exception
                Debug.WriteLine(ex.Message)
                Debug.Assert(False, "No se puede eliminar " & outputPath)
            End Try
        End If

        Dim lenguaje = If(extension = ".vb", LenguajeVisualBasic, LenguajeCSharp)
        Dim result As EmitResult = Nothing
        Try
            result = GenerarCompilacion(sourceCode, outputDll, lenguaje).Emit(outputPath)

            ' para ejecutar una DLL usando dotnet, necesitamos un fichero de configuración
            CrearJson((result, outputPath, EsWinForm))
        Catch ex As Exception

        End Try


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

        Dim jsonText As String

        If res.EsWinForm Then
            Dim version = WindowsDesktopApp().Version
            ' Aplicación de escritorio (Windows Forms)
            ' Microsoft.WindowsDesktop.App
            ' 5.0.0-preview.8.20411.6
            ' 5.0.0-rc.1.20452.2
            ' 5.0.0-rc.2.20475.6
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
            ' 5.0.0-rc.2.20475.5
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
    Public Shared Function EvaluaCodigo(sourceCode As String, lenguaje As String) As _
                                            Dictionary(Of String, Dictionary(Of String, ClassifSpanInfo))

        Dim colSpans = GetClasSpans(sourceCode, lenguaje)
        Return EvaluaCodigo(sourceCode, colSpans)
    End Function

    ''' <summary>
    ''' Evalúa el código indicado para contar las clases/módulos y palabras clave.
    ''' </summary>
    ''' <param name="sourceCode">El código a evaluar</param>
    ''' <param name="colSpans">Un enumerador del tipo IEnumerable(Of ClassifiedSpan)</param>
    ''' <returns>Una colección del tipo <see cref="Dictionary(Of String, List(Of String))"/> con los elementos del código</returns>
    Private Shared Function EvaluaCodigo(sourceCode As String, colSpans As IEnumerable(Of ClassifiedSpan)) As _
                                                Dictionary(Of String, Dictionary(Of String, ClassifSpanInfo))

        Dim colCodigo As New Dictionary(Of String, Dictionary(Of String, ClassifSpanInfo))

        Dim source = SourceText.From(sourceCode)

        ClassifSpanInfo.SourceText = source

        ' No distinguir entre mayúsculas/minúsculas al comparar     (05/Oct/20)
        CompararString.IgnoreCase = True
        Dim compStr As New CompararString

        For Each classifSpan In colSpans
            Dim word = source.ToString(classifSpan.TextSpan)
            'Dim o = source.ToString() ' classifSpan.ClassificationType)

            If Not colCodigo.Keys.Contains(classifSpan.ClassificationType, compStr) Then
                colCodigo.Add(classifSpan.ClassificationType, New Dictionary(Of String, ClassifSpanInfo))
            End If
            If Not colCodigo(classifSpan.ClassificationType).Keys.Contains(word, compStr) Then
                Dim csI = New ClassifSpanInfo(classifSpan, word)
                colCodigo(classifSpan.ClassificationType).Add(word, csI)

                'colCodigo(classifSpan.ClassificationType)(word).Add(csI)

                ' No añadir más, solo una
                'Else
                '    Dim dcsI As New ClassifSpanInfo(classifSpan, word)
                '    colCodigo(classifSpan.ClassificationType)(word).Add(dcsI)
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
        'Dim errores = comp.GetDiagnostics
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

        Try
            dll.AddRange(Directory.GetFiles(dirCore, "System*.dll"))
            dll.AddRange(Directory.GetFiles(dirCore, "Microsoft*.dll"))
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
            MessageBox.Show($"Error en Compilar.ReferenciasDir:{vbCrLf}{ex.Message}{vbCrLf}{vbCrLf}" &
                            $"Seguramente no tienes instalado .NET 5.0.",
                            "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try

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
        Dim mayor As String = "5.0.0"
        Dim dirSep = Path.DirectorySeparatorChar
        Dim j = dirCore.IndexOf($"dotnet{dirSep}shared{dirSep}")

        If j = -1 Then
            'mayor = "5.0.0-rc.1.20452.2"
            'mayor = "5.0.0-rc.2.20475.6"
            'mayor = "5.0.0"
            dirWinDesk = $"C:\Program Files\dotnet\shared\Microsoft.WindowsDesktop.App\{mayor}"
        Else
            j += ($"dotnet{dirSep}shared{dirSep}").Length
            dirWinDesk = Path.Combine(dirCore.Substring(0, j), "Microsoft.WindowsDesktop.App")
            Try
                Dim dirs = Directory.GetDirectories(dirWinDesk).ToList()
                dirs.Sort()
                mayor = Path.GetFileName(dirs.Last())
            Catch ex As Exception
                Debug.WriteLine(ex.Message)
            End Try
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
        Dim mayor As String = "5.0.0"

        Dim j = If(String.IsNullOrEmpty(dirCore), -1, dirCore.IndexOf("Microsoft.NETCore.App"))

        If j = -1 Then
            'mayor = "5.0.0-rc.1.20451.14"
            'mayor = "5.0.0-rc.2.20475.5"
            mayor = "5.0.0"
        Else
            j += ("Microsoft.NETCore.App").Length
            Dim dirCoreApp = dirCore.Substring(0, j)
            Try
                Dim dirs = Directory.GetDirectories(dirCoreApp).ToList()
                dirs.Sort()
                mayor = Path.GetFileName(dirs.Last())
            Catch ex As Exception
                Debug.WriteLine(ex.Message)
            End Try
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

        sourceCode = sourceCode.Replace("<b", "&lt;b").Replace("<s", "&lt;s").Replace("</b", "&lt;/b").Replace("</s", "&lt;/s")

        Dim colSpans = GetClasSpans(sourceCode:=sourceCode, lenguaje)

        Dim codigoHTML = sourceCode
        Dim source = SourceText.From(sourceCode)
        'Dim selectionStart As Integer
        'Dim selectionLength As Integer

        Dim colCodigo = EvaluaCodigo(sourceCode, colSpans)

        ' Se ve que el texto no siempre tiene los mismos retornos   (18/Oct/20)
        ' hasta ahora (18-oct) siempre eran vbCr (en el RichTextBox)
        ' pero hoy son vbLf.
        Dim elCambioLinea = codigoHTML.ComprobarFinLinea
        Dim htmlL = codigoHTML.Split(elCambioLinea)
        'If htmlL.Length < 2 Then
        '    htmlL = codigoHTML.Split(vbLf)
        '    elCambioLinea = vbLf
        'End If

        ' Usar estos valores para que no ecuentre palabras          (03/Oct/20)
        ' contenidas en los span
        ' Es importante que NO contengan palabras o letras          (04/Oct/20)
        ' y, por supuesto, sean diferentes.
        Const tagSpanColor = "<$$###+$$###='$$##$$:" ' <span style='color:
        Const tagSpanEnd = "</$$###-$$###>" '</span>
        Const tagBold = "<$$##+$$##>" ' <b> {tagBold}
        Const tagBoldEnd = "</$$##-$$##>" ' </b> {tagBoldEnd}

        For Each classifSpan In colSpans
            Dim word = source.ToString(classifSpan.TextSpan)
            Dim linea = source.Lines.GetLinePosition(classifSpan.TextSpan.Start)
            'Debug.Assert(Not word.Contains("///"))
            Dim colRGB = GetStringColorFromName(classifSpan.ClassificationType)

            'selectionStart = classifSpan.TextSpan.Start
            'selectionLength = word.Length
            Dim span = word
            Dim csI = New ClassifSpanInfo(classifSpan, word)

            ' Si es el color negro, no añadir los <span
            If colRGB <> "#000000" AndAlso word.Length > 1 Then

                'Debug.Assert(word.Contains("Inherits") = False)

                ' Si la palabra está en static symbol,              (25/Sep/20)
                ' colorear como static symbol y poner en negrita
                ' Si está en class name, ídem, pero teniendo en cuenta que
                ' solo se pone en negrita si es C#
                ' Si está en property name cambiarlo al color de property name

                ' Comprobaciones por si no existe en la colección
                Dim esClassName = colCodigo.Keys.Contains("class name") AndAlso colCodigo("class name").Keys.Contains(word)
                Dim esStaticSymbol = colCodigo.Keys.Contains("static symbol") AndAlso colCodigo("static symbol").Keys.Contains(word)
                Dim esPropertyName = colCodigo.Keys.Contains("property name") AndAlso colCodigo("property name").Keys.Contains(word)
                If esClassName Then
                    ' En C# poner las clases en negrita
                    If lenguaje = "C#" Then
                        span = $"{tagBold}{tagSpanColor}{GetStringColorFromName("class name")}'>{word}{tagSpanEnd}{tagBoldEnd}"
                    Else
                        span = $"{tagSpanColor}{GetStringColorFromName("class name")}'>{word}{tagSpanEnd}"
                    End If
                ElseIf esStaticSymbol Then
                    span = $"{tagBold}{tagSpanColor}{GetStringColorFromName("static symbol")}'>{word}{tagSpanEnd}{tagBoldEnd}"
                ElseIf esPropertyName Then
                    span = $"{tagSpanColor}{GetStringColorFromName("property name")}'>{word}{tagSpanEnd}"
                ElseIf classifSpan.ClassificationType = "method name" Then
                    span = $"{tagBold}{tagSpanColor}{colRGB}'>{word}{tagSpanEnd}{tagBoldEnd}"
                Else
                    span = $"{tagSpanColor}{colRGB}'>{word}{tagSpanEnd}"
                End If
            End If

            If word.Contains(elCambioLinea) Then
                Dim wordL = word.Split(elCambioLinea)
                For i = 1 To wordL.Length - 1
                    If htmlL(linea.Line + i) = "" Then htmlL(linea.Line + i) = "<BoRrAr EstO>"
                    If wordL(i) <> "" Then
                        htmlL(linea.Line + i) = htmlL(linea.Line + i).Replace(wordL(i), "")
                    End If
                    If htmlL(linea.Line + i) = "" Then htmlL(linea.Line + i) = "<BoRrAr EstO>"
                Next
                htmlL(linea.Line) = htmlL(linea.Line).ReplaceSiNoEstaPoner(wordL(0), word)
            End If
            If word <> span Then
                'Debug.Assert(word.Contains("lt") = False)
                htmlL(linea.Line) = htmlL(linea.Line).ReplaceSiNoEstaPoner(word, span)
            End If

            Application.DoEvents()
        Next

        codigoHTML = ""
        ' El color de fondo y de las letras de los números de línea
        Const cFondoNum = "#f5f5f5" ' "#eaeaea"
        Dim cNum = GetStringColorFromName("class name")
        Dim l1 = 1
        For i = 0 To htmlL.Length - 1
            If htmlL(i) = "<BoRrAr EstO>" Then Continue For
            If mostrarLineas Then
                codigoHTML &= $"{tagSpanColor}{cNum}; background:{cFondoNum}'>{(l1),4:0} {tagSpanEnd}{htmlL(i)}{elCambioLinea}"
                l1 += 1
            Else
                'Debug.Assert(htmlL(i).Contains("Extension") = False)
                'Debug.Assert(htmlL(i).Contains("&lt;") = False)
                codigoHTML &= htmlL(i) & elCambioLinea
            End If

        Next
        ' En la documentación XML se lía un poco...                 (02/Oct/20)
        ' Cambiar los tags cambiados por los correctos              (03/Oct/20)
        ' Usando ReplaceSiNoEstaPoner                               (04/Oct/20)
        codigoHTML = codigoHTML.ReplaceSiNoEstaPoner($"{tagSpanColor}#808080'></{tagSpanEnd}span> ", " &lt;/{tagSpanEnd}").
                                ReplaceSiNoEstaPoner($"{tagSpanEnd} <{tagSpanColor}#808080'>summary{tagSpanEnd}>", " &lt;summary>{tagSpanEnd}").
                                ReplaceSiNoEstaPoner(tagSpanColor, "<span style='color:").
                                ReplaceSiNoEstaPoner(tagSpanEnd, "</span>").
                                ReplaceSiNoEstaPoner(tagBold, "<b>").
                                ReplaceSiNoEstaPoner(tagBoldEnd, "</b>")

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
    Public Shared Function ColoreaRichTextBox(richtb As RichTextBox, lenguaje As String) As _
                                                    Dictionary(Of String, Dictionary(Of String, ClassifSpanInfo))

        Dim colSpans = GetClasSpans(sourceCode:=richtb.Text, lenguaje)

        Dim colCodigo = EvaluaCodigo(richtb.Text, colSpans)

        Dim selectionStartAnt = 0

        With richtb
            Dim source = SourceText.From(.Text)
            For Each classifSpan In colSpans
                Dim word = source.ToString(classifSpan.TextSpan)
                Dim csI = New ClassifSpanInfo(classifSpan, word)

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
                    Dim esClassName = colCodigo.Keys.Contains("class name") AndAlso colCodigo("class name").Keys.Contains(word)
                    Dim esStaticSymbol = colCodigo.Keys.Contains("static symbol") AndAlso colCodigo("static symbol").Keys.Contains(word)
                    Dim esPropertyName = colCodigo.Keys.Contains("property name") AndAlso colCodigo("property name").Keys.Contains(word)
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


    ''' <summary>
    ''' Colorea el texto seleccionado del richTextBox (<paramref name="richtb"/>) y 
    ''' asigna el resultado nuevamente en el mismo control,
    ''' </summary>
    ''' <param name="richtb">Un RichTextBox al que se asignará el código coloreado</param>
    ''' <param name="lenguaje">El lenguaje a usar (Visual Basic o C#)</param>
    ''' <returns>Una colección del tipo <see cref="Dictionary(Of String, List(Of String))"/> con los elementos del código
    ''' sacados de la evaluación de <see cref="ClassifiedSpan"/>.</returns>
    Public Shared Function ColoreaSeleccionRichTextBox(richtb As RichTextBox, lenguaje As String) As _
                                                            Dictionary(Of String, Dictionary(Of String, ClassifSpanInfo))

        With richtb
            Dim selStart = richtb.SelectionStart
            Dim selStartFin = selStart + richtb.SelectionLength

            Dim colSpans = GetClasSpans(sourceCode:= .Text, lenguaje)
            Dim colCodigo = EvaluaCodigo(.Text, colSpans)
            Dim selectionStartAnt = 0

            Dim source = SourceText.From(.Text)

            For Each classifSpan In colSpans
                If Not (classifSpan.TextSpan.Start >= selStart AndAlso classifSpan.TextSpan.End <= selStartFin) Then
                    Continue For
                End If

                Dim word = source.ToString(classifSpan.TextSpan)
                Dim csI = New ClassifSpanInfo(classifSpan, word)

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
                    Dim esClassName = colCodigo.Keys.Contains("class name") AndAlso colCodigo("class name").Keys.Contains(word)
                    Dim esStaticSymbol = colCodigo.Keys.Contains("static symbol") AndAlso colCodigo("static symbol").Keys.Contains(word)
                    Dim esPropertyName = colCodigo.Keys.Contains("property name") AndAlso colCodigo("property name").Keys.Contains(word)
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

            Return colCodigo
        End With

    End Function


    '
    ' El resto de funciones está en Compilar.Partial
    '

#End Region

End Class
