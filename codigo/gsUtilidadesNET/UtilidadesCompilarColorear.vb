'------------------------------------------------------------------------------
' gsUtilidadesCompilarColorear                                      (27/Nov/20)
' Biblioteca de clases para .NET 5.0 (en realidad es una aplicación de Windows Forms)
'
'v1.0.0.0   27/Nov/20   Primera versión ;-)
'v1.0.0.2               comprobaciones de error al obtener el directorio de dotnet
'v1.0.0.3               Se comprueba si está la carpeta de dotnet, se avisa mediante el evento Compilar.TieneDotnetEvent
'
' (c) Guillermo (elGuille) Som, 2020
'------------------------------------------------------------------------------
Option Strict On
Option Infer On

Imports System
Imports System.Data
Imports System.Collections.Generic
Imports System.Text
Imports System.Linq
Imports Microsoft.VisualBasic
'Imports vb = Microsoft.VisualBasic

''' <summary>
''' Clase con definición de la versión y las clases expuestas.
''' </summary>
Public Class UtilidadesCompilarColorear

    ''' <summary>
    ''' Devuelve la versión de esta biblioteca de clases.
    ''' </summary>
    ''' <returns>Una cadena con la información de la biblioteca</returns>
    Public Shared Function Version() As String
        ' Añadir la versión de esta utilidad                        (15/Sep/20)
        Dim ensamblado As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly
        Dim fvi = System.Diagnostics.FileVersionInfo.GetVersionInfo(ensamblado.Location)

        Dim versionAttr = ensamblado.GetCustomAttributes(GetType(System.Reflection.AssemblyVersionAttribute), False)
        Dim vers = If(versionAttr.Length > 0,
                                (TryCast(versionAttr(0), System.Reflection.AssemblyVersionAttribute)).Version,
                                "1.0.0.0")
        Dim prodAttr = ensamblado.GetCustomAttributes(GetType(System.Reflection.AssemblyProductAttribute), False)
        Dim producto = If(prodAttr.Length > 0,
                                (TryCast(prodAttr(0), System.Reflection.AssemblyProductAttribute)).Product,
                                "gsUtilidadesCompilarColorear")
        Dim descAttr = ensamblado.GetCustomAttributes(GetType(System.Reflection.AssemblyDescriptionAttribute), False)
        Dim desc = If(descAttr.Length > 0,
                                (TryCast(descAttr(0), System.Reflection.AssemblyDescriptionAttribute)).Description,
                                "(para .NET 5.0 revisión del 27/Nov/2020)")

        Dim k = desc.IndexOf("(para .NET")
        Dim desc1 = desc.Substring(k)
        Dim descL = desc.Substring(0, k - 1).TrimEnd
        Dim otrosCopyR = "Parte del código para colorear está basado en ColorSelector usada en CSharpToVB de Paul1956."
        otrosCopyR &= $"{vbCrLf}Algunas funciones y código están adaptados de ejemplos encontrados en Internet, principalmente en C#."

        'MessageBox.Show($"{producto} v{vers} ({fvi.FileVersion})" & vbCrLf & vbCrLf &
        '                $"{descL}" & vbCrLf &
        '                $"{desc1}",
        '                $"Acerca de {producto}",
        '                MessageBoxButtons.OK, MessageBoxIcon.Information)

        Dim ret = $"{producto} v{vers} ({fvi.FileVersion}){vbCrLf}{vbCrLf}{descL}{vbCrLf}{vbCrLf}{desc1}" &
                  If(Compilar.TieneDotnet, "", $"{vbCrLf}{vbCrLf}{Compilar.TieneDotNetMsg}") &
                  $"{vbCrLf}{vbCrLf}" &
                  $"Reconocimientos:{vbCrLf}{otrosCopyR}"

        Return ret
    End Function

    ''' <summary>
    ''' Devuelve un <see cref="HashSet(Of String)"/> con los nombres de las clases y formularios
    ''' contenidos en esta biblioteca.
    ''' </summary>
    ''' <returns></returns>
    Public Shared Function ClasesExpuestas() As HashSet(Of String)
        Dim col = New HashSet(Of String)

        ' Usar reflection para mostrar las clases definidas

        col.Add("CompararString - Clase para hacer las comparaciones de cadenas que implementa IComparer(Of String) y IEqualityComparer(Of String).")
        col.Add("Compilar - Clase para evaluar, compilar y colorear código de VB y C#.")
        col.Add("Config - Clase para manejar ficheros de configuración.")
        col.Add("DiagClassifSpanInfo.ClassifSpanInfo - Clase para manejar información del tipo ClassifiedSpan.")
        col.Add("DiagClassifSpanInfo.DiagInfo - Clase para manejar información del tipo Diagnostic.")
        col.Add("Extensiones - Módulo con extensiones para varias clases y tipos de datos.")
        col.Add("Form1 - Formulario para mostrar la información de esta utilidad y las clases que la componen.")
        col.Add("FormRecortes - Formulario para mostrar recortes (texto) y pegarlos.")
        col.Add("FormProcesando - Formulario para mostrar un diálogo cuando se está procesando (acciones largas).")
        col.Add("Marcadores - Clase para manejar marcadores (Bookmarks).")
        col.Add("UtilEnum - Clase con la definición de la enumeración FormatosEncoding y utilidades para manejar enumeraciones.")
        col.Add("UtilidadesCompilarColorear - Clase con definición de la versión y las clases expuestas.")

        Return col
    End Function

End Class
