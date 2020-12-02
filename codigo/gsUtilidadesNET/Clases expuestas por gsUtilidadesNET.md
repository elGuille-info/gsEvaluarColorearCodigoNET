# Contenido del ensamblado 'gsUtilidadesNET.dll'

Información de los tipos definidos en el ensamblado.

>
>**NOTA:**
>Esta información la he conseguido usando la utilidad [Mostrar contenido ensamblado](https://github.com/elGuille-info/Mostrar-contenido-ensamblado)
>

## Clase: VB$AnonymousDelegate_0`2
    IsGenericType = True
    Constructores:
        IsPublic = True
            Parámetros: Object TargetObject, IntPtr TargetMethod
    VB$AnonymousDelegate_0`2.Métodos:
        Method BeginInvoke
            Parámetros: TArg0 c, AsyncCallback DelegateCallback, Object DelegateAsyncState
        Method EndInvoke
            Parámetros: IAsyncResult DelegateAsyncResult
        Method Invoke
            Parámetros: TArg0 c
        Method GetObjectData
            Parámetros: SerializationInfo info, StreamingContext context
        Method Equals
            Parámetros: Object obj
        Method GetInvocationList

        Method GetHashCode

        Method Clone

        Method DynamicInvoke
            Parámetros: Object[] args
        Method GetType

        Method ToString


## Clase: MyApplication
    Constructores:
        IsPublic = True
            Sin parámetros
    MyApplication.Métodos:
        Method GetType

        Method ToString

        Method Equals
            Parámetros: Object obj
        Method GetHashCode


## Clase: CompararString
    Constructores:
        IsPublic = True
            Sin parámetros
    CompararString.Métodos:
        Method Compare
            Parámetros: String x, String y
        Method Equals
            Parámetros: String x, String y
        Method GetHashCode
            Parámetros: String obj
        Method GetType

        Method ToString

        Method Equals
            Parámetros: Object obj
        Method GetHashCode


## Clase: Compilar
    Constructores:
        IsPublic = True
            Sin parámetros
    Compilar.Métodos:
        Method GetColorFromName
            Parámetros: String classificationTypeName
        Method GetStringColorFromName
            Parámetros: String classificationTypeName
        Method WriteColorDictionaryToFile

        Method WriteColorDictionaryToFile
            Parámetros: String ficPath
        Method UpdateColorDictionaryFromFile

        Method UpdateColorDictionaryFromFile
            Parámetros: String ficPath
        Method ComprobarCodigo
            Parámetros: String sourceCode, String lenguaje
        Method CompileRun
            Parámetros: String file, [Optional] Boolean run = True
        Method CompileFile
            Parámetros: String filepath
        Method CrearJson
            Parámetros: ValueTuple`3 res
        Method EvaluaCodigo
            Parámetros: String sourceCode, String lenguaje
        Method WindowsDesktopApp

        Method NETCoreApp

        Method ColoreaHTML
            Parámetros: String sourceCode, String lenguaje, Boolean mostrarLineas
        Method ColoreaRichTextBox
        Method ColoreaSeleccionRichTextBox
        Method GetType

        Method ToString

        Method Equals
            Parámetros: Object obj
        Method GetHashCode


## Clase: Config
    Constructores:
        IsPublic = True
            Parámetros: String fic
        IsPublic = True
            Parámetros: String fic, Boolean guardarAlAsignar
    Config.Métodos:
        Method GetValue
            Parámetros: String seccion, String clave
        Method GetValue
            Parámetros: String seccion, String clave, String predeterminado
        Method GetValue
            Parámetros: String seccion, String clave, Int32 predeterminado
        Method GetValue
            Parámetros: String seccion, String clave, Double predeterminado
        Method GetValue
            Parámetros: String seccion, String clave, Boolean predeterminado
        Method SetValue
            Parámetros: String seccion, String clave, String valor
        Method SetValue
            Parámetros: String seccion, String clave, Int32 valor
        Method SetValue
            Parámetros: String seccion, String clave, Double valor
        Method SetKeyValue
            Parámetros: String seccion, String clave, Double valor
        Method SetValue
            Parámetros: String seccion, String clave, Boolean valor
        Method SetKeyValue
            Parámetros: String seccion, String clave, String valor
        Method SetKeyValue
            Parámetros: String seccion, String clave, Int32 valor
        Method SetKeyValue
            Parámetros: String seccion, String clave, Boolean valor
        Method RemoveSection
            Parámetros: String seccion
        Method Secciones

        Method Claves
            Parámetros: String seccion
        Method Save

        Method Read

        Method GetType

        Method ToString

        Method Equals
            Parámetros: Object obj
        Method GetHashCode


## Clase: ClassifSpanInfo
    Constructores:
        IsPublic = True
            Sin parámetros
        IsPublic = True
            Parámetros: ClassifiedSpan classifSpan, TrueString word
    ClassifSpanInfo.Métodos:
        Method ToString

        Method SetClassifiedSpan
            Parámetros: ClassifiedSpan classifSpan, [Optional] String word = <Nada>
        Method GetType

        Method Equals
            Parámetros: Object obj
        Method GetHashCode


## Clase: DiagInfo
    Constructores:
        IsPublic = True
            Sin parámetros
        IsPublic = True
            Parámetros: Diagnostic diag
    DiagInfo.Métodos:
        Method ToString

        Method SetDiagnostic
            Parámetros: Diagnostic diag
        Method GetType

        Method Equals
            Parámetros: Object obj
        Method GetHashCode


## Clase: Extensiones
    Extensiones.Métodos:
        Method ContienePalabra
            Parámetros: String texto, String palabra
        Method FindString
        Method ComprobarFinLinea
            Parámetros: String selT
        Method QuitarPredeterminado
            Parámetros: String texto, String predeterminado
        Method ContieneLetras
            Parámetros: String texto
        Method QuitarTildes
            Parámetros: String s
        Method ReplaceSiNoEstaPoner
            Parámetros: String texto, String buscar, String poner
        Method ReplaceSiNoEstaPoner
            Parámetros: String texto, String buscar, String poner, StringComparison comparar
        Method ReplaceWordSiNoEstaPoner
            Parámetros: String texto, String buscar, String poner, StringComparison comparar
        Method QuitarTodosLosEspacios
            Parámetros: String texto
        Method CuantasPalabras
            Parámetros: String texto
        Method CambiarCase
            Parámetros: String text, CasingValues queCase
        Method ToTitle
            Parámetros: String text
        Method ToLowerFirstChar
            Parámetros: String text
        Method ToLowerFirst
            Parámetros: String text
        Method Palabras
            Parámetros: String text
        Method AsInteger
        Method Clonar
        Method ToList
        Method ToList
        Method ToList
        Method ReplaceWordOrdinal
            Parámetros: String texto, String buscar, String poner
        Method ReplaceWordIgnoreCase
            Parámetros: String texto, String buscar, String poner
        Method ReplaceWord
            Parámetros: String texto, String buscar, String poner, StringComparison comparar
        Method GetType

        Method ToString

        Method Equals
            Parámetros: Object obj
        Method GetHashCode


## Clase: Marcadores
    Constructores:
        IsPublic = True
            Parámetros: Int32 selStart, String fic
    Marcadores.Métodos:
        Method Add
            Parámetros: Int32 inicio, Int32 selStart
        Method GetSelectionStart
            Parámetros: Int32 inicio
        Method Contains
            Parámetros: Int32 inicio
        Method Remove
            Parámetros: Int32 inicio
        Method Count

        Method ToList

        Method Clear

        Method Where
            Parámetros: Func`2 p
        Method Sort

        Method GetType

        Method ToString

        Method Equals
            Parámetros: Object obj
        Method GetHashCode


## Clase: UtilEnum
    UtilEnum.Métodos:
        Method CheckValidEnumValue
            Parámetros: String arg, Object val, Type class, [Optional] Boolean exception = False
        Method GetType

        Method ToString

        Method Equals
            Parámetros: Object obj
        Method GetHashCode


## Clase: UtilidadesCompilarColorear
    Constructores:
        IsPublic = True
            Sin parámetros
    UtilidadesCompilarColorear.Métodos:
        Method Main

        Method Version

        Method ClasesExpuestas

        Method GetType

        Method ToString

        Method Equals
            Parámetros: Object obj
        Method GetHashCode


## Clase: CasingValues
    CasingValues.Métodos:
        Method Equals
            Parámetros: Object obj
        Method HasFlag
            Parámetros: Enum flag
        Method GetHashCode

        Method ToString

        Method CompareTo
            Parámetros: Object target
        Method ToString
            Parámetros: String format, IFormatProvider provider
        Method ToString
            Parámetros: String format
        Method ToString
            Parámetros: IFormatProvider provider
        Method GetTypeCode

        Method GetType


## Clase: _Closure$__
    Constructores:
        IsPublic = True
            Sin parámetros
    _Closure$__.Métodos:
        Method GetType

        Method ToString

        Method Equals
            Parámetros: Object obj
        Method GetHashCode


## Clase: FormatosEncoding
    FormatosEncoding.Métodos:
        Method Equals
            Parámetros: Object obj
        Method HasFlag
            Parámetros: Enum flag
        Method GetHashCode

        Method ToString

        Method CompareTo
            Parámetros: Object target
        Method ToString
            Parámetros: String format, IFormatProvider provider
        Method ToString
            Parámetros: String format
        Method ToString
            Parámetros: IFormatProvider provider
        Method GetTypeCode

        Method GetType
