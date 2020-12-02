# gsUtilidadesNET
Utilidades para colorear, evaluar y compilar con dotnet Evaluar si tiene fallos, Compilar, Ejecutar y Colorear el código y para HTML  (para .NET 5.0 revisión del 01-dic-2020)<br>

> __NOTA__
> Para usar las opciones de colorear y compilar debes [instalar .NET Core (5.0  o superior)](https://dotnet.microsoft.com/download/dotnet).
<br>
<br>
Si quieres ver un resumen de las clases, métodos, propiedades, etc de esta "biblioteca" mira esto:

[Clases expuestas por gsUtilidadesNET](https://github.com/elGuille-info/gsUtilidadesNET/blob/master/Clases%20expuestas%20por%20gsUtilidadesNET.md)

<br>

## Clases expuestas
Esta "biblioteca" es una aplicación para WindowsForms que expone varias clases.

**CompararString** - Clase para hacer las comparaciones de cadenas que implementa IComparer(Of String) y IEqualityComparer(Of String).
**Compilar** - Clase para evaluar, compilar y colorear código de VB y C#.<br>
**Config** - Clase para manejar ficheros de configuración.<br>
**DiagClassifSpanInfo.ClassifSpanInfo** - Clase para manejar información del tipo ClassifiedSpan.<br>
**DiagClassifSpanInfo.DiagInfo** - Clase para manejar información del tipo Diagnostic.<br>
**Extensiones** - Módulo con extensiones para varias clases y tipos de datos.<br>
**FormInfo** - Formulario para mostrar la información de esta utilidad y las clases que la componen.<br>
**FormVisorHTML** - Formulario visor de páginas HTML.<br>
**FormRecortes** - Formulario para mostrar recortes (texto) y pegarlos.<br>
**FormProcesando** - Formulario para mostrar un diálogo cuando se está procesando (acciones largas).<br>
**Marcadores** - Clase para manejar marcadores (Bookmarks).<br>
**UtilEnum** - Clase con la definición de la enumeración FormatosEncoding y utilidades para manejar enumeraciones.<br>
**UtilidadesCompilarColorear** - Clase con definición de la versión y las clases expuestas.<br>

<br>

## Proyectos disponibles
En el código fuente verás dos ficheros de proyectos, uno lo utilizo para compilar a 32 bits (x86) y el otro para compilar a 64 bits (x64).<br>
Esto es porque esta biblioteca la utilizo en gsEvaluarColorearCodigoNET y tengo dos versiones diferentes también.<br>
Estos proyectos están firmados con mi fichero de claves (__elGuille.snk__) que no se incluye por razones obvias.<br>
<br>
### Nota sobre la versión de 32 bits (x86)
La de 32 bits la he probado en una máquina virtual con Windows 7 Enterpise SP1 (32 bits) y funciona bien, salvo que al compilar da error.<br>
Es porque falta una actualización de Microsoft (windows6.1-kb2533623), pero que no se puede descargar porque Microsoft lo retiró.<br>
Buscando con Bing me encontré con el aviso de que se podía descargar la actualización de kb4457144.<br>
La instalé, pero seguía dando el mismo error.<br>
Se ve que no está tan bien preparado el DOTNET para trabajar con Windows 7, por tanto, será recomendable usar Windows 10 de 32 bits para que las aplicaciones de .NET funcionen.<br>
Si alguien prueba la utilidad en un sistema de 32 bits y le funciona, que lo comente en la entrada que pondré en el blog.<br>
Gracias.<br>
<br>
