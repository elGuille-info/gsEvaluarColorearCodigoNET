'------------------------------------------------------------------------------
' Marcador, clase para los marcadores (bookmark) y última posición  (20/Oct/20)
'
' Usarla para los marcadores (Bookmarks)
' Posiciones en las ventanas abiertas
'
'
' (c) Guillermo (elGuille) Som, 2020
'------------------------------------------------------------------------------
Option Strict On
Option Infer On

Imports System
'Imports System.Data
Imports System.Collections.Generic
Imports System.Text
Imports System.Linq
Imports Microsoft.VisualBasic
'Imports vb = Microsoft.VisualBasic

''' <summary>
''' Clase para manejar marcadores (Bookmarks)
''' </summary>
Public Class Marcadores

    '''' <summary>
    '''' Inicializar la clase con el formulario al que pertenecerá.
    '''' El formulario solo se usa para hacer referencia a la posición actual y el nombre del fichero.
    '''' Ese formulario, llamado Form1 debe tener un RichTextBox llamado richTextBoxCodigo y
    '''' una propiedad llamada NombreFichero.
    '''' Se puede usar también la otra sobrecarga sin indicar el formulario.
    '''' </summary>
    '''' <param name="frm">El formulario que contiene un RichTextBox llamado richTextBoxCodigo</param>
    '''' <remarks>
    '''' Se puede usar otra sobrecarga indicando el RichTextBox.SelectionStart para no depender de un formulario llamado Form1
    '''' </remarks>
    'Public Sub New(frm As Form1)
    '    Me.New(frm.richTextBoxCodigo.SelectionStart, frm.NombreFichero)
    'End Sub

    ''' <summary>
    ''' Inicializa la instancia con la posición de richTextBox.SelectionStart y
    ''' el <see cref="Form1.NombreFichero"/> del formulario con el path completo,
    ''' que después se dejará en el nombre del fichero sin el path.
    ''' </summary>
    ''' <param name="selStart">El valor de richTextBox.SelectionStart</param>
    ''' <param name="fic">El nombreFichero del formulario.</param>
    Public Sub New(selStart As Integer, fic As String)
        Me.Posicion = selStart
        Me.Fichero = System.IO.Path.GetFileName(fic)
        PosInicio = New List(Of Integer)
        PosSelStart = New List(Of Integer)
    End Sub

    ''' <summary>
    ''' El fichero al que hace referencia la lista de Posiciones.
    ''' Se devuelve sin el path.
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Fichero As String

    ''' <summary>
    ''' La posición actual del texto en el richTextBoxCodigo
    ''' (cuando se creó la instancia)
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Posicion As Integer

    '
    ' Para acceder a las colecciones usadas en los Marcadores / Bookmarks
    '

    ''' <summary>
    ''' Añadir una nueva posición del marcador/bookmark
    ''' </summary>
    ''' <param name="inicio">El inicio de la línea</param>
    ''' <param name="selStart">La posición indicada por richTextBoxCodigo.SelectionStart</param>
    Public Sub Add(inicio As Integer, selStart As Integer)
        PosInicio.Add(inicio)
        PosSelStart.Add(selStart)
    End Sub

    ''' <summary>
    ''' Obtiene el valor de SelectionStart de la posición del inicio de línea
    ''' </summary>
    ''' <param name="inicio">El inicio de la línea</param>
    ''' <returns>Devuelve el valor indicado en inicio si no está asignada esa posición</returns>
    Public Function GetSelectionStart(inicio As Integer) As Integer
        Dim index = PosInicio.IndexOf(inicio)
        If index = -1 Then Return inicio
        Return PosSelStart(index)
    End Function

    ''' <summary>
    ''' Comprueba si la posición indicada está en la colección de posInicio.
    ''' Devuelve True o False según esté o no esté respectivamente.
    ''' </summary>
    ''' <param name="inicio"></param>
    ''' <returns></returns>
    Public Function Contains(inicio As Integer) As Boolean
        Return PosInicio.Contains(inicio)
    End Function

    ''' <summary>
    ''' Quita la posición indicada, tanto en Posiciones como en PosSelStart
    ''' </summary>
    ''' <param name="inicio"></param>
    ''' <returns>True si se quitó, False si no está esa posición y no se ha quitado</returns>
    Public Function Remove(inicio As Integer) As Boolean
        Dim index = PosInicio.IndexOf(inicio)
        If index = -1 Then Return False
        PosInicio.RemoveAt(index)
        PosSelStart.RemoveAt(index)
        Return True
    End Function

    ''' <summary>
    ''' El número de valores en PosInicio
    ''' </summary>
    ''' <returns></returns>
    Public Function Count() As Integer
        Return PosInicio.Count
    End Function

    ''' <summary>
    ''' La posición por el índice
    ''' </summary>
    ''' <param name="index">El índice en base cero dentro de la colección</param>
    ''' <returns>Devuelve una Tupla con el Inicio y SelStart</returns>
    Default Public ReadOnly Property Item(index As Integer) As (Inicio As Integer, SelStart As Integer)
        Get
            Return (PosInicio.Item(index), PosSelStart.Item(index))
        End Get
    End Property

    ''' <summary>
    ''' Devuelve la colección PosSelStart
    ''' </summary>
    ''' <returns></returns>
    Public Function ToList() As List(Of Integer)
        ' Llamar a ToList para crear una copia                      (23/Oct/20)
        Return PosSelStart.ToList
    End Function

    ''' <summary>
    ''' Elimina el contenido de las colecciones
    ''' </summary>
    Public Sub Clear()
        PosInicio.Clear()
        PosSelStart.Clear()
    End Sub

    ''' <summary>
    ''' Filtra el contenido de Posiciones según el predicado indicado.
    ''' Para buscar la posición anterior: Where(Function(x) x &lt; posActual.Inicio)
    ''' Para buscar la posición siguiente: Where(Function(x) x > posActual.Inicio)
    ''' </summary>
    ''' <param name="p">Una función del tipo (Integer, Boolean)</param>
    ''' <returns></returns>
    Public Function Where(p As Func(Of Integer, Boolean)) As IEnumerable(Of Integer)
        Return PosInicio.Where(p)
    End Function

    ''' <summary>
    ''' Clasifica el contenido de las dos colecciones.
    ''' </summary>
    ''' <remarks>Se supone que no debería haber desajustes ya que
    ''' PosInicio contiene la posición del primer carácter de la línea y
    ''' PosSelStart la posición en esa línea,
    ''' por tanto los dos valores estarán más o menos cercanos y ambos en la misma línea.</remarks>
    Public Sub Sort()
        PosInicio.Sort()
        PosSelStart.Sort()
    End Sub

    ''' <summary>
    ''' Las posiciones para el formulario indicado.
    ''' </summary>
    ''' <returns></returns>
    Private ReadOnly Property PosInicio As List(Of Integer)

    ''' <summary>
    ''' Para almacenar la posición real del cursor.
    ''' </summary>
    ''' <returns></returns>
    Private ReadOnly Property PosSelStart As List(Of Integer)

End Class
