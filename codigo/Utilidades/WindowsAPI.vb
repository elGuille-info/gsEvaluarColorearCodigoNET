'------------------------------------------------------------------------------
' WindowsAPI                                                        (17/Oct/20)
'
' Funciones para acceder a las funciones del API de Windows
'
' (c) Guillermo (elGuille) Som, 2020
'------------------------------------------------------------------------------
Option Strict On
Option Infer On

Imports System
Imports System.Windows.Forms
Imports System.Linq
Imports System.Collections.Generic
'Imports System.IO
'Imports System.Text
'Imports System.ComponentModel
'Imports System.Security.Cryptography
Imports System.Runtime.InteropServices

Public Module WindowsAPI

    Public Function FindTopWindowTitle(titulo As String) As IntPtr ' Integer
        ' Buscará la ventana indicada en el título                      (25/Sep/02)
        ' y devolverá el handle de la misma o un cero si no se ha hayado.
        Dim lpClassName As String = ""
        Return FindWindow(lpClassName, titulo)
    End Function

    'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer

    Public Const WM_COMMAND As Integer = &H112
    Public Const WM_CLOSE As Integer = &HF060

    '' FindWindowEx también busca en las ventanas hijas                  (25/Sep/02)
    '' Esta constante se puede usar sólo con Windows 2000/XP para indicar que sólo
    '' se buscará en las ventanas que reciban mensajes.
    'Public Const HWND_MESSAGE As Integer = (-3)
    'Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Integer, ByVal hWnd2 As Integer, ByVal lpsz1 As String, ByVal lpsz2 As String) As Integer

    <DllImport("user32.dll")>
    Public Function FindWindow(lpClassName As String, lpWindowName As String) As IntPtr ' Integer
    End Function

    <System.Runtime.InteropServices.DllImport("User32.dll")>
    Public Function SendMessage(hWnd As IntPtr, msg As Long, wParam As Long, lParam As Long) As Integer
    End Function


    <DllImport("User32.dll", ExactSpelling:=True, CharSet:=System.Runtime.InteropServices.CharSet.Auto)>
    Public Function MoveWindow(hWnd As IntPtr,
                               x As Integer, y As Integer,
                               cx As Integer, cy As Integer,
                               repaint As Boolean) As Boolean
    End Function

    ''' <summary>
    ''' Estructura para la posición y tamaño de la ventana
    ''' </summary>
    Public Structure RECTAPI
        Dim Left As Integer
        Dim Top As Integer
        Dim Right As Integer
        Dim Bottom As Integer
    End Structure

    Private Structure POINTAPI
        Public x As Integer
        Public y As Integer
    End Structure

    Private Structure WINDOWPLACEMENT
        Public Length As Integer
        Public flags As Integer
        Public showCmd As Integer
        Public ptMinPosition As POINTAPI
        Public ptMaxPosition As POINTAPI
        Public rcNormalPosition As RECTAPI
    End Structure

    ''' <summary>
    ''' La posición normal de la ventana al llamarla con GetWindowState
    ''' </summary>
    ''' <returns></returns>
    Public Property PosicionNormal As RECTAPI

    ''' <summary>
    ''' Posicionar la ventana con el Handle indicado encima y en el centro
    ''' </summary>
    ''' <param name="hWnd"></param>
    ''' <returns></returns>
    Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As IntPtr) As Integer

    ''' <summary>
    ''' Traer la ventana al frente la ventana indicada por el Handle.
    ''' </summary>
    ''' <param name="hWnd"></param>
    Public Sub BringToTop(hWnd As IntPtr)
        BringWindowToTop(hWnd)
    End Sub

    Const HWND_TOP As Integer = 0
    Const SWP_NOSIZE As Integer = &H1
    Const SWP_NOMOVE As Integer = &H2
    Const SWP_NOACTIVATE As Integer = &H10
    Const SWP_SHOWWINDOW As Integer = &H40
    Const SWP_FLAGS As Integer = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE

    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As IntPtr,
                                                        ByVal hWndInsertAfter As Integer,
                                                        ByVal x As Integer, ByVal y As Integer,
                                                        ByVal cx As Integer, ByVal cy As Integer,
                                                        ByVal wFlags As Integer) As Integer

    ''' <summary>
    ''' Posiciona y asigna el tamaño a la ventana indicada por el Handle.
    ''' Llama a SetWindowPos del API con HWND_TOPMOST y SWP_NOZORDER Or SWP_NOACTIVATE
    ''' </summary>
    ''' <param name="hWnd">El Handle de la ventana a posicionar</param>
    ''' <param name="left">La posición izquierda</param>
    ''' <param name="top">La posición arriba</param>
    ''' <param name="width">El ancho de la ventana</param>
    ''' <param name="height">El alto de la ventana</param>
    ''' <remarks>15/Sep/2020</remarks>
    Public Function SetWindowPosition(hWnd As IntPtr,
                                      left As Integer, top As Integer,
                                      width As Integer, height As Integer) As Integer
        Return SetWindowPos(hWnd, HWND_TOP, left, top, width, height, SWP_FLAGS)
    End Function

    Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As IntPtr, ByRef lpRect As RECTAPI) As Integer

    ''' <summary>
    ''' Obtiene la posición de la ventana indicada por el Handle.
    ''' </summary>
    ''' <param name="hWnd">El Handle de la ventana a posicionar</param>
    ''' <returns>Una tupla con los valores de Left, Top, Width y Height</returns>
    ''' <remarks>15/Sep/2020</remarks>
    Public Function GetWindowPosition(hWnd As IntPtr) As (Left As Integer, Top As Integer,
                                                          Width As Integer, Height As Integer)

        Dim ra As RECTAPI
        GetWindowRect(hWnd, ra)
        PosicionNormal = ra

        Dim ret As (Left As Integer, Top As Integer, Width As Integer, Height As Integer)

        ret.Left = ra.Left
        ret.Top = ra.Top
        ret.Width = CInt(Math.Abs(ra.Left - ra.Right))
        ret.Height = CInt(Math.Abs(ra.Bottom - ra.Top))

        Return ret
    End Function

    Private Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Integer

    ''' <summary>
    ''' Obtiene el estado de la ventana indicada por el Handle.
    ''' </summary>
    ''' <param name="hWnd"></param>
    ''' <returns></returns>
    Public Function GetWindowState(hWnd As IntPtr) As FormWindowState
        Dim intRet As Integer
        Dim wpTemp As WINDOWPLACEMENT

        wpTemp.Length = System.Runtime.InteropServices.Marshal.SizeOf(wpTemp)
        intRet = GetWindowPlacement(hWnd, wpTemp)

        PosicionNormal = wpTemp.rcNormalPosition

        Dim ws As FormWindowState
        If wpTemp.showCmd = 1 Then
            'normal
            ws = FormWindowState.Normal
        ElseIf wpTemp.showCmd = 2 Then
            'minimized
            ws = FormWindowState.Minimized
        ElseIf wpTemp.showCmd = 3 Then
            'maximized
            ws = FormWindowState.Maximized
        End If
        Return ws
    End Function

    Private Declare Function SetWindowPlacement Lib "user32" (ByVal hWnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Integer

    ''' <summary>
    ''' Asigna el estado a la ventana indicada por el Handle.
    ''' <see cref="PosicionNormal"/> debe estar asignado con la posición y tamaño de la ventana
    ''' </summary>
    ''' <param name="hWnd"></param>
    ''' <param name="ws"></param>
    Public Sub SetWindowState(hWnd As IntPtr, ws As FormWindowState)
        Dim wpTemp As WINDOWPLACEMENT
        wpTemp.Length = System.Runtime.InteropServices.Marshal.SizeOf(wpTemp)
        If ws = FormWindowState.Maximized Then
            wpTemp.showCmd = 3
        ElseIf ws = FormWindowState.Minimized Then
            wpTemp.showCmd = 2
        ElseIf ws = FormWindowState.Normal Then
            wpTemp.showCmd = 1
        End If
        wpTemp.rcNormalPosition = PosicionNormal
        SetWindowPlacement(hWnd, wpTemp)
    End Sub

End Module