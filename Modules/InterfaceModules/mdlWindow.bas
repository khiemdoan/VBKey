Attribute VB_Name = "mdlWindow"
Option Explicit

' --------- KHAI BÁO CÁC HÀM WIN API -----------

Private Declare Function GetClassLong Lib "user32.dll" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetClassLong Lib "user32.dll" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowW" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function DefWindowProcW Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long



' --------- KHAI BÁO CÁC KIÊ?U -----------


Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

' --------- KHAI BÁO CÁC ENUM ---------------



' --------- KHAI BÁO CÁC HA*`NG --------

Private Const GCL_STYLE = -26
Private Const CS_DROPSHADOW = &H20000
Private Const ULW_COLORKEY = &H2
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SPI_GETWORKAREA As Long = 48


' -------- KHAI BÁO BIÊ'N --------

Private OldLong As Long

' ======================== THÂN CHU*O*NG TRÌNH =================================

Public Function TransparentWindow(hwnd As Long, TransLevel As Byte, Optional Trans As Boolean = False) As Long
  If Trans Then
    OldLong = GetWindowLong(hwnd, GWL_EXSTYLE)
    SetWindowLong hwnd, GWL_EXSTYLE, WS_EX_LAYERED
    TransparentWindow = SetLayeredWindowAttributes(hwnd, &HFFCCCC, TransLevel, ULW_COLORKEY)
  Else
    If OldLong Then TransparentWindow = SetWindowLong(hwnd, GWL_EXSTYLE, OldLong)
  End If
End Function


Public Function DragWindow(ByVal hwnd As Long) As Long
  ReleaseCapture
  DragWindow = SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Function


Public Function ShadowWindow(hwnd As Long, Optional Bol As Boolean = False)
    If Bol = True Then
        ShadowWindow = SetClassLong(hwnd, GCL_STYLE, GetClassLong(hwnd, GCL_STYLE) Or CS_DROPSHADOW)
    Else
        ShadowWindow = SetClassLong(hwnd, GCL_STYLE, GetClassLong(hwnd, GCL_STYLE) And Not CS_DROPSHADOW)
    End If
    ShowWindow hwnd, 6
    ShowWindow hwnd, 9
End Function



Public Function KeepWindowOnTop(ByVal hwnd As Long, Optional OnTop As Boolean = False)
       KeepWindowOnTop = SetWindowPos(hwnd, IIf(OnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE)
End Function


Public Function SetUniText(ByVal hwnd As Long, ByVal sUni As String) As Long
    SetUniText = DefWindowProcW(hwnd, &HC, 0, StrPtr(TV(sUni)))
End Function

