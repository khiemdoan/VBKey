Attribute VB_Name = "mdlMouseCursor"
Option Explicit

' ------- KHAI BÁO CÁC HÀM WIN API --------

Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetsystemCursor Lib "user32" Alias "SetSystemCursor" (ByVal hCur As Long, ByVal ID As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function GetCursor Lib "user32" () As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

' ------- KHAI BÁO CÁC HA*`NG -------

Private Const OCR_NORMAL As Long = 32512


' ------- KHAI BÁO CÁC BIÊ'N ------


' ------- KHAI BÁO CÁC ENUM -------

Public Enum CURSOR_ICON

    IDC_APPSTARTING = 32650
    IDC_ARROW = 32512
    IDC_CROSS = 32515
    IDC_HAND = 32649
    IDC_HELP = 32651
    IDC_NO = 32648
    IDC_SIZEALL = 32646
    IDC_WAIT = 32514
        
End Enum

Public Enum SPLIT_CURSOR
    VERTICAL_SPLIT = 0
    HORIZONTAL_SPLIT = 1
End Enum

Public Enum SCREEN_SCALE
    SCREEN_PIXEL
    CLIENT_PIXEL
End Enum

' ------ KHAI BÁO CÁC KIÊ?U ------

Public Type POINTAPI
        X As Long
        Y As Long
End Type




'=================================== THÂN CHU*O*NG TRÌNH =========================================================


Public Function CursorPosition(Optional Mode As SCREEN_SCALE = CLIENT_PIXEL) As POINTAPI
    Dim P As POINTAPI
    GetCursorPos P
    CursorPosition.X = IIf(Mode = SCREEN_PIXEL, P.X, P.X * Screen.TwipsPerPixelX)
    CursorPosition.Y = IIf(Mode = SCREEN_PIXEL, P.Y, P.Y * Screen.TwipsPerPixelY)
End Function


Public Function SetCursorIcon(Optional cur As CURSOR_ICON) As Long
      SetCursorIcon = SetCursor(LoadCursor(ByVal 0&, cur))
End Function

Public Sub SetSplitCursor(Obj As Object, ByVal frSplit As SPLIT_CURSOR)
    On Error Resume Next
    If frSplit = VERTICAL_SPLIT Then
        Obj.MouseIcon = LoadResPicture("V_SPLIT", vbResCursor)
    Else
        Obj.MouseIcon = LoadResPicture("H_SPLIT", vbResCursor)
    End If
    Obj.MousePointer = vbCustom
End Sub

