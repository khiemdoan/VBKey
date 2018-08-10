Attribute VB_Name = "mdlNotifyIcon"
Option Explicit


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Type NOTIFYICONDATA
    cbSize  As Long
    hWnd  As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeout As Long
    uVersion As Long
    szInfoTitle As String * 64
    dwInfoFlags  As Long
End Type

Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const NIM_MODIFY = &H1
Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const WM_LBUTTONDOWN = &H201

Private NF As NOTIFYICONDATA



Public Sub InitTrayIcon(Optional Icon As Picture, Optional sTip As String = "")
    With NF
        .cbSize = Len(NF)
        .hIcon = Icon
        .hWnd = frmMain.hWnd
        .szTip = (sTip & Chr(0))
        .uCallbackMessage = WM_LBUTTONDOWN
        .uFlags = NIF_ICON Or NIF_TIP
        .uID = 100
    End With
    
    Shell_NotifyIcon NIM_ADD, NF
End Sub

