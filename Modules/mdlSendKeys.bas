Attribute VB_Name = "mdlSendKeys"
Option Explicit

'   ============== API FUNCTION =================
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer


'   ============= CONSTS ================
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2



'   ============= VARIABLES =============
Public PushingBack As Boolean


Public Sub PushBacks(Backs As Long)
    If BackNumbers > 0 Then
        Dim I As Integer
        For I = 1 To Backs
            PushingBack = True
            keybd_event VK_BACK, VK_BACK_SCAN, 0, 0
            keybd_event VK_BACK, VK_BACK_SCAN, KEYEVENTF_KEYUP, 0
            BackNumbers = BackNumbers - 1
        Next I
        PushingBack = False
        PushBuffer UniBuf
    End If
End Sub


Public Function CheckBackKey(wParam As Long) As Boolean
    If (wParam = VK_BACK And PushingBack = True) Then
        CheckBackKey = True
        Exit Function
    End If
    CheckBackKey = False
End Function


Public Sub SendPasteCommand()
    
    Dim Sh As Integer
    Sh = GetKeyState(VK_SHIFT) And &H80
    
    If Sh Then
        keybd_event VK_SHIFT, VK_SHIFT_SCAN, KEYEVENTF_KEYUP, 0
        
        keybd_event VK_SHIFT, VK_SHIFT_SCAN, 0, 0
        keybd_event VK_INSERT, VK_INSERT_SCAN, KEYEVENTF_EXTENDEDKEY, 0
        keybd_event VK_INSERT, VK_INSERT_SCAN, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
        keybd_event VK_SHIFT, VK_SHIFT_SCAN, KEYEVENTF_KEYUP, 0
        
        keybd_event VK_SHIFT, VK_SHIFT_SCAN, 0, 0
    Else
        keybd_event VK_SHIFT, VK_SHIFT_SCAN, 0, 0
        keybd_event VK_INSERT, VK_INSERT_SCAN, KEYEVENTF_EXTENDEDKEY, 0
        keybd_event VK_INSERT, VK_INSERT_SCAN, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
        keybd_event VK_SHIFT, VK_SHIFT_SCAN, KEYEVENTF_KEYUP, 0
    End If
End Sub



