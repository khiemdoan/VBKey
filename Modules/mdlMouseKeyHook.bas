Attribute VB_Name = "mdlMouseKeyHook"

Option Explicit

'   ========================= API FUNCTIONS =================

Private Declare Function SetWindowsHookEx Lib "user32.dll" Alias "SetWindowsHookExW" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As KeyboardBytes) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpbKeyState As KeyboardBytes, lpwTransKey As Long, ByVal fuState As Long) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function EndMenu Lib "user32.dll" () As Long
Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long


'   ===================== API TYPES =================

Private Type KBDLLHOOKSTRUCT
    VkCode As Long
    ScanCode As Long
    Time As Long
    Flags As Long
    dwExtraInfo As Long
End Type


Public Type KeyboardBytes
    KB(0 To 255) As Byte
End Type


'   ===================== ENUMES =====================


'   ================== CONST ======================

Private Const WM_LBUTTONDOWN = &H201
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_LBUTTONUP = &H202
Private Const WM_MBUTTONUP = &H208
Private Const WM_RBUTTONUP = &H205

Private Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101

Private Const HC_ACTION = 0
Private Const WH_KEYBOARD_LL As Long = 13
Private Const WH_MOUSE_LL As Long = 14
Private Const WH_KEYBOARD = 2

Private Const VK_NUMPAD0 = &H60
Private Const VK_NUMPAD9 = &H69
Private Const VK_TAB = &H9
Private Const VK_ESCAPE = &H1B

'   ================== VARIABLES ==================

Private hMouseHook As Long
Private hKeyHook As Long
Public KeyState As KeyboardBytes
Private CheckingSwitch As Boolean

Private Function GetActiveWindowClassName() As String
    Dim S As String
    S = String$(255, Chr$(0))
    GetClassName GetForegroundWindow, S, Len(S)
    GetActiveWindowClassName = Trim$(Left$(S, InStr(1, S, Chr$(0), vbTextCompare) - 1))
End Function



Private Function HookMouseProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If nCode < 0 Then
        HookMouseProc = CallNextHookEx(hMouseHook, nCode, wParam, lParam)
        Exit Function
    End If
    
    If nCode = HC_ACTION And VietNameseKeyboard Then
        
        If GetForegroundWindow = frmMain.hwnd Then
            HookMouseProc = CallNextHookEx(hMouseHook, nCode, wParam, lParam)
            Exit Function
        End If
        
        If (wParam = WM_LBUTTONDOWN) Or (wParam = WM_MBUTTONDOWN) Or (wParam = WM_RBUTTONDOWN) Then
            'Thuc hien xoa bo dem tai day
            ClearBuffer
        End If
    Else
        HookMouseProc = CallNextHookEx(hMouseHook, nCode, wParam, lParam)
    End If
    HookMouseProc = CallNextHookEx(hMouseHook, nCode, wParam, lParam)
End Function


Private Function HookKeyboardProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim keyType As Long, CharBuf As Long, KBD As KBDLLHOOKSTRUCT
    
    If (nCode < 0) Then
        HookKeyboardProc = CallNextHookEx(hKeyHook, nCode, wParam, lParam)
        Exit Function
    End If
    
    
    If GetActiveWindowClassName = "ConsoleWindowClass" Then
        HookKeyboardProc = CallNextHookEx(hKeyHook, nCode, wParam, lParam)
        Exit Function
    End If
    
    CopyMemory KBD, ByVal lParam, Len(KBD)
            
    If nCode = HC_ACTION And wParam = WM_KEYUP Then

        If (KBD.VkCode = 160 Or KBD.VkCode = 161 Or KBD.VkCode = 162 Or KBD.VkCode = 163) And switchMode = CONTROL_SHIFT Then
            If CheckingSwitch Then
                CheckingSwitch = False
                ChangeKeyMode
            End If
        ElseIf (KBD.VkCode = 164 Or KBD.VkCode = 165 Or KBD.VkCode = 162 Or KBD.VkCode = 163) And switchMode = CONTROL_ALT Then
            If CheckingSwitch = True Then
                CheckingSwitch = False
                ChangeKeyMode
            End If
        End If
    End If
    
    If nCode = HC_ACTION And wParam = WM_KEYDOWN Then
            If GetKeyState(SwitchKey1) And &H80 And GetKeyState(SwitchKey2) And &H80 Then
                CheckingSwitch = True
            Else
                CheckingSwitch = False
            End If
    End If
    
    If (nCode = HC_ACTION) And wParam = WM_KEYDOWN And VietNameseKeyboard Then
            
            If GetKeyState(VK_CONTROL) And &H80 Then ClearBuffer
            If (CheckBackKey(KBD.VkCode)) Then
                Exit Function
            End If
                                    
                                    
            If (PushingBack) Then
                If (KBD.VkCode <> VK_BACK) Then
                    keybd_event KBD.VkCode, KBD.ScanCode, KBD.Flags, KBD.dwExtraInfo
                    HookKeyboardProc = 1
                    Exit Function
                End If
                HookKeyboardProc = CallNextHookEx(hKeyHook, nCode, wParam, lParam)
            End If
        
            GetKeyboardState KeyState

            keyType = ToAscii(KBD.VkCode, KBD.ScanCode, KeyState, CharBuf, 0)
            
            
            If (keyType = 1) Then
                If (KBD.VkCode >= VK_NUMPAD0 And KBD.VkCode <= VK_NUMPAD9) Then
                    PutToBuffer Chr$(CharBuf)
                Else
                    If CharBuf = VK_ESCAPE Then
                        Process_Escape_Key CharBuf
                        Exit Function
                    ElseIf CharBuf = VK_TAB Then
                        ClearBuffer
                    Else
                        ProcessKey CharBuf
                        If BackNumbers > 0 And KeyPushed > 0 Then
                            PushBacks BackNumbers
                            HookKeyboardProc = 1
                            Exit Function
                        End If
                    End If
                End If
            ElseIf (KBD.VkCode <> VK_SHIFT And KBD.VkCode <> VK_LSHIFT And KBD.VkCode <> VK_RSHIFT And KBD.VkCode <> VK_INSERT And KBD.VkCode <> VK_CONTROL And KBD.VkCode <> VK_LCONTROL And KBD.VkCode <> VK_RCONTROL) Then
                ClearBuffer
            End If
    ElseIf nCode = HC_ACTION And wParam <> WM_KEYDOWN And wParam <> WM_KEYUP Then
        ClearBuffer
    End If
    
    HookKeyboardProc = CallNextHookEx(hKeyHook, nCode, wParam, lParam)

    
End Function


Public Sub InitKeyHook(iHook As Boolean, Optional lThread As Long = 0)
    If iHook = True Then
        If Not hKeyHook Then
            hKeyHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf HookKeyboardProc, App.hInstance, lThread)
        End If
    Else
        If hKeyHook Then
            UnhookWindowsHookEx hKeyHook
            hKeyHook = 0
        End If
    End If
End Sub


Public Sub InitMouseHook(iHook As Boolean, Optional lThread As Long = 0)
    If iHook = True Then
        If Not hMouseHook Then
            hMouseHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf HookMouseProc, App.hInstance, lThread)
        End If
    Else
        If hMouseHook Then
            UnhookWindowsHookEx hMouseHook
            hMouseHook = 0
        End If
    End If
End Sub
