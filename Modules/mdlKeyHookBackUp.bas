Attribute VB_Name = "mdlKeyHookDemo"
Option Explicit


'   ================= FUNCTIONS ============


Private Function ThuTuc_Hook_Chuot(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If nCode < 0 Then
        ThuTuc_Hook_Chuot = CallNextHookEx(hMouseHook, nCode, wParam, lParam)
        Exit Function
    End If
    If wParam = WM_LBUTTONDOWN Then
        'Thuc hien xoa bo dem tai day
        ResetBuffer
        frmMain.Caption = "KhoiTaoLaiBoDem"
    End If
    ThuTuc_Hook_Chuot = CallNextHookEx(hMouseHook, nCode, wParam, lParam)
End Function


Private Function HookKeyboardProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim KeyType As Long
    If nCode < 0 Then HookKeyboardProc = CallNextHookEx(hKeyHook, nCode, wParam, lParam)
    
    Dim KBD As KBDLLHOOKSTRUCT
    CopyMemory KBD, ByVal lParam, Len(KBD)
    
    
    If (nCode = HC_ACTION) Then
        GetKeyboardState KeyState
        Dim KeyResult As Long
        KeyType = ToAscii(KBD.VkCode, KBD.ScanCode, KeyState, KeyResult, 0)
        If KeyType = 1 Then
            If wParam = WM_KEYDOWN Then
                    ProcessKey KeyResult
                    frmMain.T2.Text = UniBuf
                    frmMain.T3.Text = AnsiBuf
                    frmMain.T4 = BackNumbers
                    frmMain.T5.Text = StringBuffer
                    If BackNumbers > 0 Then
                        PushBacks BackNumbers
                        SetClipboard UniBuf
                        PasteCommand
                    End If
                    Exit Function
            ElseIf wParam = WM_KEYUP Then
                PushBacks BackNumbers
            End If
        End If
    Else
        HookKeyboardProc = CallNextHookEx(hKeyHook, nCode, wParam, lParam)
    End If

    
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
            hMouseHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf ThuTuc_Hook_Chuot, App.hInstance, lThread)
        End If
    Else
        If hMouseHook Then
            UnhookWindowsHookEx hMouseHook
            hMouseHook = 0
        End If
    End If
End Sub

