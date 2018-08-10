Attribute VB_Name = "mdlHook"
Option Explicit

'   ================= FUNCTIONS ============


Private Function HookMouseProc(ByVal nnCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If nnCode < 0 Then
        HookMouseProc = CallNextHookEx(hMouseHook, nnCode, wParam, lParam)
        Exit Function
    End If
    If wParam = WM_LBUTTONDOWN Then
        'Thuc hien xoa bo dem tai day
        ResetBuffer
        frmMain.Caption = "KhoiTaoLaiBoDem"
    End If
    HookMouseProc = CallNextHookEx(hMouseHook, nnCode, wParam, lParam)
End Function


Private Function HookKeyboardProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim keyType As Long, CharBuf As Long, KBD As KBDLLHOOKSTRUCT
    
    If (nCode < 0) Then HookKeyboardProc = CallNextHookEx(hKeyHook, nCode, wParam, lParam)
    
    CopyMemory KBD, ByVal lParam, Len(KBD)
    If nCode = HC_ACTION And wParam = WM_SYSKEYDOWN Then ResetBuffer
    
    If nCode = HC_ACTION And Not ClipboardIsEmpty Then
        
    End If
    
    
    If (nCode = HC_ACTION) And wParam = WM_KEYDOWN Then
            If (CheckBack(KBD.VkCode)) Then
                Exit Function
            End If
            
            Dim Sh As Integer, Ctrl As Integer, Cap As Integer
            
            Sh = GetKeyState(VK_SHIFT)
            Ctrl = GetKeyState(VK_CONTROL)
            Cap = GetKeyState(VK_CAPITAL)
            
            
            
            frmMain.Caption = "Shift : " & Sh & "     -       Control : " & Ctrl & "      -       Capital : " & Cap
            
            GetKeyboardState KeyState
            
            If (KeyState.KB(VK_CONTROL) And &H80) > 0 Then ResetBuffer
                keyType = ToAscii(KBD.VkCode, KBD.ScanCode, KeyState, CharBuf, 0)
                If (keyType = 1) Then
                    Dim Ch As String
                    Ch = Chr$(CharBuf)
                    If (KBD.VkCode >= VK_NUMPAD0 And KBD.VkCode <= VK_NUMPAD9) Then
                        PutToBuffer Ch
                    Else
                        ProcessKey Ch
                        If BackNumbers > 0 And KeysPressed > 0 Then
                            PushBacks BackNumbers
                            HookKeyboardProc = 1
                            Exit Function
                        End If
                    End If
                ElseIf (KBD.VkCode <> VK_SHIFT And KBD.VkCode <> VK_LSHIFT And KBD.VkCode <> VK_RSHIFT And KBD.VkCode <> VK_INSERT) Then
                    ClearBuffer
                End If
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
