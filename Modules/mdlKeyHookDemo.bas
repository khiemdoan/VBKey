Attribute VB_Name = "mdlKeyHookDemo"
Option Explicit

Private Function HookKeyDemo(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If nCode < 0 Then
        HookKeyDemo = CallNextHookEx(hKeyHook, nCode, wParam, lParam)
        Exit Function
    End If
    
    Dim KBD As KBDLLHOOKSTRUCT, KeyType As Long, KeyResult As Long
    CopyMemory KBD, ByVal lParam, Len(KBD)
    
    
    
    
    
    If wParam = WM_KEYUP And KBD.VkCode = VK_INSERT Then
        ClearClipboard
    End If
    
    
    
    
    
    
    If nCode = HC_ACTION Then
        If CheckBack(KBD.VkCode) = True Then
            frmMain.Caption = "Checking Back"
            HookKeyDemo = 1
            Exit Function
        End If
        
        GetKeyboardState KeyState
        
        If wParam = WM_KEYDOWN Then
            KeyType = ToAscii(KBD.VkCode, KBD.ScanCode, KeyState, KeyResult, 0)
            If (KeyType = 1) Then
                    ProcessKey KeyResult
                    frmMain.T5.Text = StringBuffer
                    frmMain.T2.Text = UniBuf
                    frmMain.T3.Text = AnsiBuf
                    frmMain.T4.Text = BackNumbers
                    frmMain.Caption = Chr$(KeyResult)
                    SetClipboard "Tuyen"
                    PushBuffer 0, UniBuf
            ElseIf (KBD.VkCode <> VK_SHIFT And KBD.VkCode <> VK_INSERT) Then
                ResetBuffer
            End If
        End If
    Else
        ResetBuffer
    End If

    HookKeyDemo = CallNextHookEx(hKeyHook, nCode, wParam, lParam)
    
End Function



Public Sub InitHookDemo(Hook As Boolean)
    If Hook Then
        If Not hKeyHook Then hKeyHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf HookKeyDemo, App.hInstance, 0)
    Else
        If hKeyHook Then UnhookWindowsHookEx hKeyHook
    End If
End Sub

