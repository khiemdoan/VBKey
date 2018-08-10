Attribute VB_Name = "mdlMsgBox"
Option Explicit

' ----------- KHAI BÁO CÁC HÀM WIN API --------------

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowTextW Lib "user32" (ByVal hwnd As Long, ByVal lpString As Long) As Long
Private Declare Function MessageBoxW Lib "user32" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal wType As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

' ------------ KHAI BÁO CÁC HA*`NG ----------

Private Const WM_SETFONT = &H30
Private Const WH_CBT = 5
Private Const HCBT_ACTIVATE = 5
Private Const MB_TASKMODAL = &H2000&
Private Const MB_ICONEXCLAMATION = &H30&

' --------- KHAI BÁO CÁC KIÊ?U ----------


' ---------- KHAI BÁO BIÊ'N ----------------

Private hMsgBoxHook As Long
Private MsgFontSize As Long
Private MsgFontName As String
Private MsgFontWeight As Long
Private MsgFontUnderline As Boolean
Private MsgFontItalic As Boolean
Private MsgFontStrike As Boolean
Private MsgButtonText1 As String
Private MsgButtonText2 As String
Private MsgButtonText3 As String
Private MsgButtonText4 As String



' -------- THÂN CHU*O*NG TRÌNH -------

Private Function HookMsgBoxProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    HookMsgBoxProc = CallNextHookEx(hMsgBoxHook, nCode, wParam, lParam)
    If nCode = WH_CBT Then
        Dim hButton As Long, hStatic As Long, hFont As Long
        hFont = CreateFont(MsgFontSize, 0, 0, 0, MsgFontWeight, MsgFontItalic, MsgFontUnderline, MsgFontStrike, 0, 0, 0, 0, 0, MsgFontName)
        UniSystemMenu wParam, "Di chuye63n", "D9o1ng             Alt + F4"
        KeepWindowOnTop wParam, True
        hButton = FindWindowEx(wParam, 0&, "STATIC", vbNullString)
        hStatic = FindWindowEx(wParam, hButton, "STATIC", vbNullString)
        If hStatic <= 0 Then hStatic = hButton
        If hStatic Then SendMessage hStatic, WM_SETFONT, hFont, 1&
        
        hButton = FindWindowEx(wParam, 0&, "BUTTON", "OK")
        If hButton Then SendMessage hButton, WM_SETFONT, hFont, 0&
        If MsgButtonText1 <> "" Then
            SetWindowTextW hButton, StrPtr(TV(MsgButtonText1))
            'SetToolTipObj hButton, MsgButtonText1
        End If
         
        hButton = FindWindowEx(wParam, 0&, "BUTTON", "Cancel")
        If hButton Then SendMessage hButton, WM_SETFONT, hFont, 0&
        If MsgButtonText2 <> "" Or MsgButtonText3 <> "" Then
            'SetToolTipObj hButton, IIf(FindWindowEx(wParam, 0&, "BUTTON", "&Yes"), MsgButtonText3, IIf(FindWindowEx(wParam, 0&, "BUTTON", "&Continue"), MsgButtonText1, MsgButtonText2))
            SetWindowTextW hButton, StrPtr(TV(IIf(FindWindowEx(wParam, 0&, "BUTTON", "&Yes"), MsgButtonText3, IIf(FindWindowEx(wParam, 0&, "BUTTON", "&Continue"), MsgButtonText1, MsgButtonText2))))
        End If
        
        hButton = FindWindowEx(wParam, 0&, "BUTTON", "&Yes")
        If hButton Then SendMessage hButton, WM_SETFONT, hFont, 0&
        If MsgButtonText1 <> "" Then
            SetWindowTextW hButton, StrPtr(TV(MsgButtonText1))
            'SetToolTipObj hButton, MsgButtonText1
        End If
        
        hButton = FindWindowEx(wParam, 0&, "BUTTON", "&No")
        If hButton Then SendMessage hButton, WM_SETFONT, hFont, 0&
        If MsgButtonText2 <> "" Then
            SetWindowTextW hButton, StrPtr(TV(MsgButtonText2))
            'SetToolTipObj hButton, MsgButtonText2
        End If
        
        hButton = FindWindowEx(wParam, 0&, "BUTTON", "&Abort")
        If hButton Then SendMessage hButton, WM_SETFONT, hFont, 0&
        If MsgButtonText1 <> "" Then
            SetWindowTextW hButton, StrPtr(TV(MsgButtonText1))
            'SetToolTipObj hButton, MsgButtonText1
        End If
        
        hButton = FindWindowEx(wParam, 0&, "BUTTON", "&Try Again")
        If hButton Then SendMessage hButton, WM_SETFONT, hFont, 0&
        If MsgButtonText2 <> "" Then
        
            SetWindowTextW hButton, StrPtr(TV(MsgButtonText2))
            'SetToolTipObj hButton, MsgButtonText2
        End If
        
        hButton = FindWindowEx(wParam, 0&, "BUTTON", "&Continue")
        If hButton Then SendMessage hButton, WM_SETFONT, hFont, 0&
        If MsgButtonText3 <> "" Then
            SetWindowTextW hButton, StrPtr(TV(MsgButtonText3))
            'SetToolTipObj hButton, MsgButtonText3
        End If
                
        hButton = FindWindowEx(wParam, 0&, "BUTTON", "&Retry")
        If hButton Then SendMessage hButton, WM_SETFONT, hFont, 0&
        If MsgButtonText1 <> "" Or MsgButtonText2 <> "" Then
            SetWindowTextW hButton, StrPtr(TV(IIf(FindWindowEx(wParam, 0&, "BUTTON", "&Abort"), MsgButtonText2, MsgButtonText1)))
            'SetToolTipObj hButton, IIf(FindWindowEx(wParam, 0&, "BUTTON", "&Abort"), MsgButtonText2, MsgButtonText1)
        End If
        
        hButton = FindWindowEx(wParam, 0&, "BUTTON", "&Ignore")
        If hButton Then SendMessage hButton, WM_SETFONT, hFont, 0&
        If MsgButtonText1 <> "" Then
            SetWindowTextW hButton, StrPtr(TV(MsgButtonText3))
            'SetToolTipObj hButton, StrPtr(TV(MsgButtonText3))
        End If
        
        hButton = FindWindowEx(wParam, 0&, "BUTTON", "&Help")
        If hButton Then SendMessage hButton, WM_SETFONT, hFont, 0&
        If MsgButtonText4 <> "" Then
            SetWindowTextW hButton, StrPtr(TV(MsgButtonText4))
            'SetToolTipObj hButton, StrPtr(TV(MsgButtonText4))
        End If
        
        UnhookWindowsHookEx hMsgBoxHook
    End If
End Function


Public Function MsgBoxTV(Optional ThongBao As String, Optional NutBam As VbMsgBoxStyle = vbOKOnly, Optional TieuDe As String = "", Optional ButtonText1 As String = "D9O1NG", Optional ButtonText2 As String = "", Optional ButtonText3 As String = "", Optional ButtonText4 As String = "", Optional hwnd As Long = &H0, Optional FontN As String = "Tahoma", Optional FontS As Long = 13, Optional FontWeight As Long = 500, Optional FontU As Boolean = False, Optional FontI As Boolean = False, Optional FontStrk As Boolean = False) As VbMsgBoxResult
    MsgFontSize = FontS
    MsgFontName = FontN
    MsgFontWeight = FontWeight
    MsgFontUnderline = FontU
    MsgFontItalic = FontI
    MsgFontStrike = FontStrk
    MsgButtonText1 = ButtonText1
    MsgButtonText2 = ButtonText2
    MsgButtonText3 = ButtonText3
    MsgButtonText4 = ButtonText4
    hMsgBoxHook = SetWindowsHookEx(WH_CBT, AddressOf HookMsgBoxProc, 0&, GetCurrentThreadId)
    If NutBam >= 32 And NutBam <= 38 Then MessageBeep MB_ICONEXCLAMATION
    MsgBoxTV = MessageBoxW(hwnd, StrPtr(TV(ThongBao)), StrPtr(TV(TieuDe)), NutBam Or MB_TASKMODAL)
End Function
