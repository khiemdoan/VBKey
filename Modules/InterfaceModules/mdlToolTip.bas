Attribute VB_Name = "mdlToolTip"
Option Explicit

' ----- KHAI BÁO CÁC HÀM WIN API -------

Private Declare Function CreateWindowEX Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDc As Long, ByVal nIndex As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GetComboBoxInfo Lib "user32" (ByVal hwndCombo As Long, CBInfo As COMBOBOXINFO) As Long


' ------- KHAI BÁO CÁC HA*`NG -----

Private Const TTS_ALWAYSTIP = &H1
Private Const TTS_NOPREFIX = &H2
Private Const CW_USEDEFAULT = &H80000000
Private Const WM_USER = &H400
Private Const TTM_ADDTOOL = (WM_USER + 4)
Private Const TTM_ADDTOOLW = (WM_USER + 50)
Private Const TTM_DELTOOLW As Long = (WM_USER + 51)
Private Const TTF_IDISHWND = &H1
Private Const TTF_CENTERTIP = &H2
Private Const TTF_SUBCLASS = &H10
Private Const WM_SETFONT = &H30
Private Const LF_FACESIZE = 32
Private Const TTS_BALLOON As Long = &H40
Private Const TTM_SETTIPBKCOLOR As Long = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR As Long = (WM_USER + 20)
Private Const TTM_SETTITLEW As Long = (WM_USER + 33)
Private Const ECM_FIRST         As Long = &H1500
Private Const EM_SHOWBALLOONTIP As Long = (ECM_FIRST + 3)
Private Const EM_HIDEBALLOONTIP As Long = (ECM_FIRST + 4)

' ------- KHAI BÁO CÁC KIÊ?U ------

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type TOOLINFO
    cbSize As Long
    uFlags As Long
    hwnd As Long
    uID As Long
    cRect As RECT
    hInst As Long
    lpszText As Long
End Type

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type BALLOONTIP
    cbStruct As Long
    pszTitle As String
    pszText As String
    tIcon As Long
End Type

Private Type COMBOBOXINFO
   cbSize As Long
   rcItem As RECT
   rcButton As RECT
   stateButton  As Long
   hwndCombo  As Long
   hwndEdit  As Long
   hwndList As Long
End Type

' ----- KHAI BÁO CÁC ENUM -----

Public Enum ToolTipStyle
    CLASSIC = 0
    BALLOON = 1
End Enum

Public Enum ToolTipIcon
    TTI_ERROR = 3
    TTI_INFO = 1
    TTI_NONE = 0
    TTI_WARNING = 2
End Enum

' ------ KHAI BÁO BIÊ'N -----

Private hWnd_ToolTip As Long



' ====================== THÂN CHU*O*NMG TRÌNH =====================

Private Sub InitToolTip(Optional ttStyle As ToolTipStyle = CLASSIC, Optional ttFontName As String = "Tahoma", Optional ttFontWeight As Long = 500, Optional ttFontSize As Long = 8, Optional ttFontUnderline As Boolean = False, Optional ttFontItalic As Boolean = False, Optional ttFontStrike As Boolean = False)
    hWnd_ToolTip = CreateWindowEX(8, "Tooltips_class32", 0&, IIf(ttStyle = CLASSIC, TTS_NOPREFIX Or TTS_ALWAYSTIP, TTS_NOPREFIX Or TTS_ALWAYSTIP Or TTS_BALLOON), CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, 0&, 0&, App.hInstance, 0&)

    Dim lF As LOGFONT, hFont As Long
     
    lF.lfHeight = -MulDiv(ttFontSize, GetDeviceCaps(GetDC(hWnd_ToolTip), 90&), 72)
    lF.lfItalic = Abs(ttFontItalic)
    lF.lfStrikeOut = Abs(ttFontStrike)
    lF.lfUnderline = Abs(ttFontUnderline)
    lF.lfWeight = ttFontWeight
   
    Dim tmpArr() As Byte
    tmpArr = StrConv(ttFontName & Chr$(0), vbFromUnicode)
    Dim I As Integer, lArr As Long
    lArr = UBound(tmpArr)
    For I = 0 To lArr
        lF.lfFaceName(I) = tmpArr(I)
    Next I
    hFont = CreateFontIndirect(lF)
    SendMessageLong hWnd_ToolTip, WM_SETFONT, hFont, 1&
End Sub

Private Sub SetTipTextColor(ByVal tColor As Long)
        SendMessageLong hWnd_ToolTip, TTM_SETTIPTEXTCOLOR, tColor, 0
End Sub

Private Sub SetTipBackColor(ByVal bColor As Long)
    SendMessageLong hWnd_ToolTip, TTM_SETTIPBKCOLOR, bColor, 0
End Sub

Private Sub SetTipTitle(ByVal sTitle As String, Optional ByVal TipIcon As ToolTipIcon = 0)
      SendMessageLong hWnd_ToolTip, TTM_SETTITLEW, TipIcon, StrPtr(TV(sTitle))
End Sub

Public Sub DestroyToolTip()
    If hWnd_ToolTip Then DestroyWindow hWnd_ToolTip
End Sub


Public Sub SetToolTipObj(objhWnd As Long, sTipText As String, Optional ToolTipTitle As String, Optional ToolTipIcon As ToolTipIcon, Optional ToolTipStyle As ToolTipStyle = CLASSIC, Optional CenterTip As Boolean = False, Optional ToolTipTextColor As Long = &HFF0000, Optional ToolTipBackColor As Long = &H1FFFF, Optional ToolTipFontName As String = "Tahoma", Optional ToolTipFontWeight As Long = 500, Optional ToolTipFontSize As Long = 8, Optional ToolTipFontUnderline As Boolean = False, Optional ToolTipFontItalic As Boolean = False, Optional ToolTipFontStrike As Boolean = False)
    If Not hWnd_ToolTip Then InitToolTip ToolTipStyle, ToolTipFontName, ToolTipFontWeight, ToolTipFontSize, ToolTipFontUnderline, ToolTipFontItalic, ToolTipFontStrike
    Dim TTI As TOOLINFO
    With TTI
        .hwnd = objhWnd
        .uFlags = TTF_IDISHWND Or TTF_SUBCLASS
        If CenterTip Then
            .uFlags = .uFlags Or TTF_CENTERTIP
        End If
        .uID = objhWnd
        .lpszText = StrPtr(TV(sTipText))
        .cbSize = Len(TTI)
    End With
    SendMessageLong hWnd_ToolTip, TTM_ADDTOOLW, 0, VarPtr(TTI)
    SetTipTextColor ToolTipTextColor
    SetTipBackColor ToolTipBackColor
    SetTipTitle ToolTipTitle, ToolTipIcon
End Sub


Public Function HideBalloonTip(Control As Control) As Boolean

    Dim hwnd As Long

    Select Case UCase(TypeName(Control))

        Case "TEXTBOX"
            hwnd = Control.hwnd
        Case "RICHTEXTBOX"
            hwnd = Control.hwnd
        Case "COMBOBOX"
            If (Control.Style = 0 Or 1) Then
                Dim CBI As COMBOBOXINFO
                CBI.cbSize = Len(CBI)
                Call GetComboBoxInfo(Control.hwnd, CBI)
                hwnd = CBI.hwndEdit
            Else
                Exit Function
            End If
        Case Else
            hwnd = Control.hwnd
    End Select

    HideBalloonTip = SendMessage(hwnd, EM_HIDEBALLOONTIP, 0&, 0&)

End Function




Public Function ShowBalloonTip(Control As Control, szText As String, Optional Title As String, Optional TitleIcon As ToolTipIcon = 1) As Boolean

    Dim BLT As BALLOONTIP
    Dim hwnd As Long
    Select Case UCase(TypeName(Control))
        Case "COMBOBOX"
            If (Control.Style = 0 Or Control.Style = 1) Then
                Dim CBI As COMBOBOXINFO
                CBI.cbSize = Len(CBI)
                Call GetComboBoxInfo(Control.hwnd, CBI)
                hwnd = CBI.hwndEdit
            Else
                Exit Function
            End If
        Case Else
            hwnd = Control.hwnd
    End Select

    With BLT
        .cbStruct = Len(BLT)
        .pszTitle = StrConv(TV(Title), vbUnicode)
        .pszText = StrConv(TV(szText), vbUnicode)
        .tIcon = TitleIcon
    End With

    ShowBalloonTip = SendMessage(hwnd, EM_SHOWBALLOONTIP, 0&, BLT)

End Function


Private Function TV(str$) As String
    Dim ansi$, UNI$, I&, sTem$, sUni$, arrUni() As String
    ansi = "a1|a2|a3|a4|a5|a6|a8|a61a62a63a64a65a81a82a83a84a85A1|A2|A3|A4|A5|A6|A8|A61A62A63A64A65A81A82A83A84A85e1|e2|e3|e4|e5|e6|e61e62e63e64e65E1|E2|E3|E4|E5|E6|E61E62E63E64E65i1|i2|i3|i4|i5|I1|I2|I3|I4|I5|o1|o2|o3|o4|o5|o6|o7|o61o62o63o64o65o71o72o73o74o75O1|O2|O3|O4|O5|O6|O7|O61O62O63O64O65O71O72O73O74O75u1|u2|u3|u4|u5|u7|u71u72u73u74u75U1|U2|U3|U4|U5|U7|U71U72U73U74U75y1|y2|y3|y4|y5|Y1|Y2|Y3|Y4|Y5|d9|D9|"
    UNI = "E1,E0,1EA3,E3,1EA1,E2,103,1EA5,1EA7,1EA9,1EAB,1EAD,1EAF,1EB1,1EB3,1EB5,1EB7,C1,C0,1EA2,C3,1EA0,C2,102,1EA4,1EA6,1EA8,1EAA,1EAC,1EAE,1EB0,1EB2,1EB4,1EB6,E9,E8,1EBB,1EBD,1EB9,EA,1EBF,1EC1,1EC3,1EC5,1EC7,C9,C8,1EBA,1EBC,1EB8,CA,1EBE,1EC0,1EC2,1EC4,1EC6,ED,EC,1EC9,129,1ECB,CD,CC,1EC8,128,1ECA,F3,F2,1ECF,F5,1ECD,F4,1A1,1ED1,1ED3,1ED5,1ED7,1ED9,1EDB,1EDD,1EDF,1EE1,1EE3,D3,D2,1ECE,D5,1ECC,D4,1A0,1ED0,1ED2,1ED4,1ED6,1ED8,1EDA,1EDC,1EDE,1EE0,1EE2,FA,F9,1EE7,169,1EE5,1B0,1EE9,1EEB,1EED,1EEF,1EF1,DA,D9,1EE6,168,1EE4,1AF,1EE8,1EEA,1EEC,1EEE,1EF0,FD,1EF3,1EF7,1EF9,1EF5,DD,1EF2,1EF6,1EF8,1EF4,111,110"
    arrUni = Split(UNI, ",")

    For I = 1 To Len(str)
        If IsNumeric(Mid(str, I + 1, 1)) = False Then
            sUni = sUni & Mid(str, I, 1)
        Else
            sTem = IIf(IsNumeric(Mid(str, I + 2, 1)), Mid(str, I, 3), Mid(str, I, 2))
            I = I + IIf(IsNumeric(Mid(str, I + 2, 1)), 2, 1)
            If InStr(ansi, sTem) > 0 Then sTem = ChrW("&h" & arrUni(InStr(ansi, sTem) \ 3))
            sUni = sUni & sTem
        End If
    Next
    TV = sUni
End Function

Public Function GetComboboxHandle(ByVal hwnd As Long) As Long
     Dim Cbo As COMBOBOXINFO
     Cbo.cbSize = Len(Cbo)
     GetComboBoxInfo hwnd, Cbo
     GetComboboxHandle = Cbo.hwndEdit
End Function
