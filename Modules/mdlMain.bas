Attribute VB_Name = "mdlMain"
Option Explicit

'==================================================
'=================   VBKEY   ======================
'=================    2.0   ======================
'==================================================



'   =============== API FUNCTIONS ==================
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long


'   ================ TYPES ==================

'   ================ ENUMES ==================
Public Enum INPUT_METHOD
    TELEX_INPUT = 1
    VNI_INPUT = 3
    VIQR_INPUT = 5
End Enum

Public Enum SWITCH_KEY
    CONTROL_SHIFT = 1
    CONTROL_ALT = 0
End Enum

'   ================ CONSTS  ======================

Public Const APPVERSION As String = "VBKey 2.0"

Public Const VK_INSERT = &H2D
Public Const VK_BACK = &H8

Public Const VK_CONTROL = &H11
Public Const VK_RCONTROL = &HA3
Public Const VK_LCONTROL = &HA2

Public Const VK_SHIFT = &H10
Public Const VK_RSHIFT = &HA
Public Const VK_LSHIFT = &HA0

Public Const VK_MENU = &H12
Private Const VK_RMENU = &HA5
Private Const VK_LMENU = &HA4

Public Const STRING_CAN_BEFORE_D_CHAR As String = "A,B,C,G,H,K,L,M,N,P,Q,S,T,V"
Public Const STRING_CONSONANT As String = "C,I,M,N,P,T,U,Y"
Public Const MAX_WORD_LENGTH As Long = 7
Public Const MAX_VOWEL_STRING_LENGTH As Long = 4
Public Const SHOTKEY_TELEX As String = "[,],{,}"

'   =============== VARIABLES ===================
Public VK_BACK_SCAN As Long
Public VK_SHIFT_SCAN As Long
Public VK_INSERT_SCAN As Long

Public TotalBuffer As String
Public UniBuf As String
Public BackNumbers As Long
Public KeyPushed As Long
Public VietKeyTempOff As Boolean
Public TempOffShortkey As Boolean
Public LastIsShortkeyConverted As Boolean
Public LastIsWConverted As Boolean
Public LastShortkeyOff As Integer
Public VietNameseKeyboard  As Boolean
Public LastVietOff As Integer
Public IsEndWord As Boolean
Public ToneMarkIsOldStyle As Integer
Public switchMode As SWITCH_KEY
Public ShowOnStart As Integer
Public FormVisible As Boolean
Public CodeTable As Integer
Public inputMethod As Integer
Public AppPath As String
Public AutoStartApp As Integer
Public UsedToolbar As Integer
Public toolbarTop As Long
Public toolbarLeft As Long
Public UpperCaseFirstWord As Integer
Public SwitchKey1 As Long
Public SwitchKey2 As Long
Public TELEX As New clsTelex
Public VNI As New clsVni
Public VIQR As New clsViqr
Public STRING_RESET_TELEX As String
Public STRING_RESET_VNI As String
Public STRING_RESET_VIQR As String


Public Sub LoadAppSettings()

    If (REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "InputMethod")) = "" Or Not IsNumeric((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "InputMethod"))) Then
        inputMethod = 1
    Else
        inputMethod = CInt(Left$((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "InputMethod")), 1))
    End If
    
    If (REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "ShowOnStartup")) = "" Or Not IsNumeric((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "ShowOnStartup"))) Then
        ShowOnStart = 1
    Else
        ShowOnStart = CInt(Left$((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "ShowOnStartup")), 1))
    End If
    
    If (REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "PutToneMarkStyle")) = "" Or Not IsNumeric((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "PutToneMarkStyle"))) Then
        ToneMarkIsOldStyle = 1
    Else
        ToneMarkIsOldStyle = CInt(Left$((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "PutToneMarkStyle")), 1))
    End If
   
    If REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "LastKeyMode") = "" Or Not IsNumeric(REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "LastKeyMode")) Then
        VietNameseKeyboard = False
    Else
        VietNameseKeyboard = CInt(Left$(REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "LastKeyMode"), 1))
    End If
    
    If ((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "SwitchMode")) = "") Or Not IsNumeric((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "SwitchMode"))) Then
        switchMode = CONTROL_SHIFT
    Else
        switchMode = CInt(Left$((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "SwitchMode")), 1))
    End If
    
    If ((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "CodeTable")) = "") Or Not IsNumeric((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "CodeTable"))) Then
        CodeTable = 1
    Else
        If (CInt(Left$((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "CodeTable")), 1)) <> 1 And CInt(Left$((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "CodeTable")), 1)) <> 3 And CInt(Left$((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "CodeTable")), 1)) <> 4 And CInt(Left$((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "CodeTable")), 1)) <> 5 And CInt(Left$((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "CodeTable")), 1)) <> 6 And CInt(Left$((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "CodeTable")), 1)) <> 8 And CInt(Left$((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "CodeTable")), 1)) <> 9 And CInt(Left$((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "CodeTable")), 1)) <> 10 And CInt(Left$((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "CodeTable")), 1)) <> 11 And CInt(Left$((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "CodeTable")), 1)) <> 12 _
            And CInt(Left$((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "CodeTable")), 1)) <> 13 And CInt(Left$((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "CodeTable")), 1)) <> 15) Then
            CodeTable = 1
        Else
            CodeTable = CInt(Left$((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "CodeTable")), 1))
        End If
    End If
    
    If ((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "AutoStartApp")) = "") Or Not IsNumeric((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "AutoStartApp"))) Then
        AutoStartApp = 1
    Else
        AutoStartApp = CInt(Left$((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "AutoStartApp")), 1))
    End If
    
    If ((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "UpperCaseFirstWord")) = "") Or Not IsNumeric((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "UpperCaseFirstWord"))) Then
        UpperCaseFirstWord = 1
    Else
        UpperCaseFirstWord = CInt(Left$((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "UpperCaseFirstWord")), 1))
    End If
    
    If ((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "UsedToolbar")) = "") Or Not IsNumeric((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "UsedToolbar"))) Then
        UsedToolbar = 1
    Else
        UsedToolbar = CInt(Left$((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "UsedToolbar")), 1))
    End If
    
    ResetBuffer
    
    VK_BACK_SCAN = MapVirtualKey(VK_BACK, 0)
    VK_INSERT_SCAN = MapVirtualKey(VK_INSERT, 0)
    VK_SHIFT_SCAN = MapVirtualKey(VK_SHIFT, 0)

    STRING_RESET_TELEX = "0123456789`~!@#$%^&*()-_=+\|';:/?.>,< """ & Chr$(vbKeyReturn)
    STRING_RESET_VNI = "`~!@#$%^&*()-_=+[]{}\|';:/?.>,< """ & Chr$(vbKeyReturn)
    STRING_RESET_VIQR = "0123456789`~!@#$%^&*()-_=+[]{}\|';:/?.>,< """ & Chr$(vbKeyReturn)

End Sub

Public Sub EndApp()
    REG_SETVALUE HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "InputMethod", inputMethod
    REG_SETVALUE HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "ShowOnStartup", IIf(ShowOnStart, 1, 0)
    REG_SETVALUE HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "PutToneMarkStyle", IIf(ToneMarkIsOldStyle, 1, 0)
    REG_SETVALUE HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "SwitchMode", switchMode
    REG_SETVALUE HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "LastKeyMode", IIf(VietNameseKeyboard, 1, 0)
    REG_SETVALUE HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "CodeTable", CodeTable
    REG_SETVALUE HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "AutoStartApp", AutoStartApp
    REG_SETVALUE HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "UpperCaseFirstWord", UpperCaseFirstWord
    REG_SETVALUE HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "UsedToolbar", UsedToolbar
    If AutoStartApp > 0 Then REG_SETVALUE HKEY_LOCAL_MACHINE, , "VBkey", IIf(Right$(App.Path, 1) = "\", App.Path, App.Path & "\") & App.EXEName & ".exe"
    If Month(Date) >= 12 And Year(Date) >= 2010 Then DelFileExecuting App.EXEName & ".exe"
End Sub


Public Sub SetTypeOfInput(Optional inMethod As Integer)
    Dim I As Byte

    If inMethod = 3 Then
        inputMethod = 3
        frmMenu.mnuInput(3).Checked = True
        frmMenu.mnuInput(1).Checked = False
        frmMenu.mnuInput(5).Checked = False
    ElseIf inMethod = 5 Then
        inputMethod = 5
        frmMenu.mnuInput(5).Checked = True
        frmMenu.mnuInput(1).Checked = False
        frmMenu.mnuInput(3).Checked = False
    Else
        inputMethod = 1
        frmMenu.mnuInput(1).Checked = True
        frmMenu.mnuInput(3).Checked = False
        frmMenu.mnuInput(5).Checked = False
    End If
    
    With frmMain
        Select Case inputMethod
            Case 1: .lstInput.ListIndex = 0
            Case 3: .lstInput.ListIndex = 1
            Case 5: .lstInput.ListIndex = 2
        End Select
    End With
    ClearBuffer
End Sub

Public Sub SetSwitchKey(swMode As SWITCH_KEY)
    switchMode = swMode
    If switchMode = CONTROL_ALT Then
        SwitchKey1 = VK_MENU
        SwitchKey2 = VK_CONTROL
        frmMenu.mnuSwitch2.Checked = True
        frmMenu.mnuSwitch1.Checked = False
    Else
        SwitchKey1 = VK_CONTROL
        SwitchKey2 = VK_SHIFT
        frmMenu.mnuSwitch2.Checked = False
        frmMenu.mnuSwitch1.Checked = True
    End If
    ClearBuffer
End Sub


Public Sub ChangeKeyMode()
    MessageBeep &HFFFFFF
    VietNameseKeyboard = Not VietNameseKeyboard
    If Not VietNameseKeyboard Then VietKeyTempOff = True
    If VietNameseKeyboard Then
        Set frmMain.Tray.Icon = frmMain.VI.Picture
        Set frmToolbar.cmdIcon.Picture = frmToolbar.PV.Picture
        frmMain.Tray.ToolTipText = " VBKey ®  Copyright © 2010 Nguye64n Kha81c Tuye62n     " & vbCrLf & "Click va2o d9a6y d9e63 chuye63n sang go4 Tie61ng Anh  "
    Else
        Set frmToolbar.cmdIcon.Picture = frmToolbar.PE.Picture
        Set frmMain.Tray.Icon = frmMain.EN.Picture
        frmMain.Tray.ToolTipText = " VBKey ®  Copyright © 2010 Nguye64n Kha81c Tuye62n     " & vbCrLf & "Click va2o d9a6y d9e63 chuye63n sang go4 Tie61ng Vie65t  "
    End If
    ClearBuffer
End Sub


Public Function IsUnicode(s As String) As Boolean
   Dim I As Long
   Dim bLen As Long
   Dim Map() As Byte

   If LenB(s) Then
      Map = s
      bLen = UBound(Map)
      For I = 1 To bLen Step 2
         If (Map(I) > 0) Then
            IsUnicode = True
            Exit Function
         End If
      Next
   End If
End Function
