VERSION 5.00
Object = "{09F8995D-E1C7-449B-B63C-D210B6410F4F}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   5205
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox EN 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   180
      Picture         =   "frmMain.frx":0ECA
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   15
      Top             =   1410
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox VI 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   180
      Picture         =   "frmMain.frx":15B4
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   14
      Top             =   2220
      Visible         =   0   'False
      Width           =   360
   End
   Begin UniControls.UniComboBox lstCode 
      Height          =   345
      Left            =   1170
      TabIndex        =   12
      Top             =   180
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   609
      Style           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ExtendedUI      =   0   'False
      DropDownWidth   =   0
   End
   Begin UniControls.UniTrayIcon Tray 
      Left            =   180
      Top             =   1770
      _ExtentX        =   741
      _ExtentY        =   741
      TooltipText     =   ""
      Icon            =   "frmMain.frx":1C9E
   End
   Begin UniControls.UniFrame UniFrame1 
      Height          =   1125
      Left            =   1170
      Top             =   1080
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   1984
      MaskColor       =   16711935
      FrameColor      =   -2147483629
      Style           =   0
      Caption         =   "&Phi1m chuye63n"
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin UniControls.UniOptionButton Op2 
         Height          =   240
         Left            =   420
         TabIndex        =   3
         Top             =   750
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   423
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Control + &Alt"
         ForeColor       =   0
      End
      Begin UniControls.UniOptionButton OP1 
         Height          =   240
         Left            =   420
         TabIndex        =   2
         Top             =   390
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   423
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
         AutoUnicode     =   0   'False
         Caption         =   "&Control + Shift"
         ForeColor       =   0
      End
   End
   Begin UniControls.UniCheckBox CK4 
      Height          =   210
      Left            =   2700
      TabIndex        =   9
      Top             =   3150
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   370
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Vie61t hoa d9a62u ca6u"
      ForeColor       =   0
   End
   Begin UniControls.UniCheckBox CK5 
      Height          =   210
      Left            =   105
      TabIndex        =   10
      Top             =   3450
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   370
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Su73 du5ng VBKey Toolbar"
      ForeColor       =   0
   End
   Begin UniControls.UniCheckBox CK3 
      Height          =   210
      Left            =   105
      TabIndex        =   11
      Top             =   3150
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   370
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Kho73i d9o65ng cu2ng Windows"
      ForeColor       =   0
   End
   Begin UniControls.UniCheckBox CK1 
      Height          =   210
      Left            =   2700
      TabIndex        =   0
      Top             =   2850
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   370
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "D9a85t da61u kie63u cu4 (o2a,o2e,u2y)"
      ForeColor       =   0
   End
   Begin UniControls.UniCheckBox CK2 
      Height          =   210
      Left            =   105
      TabIndex        =   1
      Top             =   2850
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   370
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Hie65n cu73a so63 lu1c kho73i d9o65ng"
      ForeColor       =   0
   End
   Begin UniControls.UniLabel lbl1 
      Height          =   255
      Left            =   105
      Top             =   650
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   450
      AutoSize        =   -1  'True
      BackStyle       =   0
      Caption         =   "&Kie63u go4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WordWrap        =   0   'False
   End
   Begin UniControls.UniLabel lbl2 
      Height          =   255
      Left            =   120
      Top             =   230
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   450
      AutoSize        =   -1  'True
      BackStyle       =   0
      Caption         =   "&Ba3ng ma4 :"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WordWrap        =   0   'False
   End
   Begin UniControls.UniButton cmdOption 
      Height          =   360
      Left            =   4000
      TabIndex        =   4
      Top             =   1860
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   635
      Style           =   1
      Caption         =   "Tu2y c&ho5n"
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UniControls.UniButton cmdHelp 
      Height          =   360
      Left            =   4000
      TabIndex        =   5
      Top             =   1440
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   635
      Style           =   1
      Caption         =   "Tro75 &giu1p"
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UniControls.UniButton cmdAbout 
      Height          =   360
      Left            =   4000
      TabIndex        =   6
      Top             =   1020
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   635
      Style           =   1
      Caption         =   "Tho6ng t&in"
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UniControls.UniButton cmdHide 
      Height          =   360
      Left            =   4000
      TabIndex        =   7
      Top             =   600
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   635
      Style           =   1
      Caption         =   "D9o1&ng"
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UniControls.UniButton cmdExit 
      Height          =   360
      Left            =   4000
      TabIndex        =   8
      Top             =   180
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   635
      Style           =   1
      Caption         =   "Ke61t &thu1c"
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UniControls.UniComboBox lstInput 
      Height          =   345
      Left            =   1170
      TabIndex        =   13
      Top             =   600
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   609
      Style           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ExtendedUI      =   0   'False
      DropDownWidth   =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long

Private Loaded As Boolean

Private Sub Ck1_Click()
    ToneMarkIsOldStyle = CInt(CK1.Value)
    frmMenu.mnuToneStype.Checked = CK1.Value
End Sub

Private Sub Ck2_Click()
    ShowOnStart = CK2.Value
    frmMenu.mnuStartup.Checked = CK2.Value
End Sub

Private Sub CK3_Click()
    If CK3.Value Then
        AutoStartApp = 1
        frmMenu.mnuAutoStart.Checked = True
    Else
        AutoStartApp = 0
        frmMenu.mnuAutoStart.Checked = False
    End If
End Sub

Private Sub CK4_Click()
    If CK4.Value Then
        UpperCaseFirstWord = 1
    Else
        UpperCaseFirstWord = 0
    End If
End Sub

Private Sub CK5_Click()
    If CK5.Value Then
        UsedToolbar = 1
        frmToolbar.Show
    Else
        UsedToolbar = 0
        Unload frmToolbar
    End If
End Sub

 

Private Sub cmdAbout_Click()
    frmAbout.Show 1, Me
End Sub


Private Sub cmdAbout_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdAbout_Click
End Sub

Private Sub cmdAbout_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetCursorIcon IDC_HAND
End Sub



Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdExit_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdExit_Click
End Sub

Private Sub cmdExit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetCursorIcon IDC_HAND
End Sub



Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetCursorIcon IDC_HAND
End Sub

Private Sub cmdHelp_Click()
    MsgBoxTV "Chu7o7ng tri2nh hie65n chu7a co1 ba3n hu7o71ng da64n hoa2n chi3nh, ba5n vui lo2ng xem hu7o71ng da64n su73 du5ng cu3a chu7o7ng tri2nh Unikey", vbInformation, APPVERSION
End Sub


Private Sub cmdHelp_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdHelp_Click
End Sub

Private Sub cmdHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetCursorIcon IDC_HAND
End Sub

Private Sub cmdHide_Click()
    Me.Hide
    FormVisible = False
End Sub



Private Sub cmdHide_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdHide_Click
End Sub

Private Sub cmdOption_Click()
    Dim I As Integer
    If Me.Height < 4300 Then
        If Me.Height < 4300 Then
            Do Until Me.Height >= 4300
                DoEvents
                If Screen.Height - Me.Top < Me.Height + 500 Then Me.Top = Me.Top - 10
                Me.Height = Me.Height + 10
            Loop
        Else
            Me.Height = 4300
        End If
    Else
        If Me.Height > 3200 Then
            Do Until Me.Height <= 3200
                DoEvents
                Me.Height = Me.Height - 10
            Loop
        Else
            Me.Height = 3200
        End If
    End If
End Sub


Private Sub cmdOption_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdOption_Click
End Sub

Private Sub cmdOption_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetCursorIcon IDC_HAND
End Sub


Private Sub Form_Initialize()
    Dim hW As Long
    hW = FindWindow(vbNullString, APPVERSION)
    If hW > 0 Then
        'SetWindowPos hW, 0, (Screen.Width \ 2 - Me.Width \ 2) \ Screen.TwipsPerPixelX, (Screen.Height \ 2 - Me.Height \ 2) \ Screen.TwipsPerPixelY, Me.Width \ Screen.TwipsPerPixelX, Me.Height \ Screen.TwipsPerPixelY, 0
        ShowWindow hW, vbNormalFocus
        SetForegroundWindow hW
        End
    End If

    If FileExisting("UniControls_v2.0.ocx") = False Then
        LoadResToFile "UniControls_v2.0.ocx", "XP_OCX", "UniControls_v2.0.ocx", True, True
    End If
End Sub


Private Sub lstCode_Change()
    Select Case lstCode.ListIndex
        Case 0: CodeTable = UNICODE_PRECOMPOSED_TABLE_ENUM
        Case 1: CodeTable = BKHCM1_TABLE_ENUM
        Case 2: CodeTable = BKHCM2_TABLE_ENUM
        Case 3: CodeTable = TCVN3_TABLE_ENUM
        Case 4: CodeTable = UTF8_TABLE_ENUM
        Case 5: CodeTable = VIETWARE_F_TABLE_ENUM
        Case 6: CodeTable = VIQR_TABLE_ENUM
        Case 7: CodeTable = VISCII_TABLE_ENUM
        Case 8: CodeTable = VNCP_1258_TABLE_ENUM
        Case 9: CodeTable = VNI_WINDOWS_TABLE_ENUM
        Case 10: CodeTable = VPS_TABLE_ENUM
        Case 11: CodeTable = UNICODE_COMPOSED_TABLE_ENUM
    End Select
    SetCodeTable CodeTable

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then frmAbout.Show 1, Me
End Sub

Private Sub Form_Load()
    'Check_CodeTable
    Me.Caption = APPVERSION
    InitMenuTV Me.hWnd
    UniSystemMenu Me.hWnd
    SetUniText Me.hWnd, Me.Caption
    UniFrame1.Caption = TV("Phi1m chuye63n")
    AppPath = IIf(Right$(App.Path, 1) = "\", App.Path, App.Path & "\")
    
    LoadAppSettings
    
    lstInput.AddItem "TELEX"
    lstInput.AddItem "VNI"
    lstInput.AddItem "VIQR"
    'lstInput.AddItem TV("Ta3i le6n")
    'lstInput.AddItem TV("The6m...")
    
    With lstCode
        .AddItem TV("Unicode du75ng sa84n")
        .AddItem TV("BKHCM 1")
        .AddItem TV("BKHCM 2")
        .AddItem TV("TCVN 3")
        .AddItem TV("UTF 8")
        .AddItem TV("VIETWARE - F")
        .AddItem TV("VIQR")
        .AddItem TV("VISCII")
        .AddItem TV("VNCP 1258")
        .AddItem TV("VNI Windows")
        .AddItem TV("VPS")
        .AddItem TV("Unicode to63 ho75p")
    End With

    If switchMode = CONTROL_SHIFT Then
        OP1.Value = True
    Else
        Op2.Value = True
    End If
    
    CK1.Value = ToneMarkIsOldStyle
    CK2.Value = ShowOnStart
    CK3.Value = AutoStartApp
    CK4.Value = UpperCaseFirstWord
    CK5.Value = UsedToolbar
     
    frmMenu.mnuToneStype.Checked = CK1.Value
    frmMenu.mnuStartup.Checked = CK2.Value
    frmMenu.mnuAutoStart.Checked = AutoStartApp
    frmMenu.mnuUppercase.Checked = UpperCaseFirstWord
    frmMenu.mnuToolbar.Checked = UsedToolbar
    
    SetSwitchKey switchMode
    SetTypeOfInput inputMethod
    SetCodeTable CodeTable
    
    If VietNameseKeyboard Then
        Set Tray.Icon = VI.Picture
        Set frmToolbar.cmdIcon.Picture = frmToolbar.PV.Picture
        frmMain.Tray.ToolTipText = APPVERSION & " Copyright © Nguye64n Kha81c Tuye62n     " & vbCrLf & "Click va2o d9a6y d9e63 chuye63n sang go4 Tie61ng Anh  "
    Else
        Set Tray.Icon = EN.Picture
        Set frmToolbar.cmdIcon.Picture = frmToolbar.PE.Picture
        frmMain.Tray.ToolTipText = APPVERSION & " Copyright © Nguye64n Kha81c Tuye62n     " & vbCrLf & "Click va2o d9a6y d9e63 chuye63n sang go4 Tie61ng Vie65t  "
    End If

    frmMenu.Hide
    Me.Show
    If UsedToolbar > 0 Then frmToolbar.Show
    
    If ShowOnStart > 0 And Loaded = False Then
        Me.Show
        FormVisible = True
    Else
        Me.Hide
        FormVisible = False
    End If
    
    Loaded = True
    
    InitMouseHook True
    InitKeyHook True
End Sub


Private Sub Check_CodeTable()
    MakeCodeTable
    MsgBox "BK HCM 1 :             " & UBound(BKHCM1_TABLE) & vbCrLf & _
            "BK HCM 2 :             " & UBound(BKHCM2_TABLE) & vbCrLf & _
            "TCVN 3 :               " & UBound(TCVN3_TABLE) & vbCrLf & _
            "UTF 8 :                " & UBound(UTF8_TABLE) & vbCrLf & _
            "VIETWARE-F :           " & UBound(VIETWARE_F_TABLE) & vbCrLf & _
            "VIQR :                 " & UBound(VIQR_TABLE) & vbCrLf & _
            "VISCII  :              " & UBound(VISCII_TABLE) & vbCrLf & _
            "VNCP 1258 :            " & UBound(VNCP_1258_TABLE) & vbCrLf & _
            "VNI WINDOWS:           " & UBound(VNI_WINDOWS_TABLE) & vbCrLf & _
            "VPS :                  " & UBound(VPS_TABLE) & vbCrLf & _
            "UNICODE_COMPOSED :     " & UBound(UNICODE_COMPOSED_TABLE) & vbCrLf & _
            "UNICODE_PRECOMPOSED :  " & UBound(UNICODE_PRECOMPOSED_TABLE)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Me.Hide
        FormVisible = False
        Cancel = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    InitKeyHook False
    InitMouseHook False
    ClearClipboar
    Tray.Remove
    Unload frmMenu
    Unload frmAbout
    Unload frmToolbar
    EndApp
End Sub

Private Sub lstInput_Change()
    Select Case lstInput.ListIndex
        Case 0:
            SetTypeOfInput 1
            frmMenu.mnuInput(1).Checked = True
            frmMenu.mnuInput(3).Checked = False
            frmMenu.mnuInput(5).Checked = False
        Case 1: SetTypeOfInput 3
            frmMenu.mnuInput(3).Checked = True
            frmMenu.mnuInput(1).Checked = False
            frmMenu.mnuInput(5).Checked = False
        Case 2: SetTypeOfInput 5
            frmMenu.mnuInput(5).Checked = True
            frmMenu.mnuInput(1).Checked = False
            frmMenu.mnuInput(3).Checked = False
        'Case 3:
            
        'Case 4:
            'frmDefineInput.Show 1, Me
            'If frmDefineInput.OK = False Then
                'Exit Sub
            'Else
                'Tai kieu go moi
            'End If
    End Select
End Sub


Private Sub op1_Click()
    SetSwitchKey CONTROL_SHIFT
End Sub

Private Sub op2_Click()
    SetSwitchKey CONTROL_ALT
End Sub

Private Sub Tray_TrayClick(Button As UniControls.stMouseEvent)
    If Button = stLeftButtonUp Then
        ChangeKeyMode
    End If
    
    If Button = stLeftButtonDoubleClick Then
        frmMain.Show
        FormVisible = True
    End If
    
    If Button = stRightButtonUp Then
        InitMenuTV frmMenu.hWnd
        frmMain.PopupMenu frmMenu.mnuMain
    End If
End Sub
