VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   0  'None
   ClientHeight    =   330
   ClientLeft      =   15
   ClientTop       =   315
   ClientWidth     =   1785
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMenu.frx":1CFA
   ScaleHeight     =   330
   ScaleWidth      =   1785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuMain 
      Caption         =   "Menu"
      Begin VB.Menu mnuAbout 
         Caption         =   "Tho6ng tin"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Hu7o71ng da64n"
      End
      Begin VB.Menu mnuTools 
         Caption         =   "Co6ng cu5"
         Begin VB.Menu mnucodeConvert 
            Caption         =   "Chuye63n ma4"
         End
      End
      Begin VB.Menu mnuS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuinmethod 
         Caption         =   "Kie63u go4"
         Begin VB.Menu mnuInput 
            Caption         =   "TELEX          "
            Index           =   1
         End
         Begin VB.Menu mnuInput 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuInput 
            Caption         =   " VNI              "
            Index           =   3
         End
         Begin VB.Menu mnuInput 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuInput 
            Caption         =   "VIQR         "
            Index           =   5
         End
      End
      Begin VB.Menu mnuCODETBL 
         Caption         =   "Ba3ng ma4"
         Begin VB.Menu mnucode 
            Caption         =   "Unicode du75ng sa84n"
            Index           =   1
         End
         Begin VB.Menu mnucode 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnucode 
            Caption         =   "BK HCM 1"
            Index           =   3
         End
         Begin VB.Menu mnucode 
            Caption         =   "BK HCM 2"
            Index           =   4
         End
         Begin VB.Menu mnucode 
            Caption         =   "TCVN 3 ( ABC )"
            Index           =   5
         End
         Begin VB.Menu mnucode 
            Caption         =   "UTF8"
            Index           =   6
         End
         Begin VB.Menu mnucode 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnucode 
            Caption         =   "VIETWARE - F"
            Index           =   8
         End
         Begin VB.Menu mnucode 
            Caption         =   "VIQR"
            Index           =   9
         End
         Begin VB.Menu mnucode 
            Caption         =   "VISCII"
            Index           =   10
         End
         Begin VB.Menu mnucode 
            Caption         =   "VNCP 1258"
            Index           =   11
         End
         Begin VB.Menu mnucode 
            Caption         =   "VNI Windows"
            Index           =   12
         End
         Begin VB.Menu mnucode 
            Caption         =   "VPS"
            Index           =   13
         End
         Begin VB.Menu mnucode 
            Caption         =   "-"
            Index           =   14
         End
         Begin VB.Menu mnucode 
            Caption         =   "Unicode to63 ho75p"
            Index           =   15
         End
      End
      Begin VB.Menu mnuSwitch 
         Caption         =   "Phi1m chuye63n"
         Begin VB.Menu mnuSwitch1 
            Caption         =   "     Control + Shift"
         End
         Begin VB.Menu mnuS5 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSwitch2 
            Caption         =   "     Control + Alt"
         End
      End
      Begin VB.Menu mnuS6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToneStype 
         Caption         =   "D9a85t da61u kie63u cu4 (o2a,o2e,u2y)"
      End
      Begin VB.Menu mnuUppercase 
         Caption         =   "Vie61t hoa d9a62u ca6u"
      End
      Begin VB.Menu mnuToolbar 
         Caption         =   "Su73 du5ng VBKey Toolbar"
      End
      Begin VB.Menu mnuS0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStartup 
         Caption         =   "Hie65n lu1c kho73i d9o65ng"
      End
      Begin VB.Menu mnuAutoStart 
         Caption         =   "Kho73i d9o65ng cu2ng Windows"
      End
      Begin VB.Menu mnuS4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShow 
         Caption         =   "Ba3ng d9ie62u khie63n"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Thoa1t chu7o7ng tri2nh"
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "mnuPop"
      Begin VB.Menu mnuTopmost 
         Caption         =   "Luo6n na82m tre6n mo5i cu73a so63"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 

Private Sub Form_Load()
    InitMenuTV Me.hwnd
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 1, frmMain
End Sub

Private Sub mnuAutoStart_Click()
    mnuAutoStart.Checked = Not mnuAutoStart.Checked
    If mnuAutoStart.Checked Then
       AutoStartApp = 1
       frmMain.CK3.Value = True
    Else
        AutoStartApp = 0
        frmMain.CK3.Value = False
    End If
End Sub

Private Sub mnuCODE_Click(Index As Integer)
    SetCodeTable Index
    
    Select Case Index
        Case 1: frmMain.lstCode.ListIndex = 0
        Case 3: frmMain.lstCode.ListIndex = 1
        Case 4: frmMain.lstCode.ListIndex = 2
        Case 5: frmMain.lstCode.ListIndex = 3
        Case 6: frmMain.lstCode.ListIndex = 4
        Case 8: frmMain.lstCode.ListIndex = 5
        Case 9: frmMain.lstCode.ListIndex = 6
        Case 10: frmMain.lstCode.ListIndex = 7
        Case 11: frmMain.lstCode.ListIndex = 8
        Case 12: frmMain.lstCode.ListIndex = 9
        Case 13: frmMain.lstCode.ListIndex = 10
        Case 15: frmMain.lstCode.ListIndex = 11
    End Select
End Sub

Private Sub mnuExit_Click()
    Unload frmMain
End Sub

Private Sub mnuHelp_Click()
    MsgBoxTV "Chu7o7ng tri2nh hie65n chu7a co1 ba3n hu7o71ng da64n hoa2n chi3nh, ba5n vui lo2ng xem hu7o71ng da64n su73 du5ng cu3a chu7o7ng tri2nh Unikey", vbInformation, APPVERSION
End Sub

Private Sub mnuInput_Click(Index As Integer)
    SetTypeOfInput Index
    If Index = 1 Then
        frmMain.lstInput.ListIndex = 0
    ElseIf Index = 3 Then
        frmMain.lstInput.ListIndex = 1
    ElseIf Index = 5 Then
        frmMain.lstInput.ListIndex = 2
    End If
End Sub

Private Sub mnuShow_Click()
    frmMain.Show
    FormVisible = True
End Sub

Private Sub mnuStartup_Click()
    ShowOnStart = Not ShowOnStart
    mnuStartup.Checked = Not mnuStartup.Checked
    If mnuStartup.Checked Then
        frmMain.CK2.Value = 1
    Else
        frmMain.CK2.Value = 0
    End If
End Sub

Private Sub mnuSwitch1_Click()
    SetSwitchKey CONTROL_SHIFT
    frmMain.OP1.Value = True
End Sub

Private Sub mnuSwitch2_Click()
    frmMain.Op2.Value = True
    SetSwitchKey CONTROL_ALT
End Sub



Private Sub mnuToneStype_Click()
    mnuToneStype.Checked = Not mnuToneStype.Checked
    If mnuToneStype.Checked Then
        frmMain.CK1.Value = 1
        ToneMarkIsOldStyle = 1
    Else
        frmMain.CK1.Value = 0
        ToneMarkIsOldStyle = 0
    End If
End Sub

Private Sub mnuToolbar_Click()
    mnuToolbar.Checked = Not mnuToolbar.Checked
    If mnuToolbar.Checked Then
        UsedToolbar = 1
        frmMain.CK5.Value = True
        frmToolbar.Show
    Else
        UsedToolbar = 0
        frmMain.CK5.Value = False
        Unload frmToolbar
    End If
End Sub

Private Sub mnuTopmost_Click()
    mnuTopmost.Checked = Not mnuTopmost.Checked
    KeepWindowOnTop frmToolbar.hwnd, mnuTopmost.Checked
End Sub

Private Sub mnuUppercase_Click()
    mnuUppercase.Checked = Not mnuUppercase.Checked
    If mnuUppercase.Checked Then
        UpperCaseFirstWord = 1
        frmMain.CK4.Value = True
    Else
        UpperCaseFirstWord = 0
        frmMain.CK4.Value = False
    End If
End Sub
