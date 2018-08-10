VERSION 5.00
Begin VB.Form frmToolbar 
   BackColor       =   &H0000C000&
   BorderStyle     =   0  'None
   ClientHeight    =   420
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   1200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   420
   ScaleWidth      =   1200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox lblMenu 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   810
      Picture         =   "frmToolbar.frx":0000
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   4
      Top             =   30
      Width           =   360
   End
   Begin VB.PictureBox cmdIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   30
      Picture         =   "frmToolbar.frx":06EA
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   3
      Top             =   30
      Width           =   360
   End
   Begin VB.PictureBox PV 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1140
      Picture         =   "frmToolbar.frx":0DD4
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   2
      Top             =   2010
      Width           =   360
   End
   Begin VB.PictureBox PE 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   60
      Picture         =   "frmToolbar.frx":14BE
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   1
      Top             =   1950
      Width           =   360
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   420
      Picture         =   "frmToolbar.frx":1BA8
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   0
      Top             =   30
      Width           =   360
   End
End
Attribute VB_Name = "frmToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private toolbarTop As Long
Private toolbarLeft As Long

Private Sub cmdIcon_Click()
    ChangeKeyMode
End Sub

Private Sub cmdIcon_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        SetCursorIcon IDC_HAND
    End If
End Sub


Private Sub cmdIcon_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetCursorIcon IDC_HAND
End Sub

Private Sub cmdIcon_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu frmMenu.mnuPop
    End If
End Sub

Private Sub Form_Load()
    KeepWindowOnTop Me.hwnd, True
    If ((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "toolbarLeft")) = "") Or Not IsNumeric((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "toolbarLeft"))) Then
        toolbarLeft = Screen.Width - frmToolbar.ScaleWidth
    Else
        If CLng(Left$((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "toolbarLeft")), 5)) < 0 Or CLng(Left$((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "toolbarLeft")), 5)) > Screen.Width - frmToolbar.ScaleWidth Then
            toolbarLeft = Screen.Width - frmToolbar.ScaleWidth
        Else
            toolbarLeft = CLng((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "toolbarLeft")))
        End If
    End If
    If ((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "toolbarTop")) = "") Or Not IsNumeric((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "toolbarTop"))) Then
        toolbarTop = 0
    Else
        If CLng(Left$(REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "toolbarTop"), 5)) < 0 Or CLng(Left$(REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "toolbarTop"), 5)) > Screen.Height - frmToolbar.ScaleHeight Then
            toolbarTop = 0
        Else
            toolbarTop = CLng((REG_GETVALUE(HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "toolbarTop")))
        End If
    End If
    
    Me.Move toolbarLeft, toolbarTop
    cmdIcon.Picture = IIf(VietNameseKeyboard, frmToolbar.PV.Picture, frmToolbar.PE.Picture)
    SetToolTipObj Picture1.hwnd, "Di chuye63n thanh co6ng cu5  ", , , , , , , "Arial", 600
    SetToolTipObj lblMenu.hwnd, "Menu chu7o7ng tri2nh  ", , , , , , , "Arial", 600
    SetToolTipObj cmdIcon.hwnd, "Chuye63n che61 d9o65 ba2n phi1m  ", , , , , , , "Arial", 600
End Sub

Private Sub Form_Unload(Cancel As Integer)
    toolbarLeft = Me.Left
    toolbarTop = Me.Top
    REG_SETVALUE HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "toolbarLeft", toolbarLeft
    REG_SETVALUE HKEY_LOCAL_MACHINE, "Software\VBKey\settings", "toolbarTop", toolbarTop
End Sub

Private Sub lblMenu_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        SetCursorIcon IDC_HAND
    End If
End Sub

Private Sub lblMenu_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetCursorIcon IDC_HAND
End Sub

Private Sub lblMenu_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    InitMenuTV frmMenu.hwnd
    PopupMenu frmMenu.mnuMain
End Sub


Private Sub Picture1_DblClick()
    frmMain.Show
    FormVisible = True
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        SetCursorIcon IDC_SIZEALL
        DragWindow Me.hwnd
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetCursorIcon IDC_SIZEALL
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu frmMenu.mnuPop
    End If
End Sub
