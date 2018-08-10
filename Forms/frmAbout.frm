VERSION 5.00
Object = "{09F8995D-E1C7-449B-B63C-D210B6410F4F}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   4935
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin UniControls.UniButton cmdCLose 
      Height          =   375
      Left            =   1967
      TabIndex        =   1
      Top             =   2910
      Width           =   1000
      _ExtentX        =   1773
      _ExtentY        =   661
      Icon            =   "frmAbout.frx":0ECA
      Style           =   1
      Caption         =   "D9o1ng"
      IconAlign       =   3
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedBordersByTheme=   0   'False
   End
   Begin UniControls.UniLabel lbl2 
      Height          =   225
      Left            =   2070
      Top             =   480
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   397
      AutoSize        =   -1  'True
      BackStyle       =   0
      Caption         =   "Email : tuyen_dt18@yahoo.com"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WordWrap        =   0   'False
   End
   Begin UniControls.UniLabel lblName 
      Height          =   225
      Left            =   2010
      Top             =   150
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   397
      AutoSize        =   -1  'True
      BackStyle       =   0
      Caption         =   "Ta1c gia3 : Nguye64n Kha81c Tuye62n"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WordWrap        =   0   'False
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1305
      Left            =   300
      Picture         =   "frmAbout.frx":0EE6
      ScaleHeight     =   1305
      ScaleWidth      =   1740
      TabIndex        =   0
      Top             =   0
      Width           =   1740
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   60
      Top             =   3480
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   240
      Top             =   3390
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   1170
      Top             =   3390
   End
   Begin UniControls.UniTextBox txtLicense 
      Height          =   1365
      Left            =   300
      TabIndex        =   2
      Top             =   1410
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2408
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Text            =   ""
      MultiLine       =   -1  'True
      Locked          =   -1  'True
      BorderLine      =   11709605
      Scrollbar       =   2
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Trans As Byte

Private Sub cmdClose_Click()
    Timer1.Enabled = True
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetCursorIcon IDC_HAND
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdClose_Click
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        cmdClose_Click
        Cancel = 1
    End If
End Sub


Private Sub Lbl2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetCursorIcon IDC_HAND
End Sub

Private Sub Lbl2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetCursorIcon IDC_HAND
End Sub


Private Sub Timer1_Timer()
    DoEvents
    Static COunt As Long
    Static Bl As Boolean
    Static Offset As Long
    
    COunt = COunt + 1
    If Timer1.Interval >= 100 Then
        Timer1.Interval = 50
    Else
        Timer1.Interval = Timer1.Interval + 1
    End If
    If Offset <= 100 Then
        Offset = 50
    Else
        Offset = Offset - 10
    End If
    Bl = Not Bl
    If Bl Then
        Me.Left = Me.Left - (Offset + 20)
        Me.Top = Me.Top + Offset
    Else
        Me.Left = Me.Left + (Offset + 20)
        Me.Top = Me.Top - Offset
    End If
    
    If COunt >= 40 Then
        Timer1.Enabled = False
        COunt = 0
        Sleep 700
        Timer2.Enabled = True
    End If
End Sub

Private Sub Timer2_Timer()
    TransparentWindow Me.hwnd, Trans, True
    If Trans <= 5 Then
        Timer2.Enabled = False
        Me.BorderStyle = 0
        Trans = 0
        Unload Me
    Else
        Trans = Trans - 5
    End If
End Sub


Private Sub cmdClose_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then SetCursorIcon IDC_HAND
End Sub


Private Sub Form_Load()
    UniSystemMenu Me.hwnd
    Me.Caption = " Tho6ng tin chu7o7ng tri2nh " & APPVERSION
    SetUniText Me.hwnd, Me.Caption
    Trans = 255
    txtLicense.Text = TV("VBKey la2 chu7o7ng tri2nh go4 Tie61ng Vie65t hoa2n toa2n mie64n phi1." & vbCrLf & "Chu7o7ng tri2nh cha5y tre6n ne62n he65 d9ie63u ha2nh Windows NT " & vbCrLf & vbCrLf & "Trong qua1 tri2nh su73 du5ng, mong ca1c ba5n quan ta6m, d9o1ng go1p y1 kie61n ve62 chu7o7ng tri2nh cho chu1ng to6i, d9e63 chu7o7ng tri2nh nga2y ca2ng hoa2n thie65n ho7n." & vbCrLf & vbCrLf & "Ta1c gia3 : tuyen_dt18@yahoo.com")
End Sub

Private Sub Timer3_Timer()
    Dim P As POINTAPI
    P = CursorPosition(SCREEN_PIXEL)
    ScreenToClient hwnd, P
    If (P.x * Screen.TwipsPerPixelX >= lbl2.Left And P.x * Screen.TwipsPerPixelX <= lbl2.Left + lbl2.Width) And (P.y * Screen.TwipsPerPixelY >= lbl2.Top And P.y * Screen.TwipsPerPixelY <= lbl2.Top + lbl2.Height) Then
        lbl2.Font.Italic = True
        lbl2.ForeColor = vbBlue
    Else
        lbl2.Font.Italic = False
        lbl2.ForeColor = vbBlack
    End If
End Sub
