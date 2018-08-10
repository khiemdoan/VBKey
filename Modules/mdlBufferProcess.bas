Attribute VB_Name = "mdlBufferProcess"
Option Explicit

'   =========== API FUNCTIONS ============

Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32.dll" () As Long
Private Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "user32.dll" () As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function PostMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindowUnicode Lib "user32" (ByVal hWnd As Long) As Long


'   ============= CONST ================

Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40
Private Const WM_IME_CHAR = &H286

Public Enum CLIPBOARD_FORMAT
    CF_TEXT = 1
    CF_UNICODETEXT = 13
End Enum
Private m_ThreadActive As Long

Public Sub ClearClipboar()
    Dim Ret As Long
    If Not Ret Then Ret = OpenClipboard(GetFocus)
    If Not Ret Then Ret = OpenClipboard(0&)
    
    EmptyClipboard
    CloseClipboard
End Sub

Public Sub ClearBuffer()
    KeyPushed = 0
    TotalBuffer = ""
    UniBuf = ""
    VietKeyTempOff = False
    BackNumbers = 0
    LastIsWConverted = False
    LastIsShortkeyConverted = False
    LastVietOff = 0
    LastShortkeyOff = 0
End Sub

Public Sub ResetBuffer()
    ClearBuffer
    ClearClipboar
End Sub

Public Function GetClipboard(wF As CLIPBOARD_FORMAT) As String
On Error Resume Next

    Dim myStrPtr As Long, myLen As Long, myLock As Long, myData As String

    OpenClipboard 0&
    myStrPtr = GetClipboardData(wF)

    myLock = GlobalLock(myStrPtr)
    myLen = GlobalSize(myStrPtr)
    myData = String$(myLen \ 2 - 1, vbNullChar)
    lstrcpy StrPtr(myData), myLock
    GlobalUnlock myStrPtr

    CloseClipboard

    GetClipboard = myData
    
End Function


Public Sub SetClipboard(s As String)
    Dim sPtr As Long, iLen As Long, iLock As Long
    
    OpenClipboard 0&
    EmptyClipboard

    iLen = LenB(s) + 2
    sPtr = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, iLen)
    iLock = GlobalLock(sPtr)
    lstrcpy iLock, StrPtr(s)
    GlobalUnlock sPtr

    SetClipboardData CF_UNICODETEXT, sPtr
    CloseClipboard
    
End Sub

Private Function GetFocusHandle() As Long
Dim m_HandleFocus As Long
Dim m_HandleActive As Long

    m_HandleActive = GetForegroundWindow    'Lay Handle cua so Active
    m_ThreadActive = GetWindowThreadProcessId(m_HandleActive, ByVal 0) 'Lay ve ThreadID cua so Active
    AttachThreadInput GetCurrentThreadId, m_ThreadActive, 1
    m_HandleFocus = GetFocus
    GetFocusHandle = m_HandleFocus
    AttachThreadInput GetCurrentThreadId, m_ThreadActive, 0
End Function


Public Sub PushBuffer(ByVal sBuf As String)
    Dim s As String, mFocus As Long
    s = sBuf
    If CodeTable <> UNICODE_PRECOMPOSED_TABLE_ENUM Then s = CodeTableConvert(s, 1, CodeTable)
    mFocus = GetFocusHandle
    SetClipboard s
    SendPasteCommand
End Sub

Private Sub SendBuffer(ByVal m_HandleFocus As Long, ByVal s As String)
    If Len(s) <= 0 Then Exit Sub
    Dim I As Long
    For I = 1 To Len(s)
        PostMessageW m_HandleFocus, WM_IME_CHAR, AscW(Mid$(s, I, 1)), 1
    Next I
    Debug.Print "m_HandleFocus="; m_HandleFocus
End Sub

Public Function UpperCaseFirstWords(str As String) As String
    If str = "" Then Exit Function
    Dim s As String
    s = str
    Dim I As Integer
    For I = 1 To Len(s)
        If (Mid$(s, I, 1) = ".") And I < Len(s) Then
            Mid$(s, I + 1, 1) = UCase$(Mid$(s, I + 1, 1))
            I = I + 1
        End If
    Next I
    
    UpperCaseFirstWords = s
End Function
