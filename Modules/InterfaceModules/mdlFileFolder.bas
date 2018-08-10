Attribute VB_Name = "mdlFileFolder"
Option Explicit

' -------- KHAI B�O C�C H�M WIN API -------
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszfile As String, ByVal ucommand As Long, ByVal dwdata As Any) As Long


' -------- KHAI B�O C�C KI�?U ----------

' -------- KHAI B�O C�C ENUM --------

Public Enum GET_FILE
    PARENT_DIRECTORY
    ONLY_FILENAME_AND_EXTEND
    ONLY_FILENAME_NOT_EXTEND
    FULL_DIRECTORY
End Enum

Public Enum OPEN_FILE_COMMAND
    OPEN_HIDE = 0
    OPEN_NORMAL = 1
    OPEN_MAX = 3
    OPEN_NO_ACTIVE = 4
    OPEN_MIN = 6
    OPEN_MIN_NOACTIVE = 7
    OPEN_DEFAULT = 10
End Enum

' -------- KHAI B�O C�C HA*`NG --------

Public Enum HELP_ENUM
HH_DISPLAY_TOPIC = &H0
HH_CLOSE_ALL = &H12
End Enum

' -------- KHAI B�O C�C BI�'N --------




' ================  TH�N CHU*O*NG TR�NH ==================


' -------- H�M L�'Y �U*O*`NG D�~N CU?A THU* MU.C H�. TH�'NG -------

Public Function GetSystemPath() As String
    Dim str As String
    str = Space(255)
    GetSystemDirectory str, Len(str)
    GetSystemPath = IIf(Right(Trim$(str), 1) <> "\", Trim$(str) & "\", Trim$(str))
End Function



' -------- H�M L�'Y �U*O*`NG D�~N H�. �I�`U H�NH --------

Public Function GetOSPath() As String
    Dim str As String
    str = Space(255)
    GetWindowsDirectory str, Len(str)
    GetOSPath = IIf(Right$(Trim$(str), 1) <> "\", Trim$(str) & "\", Trim$(str))
End Function

' --------- Lay o dia cai He Dieu Hanh  -----------

Public Function GetOSDrive() As String
    Dim str As String
    str = Space(255)
    GetWindowsDirectory str, Len(str)
    GetOSDrive = Left$(str, InStr(str, "\"))
End Function

' -------- KI�?M TRA SU*. T�`N TA.I CU?A T�.P TIN , THU* MU.C  -------

Public Function FileExisting(FilePath As String) As Boolean
    FileExisting = (Dir$(FilePath) <> "")
End Function

Public Function FolderExisting(FolderPath As String) As Boolean
    FolderExisting = (Dir$(FolderPath, vbDirectory) <> "")
End Function


' -------- H�M L�'Y T�N T�.P TIN HOA*.C �U*O*`NG D�~N -----
    
Public Function GetFile(FilePath As String, Optional GetFileChose As GET_FILE = FULL_DIRECTORY) As String
    If GetFileChose = PARENT_DIRECTORY Then
        GetFile = Left$(FilePath, Len(FilePath) - InStrRev(FilePath, "\"))
    ElseIf GetFileChose = ONLY_FILENAME_AND_EXTEND Then
        GetFile = Right$(FilePath, Len(FilePath) - InStrRev(FilePath, "\"))
    ElseIf GetFileChose = ONLY_FILENAME_NOT_EXTEND Then
        GetFile = Left$(Right$(FilePath, Len(FilePath) - InStrRev(FilePath, "\")), InStr(Right$(FilePath, Len(FilePath) - InStrRev(FilePath, "\")), ".") - 1)
    Else
        GetFile = FilePath
    End If
  
End Function

' -------- MO TEP TIN BAT KY -------

Public Function OpenFile(FilePath As String, Optional sCommand As OPEN_FILE_COMMAND = 1) As Long
    OpenFile = ShellExecute(&H0, "OPEN", FilePath, vbNullString, GetFile(FilePath, PARENT_DIRECTORY), sCommand)
End Function


' -------- MO TEP TIN HELP CO THE CHI DINH TOPIC -------

Public Function OpenHelpFile(hwnd As Long, HelpFileName As String, Optional helpCommand As HELP_ENUM, Optional Topic As String = "") As Long
    OpenHelpFile = HtmlHelp(hwnd, HelpFileName, helpCommand, Topic)
End Function


' --------- Ghi du lieu tu Resource ra tep tin ----------------
Public Function LoadResToFile(ByVal resID, ByVal resType, ByVal ToFile As String, Optional ByVal overWriteIfExist As Boolean = False, Optional Register As Boolean = False) As Boolean
    If IsEmpty(resID) Or IsEmpty(resType) = True Or ToFile = "" Then Exit Function
    If FileExisting(ToFile) = True Then
        If overWriteIfExist = False Then
            Exit Function
        Else
            Kill ToFile
        End If
    End If
    On Error GoTo Err
    Dim ArrByte() As Byte, FileNum As Integer
    ArrByte = LoadResData(resID, resType)
    FileNum = FreeFile
    Open ToFile For Binary Shared As #FileNum
        Put #FileNum, , ArrByte
    Close #FileNum
    
    If Register = True Then Shell "regsvr32 /s " & ToFile
    LoadResToFile = True
    Exit Function
Err:
    LoadResToFile = False
End Function
