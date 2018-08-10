Attribute VB_Name = "mdlDate"
Option Explicit

DefInt A-Z

' ----------------- KHAI BÁO HÀM API ----------------

Private Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SYSTEMTIME, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal lLocale As Long, ByVal lLocaleType As Long, ByVal sLCData As String, ByVal lBufferLength As Long) As Long
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function EnumDateFormats Lib "kernel32" Alias "EnumDateFormatsA" (ByVal lpDateFmtEnumProc As Long, ByVal Locale As Long, ByVal dwFlags As Long) As Long
Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long


' ----------------- KHAI BÁO CÁC KIÊ?U ------------

Public Type DateType
        Day_ As String
        Month_ As String
        Year_ As String
End Type

Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

' -------------- KHAI BÁO CÁC ENUM -------------




' --------------- KHAI BÁO CÁC CONST ----------

Private Const LOCALE_SLONGDATE = &H20
Private Const LOCALE_STIMEFORMAT = &H1003
Private Const LOCALE_USER_DEFAULT As Long = &H400
Private Const LOCALE_SLANGUAGE As Long = &H2
Private Const LOCALE_SSHORTDATE As Long = &H1F
Private Const DATE_LONGDATE As Long = &H2
Private Const DATE_SHORTDATE As Long = &H1
Private Const HWND_BROADCAST As Long = &HFFFF&
Private Const WM_SETTINGCHANGE As Long = &H1A


' -------------- KHAI BÁO BIÊ'N ------------


' -------- THÂN CHU*O*NG TRÌNH -------

Public Function GetDateStringFormat() As String
    Dim BuffLen As Long, Result As Long
    Dim Buffer As String
    On Error Resume Next
    BuffLen = 128
    Buffer = String$(BuffLen, vbNullChar)
    Result = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE, Buffer, BuffLen)
    If Result > 0 Then
        GetDateStringFormat = Left$(Buffer, Result - 1)
    Else
        GetDateStringFormat = "dd/mm/yyyy"
    End If
    On Error GoTo 0
End Function

Public Function DateValidate(ByVal theDay As Integer, theMonth As Integer, theYear As Integer) As Boolean
    Dim ST As SYSTEMTIME
    With ST
      .wDay = theDay
      .wMonth = theMonth
      .wYear = theYear
    End With
Dim str As String
If theYear > Year(Date) Then
    DateValidate = False
    Exit Function
End If
If theMonth > Month(Date) And (theYear >= Year(Date)) Then
    DateValidate = False
    Exit Function
End If

If theDay > Day(Date) And (theMonth >= Month(Date)) And theYear >= Year(Date) Then
    DateValidate = False
    Exit Function
End If
DateValidate = GetDateFormat(0, 0, ST, vbNullString, vbNullString, 0&)
End Function

Public Function GetDateElement(ByVal sDate As String) As DateType
    Dim Arr() As String
    Arr = Split(sDate, "/")
    GetDateElement.Day_ = Arr(0)
    GetDateElement.Month_ = Arr(1)
    GetDateElement.Year_ = Arr(2)
End Function

Public Function ChangeDateFormat(sFormat As String) As Long
    Dim xCID  As Long
    xCID = GetSystemDefaultLCID
    If sFormat <> "" Then
        ChangeDateFormat = SetLocaleInfo(xCID, LOCALE_SSHORTDATE, sFormat)
        Call PostMessage(HWND_BROADCAST, WM_SETTINGCHANGE, 0&, ByVal 0&)
        Call EnumDateFormats(AddressOf theEnumDates, xCID, DATE_SHORTDATE)
    End If
End Function
