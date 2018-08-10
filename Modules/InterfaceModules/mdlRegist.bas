Attribute VB_Name = "mdlRegist"
Option Explicit

' -------------- KHAI BÁO CÁC HÀM  WIN API -------------

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long


' ----------- KHAI BÁO CÁC ENUM -----------------

Public Enum REGTYPE
    REG_NONE = 0
    REG_SZ = 1
    REG_EXPAND_SZ = 2
    REG_BINARY = 3
    REG_DWORD = 4
    REG_LINK = 6
    REG_MULTI_SZ = 7
    REG_QWORD = 11
End Enum

Public Enum REGKEY
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_DYN_DATA = &H80000006
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_PERF_ROOT = &H80000002
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_USERS = &H80000003
End Enum



' -------- THÂN CHU*O*NG TRÌNH -------


Private Function RegGetString(ByVal hKey As Long, ByVal strValueName As String) As String
    Dim Result As Long, strBuf As String, DataBuffer As Long
    Result = RegQueryValueEx(hKey, strValueName, 0, REGTYPE.REG_SZ, ByVal 0, DataBuffer)
    If Result = 0 Then
            strBuf = String(DataBuffer, Chr$(0))
            Result = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, DataBuffer)
            If Result = 0 Then
                RegGetString = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If
    End If
End Function


Public Function REG_GETVALUE(hKey As REGKEY, Optional subKey As String = "Software\Microsoft\Windows\CurrentVersion\Run", Optional ValueName As String = "") As String
    Dim Ret As Long
    ' Mo mot Key
    RegOpenKey hKey, subKey, Ret
    ' Doc noi dung Key
    REG_GETVALUE = RegGetString(Ret, ValueName)
    ' Dong Key
    RegCloseKey Ret
End Function

Public Function REG_SETVALUE(hKey As REGKEY, Optional subKey = "Software\Microsoft\Windows\CurrentVersion\Run", Optional ValueName, Optional strData)
    Dim Ret
    ' Tao Key moi
    RegCreateKey hKey, CStr(subKey), Ret
    ' Ghi chuoi vao Key
    REG_SETVALUE = RegSetValueEx(Ret, CStr(ValueName), 0, REG_SZ, ByVal CStr(strData), Len(CStr(strData)))
    ' Dong Key lai
    RegCloseKey Ret
End Function


Public Function REG_DELETE(hKey As REGKEY, Optional subKey As String = "Software\Microsoft\Windows\CurrentVersion\Run", Optional DelsubKey As Boolean = False, Optional ValueName As String = "")
    Dim Ret
    ' Tao Key moi
    RegCreateKey hKey, StrPtr(subKey), Ret
    ' Xoa gia tri cua Key
      If DelsubKey Then
          REG_DELETE = RegDeleteKey(hKey, subKey)      ' Xoa key
      Else
          REG_DELETE = RegDeleteValue(Ret, ValueName)
      End If
    ' Dong Key lai
    RegCloseKey Ret
End Function


