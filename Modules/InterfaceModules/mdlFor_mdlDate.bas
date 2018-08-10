Attribute VB_Name = "mdlFor_mdlDate"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Function fGetUserLocaleInfo(ByVal lLocaleID As Long, ByVal lLCType As Long) As String
    Dim sReturn As String
    Dim lReturn As Long
    lReturn = GetLocaleInfo(lLocaleID, lLCType, sReturn, Len(sReturn))
    If lReturn Then
        sReturn = Space$(lReturn)
        If lReturn Then
            fGetUserLocaleInfo = Left$(sReturn, lReturn - 1)
        End If
    End If
End Function

Public Function theEnumDates() As Long
    theEnumDates = 1
End Function

