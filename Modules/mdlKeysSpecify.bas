Attribute VB_Name = "mdlKeysSpecify"
Option Explicit



Public Function TypeOfVowel(C As String) As KIEU_NGUYEN_AM
    If InStr(1, "ueoaiy", C, vbTextCompare) > 0 Then
        TypeOfVowel = NGUYENAM_KHONGDAU
        Exit Function
    ElseIf InStr(1, ChrW$(226) & ChrW$(259) & ChrW$(234) & ChrW$(244) & ChrW$(417) & ChrW$(432), C, vbTextCompare) > 0 Then
        TypeOfVowel = NGUYENAM_DAU_TRANG
        Exit Function
    ElseIf InStr(1, ChrW$(224) & ChrW$(225) & ChrW$(7843) & ChrW$(227) & ChrW$(7841) & ChrW$(232) & ChrW$(233) & ChrW$(7867) & ChrW$(7869) & ChrW$(7865) & ChrW$(242) & ChrW$(243) & ChrW$(7887) & ChrW$(245) & ChrW$(7885) & ChrW$(236) & ChrW$(237) & ChrW$(7881) & ChrW$(297) & ChrW$(7883) & ChrW$(249) & ChrW$(250) & ChrW$(7911) & ChrW$(361) & ChrW$(7909) & ChrW$(7923) & ChrW$(253) & ChrW$(7927) & ChrW$(7929) & ChrW$(7925), C, vbTextCompare) > 0 Then
        TypeOfVowel = NGUYENAM_DAU_THANH
        Exit Function
    ElseIf InStr(1, ChrW$(7847) & ChrW$(7845) & ChrW$(7849) & ChrW$(7851) & ChrW$(7853) & ChrW$(7857) & ChrW$(7855) & ChrW$(7859) & ChrW$(7861) & ChrW$(7863) & ChrW$(7873) & ChrW$(7871) & ChrW$(7875) & ChrW$(7877) & ChrW$(7879) & ChrW$(7891) & ChrW$(7889) & ChrW$(7893) & ChrW$(7895) & ChrW$(7897) & ChrW$(7901) & ChrW$(7899) & ChrW$(7903) & ChrW$(7905) & ChrW$(7907) & ChrW$(7915) & ChrW$(7913) & ChrW$(7917) & ChrW$(7919) & ChrW$(7921), C, vbTextCompare) > 0 Then
        TypeOfVowel = NGUYENAM_DAUTRANG_DAU_THANH
        Exit Function
    Else
        TypeOfVowel = 0
    End If
End Function


Public Function StringIncludeVowel(s As String) As Boolean
    StringIncludeVowel = False
    Dim I As Long
    For I = 1 To Len(s)
        If TypeOfVowel(Mid$(s, I, 1)) = NGUYENAM_DAUTRANG_DAU_THANH Or TypeOfVowel(Mid$(s, I, 1)) = NGUYENAM_DAU_THANH Or TypeOfVowel(Mid$(s, I, 1)) = NGUYENAM_DAU_TRANG Or TypeOfVowel(Mid$(s, I, 1)) = NGUYENAM_KHONGDAU Then
            StringIncludeVowel = True
            Exit Function
        End If
    Next I
    
End Function



Public Function TypeOfKey(C As String) As TYPE_OF_KEY
    
    If InStr(1, "fsrxjz", C, vbTextCompare) > 0 Then
        TypeOfKey = DAU_THANH
        Exit Function
    ElseIf InStr(1, "aeo", C, vbTextCompare) > 0 Then
        TypeOfKey = KYTU_DOI
        Exit Function
    ElseIf C = "D" Or C = "d" Then
        TypeOfKey = KYTU_D
    ElseIf C = "w" Or C = "W" Then
        TypeOfKey = KYTU_W
        Exit Function
    ElseIf Asc(C) = 8 Then
        TypeOfKey = KYTU_BACK
        Exit Function
    Else
        TypeOfKey = 0
    End If
End Function


Public Sub ProcessKey(C As String)
    If TypeOfKey(C) = DAU_THANH Then
        ProcessToneMark C
        Exit Sub
    ElseIf TypeOfKey(C) = KYTU_BACK Then
        ProcessBackChar 8
        Exit Sub
    ElseIf TypeOfKey(C) = KYTU_DOI Then
        ProcessDoubleChar C
        Exit Sub
    ElseIf TypeOfKey(C) = KYTU_W Then
        Process_W_Char C
        Exit Sub
    ElseIf TypeOfKey(C) = KYTU_D Then
        Process_D_Char C
        Exit Sub
    Else
        PutToBuffer C
    End If
    
End Sub

