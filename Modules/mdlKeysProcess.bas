Attribute VB_Name = "mdlKeysProcess"
Option Explicit

Public Function IsUpper(C As String) As Boolean
    IsUpper = (C = UCase$(C))
End Function

Public Sub PutToBuffer(C As String)
    StringBuffer = StringBuffer & C
    ReDim Preserve UpperCase(KeysPressed)
    UpperCase(KeysPressed) = IsUpper(C)
    KeysPressed = KeysPressed + 1
End Sub


Public Sub ProcessBackChar(C As Long)
    If C <> 8 Then
        Exit Sub
    End If
    If KeysPressed <= 0 Then Exit Sub
    If KeysPressed > 1 Then
        KeysPressed = KeysPressed - 1
        StringBuffer = Left$(StringBuffer, KeysPressed)
    Else
        KeysPressed = 0
        StringBuffer = ""
    End If
End Sub



Public Sub ProcessDoubleChar(Ch As String)

    Dim Pos As Integer, Pos1 As Integer, Pos2 As Integer, Founds As Boolean
    Founds = False
    
    If ((StringIncludeVowel(StringBuffer) = False) Or (KeysPressed <= 0)) Then
        PutToBuffer Ch
        Exit Sub
    End If
    
    Pos = KeysPressed
    
    
    Do While Pos > 0
        If InStr(1, STRING_RESET, Mid$(StringBuffer, Pos, 1), vbTextCompare) > 0 Then
            Founds = True
            Exit Do
        End If
        Pos = Pos - 1
    Loop

    If Founds = False Then Pos = 1
    
    Pos1 = Pos
    Founds = False
    Do While Pos1 <= KeysPressed
        If TypeOfVowel(Mid$(StringBuffer, Pos1, 1)) = NGUYENAM_DAUTRANG_DAU_THANH And UCase$(Left$(UniToTelex(Mid$(StringBuffer, Pos1, 1)), 1)) = UCase$(Ch) Then
            Founds = True
            Exit Do
        End If
        Pos1 = Pos1 + 1
    Loop
    
    If Not Founds Then
        Pos1 = Pos
        Founds = False
        Do While Pos1 <= KeysPressed
            If TypeOfVowel(Mid$(StringBuffer, Pos1, 1)) = NGUYENAM_DAU_THANH And UCase$(Left$(UniToTelex(Mid$(StringBuffer, Pos1, 1)), 1)) = UCase$(Ch) Then
                Founds = True
                Exit Do
            End If
            Pos1 = Pos1 + 1
        Loop
    End If
    
    If Not Founds Then
        Pos1 = Pos
        Founds = False
        Do While Pos1 <= KeysPressed
            If TypeOfVowel(Mid$(StringBuffer, Pos1, 1)) = NGUYENAM_DAU_TRANG And UCase$(Left$(UniToTelex(Mid$(StringBuffer, Pos1, 1)), 1)) = UCase$(Ch) Then
                Founds = True
                Exit Do
            End If
            Pos1 = Pos1 + 1
        Loop
    End If
    
    
    If Not Founds Then
        Pos1 = Pos
        Founds = False
        Do While Pos1 <= KeysPressed
            If TypeOfVowel(Mid$(StringBuffer, Pos1, 1)) = NGUYENAM_KHONGDAU And UCase$(Left$(UniToTelex(Mid$(StringBuffer, Pos1, 1)), 1)) = UCase$(Ch) Then
                Founds = True
                Exit Do
            End If
            Pos1 = Pos1 + 1
        Loop
    End If
    
    
    '--------- Da tim thay Vi tri dau tien can lay trong chuoi bo dem-----------
    '--------- Tim vi tri cuoi cung ----------------
    
    
    Founds = False
    Pos2 = KeysPressed
    
    Do While Pos2 >= Pos1
        If TypeOfVowel(Mid$(StringBuffer, Pos2, 1)) = NGUYENAM_DAUTRANG_DAU_THANH And UCase$(Left$(UniToTelex(Mid$(StringBuffer, Pos2, 1)), 1)) = UCase$(Ch) Then
            Founds = True
            Exit Do
        End If
        Pos2 = Pos2 - 1
    Loop
    
    
    If Not Founds Then
        Pos2 = KeysPressed
        Founds = False
        Do While Pos2 >= Pos1
            If TypeOfVowel(Mid$(StringBuffer, Pos2, 1)) = NGUYENAM_DAU_THANH And UCase$(Left$(UniToTelex(Mid$(StringBuffer, Pos2, 1)), 1)) = UCase$(Ch) Then
                Founds = True
                Exit Do
            End If
            Pos2 = Pos2 - 1
        Loop
    End If
    
    If Not Founds Then
        Pos2 = KeysPressed
        Founds = False
        Do While Pos2 >= Pos1
            If TypeOfVowel(Mid$(StringBuffer, Pos2, 1)) = NGUYENAM_DAU_TRANG And UCase$(Left$(UniToTelex(Mid$(StringBuffer, Pos2, 1)), 1)) = UCase$(Ch) Then
                Founds = True
                Exit Do
            End If
            Pos2 = Pos2 - 1
        Loop
    End If

    If Not Founds Then
        Pos2 = KeysPressed
        Founds = False
        Do While Pos2 >= Pos1
            If TypeOfVowel(Mid$(StringBuffer, Pos2, 1)) = NGUYENAM_KHONGDAU And UCase$(Left$(UniToTelex(Mid$(StringBuffer, Pos2, 1)), 1)) = UCase$(Ch) Then
                Founds = True
                Exit Do
            End If
            Pos2 = Pos2 - 1
        Loop
    End If

    
    If Not Founds Then
        PutToBuffer Ch
        Exit Sub
    End If
    
    '--------------- Tim duoc Vi tri cuoi cung roi --------------
    

    Pos = Pos2
    
    Founds = False
    Do While Pos >= Pos1
        If TypeOfVowel(Mid$(StringBuffer, Pos, 1)) = NGUYENAM_DAUTRANG_DAU_THANH Or InStr(1, STRING_RESET, Mid$(StringBuffer, Pos, 1), vbTextCompare) > 0 Then
            Founds = True
        End If
        Pos = Pos - 1
    Loop
    
    If Not Founds Then
        Pos = Pos2
        Founds = False
        Do While Pos >= Pos1
            If TypeOfVowel(Mid$(StringBuffer, Pos, 1)) = NGUYENAM_DAU_THANH Or InStr(1, STRING_RESET, Mid$(StringBuffer, Pos, 1), vbTextCompare) > 0 Then
                Founds = True
            End If
            Pos = Pos - 1
        Loop
    End If
    
    If Not Founds Then
        Pos = Pos2
        Founds = False
        Do While Pos >= Pos1
            If TypeOfVowel(Mid$(StringBuffer, Pos, 1)) = NGUYENAM_DAU_TRANG Then
                Founds = True
            End If
            Pos = Pos - 1
        Loop
    End If
    
    If Not Founds Then
        Pos = Pos2
        Founds = False
        Do While Pos >= Pos1
            If TypeOfVowel(Mid$(StringBuffer, Pos, 1)) = NGUYENAM_KHONGDAU Then
                Founds = True
            End If
            Pos = Pos - 1
        Loop
    End If
    
    If Pos < Pos1 Then Pos = Pos1
    
    AnsiBuf = Mid$(StringBuffer, Pos, KeysPressed - Pos + 1)
    UniBuf = AnsiBuf
    BackNumbers = Len(AnsiBuf)
    Dim sAnsi As String
    
    
    If TypeOfVowel(Mid$(UniBuf, 1, 1)) = NGUYENAM_DAUTRANG_DAU_THANH Then
        sAnsi = UniToTelex(Mid$(UniBuf, 1, 1))
        
        
        If UCase$(Mid$(sAnsi, 2, 1)) = UCase$(Ch) Then
            If Pos > 1 Then
                If ((Ch = "o" Or Ch = "O") And (Mid$(StringBuffer, Pos - 1, 1) = "u" Or Mid$(StringBuffer, Pos - 1, 1) = "U")) Then
                    Mid$(UniBuf, 1, 1) = Left$(sAnsi, 1)
                End If
            Else
                Mid$(UniBuf, 1, 1) = TelexToUni(Left$(sAnsi, 1) & Right$(sAnsi, 1))
            End If
            
            Mid$(StringBuffer, Pos, KeysPressed - Pos + 1) = UniBuf
            PutToBuffer Ch
            UniBuf = Mid$(StringBuffer, Pos, KeysPressed - Pos + 1)
            Exit Sub
        Else
            If Pos > 1 Then
                If ((Ch = "o" Or Ch = "O") And (Mid$(StringBuffer, Pos - 1, 1) = ChrW$(432) Or Mid$(StringBuffer, Pos - 1, 1) = ChrW$(431))) Then
                    Mid$(StringBuffer, Pos - 1, 1) = Left$(UniToTelex(Mid$(StringBuffer, Pos - 1, 1)), 1)
                End If
            End If
    
            Mid$(UniBuf, 1, 1) = TelexToUni(Left$(sAnsi, 1) & Ch & Right$(sAnsi, 1))
            Mid$(StringBuffer, Pos, KeysPressed - Pos + 1) = UniBuf
            Exit Sub
        End If
    ElseIf TypeOfVowel(Mid$(UniBuf, 1, 1)) = NGUYENAM_DAU_THANH Then
        sAnsi = UniToTelex(Mid$(UniBuf, 1, 1))
        Mid$(UniBuf, 1, 1) = TelexToUni(Left$(sAnsi, 1) & Ch & Right$(sAnsi, 1))
        Mid$(StringBuffer, Pos, KeysPressed - Pos + 1) = UniBuf
        Exit Sub
    ElseIf TypeOfVowel(Mid$(UniBuf, 1, 1)) = NGUYENAM_DAU_TRANG Then
        sAnsi = UniToTelex(Mid$(UniBuf, 1, 1))
        If UCase$(Right$(sAnsi, 1)) = UCase$(Ch) Then
            Mid$(UniBuf, 1, 1) = Left$(sAnsi, 1)
            Mid$(StringBuffer, Pos, KeysPressed - Pos + 1) = UniBuf
            PutToBuffer Ch
            UniBuf = Mid$(StringBuffer, Pos, KeysPressed - Pos + 1)
            Exit Sub
        Else
            If Pos > 1 Then
                If ((Ch = "o" Or Ch = "O") And (Mid$(StringBuffer, Pos - 1, 1) = ChrW$(432) Or Mid$(StringBuffer, Pos - 1, 1) = ChrW$(431))) Then
                    Mid$(StringBuffer, Pos - 1, 1) = Left$(UniToTelex(Mid$(StringBuffer, Pos - 1, 1)), 1)
                End If
            End If
        
            Mid$(UniBuf, 1, 1) = TelexToUni(Left$(sAnsi, 1) & Ch)
            Mid$(StringBuffer, Pos, KeysPressed - Pos + 1) = UniBuf
            Exit Sub
        End If
    ElseIf TypeOfVowel(Mid$(UniBuf, 1, 1)) = NGUYENAM_KHONGDAU Then
        Mid$(UniBuf, 1, 1) = TelexToUni(Mid$(UniBuf, 1, 1) & Ch)
        Mid$(StringBuffer, Pos, KeysPressed - Pos + 1) = UniBuf
        Exit Sub
    End If
End Sub



Public Sub Process_D_Char(Ch As String)
    If Ch <> "D" And Ch <> "d" Then Exit Sub
    
    If ((InStr(1, StringBuffer, "d", vbTextCompare) <= 0) Or (KeysPressed <= 0)) Then
        PutToBuffer Ch
        Exit Sub
    End If
    
    Dim Pos As Integer, Pos1 As Integer, Pos2 As Integer, Founs As Boolean
    Founs = False
    Pos = KeysPressed
    
    Do While Pos > 0
        If (InStr(1, STRING_RESET, Mid$(StringBuffer, Pos, 1), vbTextCompare) > 0 Or Mid$(StringBuffer, Pos, 1) = ChrW$(273) Or Mid$(StringBuffer, Pos, 1) = ChrW$(272)) Then
            Founs = True
            Exit Do
        End If
        Pos = Pos - 1
    Loop
    
    If Founs = False Then Pos = 1
    
    '-------------- Di tim vi tri dau tien can lay cua chuoi bo dem -------------------
    
    
    Pos1 = Pos
    Founs = False
    
    Do While Pos1 <= KeysPressed
        If (Mid$(StringBuffer, Pos1, 1) = ChrW$(272) Or Mid$(StringBuffer, Pos1, 1) = ChrW$(273)) Then
            Founs = True
            Exit Do
        End If
        Pos1 = Pos1 + 1
    Loop
    
    If Not Founs Then
        Pos1 = Pos
        Founs = False
        Do While Pos1 <= KeysPressed
            If (Mid$(StringBuffer, Pos1, 1) = "d" Or Mid$(StringBuffer, Pos1, 1) = "D") Then
                Founs = True
                Exit Do
            End If
            Pos1 = Pos1 + 1
        Loop
    End If
    
    
    '----------- Tim thay Vi tri dau tien cua chuoi can xu ly ----------
    '----------- Tim vi ri cuoi cung ----------------
    
    Pos2 = KeysPressed
    Founs = False
    Do While Pos2 >= Pos1
        If (Mid$(StringBuffer, Pos2, 1) = ChrW$(272) Or Mid$(StringBuffer, Pos2, 1) = ChrW$(273)) Then
            Founs = True
            Exit Do
        End If
        Pos2 = Pos2 - 1
    Loop
    
    If Not Founs Then
        Pos2 = KeysPressed
        Founs = False
        Do While Pos2 >= Pos1
            If (Mid$(StringBuffer, Pos2, 1) = "D" Or Mid$(StringBuffer, Pos2, 1) = "d") Then
                Founs = True
                Exit Do
            End If
            Pos2 = Pos2 - 1
        Loop
    End If
    
    If (Mid$(StringBuffer, Pos2, 1) = ChrW$(272) Or Mid$(StringBuffer, Pos2, 1) = ChrW$(273)) Then
        Pos = Pos2
    ElseIf (Mid$(StringBuffer, Pos1, 1) = ChrW$(272) Or Mid$(StringBuffer, Pos1, 1) = ChrW$(273)) Then
            Pos = Pos1
    Else
        Pos = Pos2
    End If
    
    AnsiBuf = Mid$(StringBuffer, Pos, KeysPressed - Pos + 1)
    UniBuf = AnsiBuf
    BackNumbers = Len(AnsiBuf)
    
    Dim sAnsi As String
    If ((Mid$(UniBuf, 1, 1) = "D") Or (Mid$(UniBuf, 1, 1) = "d")) Then
        Mid$(UniBuf, 1, 1) = TelexToUni(Mid$(UniBuf, 1, 1) & Ch)
        Mid$(StringBuffer, Pos, KeysPressed - Pos + 1) = UniBuf
        Exit Sub
    Else
        sAnsi = UniToTelex(Mid$(UniBuf, 1, 1))
        Mid$(UniBuf, 1, 1) = Left$(sAnsi, 1)
        Mid$(StringBuffer, Pos, KeysPressed - Pos + 1) = UniBuf
        PutToBuffer Ch
        UniBuf = Mid$(StringBuffer, Pos, KeysPressed - Pos + 1)
        Exit Sub
    End If
    
End Sub




Public Sub Process_W_Char(Ch As String)

    Dim strCanWith_W As String

    strCanWith_W = "aou" & ChrW$(224) & ChrW$(225) & ChrW$(7843) & ChrW$(227) & ChrW$(7841) & ChrW$(259) & ChrW$(7857) & ChrW$(7855) & ChrW$(7859) & ChrW$(7861) & ChrW$(7863) & ChrW$(226) & ChrW$(7847) & ChrW$(7845) & ChrW$(7849) & ChrW$(7851) & ChrW$(7853) & ChrW$(242) & ChrW$(243) & ChrW$(7887) & ChrW$(245) & ChrW$(7885) & ChrW$(244) & ChrW$(7891) & ChrW$(7889) & ChrW$(7893) & ChrW$(7895) & ChrW$(7897) & ChrW$(417) & ChrW$(7901) & ChrW$(7899) & ChrW$(7903) & ChrW$(7905) & ChrW$(7907) & ChrW$(249) & ChrW$(250) & ChrW$(7911) & ChrW$(361) & ChrW$(7909) & ChrW$(432) & ChrW$(7915) & ChrW$(7913) & ChrW$(7917) & ChrW$(7919) & ChrW$(7921)
    
    If StringIncludeVowel(StringBuffer) = False Or KeysPressed <= 0 Then
        PutToBuffer IIf(Ch = UCase$(Ch), TelexToUni("UW"), TelexToUni("uw"))
        lastIsWConverted = True
        Exit Sub
    End If
    
    Dim Pos As Integer, Pos1 As Integer, Pos2 As Integer, Found As Boolean
    
    Pos = Len(StringBuffer)
    Found = False
    Do While Pos > 0
        If InStr(1, STRING_RESET, Mid$(StringBuffer, Pos, 1), vbTextCompare) > 0 Then
            Found = True
            Exit Do
        End If
        Pos = Pos - 1
    Loop
    
    If Not Found Then Pos = 1
    
    Pos1 = Pos
    Found = False
    Do While Pos1 < KeysPressed
        If (InStr(1, strCanWith_W, Mid$(StringBuffer, Pos1, 1), vbTextCompare) > 0) And TypeOfVowel(Mid$(StringBuffer, Pos1, 1)) = NGUYENAM_DAUTRANG_DAU_THANH Then
            Found = True
            Exit Do
        End If
        Pos1 = Pos1 + 1
    Loop
    
    If Not Found Then
        Pos1 = Pos
        Found = False
        Do While Pos1 < KeysPressed
            If (InStr(1, strCanWith_W, Mid$(StringBuffer, Pos1, 1), vbTextCompare) > 0) And TypeOfVowel(Mid$(StringBuffer, Pos1, 1)) = NGUYENAM_DAU_THANH Then
                Found = True
                Exit Do
            End If
            Pos1 = Pos1 + 1
        Loop
    End If
    
    If Not Found Then
        Pos1 = Pos
        Found = False
        Do While Pos1 < KeysPressed
            If (InStr(1, strCanWith_W, Mid$(StringBuffer, Pos1, 1), vbTextCompare) > 0) And TypeOfVowel(Mid$(StringBuffer, Pos1, 1)) = NGUYENAM_DAU_TRANG Then
                Found = True
                Exit Do
            End If
            Pos1 = Pos1 + 1
        Loop
    End If
    
    If Not Found Then
        Pos1 = Pos
        Found = False
        Do While Pos1 < KeysPressed
            If (InStr(1, strCanWith_W, Mid$(StringBuffer, Pos1, 1), vbTextCompare) > 0) And TypeOfVowel(Mid$(StringBuffer, Pos1, 1)) = NGUYENAM_KHONGDAU Then
                Found = True
                Exit Do
            End If
            Pos1 = Pos1 + 1
        Loop
    End If
    
    'MsgBox Pos1
    '--------- Tim duoc vi tri dau tien -----------
    '--------- Tim tiep vi tri cuoi cung ----------
    
    Pos2 = KeysPressed
    Found = False
    Do While Pos2 >= Pos1
        If (InStr(1, strCanWith_W, Mid$(StringBuffer, Pos2, 1), vbTextCompare) > 0) And TypeOfVowel(Mid$(StringBuffer, Pos2, 1)) = NGUYENAM_DAUTRANG_DAU_THANH Then
            Found = True
            Exit Do
        End If
        Pos2 = Pos2 - 1
    Loop
    
    If Not Found Then
        Pos2 = KeysPressed
        Found = False
        Do While Pos2 >= Pos1
            If (InStr(1, strCanWith_W, Mid$(StringBuffer, Pos2, 1), vbTextCompare) > 0) And TypeOfVowel(Mid$(StringBuffer, Pos2, 1)) = NGUYENAM_DAU_THANH Then
                Found = True
                Exit Do
            End If
            Pos2 = Pos2 - 1
        Loop
    End If
    
    If Not Found Then
        Pos2 = KeysPressed
        Found = False
        Do While Pos2 >= Pos1
            If (InStr(1, strCanWith_W, Mid$(StringBuffer, Pos2, 1), vbTextCompare) > 0) And TypeOfVowel(Mid$(StringBuffer, Pos2, 1)) = NGUYENAM_DAU_TRANG Then
                Found = True
                Exit Do
            End If
            Pos2 = Pos2 - 1
        Loop
    End If
    
    If Not Found Then
        Pos2 = KeysPressed
        Found = False
        Do While Pos2 >= Pos1
            If (InStr(1, strCanWith_W, Mid$(StringBuffer, Pos2, 1), vbTextCompare) > 0) And TypeOfVowel(Mid$(StringBuffer, Pos2, 1)) = NGUYENAM_KHONGDAU Then
                Found = True
                Exit Do
            End If
            Pos2 = Pos2 - 1
        Loop
    End If
    
    If Not Found Then
        PutToBuffer Ch
        Exit Sub
    End If
    
    '-------- Tim duoc Vi tri cuoi -----------
    '-------- Chon vi tri thuc su tu Vi tri dau va vi tri cuoi -------
    
    
    Pos = Pos2
    Found = False
    Do While Pos >= Pos1
        If TypeOfVowel(Mid$(StringBuffer, Pos, 1)) = NGUYENAM_DAUTRANG_DAU_THANH Then
            Found = True
            Exit Do
        End If
        Pos = Pos - 1
    Loop
    
    If Not Found Then
        Pos = Pos2
        Found = False
        Do While Pos >= Pos1
            If TypeOfVowel(Mid$(StringBuffer, Pos, 1)) = NGUYENAM_DAU_THANH Then
                Found = True
                Exit Do
            End If
            Pos = Pos - 1
        Loop
    End If
    
    If Not Found Then
        Pos = Pos2
        Found = False
        Do While Pos >= Pos1
            If TypeOfVowel(Mid$(StringBuffer, Pos, 1)) = NGUYENAM_DAU_TRANG Then
                Found = True
                Exit Do
            End If
            Pos = Pos - 1
        Loop
    End If
    
    If Not Found Then
        Pos = Pos2
        Found = False
        Do While Pos >= Pos1
            If TypeOfVowel(Mid$(StringBuffer, Pos, 1)) = NGUYENAM_KHONGDAU Then
                Found = True
                Exit Do
            End If
            Pos = Pos - 1
        Loop
    End If
    
    
    
    AnsiBuf = Mid$(StringBuffer, Pos, KeysPressed - Pos + 1)
    UniBuf = AnsiBuf
    BackNumbers = Len(AnsiBuf)
    'MsgBox AnsiBuf
    
    Dim sAnsi As String
    If TypeOfVowel(Mid$(UniBuf, 1, 1)) = NGUYENAM_DAUTRANG_DAU_THANH Then
        sAnsi = UniToTelex(Mid$(UniBuf, 1, 1))
        If (Mid$(sAnsi, 2, 1) = "W" Or Mid$(sAnsi, 2, 1) = "w") Then
            If Pos > 1 Then
                If ((Left$(sAnsi, 1) = "o" Or Left$(sAnsi, 1) = "O") And (Mid$(StringBuffer, Pos - 1, 1) = ChrW$(432) Or Mid$(StringBuffer, Pos - 1, 1) = ChrW$(431))) Then
                    Mid$(UniBuf, 1, 1) = Left$(UniToTelex(Mid$(StringBuffer, Pos - 1, 1)), 1)
                    Mid$(UniBuf, 2, 1) = Left$(sAnsi, 1)
                Else
                    Mid$(UniBuf, 1, 1) = TelexToUni(Left$(sAnsi, 1) & Right$(sAnsi, 1))
                End If
            Else
                Mid$(UniBuf, 1, 1) = TelexToUni(Left$(sAnsi, 1) & Right$(sAnsi, 1))
            End If
        
            Mid$(StringBuffer, Pos, KeysPressed - Pos + 1) = UniBuf
            PutToBuffer Ch
            UniBuf = Mid$(StringBuffer, Pos, KeysPressed - Pos + 1)
            Exit Sub
        Else
            Mid$(UniBuf, 1, 1) = TelexToUni(Left$(sAnsi, 1) & Ch & Right$(sAnsi, 1))
            Mid$(StringBuffer, Pos, KeysPressed - Pos + 1) = UniBuf
            Exit Sub
        End If
        
    ElseIf TypeOfVowel(Mid$(UniBuf, 1, 1)) = NGUYENAM_DAU_THANH Then
        sAnsi = UniToTelex(Mid$(UniBuf, 1, 1))
        Mid$(UniBuf, 1, 1) = TelexToUni(Left$(sAnsi, 1) & Ch & Right$(sAnsi, 1))
        Mid$(StringBuffer, Pos, KeysPressed - Pos + 1) = UniBuf
        Exit Sub
    ElseIf TypeOfVowel(Mid$(UniBuf, 1, 1)) = NGUYENAM_DAU_TRANG Then
        sAnsi = UniToTelex(Mid$(UniBuf, 1, 1))
        
        If Pos > 1 Then
            If (Mid$(UniBuf, 1, 1) = ChrW$(416) Or Mid$(UniBuf, 1, 1) = ChrW$(417)) And (Mid$(StringBuffer, Pos - 1, 1) = ChrW$(432) Or Mid$(StringBuffer, Pos - 1, 1) = ChrW$(431)) Then
                Mid$(StringBuffer, Pos - 1, 1) = Left$(UniToTelex(Mid$(StringBuffer, Pos - 1, 1)), 1)
            End If
        End If
        
        If (Right$(sAnsi, 1) = "w" Or Right$(sAnsi, 1) = "W") Then
            If (Left$(sAnsi, 1) = "u" Or Left$(sAnsi, 1) = "U") And lastIsWConverted Then
                Mid$(UniBuf, 1, 1) = IIf(Mid$(UniBuf, 1, 1) = UCase$(Mid$(UniBuf, 1, 1)), "W", "w")
                Mid$(StringBuffer, Pos, KeysPressed - Pos + 1) = UniBuf
                Exit Sub
            End If
        
            Mid$(UniBuf, 1, 1) = Left$(sAnsi, 1)
            Mid$(StringBuffer, Pos, KeysPressed - Pos + 1) = UniBuf
            PutToBuffer Ch
            UniBuf = Mid$(StringBuffer, Pos, KeysPressed - Pos + 1)
            Exit Sub
        Else
            Mid$(UniBuf, 1, 1) = TelexToUni(Left$(sAnsi, 1) & Ch)
            Mid$(StringBuffer, Pos, KeysPressed - Pos + 1) = UniBuf
            Exit Sub
        End If
    ElseIf TypeOfVowel(Mid$(UniBuf, 1, 1)) = NGUYENAM_KHONGDAU Then
        If Mid$(UniBuf, 1, 1) = "u" Or Mid$(UniBuf, 1, 1) = "U" Then lastIsWConverted = False
        If Pos > 1 Then
            If (Mid$(UniBuf, 1, 1) = "o" Or Mid$(UniBuf, 1, 1) = "O") And (Mid$(StringBuffer, Pos - 1, 1) = "U" Or Mid$(StringBuffer, Pos - 1, 1) = "u") Then
                Pos = Pos - 1
                AnsiBuf = Mid$(StringBuffer, Pos, KeysPressed - Pos + 1)
                UniBuf = AnsiBuf
                Mid$(UniBuf, 2, 1) = TelexToUni(Mid$(UniBuf, 2, 1) & Ch)
                BackNumbers = Len(AnsiBuf)
            End If
        End If
        Mid$(UniBuf, 1, 1) = TelexToUni(Mid$(UniBuf, 1, 1) & Ch)
        Mid$(StringBuffer, Pos, KeysPressed - Pos + 1) = UniBuf
        Exit Sub
    End If

End Sub




Public Sub ProcessToneMark(Ch As String)

    If (InStr(1, "fsrxjz", Ch, vbTextCompare) <= 0 Or KeysPressed <= 0) Or StringIncludeVowel(StringBuffer) = False Then
        PutToBuffer Ch
        Exit Sub
    End If
    
    If UCase$(Right$(StringBuffer, 1)) = UCase$(Ch) Then
        PutToBuffer Ch
        Exit Sub
    End If
    
    Dim Pos As Integer, Pos1 As Integer, Pos2 As Integer, Founds As Boolean, sVowel As String
    
    Pos = KeysPressed
    Founds = False
    Do While Pos > 0
        If InStr(1, STRING_RESET, Mid$(StringBuffer, Pos, 1), vbTextCompare) > 0 Then
            Founds = True
            Exit Do
        End If
        Pos = Pos - 1
    Loop
    
    If Not Founds Then Pos = 1
    
    
    '----------- Tim vi tri dau tien cua day nguyen am ----------
    
    
    Pos1 = Pos
    Founds = False
    Do While Pos1 <= KeysPressed
        If TypeOfVowel(Mid$(StringBuffer, Pos1, 1)) = NGUYENAM_DAUTRANG_DAU_THANH Then
            Founds = True
            Exit Do
        End If
        Pos1 = Pos1 + 1
    Loop
        
    If Not Founds Then
        Pos1 = Pos
        Founds = False
        Do While Pos1 <= KeysPressed
            If TypeOfVowel(Mid$(StringBuffer, Pos1, 1)) = NGUYENAM_DAU_THANH Then
                Founds = True
                Exit Do
            End If
            Pos1 = Pos1 + 1
        Loop
    End If
        
    If Not Founds Then
        Pos1 = Pos
        Founds = False
        Do While Pos1 <= KeysPressed
            If TypeOfVowel(Mid$(StringBuffer, Pos1, 1)) = NGUYENAM_DAU_TRANG Then
                Founds = True
                Exit Do
            End If
            Pos1 = Pos1 + 1
        Loop
    End If
        
    
    If Not Founds Then
        Pos1 = Pos
        Founds = False
        Do While Pos1 <= KeysPressed
            If TypeOfVowel(Mid$(StringBuffer, Pos1, 1)) = NGUYENAM_KHONGDAU Then
                Founds = True
                Exit Do
            End If
            Pos1 = Pos1 + 1
        Loop
    End If
    
    '--------- Tim thay vi tri dau tien cua chuoi nguyen am ---------
    '--------- Tim tiep vi tri cuoi cua day nguyen am --------
    
    
    Pos2 = KeysPressed
    Founds = False
    Do While Pos2 >= Pos1
        If TypeOfVowel(Mid$(StringBuffer, Pos2, 1)) = NGUYENAM_DAUTRANG_DAU_THANH Then
            Founds = True
            Exit Do
        End If
        Pos2 = Pos2 - 1
    Loop
        
    If Not Founds Then
        Pos2 = KeysPressed
        Founds = False
        Do While Pos2 >= Pos1
            If TypeOfVowel(Mid$(StringBuffer, Pos2, 1)) = NGUYENAM_DAU_THANH Then
                Founds = True
                Exit Do
            End If
            Pos2 = Pos2 - 1
        Loop
    End If
        
    If Not Founds Then
        Pos2 = KeysPressed
        Founds = False
        Do While Pos2 >= Pos1
            If TypeOfVowel(Mid$(StringBuffer, Pos2, 1)) = NGUYENAM_DAU_TRANG Then
                Founds = True
                Exit Do
            End If
            Pos2 = Pos2 - 1
        Loop
    End If
        
    If Not Founds Then
        Pos2 = KeysPressed
        Founds = False
        Do While Pos2 >= Pos1
            If TypeOfVowel(Mid$(StringBuffer, Pos2, 1)) = NGUYENAM_KHONGDAU Then
                Founds = True
                Exit Do
            End If
            Pos2 = Pos2 - 1
        Loop
    End If
    
    If Not Founds Then
        PutToBuffer Ch
        Exit Sub
    End If
        
    Pos1 = Pos
    Pos = Pos2
    
    Do While Pos1 <= Pos2
        If TypeOfVowel(Mid$(StringBuffer, Pos1, 1)) <> 0 Then
            Exit Do
        End If
        Pos1 = Pos1 + 1
    Loop
    
    sVowel = Mid$(StringBuffer, Pos1, Pos2 - Pos1 + 1)
    'MsgBox sVowel
    
    If Len(sVowel) > 1 Then
        If (Right$(sVowel, 2) = "ay" Or Right$(sVowel, 2) = "Ay" Or Right$(sVowel, 2) = "aY" Or Right$(sVowel, 2) = "AY" Or Right$(sVowel, 2) = "ai" Or Right$(sVowel, 2) = "AI" Or Right$(sVowel, 2) = "aI" Or Right$(sVowel, 2) = "Ai" Or Right$(sVowel, 2) = "ui" Or Right$(sVowel, 2) = "UI" Or Right$(sVowel, 2) = "Ui" Or Right$(sVowel, 2) = "uI" Or Right$(sVowel, 2) = "oi" Or Right$(sVowel, 2) = "OI" Or Right$(sVowel, 2) = "oI" Or _
        Right$(sVowel, 2) = "Oi" Or Right$(sVowel, 2) = "UA" Or Right$(sVowel, 2) = "ua" Or Right$(sVowel, 2) = "Ua" Or Right$(sVowel, 2) = "uA" Or Right$(sVowel, 2) = "UI" Or Right$(sVowel, 2) = "ui" Or Right$(sVowel, 2) = "Ui" Or Right$(sVowel, 2) = "uI" Or Right$(sVowel, 2) = "uy" Or Right$(sVowel, 2) = "UY" Or Right$(sVowel, 2) = "Uy" Or Right$(sVowel, 2) = "uY" Or Right$(sVowel, 2) = "ao" Or Right$(sVowel, 2) = "AO" Or Right$(sVowel, 2) = "Ao" Or Right$(sVowel, 2) = "aO") Then
            If Pos > 1 Then Pos = Pos - 1
        End If
    End If
    
    
    
    AnsiBuf = Mid$(StringBuffer, Pos, KeysPressed - Pos + 1)
    UniBuf = AnsiBuf
    BackNumbers = Len(AnsiBuf)
    Dim sAnsi As String
    
    If TypeOfVowel(Mid$(UniBuf, 1, 1)) = NGUYENAM_DAUTRANG_DAU_THANH Then
        sAnsi = UniToTelex(Mid$(UniBuf, 1, 1))
        If Ch = "z" Or Ch = "Z" Then
            Mid$(UniBuf, 1, 1) = TelexToUni(Left$(sAnsi, 2))
            Mid$(StringBuffer, Pos, KeysPressed - Pos + 1) = UniBuf
            UniBuf = Mid$(StringBuffer, Pos, KeysPressed - Pos + 1)
            Exit Sub
        End If
        If UCase$(Right$(sAnsi, 1)) = UCase$(Ch) Then
            Mid$(UniBuf, 1, 1) = TelexToUni(Left$(sAnsi, 2))
            Mid$(StringBuffer, Pos, KeysPressed - Pos + 1) = UniBuf
            PutToBuffer Ch
            UniBuf = Mid$(StringBuffer, Pos, KeysPressed - Pos + 1)
            Exit Sub
        Else
            Mid$(UniBuf, 1, 1) = TelexToUni(Left$(sAnsi, 2) & Ch)
            Mid$(StringBuffer, Pos, KeysPressed - Pos + 1) = UniBuf
            Exit Sub
        End If
    ElseIf TypeOfVowel(Mid$(UniBuf, 1, 1)) = NGUYENAM_DAU_THANH Then
        sAnsi = UniToTelex(Mid$(UniBuf, 1, 1))
        If Ch = "z" Or Ch = "Z" Then
            Mid$(UniBuf, 1, 1) = Left$(sAnsi, 1)
            Mid$(StringBuffer, Pos, KeysPressed - Pos + 1) = UniBuf
            UniBuf = Mid$(StringBuffer, Pos, KeysPressed - Pos + 1)
            Exit Sub
        End If
        If UCase$(Right$(Mid$(UniBuf, 1, 1), 1)) = UCase$(Ch) Then
            Mid$(UniBuf, 1, 1) = Left$(sAnsi, 1)
            Mid$(StringBuffer, Pos, KeysPressed - Pos + 1) = UniBuf
            PutToBuffer Ch
            UniBuf = Mid$(StringBuffer, Pos, KeysPressed - Pos + 1)
            Exit Sub
        Else
            Mid$(UniBuf, 1, 1) = TelexToUni(Left$(sAnsi, 1) & Ch)
            Mid$(StringBuffer, Pos, KeysPressed - Pos + 1) = UniBuf
            Exit Sub
        End If
    ElseIf TypeOfVowel(Mid$(UniBuf, 1, 1)) = NGUYENAM_DAU_TRANG Then
        sAnsi = UniToTelex(Mid$(UniBuf, 1, 1))
        If Ch = "z" Or Ch = "Z" Then
            Mid$(UniBuf, 1, 1) = Left$(sAnsi, 1)
            Mid$(StringBuffer, Pos, KeysPressed - Pos + 1) = UniBuf
            UniBuf = Mid$(StringBuffer, Pos, KeysPressed - Pos + 1)
            Exit Sub
        End If
        Mid$(UniBuf, 1, 1) = TelexToUni(sAnsi & Ch)
        Mid$(StringBuffer, Pos, KeysPressed - Pos + 1) = UniBuf
        Exit Sub
    ElseIf TypeOfVowel(Mid$(UniBuf, 1, 1)) = NGUYENAM_KHONGDAU Then
        If Ch = "z" Or Ch = "Z" Then
            PutToBuffer Ch
            Exit Sub
        End If
        Mid$(UniBuf, 1, 1) = TelexToUni(Mid$(UniBuf, 1, 1) & Ch)
        Mid$(StringBuffer, Pos, KeysPressed - Pos + 1) = UniBuf
        Exit Sub
    Else
        PutToBuffer Ch
    End If
    
End Sub
