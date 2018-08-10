Attribute VB_Name = "mdlKeyProcess"
Option Explicit


Private Enum TYPE_OF_KEYS
    BACK_CHAR = 8
    ESCAPE_KEY = 27
    
    DOUBLE_CHAR_TELEX = 1
    D_CHAR_TELEX_VIQR = 2
    BREVE_MARK_TELEX = 3
    TONE_MARK_TELEX = 4
    UNMARK_TELEX = 5
    SHORT_KEY_TELEX = 6
    
    DOUBLE_CHAR_VNI = 9
    D_CHAR_VNI = 10
    BREVE_MARK_VNI = 11
    TONE_MARK_VNI = 12
    UNMARK_VNI = 13
    SHORT_KEY_VNI = 14
    
    
    DOUBLE_CHAR_VIQR = 15
    BREVE_MARK_VIQR = 16
    TONE_MARK_VIQR = 17
    UNMARK_VIQR = 18
    SHORT_KEY_VIQR = 19
             
End Enum


Public Enum TYPE_OF_VOWELS
    NONE_MARK_VOWEL = 1
    BREVE_MARK_VOWEL = 2
    TONE_MARK_VOWEL = 3
    TONE_AND_BREVE_MARK_VOWEL = 4
End Enum


'   Specify type of keys to process

Private Function KeySpecify(kCode As Long) As TYPE_OF_KEYS
    
    If kCode = VK_BACK Then
        KeySpecify = BACK_CHAR
        Exit Function
    ElseIf kCode = 27 Then
        KeySpecify = ESCAPE_KEY
        Exit Function
        
'   -------- TELEX -----------

    ElseIf ((kCode = 65) Or (kCode = 97) Or (kCode = 69) Or (kCode = 101) Or (kCode = 111) Or (kCode = 79)) Then
        KeySpecify = DOUBLE_CHAR_TELEX
        Exit Function
    ElseIf ((kCode = 100) Or (kCode = 68)) Then
        KeySpecify = D_CHAR_TELEX_VIQR
        Exit Function
    ElseIf ((kCode = 119) Or (kCode = 87)) Then
        KeySpecify = BREVE_MARK_TELEX
        Exit Function
    ElseIf ((kCode = 70) Or (kCode = 83) Or (kCode = 82) Or (kCode = 88) Or (kCode = 74) Or (kCode = 102) Or (kCode = 115) Or (kCode = 114) Or (kCode = 120) Or (kCode = 106)) Then
        KeySpecify = TONE_MARK_TELEX
        Exit Function
    ElseIf InStr(1, SHOTKEY_TELEX, Chr$(kCode), vbTextCompare) > 0 Then
        KeySpecify = SHORT_KEY_TELEX
        Exit Function
    ElseIf ((kCode = 90) Or (kCode = 122)) Then
        KeySpecify = UNMARK_TELEX
        Exit Function
    
'   ---------- VNI ---------------

    ElseIf kCode = 54 Then
        KeySpecify = DOUBLE_CHAR_VNI
        Exit Function
    ElseIf kCode = 57 Then
        KeySpecify = D_CHAR_VNI
        Exit Function
    ElseIf ((kCode = 55) Or (kCode = 56)) Then
        KeySpecify = BREVE_MARK_VNI
        Exit Function
    ElseIf kCode <= 53 And kCode >= 49 Then
        KeySpecify = TONE_MARK_VNI
        Exit Function
    ElseIf kCode = 48 Then
        KeySpecify = UNMARK_VNI
        Exit Function
    
'   ----------- VIQR --------------
        
    ElseIf kCode = 94 Then
        KeySpecify = DOUBLE_CHAR_VIQR
        Exit Function
    ElseIf ((kCode = 40) Or (kCode = 43)) Then
        KeySpecify = BREVE_MARK_VIQR
        Exit Function
    ElseIf ((kCode = 96) Or (kCode = 39) Or (kCode = 63) Or (kCode = 126) Or (kCode = 46)) Then
        KeySpecify = TONE_MARK_VIQR
        Exit Function
    ElseIf kCode = 47 Then
        KeySpecify = UNMARK_VIQR
        Exit Function
    End If
    KeySpecify = 0
End Function




Public Function VowelSpecify(Ch As String) As TYPE_OF_VOWELS
    If InStr(1, "aeiouyAEIOUY", Ch, vbBinaryCompare) > 0 Then
        VowelSpecify = NONE_MARK_VOWEL
        Exit Function
        
    ElseIf InStr(1, ChrW$(226) & ChrW$(259) & ChrW$(234) & ChrW$(244) & ChrW$(417) & ChrW$(432) & ChrW$(194) & ChrW$(258) & ChrW$(202) & ChrW$(212) & ChrW$(416) & ChrW$(431), Ch, vbTextCompare) > 0 Then
        VowelSpecify = BREVE_MARK_VOWEL
        Exit Function
        
    ElseIf InStr(1, ChrW$(224) & ChrW$(225) & ChrW$(7843) & ChrW$(227) & _
                    ChrW$(7841) & ChrW$(232) & ChrW$(233) & ChrW$(7867) & _
                    ChrW$(7869) & ChrW$(7865) & ChrW$(236) & ChrW$(237) & _
                    ChrW$(7881) & ChrW$(297) & ChrW$(7883) & ChrW$(242) & _
                    ChrW$(243) & ChrW$(7887) & ChrW$(245) & ChrW$(7885) & _
                    ChrW$(249) & ChrW$(250) & ChrW$(7911) & ChrW$(361) & _
                    ChrW$(7909) & ChrW$(7923) & ChrW$(253) & ChrW$(7927) & _
                    ChrW$(7929) & ChrW$(7925) & ChrW$(192) & ChrW$(193) & _
                    ChrW$(7842) & ChrW$(195) & ChrW$(7840) & ChrW$(200) & _
                    ChrW$(201) & ChrW$(7866) & ChrW$(7868) & ChrW$(7864) & _
                    ChrW$(204) & ChrW$(205) & ChrW$(7880) & ChrW$(296) & _
                    ChrW$(7882) & ChrW$(210) & ChrW$(211) & ChrW$(7886) & _
                    ChrW$(213) & ChrW$(7884) & ChrW$(217) & ChrW$(218) & _
                    ChrW$(7910) & ChrW$(360) & ChrW$(7908) & ChrW$(7922) & _
                    ChrW$(221) & ChrW$(7926) & ChrW$(7928) & ChrW$(7924), Ch, vbTextCompare) > 0 Then
        VowelSpecify = TONE_MARK_VOWEL
        Exit Function
        
    ElseIf InStr(1, ChrW$(7847) & ChrW$(7845) & ChrW$(7849) & ChrW$(7851) & _
                    ChrW$(7853) & ChrW$(7857) & ChrW$(7855) & ChrW$(7859) & _
                    ChrW$(7861) & ChrW$(7863) & ChrW$(7873) & ChrW$(7871) & _
                    ChrW$(7875) & ChrW$(7877) & ChrW$(7879) & ChrW$(7891) & _
                    ChrW$(7889) & ChrW$(7893) & ChrW$(7895) & ChrW$(7897) & _
                    ChrW$(7901) & ChrW$(7899) & ChrW$(7903) & ChrW$(7905) & _
                    ChrW$(7907) & ChrW$(7915) & ChrW$(7913) & ChrW$(7917) & _
                    ChrW$(7919) & ChrW$(7921) & ChrW$(7846) & ChrW$(7844) & _
                    ChrW$(7848) & ChrW$(7850) & ChrW$(7852) & ChrW$(7856) & _
                    ChrW$(7854) & ChrW$(7858) & ChrW$(7860) & ChrW$(7862) & _
                    ChrW$(7872) & ChrW$(7870) & ChrW$(7874) & ChrW$(7876) & _
                    ChrW$(7878) & ChrW$(7890) & ChrW$(7888) & ChrW$(7892) & _
                    ChrW$(7894) & ChrW$(7896) & ChrW$(7900) & ChrW$(7898) & _
                    ChrW$(7902) & ChrW$(7904) & ChrW$(7906) & ChrW$(7914) & _
                    ChrW$(7912) & ChrW$(7916) & ChrW$(7918) & ChrW$(7920), Ch, vbTextCompare) > 0 Then
        VowelSpecify = TONE_AND_BREVE_MARK_VOWEL
        Exit Function
    End If
    
    VowelSpecify = 0
End Function


Public Function LenX(ByVal sUni As String) As Integer
    If IsDoubleCharSet(CodeTable) Then
        Dim sTemp As String
        sTemp = CodeTableConvert(sUni, 1, CodeTable)
        LenX = Len(sTemp)
    Else
        LenX = Len(sUni)
    End If
End Function

Public Sub PutToBuffer(Ch As String)

    KeyPushed = KeyPushed + 1
    
    If KeyPushed >= 3 Then
        If Right$(TotalBuffer, 2) = ". " And Mid$(TotalBuffer, KeyPushed - 2) <> "." Then
            IsEndWord = True
        Else
            IsEndWord = False
        End If
    ElseIf KeyPushed = 2 Then
        If Right$(TotalBuffer, 2) = ". " Then
            IsEndWord = True
        Else
            IsEndWord = False
        End If
    End If
    
    If IsEndWord Then
        UniBuf = Right$(TotalBuffer, 1)
        BackNumbers = LenX(UniBuf)
        TotalBuffer = TotalBuffer & UCase$(Ch)
        UniBuf = Right$(TotalBuffer, 2)
        Exit Sub
    End If
        
    TotalBuffer = TotalBuffer & Ch

    If VietNameseKeyboard Then
        If KeyPushed > 1 Then
            Dim S1 As String, S2 As String
            S1 = Mid$(TotalBuffer, KeyPushed - 1, 1)
            S2 = Mid$(TotalBuffer, KeyPushed, 1)
            If (UCase$(Left$(TELEX.UniToTelex(S1), 1)) = "O" And UCase$(Left$(TELEX.UniToTelex(S2), 1)) = "A") Or (UCase$(Left$(TELEX.UniToTelex(S1), 1)) = "O" And UCase$(Left$(TELEX.UniToTelex(S2), 1)) = "E") Or (UCase$(Left$(TELEX.UniToTelex(S1), 1)) = "U" And UCase$(Left$(TELEX.UniToTelex(S2), 1)) = "Y") Then
                If (VowelSpecify(S1) = TONE_MARK_VOWEL) And (VowelSpecify(S2) = NONE_MARK_VOWEL) Then
                    If ToneMarkIsOldStyle Then Exit Sub
                    If Not ToneMarkIsOldStyle Then
                        Mid$(TotalBuffer, KeyPushed - 1, 1) = Left$(TELEX.UniToTelex(S1), 1)
                        Mid$(TotalBuffer, KeyPushed, 1) = TELEX.TelexToUni(S2 & Right$(TELEX.UniToTelex(S1), 1))
                        UniBuf = Mid$(TotalBuffer, KeyPushed - 1, KeyPushed - (KeyPushed - 1) + 1)
                        'BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
                        BackNumbers = 1
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
    If VietNameseKeyboard Then
        If InStr(1, STRING_RESET_TELEX, Ch, vbTextCompare) > 0 And inputMethod = TELEX_INPUT Then VietKeyTempOff = False: TempOffShortkey = False
        If InStr(1, STRING_RESET_VIQR, Ch, vbTextCompare) > 0 And inputMethod = VIQR_INPUT Then VietKeyTempOff = False: TempOffShortkey = False
        If InStr(1, STRING_RESET_VNI, Ch, vbTextCompare) > 0 And inputMethod = VNI_INPUT Then VietKeyTempOff = False: TempOffShortkey = False
    End If
End Sub


Public Sub ProcessKey(Ch As Long)
    
    If KeySpecify(Ch) = BACK_CHAR Then
        ProcessBackChar Ch
    ElseIf KeySpecify(Ch) = ESCAPE_KEY Then
        Process_Escape_Key Ch

  ' ----------- TELEX -----------------
  
    ElseIf KeySpecify(Ch) = DOUBLE_CHAR_TELEX Then
        TELEX.ProcessDoubleChar Chr$(Ch)
    ElseIf KeySpecify(Ch) = D_CHAR_TELEX_VIQR Then    '  -D in TELEX METHOD AND VIQR METHOD
        TELEX.Process_D_Char Chr$(Ch)
    ElseIf KeySpecify(Ch) = BREVE_MARK_TELEX Then
        TELEX.ProcessBreveMark Chr$(Ch)
    ElseIf KeySpecify(Ch) = TONE_MARK_TELEX Then
        TELEX.ProcessToneMark Chr$(Ch)
    ElseIf KeySpecify(Ch) = UNMARK_TELEX Then
        ProcessUnMark Chr$(Ch)
    ElseIf KeySpecify(Ch) = SHORT_KEY_TELEX Then
        TELEX.ProcessShortkey Chr$(Ch)
        
  ' ------------- VNI ---------------
   
    ElseIf KeySpecify(Ch) = DOUBLE_CHAR_VNI Then
        VNI.ProcessDoubleChar Chr$(Ch)
    ElseIf KeySpecify(Ch) = D_CHAR_VNI Then
        VNI.Process_D_Char Chr$(Ch)
    ElseIf KeySpecify(Ch) = BREVE_MARK_VNI Then
        VNI.ProcessBreveMark Chr$(Ch)
    ElseIf KeySpecify(Ch) = TONE_MARK_VNI Then
        VNI.ProcessToneMark Chr$(Ch)
    ElseIf KeySpecify(Ch) = UNMARK_VNI Then
        ProcessUnMark Chr$(Ch)
    'ElseIf KeySpecify(Ch) = SHORT_KEY_VNI Then
        'ProcessShortKey_Vni Chr$(ch)
            
  ' -------------- VIQR --------------
    
    ElseIf KeySpecify(Ch) = DOUBLE_CHAR_VIQR Then
        VIQR.ProcessDoubleChar Chr$(Ch)
    ElseIf KeySpecify(Ch) = BREVE_MARK_VIQR Then
        VIQR.ProcessBreveMark Chr$(Ch)
    ElseIf KeySpecify(Ch) = TONE_MARK_VIQR Then
        VIQR.ProcessToneMark Chr$(Ch)
    ElseIf KeySpecify(Ch) = UNMARK_VIQR Then
        ProcessUnMark Chr$(Ch)
    ElseIf InStr(1, STRING_CONSONANT, Chr$(Ch), vbTextCompare) > 0 Then
        ProcessLastConSoNant Chr$(Ch)
    'ElseIf KeySpecify(Ch) = SHORT_KEY_VIQR Then
        'ProcessShortKey_Viqr Chr$(ch)
    Else
        PutToBuffer Chr$(Ch)
    End If

End Sub

Public Sub Process_Escape_Key(Ch As Long)
    'Chua su dung
End Sub

Private Function LenOfEndWord(s As String, iStart As Long) As Long
    
    LenOfEndWord = 0
    If iStart <= 0 Then Exit Function
    If s = "" Then Exit Function
    Dim I As Long
    For I = iStart To 1 Step -1
        If Mid$(s, I, 1) <> "." Then
            Exit Function
        Else
            LenOfEndWord = LenOfEndWord + 1
        End If
    Next I
End Function

Public Function GetLastWord(s As String) As Long
    Dim Counts As Byte, I As Long
    If s = "" Then GetLastWord = 0
    
    For I = Len(s) To 1 Step -1
        Counts = Counts + 1
        If ((Counts > MAX_WORD_LENGTH) Or (InStr(1, STRING_RESET_TELEX, Mid$(s, I, 1), vbTextCompare) > 0)) Then
            Exit For
        End If
    Next I
        
    GetLastWord = IIf(I > 0, I, 1)
End Function


Private Sub ProcessBackChar(Ch As Long)
    If Ch <> VK_BACK Then Exit Sub
    If KeyPushed > 0 Then
    
        KeyPushed = KeyPushed - 1
        
        If KeyPushed >= 3 Then
            If Right$(TotalBuffer, 2) = ". " And Mid$(TotalBuffer, KeyPushed - 2) <> "." Then
                IsEndWord = True
            Else
                IsEndWord = False
            End If
        ElseIf KeyPushed = 2 Then
            If Right$(TotalBuffer, 2) = ". " Then
                IsEndWord = True
            Else
                IsEndWord = False
            End If
        End If
        
        If KeyPushed < LastVietOff Then
            LastVietOff = 0
            VietKeyTempOff = False
        End If
        
        If KeyPushed < LastShortkeyOff Then
            TempOffShortkey = False
            LastShortkeyOff = 0
        End If
        
        TotalBuffer = Left$(TotalBuffer, KeyPushed)
        If VietNameseKeyboard Then
            If KeyPushed > 1 Then
                Dim S1 As String, S2 As String
                S1 = Mid$(TotalBuffer, KeyPushed - 1, 1)
                S2 = Mid$(TotalBuffer, KeyPushed, 1)
                If (UCase$(Left$(TELEX.UniToTelex(S1), 1)) = "O" And UCase$(Left$(TELEX.UniToTelex(S2), 1)) = "A") Or (UCase$(Left$(TELEX.UniToTelex(S1), 1)) = "O" And UCase$(Left$(TELEX.UniToTelex(S2), 1)) = "E") Then
                    If (VowelSpecify(S1) = NONE_MARK_VOWEL) And (VowelSpecify(S2) = TONE_MARK_VOWEL) Then
                        If ToneMarkIsOldStyle Then
                            UniBuf = Mid$(TotalBuffer, KeyPushed - 1, KeyPushed - (KeyPushed - 1) + 1)
                            BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
                            Mid$(TotalBuffer, KeyPushed, 1) = Left$(TELEX.UniToTelex(S2), 1)
                            Mid$(TotalBuffer, KeyPushed - 1, 1) = TELEX.TelexToUni(S1 & Right$(TELEX.UniToTelex(S2), 1))
                            UniBuf = Mid$(TotalBuffer, KeyPushed - 1, KeyPushed - (KeyPushed - 1) + 1)
                            BackNumbers = BackNumbers + 1
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    Else
        KeyPushed = 0
        TotalBuffer = ""
        UniBuf = ""
        ClearBuffer
    End If
End Sub


Private Sub ProcessUnMark(Ch As String)
    If ((VietKeyTempOff = True) Or (KeyPushed <= 0)) Then
        PutToBuffer Ch
        Exit Sub
    End If
    
    If inputMethod = TELEX_INPUT Then
        If Ch <> "z" And Ch <> "Z" Then
            PutToBuffer Ch
            Exit Sub
        End If
    End If
    
    If inputMethod = VNI_INPUT Then
        If Ch <> "0" Then
            PutToBuffer Ch
            Exit Sub
        End If
    End If
    
    If inputMethod = VIQR_INPUT Then
        If Ch <> "/" Then
            PutToBuffer Ch
            Exit Sub
        End If
    End If
    
    
    Dim FirstPos As Integer, LastPos As Integer, RealPos As Integer, FoundChar As Boolean
    
    FirstPos = GetLastWord(TotalBuffer)
    If LastVietOff > FirstPos Then FirstPos = LastVietOff
    LastPos = KeyPushed
    
    FoundChar = False
    Do While FirstPos <= LastPos
        If VowelSpecify(Mid$(TotalBuffer, FirstPos, 1)) <> 0 Then
            FoundChar = True
            Exit Do
        End If
        FirstPos = FirstPos + 1
    Loop
    
    If Not FoundChar Then
        PutToBuffer Ch
        Exit Sub
    End If
    
    FoundChar = False
    Do While LastPos >= FirstPos
        If VowelSpecify(Mid$(TotalBuffer, LastPos, 1)) = TONE_AND_BREVE_MARK_VOWEL Then
            FoundChar = True
            Exit Do
        End If
        LastPos = LastPos - 1
    Loop
        
    If Not FoundChar Then
        LastPos = KeyPushed
        FoundChar = False
        Do While LastPos >= FirstPos
            If VowelSpecify(Mid$(TotalBuffer, LastPos, 1)) = TONE_MARK_VOWEL Then
                FoundChar = True
                Exit Do
            End If
            LastPos = LastPos - 1
        Loop
    End If
    
    If Not FoundChar Then
        LastPos = KeyPushed
        FoundChar = False
        Do While LastPos >= FirstPos
            If VowelSpecify(Mid$(TotalBuffer, LastPos, 1)) = BREVE_MARK_VOWEL Then
                FoundChar = True
                Exit Do
            End If
            LastPos = LastPos - 1
        Loop
    End If
    
    If Not FoundChar Then
        LastPos = KeyPushed
        FoundChar = False
        Do While LastPos >= FirstPos
            If VowelSpecify(Mid$(TotalBuffer, LastPos, 1)) = TONE_AND_BREVE_MARK_VOWEL Then
                FoundChar = True
                Exit Do
            End If
            LastPos = LastPos - 1
        Loop
    End If
    
    If Not FoundChar Then
        PutToBuffer Ch
        Exit Sub
    End If
    
    If LastPos < FirstPos Then LastPos = FirstPos
    
    RealPos = LastPos
    UniBuf = Mid$(TotalBuffer, RealPos, KeyPushed - RealPos + 1)
    BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
    Dim sAnsi As String
    sAnsi = TELEX.UniToTelex(Mid$(TotalBuffer, RealPos, 1))
    If VowelSpecify(Mid$(TotalBuffer, RealPos, 1)) = TONE_AND_BREVE_MARK_VOWEL Then
        Mid$(TotalBuffer, RealPos, 1) = TELEX.TelexToUni(Left$(sAnsi, 2))
        UniBuf = Mid$(TotalBuffer, RealPos, KeyPushed - RealPos + 1)
        Exit Sub
    ElseIf VowelSpecify(Mid$(TotalBuffer, RealPos, 1)) = TONE_MARK_VOWEL Then
        Mid$(TotalBuffer, RealPos, 1) = Left$(sAnsi, 1)
        UniBuf = Mid$(TotalBuffer, RealPos, KeyPushed - RealPos + 1)
        Exit Sub
    ElseIf VowelSpecify(Mid$(TotalBuffer, RealPos, 1)) = BREVE_MARK_VOWEL Then
        Mid$(TotalBuffer, RealPos, 1) = Left$(sAnsi, 1)
        UniBuf = Mid$(TotalBuffer, RealPos, KeyPushed - RealPos + 1)
        Exit Sub
    End If
End Sub



Private Sub ProcessLastConSoNant(Ch As String)
    If (Not VietNameseKeyboard) Or (VietKeyTempOff) Or (KeyPushed <= 0) Then
        PutToBuffer Ch
        Exit Sub
    End If
    

    Dim Pos As Integer, FoundChar As Boolean
    
    FoundChar = False
    Pos = GetLastWord(TotalBuffer)
    If Pos < LastVietOff Then Pos = LastVietOff
    
    Do While (Pos <= KeyPushed)
        If VowelSpecify(Mid$(TotalBuffer, Pos, 1)) = TONE_MARK_VOWEL Then
            FoundChar = True
            Exit Do
        End If
        Pos = Pos + 1
    Loop
    
    If Not FoundChar Then
        PutToBuffer Ch
        Exit Sub
    End If
    If Pos > KeyPushed Then Pos = KeyPushed

    Dim I As Integer
    I = KeyPushed
    
    Do While I >= Pos
        If VowelSpecify(Mid$(TotalBuffer, I, 1)) <> 0 Then
            Exit Do
        End If
        I = I - 1
    Loop
    
    If I < Pos Then I = Pos
    
    Dim s As String
    s = Mid$(TotalBuffer, Pos, I - Pos + 1)

    UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
    BackNumbers = Len(UniBuf)
    PutToBuffer Ch
    UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
    
    Select Case Len(s)
        Case 2:
            Dim S1 As String
            S1 = TELEX.UniToTelex(Mid$(TotalBuffer, Pos, 1))
            
            If Pos < KeyPushed And UCase$(Left$(S1, 1)) = "O" Then
                If VowelSpecify(Mid$(TotalBuffer, Pos + 1, 1)) = NONE_MARK_VOWEL Then
                    Mid$(TotalBuffer, Pos + 1, 1) = TELEX.TelexToUni(Mid$(TotalBuffer, Pos + 1, 1) & Right$(S1, 1))
                    Mid$(TotalBuffer, Pos, 1) = Left$(S1, 1)
                    UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
                    Exit Sub
                End If
            End If
        Case 3:
            'Chua xu ly
    End Select
    
End Sub
