VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTelex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Option Explicit

Public Function TelexToUni(S As String) As String
    Dim I As Long, J As Long, sResult As String
    sResult = S
    For I = UBound(UNICODE_PRECOMPOSED_TABLE) To LBound(UNICODE_PRECOMPOSED_TABLE) Step -1
        For J = 1 To Len(S)
            If (LCase$(TelexArr(I)) = LCase$(Mid$(S, J, 3))) And (Mid$(S, J, 1) = Left$(TelexArr(I), 1)) Then sResult = Replace$(sResult, Mid$(S, J, 3), UNICODE_PRECOMPOSED_TABLE(I))
        Next J
        
        For J = 1 To Len(S)
            If (LCase$(TelexArr(I)) = LCase$(Mid$(S, J, 2))) And (Mid$(S, J, 1) = Left$(TelexArr(I), 1)) Then sResult = Replace$(sResult, Mid$(S, J, 2), UNICODE_PRECOMPOSED_TABLE(I))
        Next J
    Next I
    
    TelexToUni = sResult
    
End Function


Public Function UniToTelex(S As String) As String
    Dim I As Long, J As Long, sResult As String
    sResult = S
        
    For I = UBound(UNICODE_PRECOMPOSED_TABLE) To LBound(UNICODE_PRECOMPOSED_TABLE) Step -1
        For J = 1 To Len(S)
            If UNICODE_PRECOMPOSED_TABLE(I) = Mid$(S, J, 1) Then sResult = Replace$(sResult, Mid$(S, J, 1), TelexArr(I))
        Next J
    Next I
    
    UniToTelex = sResult
    
End Function


Public Sub Process_D_Char(Ch As String)
    If LastVietOff > 0 Then
        If UCase$(Ch) <> UCase$(Mid$(TotalBuffer, LastVietOff, 1)) Then VietKeyTempOff = False
    End If

    If ((inputMethod <> TELEX_INPUT And inputMethod <> VIQR_INPUT) Or (VietKeyTempOff = True) Or (KeyPushed <= 0) Or (Not VietNameseKeyboard)) Then
        PutToBuffer Ch
        Exit Sub
    End If
                                            
                                                
    Dim FirstPos As Integer, LastPos As Integer, RealPos As Integer
        
    FirstPos = GetLastWord(TotalBuffer)
    LastPos = KeyPushed
    
    Dim FoundChar As Boolean
    
    FoundChar = False
    Do While FirstPos <= LastPos
        If (Mid$(TotalBuffer, FirstPos, 1) = "d" Or Mid$(TotalBuffer, FirstPos, 1) = "D" Or Mid$(TotalBuffer, FirstPos, 1) = ChrW$(272) Or Mid$(TotalBuffer, FirstPos, 1) = ChrW$(273)) Then
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
        If (Mid$(TotalBuffer, LastPos, 1) = ChrW$(272) Or Mid$(TotalBuffer, LastPos, 1) = ChrW$(273)) Then
            FoundChar = True
            Exit Do
        End If
        LastPos = LastPos - 1
    Loop
    
    If Not FoundChar Then
        FoundChar = False
        Do While LastPos >= FirstPos
            If (Mid$(TotalBuffer, LastPos, 1) = "D" Or Mid$(TotalBuffer, LastPos, 1) = "d") Then
                FoundChar = True
                Exit Do
            End If
            LastPos = LastPos - 1
        Loop
    End If
    
    If Not FoundChar Then LastPos = FirstPos
    
    RealPos = LastPos
    
    If RealPos > 1 Then
        If InStr(1, STRING_RESET_TELEX & STRING_CAN_BEFORE_D_CHAR, Mid$(TotalBuffer, RealPos - 1, 1), vbTextCompare) <= 0 Then
            PutToBuffer Ch
            Exit Sub
        End If
    End If
    
    UniBuf = Mid$(TotalBuffer, RealPos, KeyPushed - RealPos + 1)
    BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
    
    If Mid$(TotalBuffer, RealPos, 1) = "d" Then
        Mid$(TotalBuffer, RealPos, 1) = ChrW$(273)
        UniBuf = Mid$(TotalBuffer, RealPos, KeyPushed - RealPos + 1)
        Exit Sub
    ElseIf Mid$(TotalBuffer, RealPos, 1) = "D" Then
        Mid$(TotalBuffer, RealPos, 1) = ChrW$(272)
        UniBuf = Mid$(TotalBuffer, RealPos, KeyPushed - RealPos + 1)
        Exit Sub
    Else
        VietKeyTempOff = True
        Mid$(TotalBuffer, RealPos, 1) = Left$(UniToTelex(Mid$(TotalBuffer, RealPos, 1)), 1)
        PutToBuffer Ch
        UniBuf = Mid$(TotalBuffer, RealPos, KeyPushed - RealPos + 1)
        Exit Sub
    End If
End Sub



Public Sub ProcessDoubleChar(Ch As String)
    If LastVietOff > 0 Then
        If UCase$(Ch) <> UCase$(Mid$(TotalBuffer, LastVietOff, 1)) Then VietKeyTempOff = False
    End If

    If ((inputMethod <> TELEX_INPUT) Or (VietKeyTempOff = True) Or (KeyPushed <= 0) Or (Not VietNameseKeyboard)) Then
        PutToBuffer Ch
        Exit Sub
    End If

    Dim FPos As Integer, LPos As Integer, Pos As Integer, sW As String, Founds As Boolean
    
    FPos = GetLastWord(TotalBuffer)
    If FPos < LastVietOff Then FPos = LastVietOff
    LPos = KeyPushed
    Founds = False
    
    Do While FPos <= LPos
        If VowelSpecify(Mid$(TotalBuffer, FPos, 1)) <> 0 Then
            Founds = True
            Exit Do
        End If
        FPos = FPos + 1
    Loop
    
    If Not Founds Then
        PutToBuffer Ch
        Exit Sub
    End If
    
    Founds = False
    Do While LPos >= FPos
        If UCase$(Left$(UniToTelex(Mid$(TotalBuffer, LPos, 1)), 1)) = UCase$(Ch) Then
            Founds = True
            Exit Do
        End If
        LPos = LPos - 1
    Loop
    If LPos < FPos Then LPos = FPos
    
    If Not Founds Then
        PutToBuffer Ch
        Exit Sub
    End If
    Pos = LPos
    
    Do While Pos >= LPos And LPos - Pos <= MAX_VOWEL_STRING_LENGTH
        If VowelSpecify(Mid$(TotalBuffer, Pos, 1)) = 0 Then
            Exit Do
        End If
        Pos = Pos - 1
    Loop
    
    If Pos < LPos Then Pos = FPos
    
    sW = Mid$(TotalBuffer, Pos, LPos - Pos + 1)
    'frmMain.Caption = sW
    FPos = Pos
    
    Select Case Len(sW)
        Case 1:
            If LPos < KeyPushed Then
                If (InStr(1, "c,p,t", Mid$(TotalBuffer, LPos + 1, 1), vbTextCompare) > 0) And InStr(1, "f,r,x", Right$(UniToTelex(Left$(sW, 1)), 1), vbTextCompare) > 0 Then
                    PutToBuffer Ch
                    Exit Sub
                End If
            End If
            Pos = LPos
        Case 2:
            If VowelSpecify(Right$(sW, 1)) = NONE_MARK_VOWEL Then
                If UCase$(Right$(sW, 1)) = "A" Then
                    If LPos < KeyPushed Then
                        If (VowelSpecify(Mid$(TotalBuffer, LPos + 1)) <> 0) Or (InStr(1, "c,m,n,p,t", Mid$(TotalBuffer, LPos + 1, 1), vbTextCompare) <= 0) Then
                            PutToBuffer Ch
                            Exit Sub
                        End If
                    End If
                    If (UCase$(Left$(UniToTelex(Left$(sW, 1)), 1)) = "U") Then
                        If (VowelSpecify(Left$(sW, 1)) = BREVE_MARK_VOWEL) Then
                            Mid$(TotalBuffer, FPos, 1) = Left$(UniToTelex(Left$(sW, 1)), 1)
                            Mid$(TotalBuffer, LPos, 1) = TelexToUni(Mid$(TotalBuffer, LPos, 1) & Ch)
                            Pos = LPos - 1
                            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
                            BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
                            Exit Sub
                        ElseIf (VowelSpecify(Left$(sW, 1)) = TONE_AND_BREVE_MARK_VOWEL) Then
                            Mid$(TotalBuffer, FPos, 1) = Left$(UniToTelex(Left$(sW, 1)), 1)
                            Mid$(TotalBuffer, LPos, 1) = TelexToUni(Mid$(TotalBuffer, LPos, 1) & Ch & Right$(UniToTelex(Left$(sW, 1)), 1))
                            Pos = LPos - 1
                            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
                            BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
                            Exit Sub
                        ElseIf (VowelSpecify(Left$(sW, 1)) = TONE_MARK_VOWEL) Then
                            If LPos < KeyPushed Then
                                If (InStr(1, "c,p,t", Mid$(TotalBuffer, LPos + 1, 1), vbTextCompare) > 0) And InStr(1, "f,r,x", Right$(UniToTelex(Left$(sW, 1)), 1), vbTextCompare) > 0 Then
                                    PutToBuffer Ch
                                    Exit Sub
                                End If
                            End If
                            Mid$(TotalBuffer, FPos, 1) = Left$(UniToTelex(Left$(sW, 1)), 1)
                            Mid$(TotalBuffer, LPos, 1) = TelexToUni(Mid$(TotalBuffer, LPos, 1) & Ch & Right$(UniToTelex(Left$(sW, 1)), 1))
                            Pos = LPos - 1
                            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
                            BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
                            Exit Sub
                        ElseIf (VowelSpecify(Left$(sW, 1)) = NONE_MARK_VOWEL) Then
                            Pos = LPos
                        End If
                    Else
                        PutToBuffer Ch
                        Exit Sub
                    End If
                ElseIf UCase$(Right$(sW, 1)) = "E" Then
                    If LPos < KeyPushed Then
                        If (InStr(1, "c,m,n,p,t,u", Mid$(TotalBuffer, LPos + 1, 1), vbTextCompare) <= 0) Then
                            PutToBuffer Ch
                            Exit Sub
                        End If
                        If (InStr(1, "c,p,t", Mid$(TotalBuffer, LPos + 1, 1), vbTextCompare) > 0) And InStr(1, "f,r,x", Right$(UniToTelex(Left$(sW, 1)), 1), vbTextCompare) > 0 Then
                            PutToBuffer Ch
                            Exit Sub
                        End If
                    End If
                
                    If (UCase$(Left$(UniToTelex(Left$(sW, 1)), 1)) = "I") Or (UCase$(Left$(UniToTelex(Left$(sW, 1)), 1)) = "Y") Then
                        If VowelSpecify(Left$(sW, 1)) = TONE_MARK_VOWEL Then

                            Mid$(TotalBuffer, FPos, 1) = Left$(UniToTelex(Left$(sW, 1)), 1)
                            Mid$(TotalBuffer, LPos, 1) = TelexToUni(Mid$(TotalBuffer, LPos, 1) & Ch & Right$(UniToTelex(Left$(sW, 1)), 1))
                            Pos = LPos - 1
                            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
                            BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
                            Exit Sub
                        ElseIf VowelSpecify(Left$(sW, 1)) = NONE_MARK_VOWEL Then
                            Pos = LPos
                        End If
                    ElseIf (UCase$(Left$(UniToTelex(Left$(sW, 1)), 1)) = "U") Then
                        If VowelSpecify(Left$(sW, 1)) = TONE_AND_BREVE_MARK_VOWEL Then
                            PutToBuffer Ch
                            Exit Sub
                        ElseIf VowelSpecify(Left$(sW, 1)) = TONE_MARK_VOWEL Then
                            If LPos < KeyPushed Then
                                If (InStr(1, "c,p,t", Mid$(TotalBuffer, LPos + 1, 1), vbTextCompare) > 0) And InStr(1, "f,r,x", Right$(UniToTelex(Left$(sW, 1)), 1), vbTextCompare) > 0 Then
                                    PutToBuffer Ch
                                    Exit Sub
                                End If
                            End If
                        
                            Mid$(TotalBuffer, FPos, 1) = Left$(UniToTelex(Left$(sW, 1)), 1)
                            Mid$(TotalBuffer, LPos, 1) = TelexToUni(Mid$(TotalBuffer, LPos, 1) & Ch & Right$(UniToTelex(Left$(sW, 1)), 1))
                            Pos = LPos - 1
                            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
                            BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
                        ElseIf VowelSpecify(Left$(sW, 1)) = BREVE_MARK_VOWEL Then
                            Mid$(TotalBuffer, FPos, 1) = Left$(UniToTelex(Left$(sW, 1)), 1)
                            Mid$(TotalBuffer, LPos, 1) = TelexToUni(Mid$(TotalBuffer, LPos, 1) & Ch)
                            Pos = LPos - 1
                            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
                            BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
                        ElseIf VowelSpecify(Left$(sW, 1)) = NONE_MARK_VOWEL Then
                            Pos = LPos
                        End If
                    Else
                        PutToBuffer Ch
                        Exit Sub
                    End If
                ElseIf UCase$(Right$(sW, 1)) = "O" Then
                    If LPos < KeyPushed Then
                        If (InStr(1, "c,m,n,p,t,i", Mid$(TotalBuffer, LPos + 1, 1), vbTextCompare) <= 0) Then
                            PutToBuffer Ch
                            Exit Sub
                        End If
                        
                        If (InStr(1, "c,p,t", Mid$(TotalBuffer, LPos + 1, 1), vbTextCompare) > 0) And InStr(1, "f,r,x", Right$(UniToTelex(Left$(sW, 1)), 1), vbTextCompare) > 0 Then
                            PutToBuffer Ch
                            Exit Sub
                        End If
                        
                    End If
                    If UCase$(Left$(UniToTelex(Left$(sW, 1)), 1)) = "U" Then
                        If VowelSpecify(Left$(sW, 1)) = TONE_AND_BREVE_MARK_VOWEL Then
                            PutToBuffer Ch
                            Exit Sub
                        ElseIf VowelSpecify(Left$(sW, 1)) = TONE_MARK_VOWEL Then
                            If LPos < KeyPushed Then
                                If (InStr(1, "c,p,t", Mid$(TotalBuffer, LPos + 1, 1), vbTextCompare) > 0) And InStr(1, "f,r,x", Right$(UniToTelex(Left$(sW, 1)), 1), vbTextCompare) > 0 Then
                                    PutToBuffer Ch
                                    Exit Sub
                                End If
                            End If
                            
                            Mid$(TotalBuffer, FPos, 1) = Left$(UniToTelex(Left$(sW, 1)), 1)
                            Mid$(TotalBuffer, LPos, 1) = TelexToUni(Mid$(TotalBuffer, LPos, 1) & Ch & Right$(UniToTelex(Left$(sW, 1)), 1))
                            Pos = LPos - 1
                            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
                            BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
                        ElseIf VowelSpecify(Left$(sW, 1)) = BREVE_MARK_VOWEL Then
                            Mid$(TotalBuffer, FPos, 1) = Left$(UniToTelex(Left$(sW, 1)), 1)
                            Mid$(TotalBuffer, LPos, 1) = TelexToUni(Mid$(TotalBuffer, LPos, 1) & Ch)
                            Pos = LPos - 1
                            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
                            BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
                        ElseIf VowelSpecify(Left$(sW, 1)) = NONE_MARK_VOWEL Then
                            Pos = LPos
                        End If
                    Else
                        PutToBuffer Ch
                        Exit Sub
                    End If
                End If
            ElseIf VowelSpecify(Right$(sW, 1)) = BREVE_MARK_VOWEL Then
                If UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "O" Then
                    If UCase$(Left$(UniToTelex(Left$(sW, 1)), 1)) = "U" Then
                        If VowelSpecify(Left$(sW, 1)) = NONE_MARK_VOWEL Then
                            Pos = LPos
                        ElseIf VowelSpecify(Left$(sW, 1)) = BREVE_MARK_VOWEL Then
                            Mid$(TotalBuffer, FPos, 1) = Left$(UniToTelex(Mid$(TotalBuffer, FPos, 1)), 1)
                            Mid$(TotalBuffer, LPos, 1) = TelexToUni(Left$(UniToTelex(Mid$(TotalBuffer, LPos, 1)), 1) & Ch)
                            Pos = LPos - 1
                            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
                            BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
                            Exit Sub
                        Else
                            Pos = LPos
                        End If
                    Else
                        PutToBuffer Ch
                        Exit Sub
                    End If
                ElseIf UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "A" Then
                    Pos = LPos
                ElseIf UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "E" Then
                    Pos = LPos
                End If
            ElseIf VowelSpecify(Right$(sW, 1)) = TONE_MARK_VOWEL Then
                If UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "A" Then
                    If UCase$(Left$(sW, 1)) <> "U" Then
                        PutToBuffer Ch
                        Exit Sub
                    Else
                        Pos = LPos
                    End If
                ElseIf UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "E" Then
                ElseIf UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "O" Then
                End If
            ElseIf VowelSpecify(Right$(sW, 1)) = TONE_AND_BREVE_MARK_VOWEL Then
                If UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "A" Then
                    PutToBuffer Ch
                    Exit Sub
                ElseIf UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "E" Then
                    If UCase$(Left$(UniToTelex(Left$(sW, 1)), 1)) = "U" Then
                        Pos = LPos
                    ElseIf UCase$(Left$(UniToTelex(Left$(sW, 1)), 1)) = "I" Then
                        Pos = LPos
                    Else
                        PutToBuffer Ch
                        Exit Sub
                    End If
                ElseIf UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "O" Then
                    If UCase$(Left$(UniToTelex(Left$(sW, 1)), 1)) = "U" Then
                        If VowelSpecify(Left$(sW, 1)) = NONE_MARK_VOWEL Then
                            Mid$(TotalBuffer, LPos, 1) = Left$(UniToTelex(Mid$(TotalBuffer, LPos, 1)), 1)
                            Pos = LPos - 1
                            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
                            BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
                            Exit Sub
                        ElseIf VowelSpecify(Left$(sW, 1)) = BREVE_MARK_VOWEL Then
                            Mid$(TotalBuffer, FPos, 1) = Left$(UniToTelex(Mid$(TotalBuffer, FPos, 1)), 1)
                            Mid$(TotalBuffer, LPos, 1) = TelexToUni(Left$(UniToTelex(Mid$(TotalBuffer, LPos, 1)), 1) & Ch & Right$(UniToTelex(Mid$(TotalBuffer, LPos, 1)), 1))
                            Pos = LPos - 1
                            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
                            BackNumbers = Len(UniBuf)
                            Exit Sub
                        'ElseIf VowelSpecify(Left$(sW, 1)) = TONE_MARK_VOWEL Then
                        'ElseIf VowelSpecify(Left$(sW, 1)) = TONE_AND_BREVE_MARK_VOWEL Then
                        End If
                    Else
                        PutToBuffer Ch
                        Exit Sub
                    End If
                End If
            Else
                Pos = LPos
            End If
        Case 3
            If VowelSpecify(Right$(sW, 1)) > NONE_MARK_VOWEL Then
                Pos = LPos
            ElseIf VowelSpecify(Mid$(sW, 2, 1)) > NONE_MARK_VOWEL Then
                Pos = LPos - 1
            ElseIf VowelSpecify(Left$(sW, 1)) > NONE_MARK_VOWEL Then
                Pos = LPos - 2
            Else
                Pos = LPos
            End If
            If UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "E" And UCase$(Left$(UniToTelex(Mid$(sW, 2, 1)), 1)) = "Y" And UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "U" Then
                PutToBuffer Ch
                Exit Sub
            Else
                If VowelSpecify(Left$(sW, 1)) = TONE_MARK_VOWEL Then
                    Mid$(TotalBuffer, FPos, 1) = Left$(UniToTelex(Left$(sW, 1)), 1)
                    Mid$(TotalBuffer, LPos, 1) = TelexToUni(Mid$(TotalBuffer, LPos, 1) & Ch & Right$(UniToTelex(Left$(sW, 1)), 1))
                    Pos = FPos
                    UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
                    BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
                    Exit Sub
                ElseIf VowelSpecify(Mid$(sW, 2, 1)) = TONE_MARK_VOWEL Then
                    Mid$(TotalBuffer, LPos - 1, 1) = Left$(UniToTelex(Mid$(sW, 2, 1)), 1)
                    Mid$(TotalBuffer, LPos, 1) = TelexToUni(Mid$(TotalBuffer, LPos, 1) & Ch & Right$(UniToTelex(Mid$(sW, 2, 1)), 1))
                    Pos = LPos - 1
                    UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
                    BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
                    Exit Sub
                End If
            End If
        Case Else
            'Chua tim ra tu co 4 nguyen am
    End Select
    
    UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
    BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
    
    Dim sAnsi As String
    sAnsi = UniToTelex(Mid$(TotalBuffer, Pos, 1))
    
    If VowelSpecify(Mid$(TotalBuffer, Pos, 1)) = NONE_MARK_VOWEL Then
        Mid$(TotalBuffer, Pos, 1) = TelexToUni(Mid$(TotalBuffer, Pos, 1) & Ch)
        UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
        Exit Sub
    ElseIf VowelSpecify(Mid$(TotalBuffer, Pos, 1)) = BREVE_MARK_VOWEL Then
        If UCase$(Right$(sAnsi, 1)) = UCase$(Ch) Then
            Mid$(TotalBuffer, Pos, 1) = Left$(sAnsi, 1)
            PutToBuffer Ch
            LastVietOff = KeyPushed
            VietKeyTempOff = True
            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
            Exit Sub
        Else
            Mid$(TotalBuffer, Pos, 1) = TelexToUni(Left$(sAnsi, 1) & Ch)
            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
            Exit Sub
        End If
    ElseIf VowelSpecify(Mid$(TotalBuffer, Pos, 1)) = TONE_MARK_VOWEL Then
        Mid$(TotalBuffer, Pos, 1) = TelexToUni(Left$(sAnsi, 1) & Ch & Right$(sAnsi, 1))
        UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
        Exit Sub
    ElseIf VowelSpecify(Mid$(TotalBuffer, Pos, 1)) = TONE_AND_BREVE_MARK_VOWEL Then
        If UCase$(Mid$(sAnsi, 2, 1)) = UCase$(Ch) Then
            Mid$(TotalBuffer, Pos, 1) = TelexToUni(Left$(sAnsi, 1) & Right$(sAnsi, 1))
            PutToBuffer Ch
            LastVietOff = KeyPushed
            VietKeyTempOff = True
            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
            Exit Sub
        Else
            Mid$(TotalBuffer, Pos, 1) = TelexToUni(Left$(sAnsi, 1) & Ch & Right$(sAnsi, 1))
            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
            Exit Sub
        End If
    End If
End Sub



Public Sub ProcessToneMark(Ch As String)
    
    If (inputMethod <> TELEX_INPUT) Or (VietKeyTempOff = True) Or (KeyPushed <= 0) Or (Not VietNameseKeyboard) Then
        PutToBuffer Ch
        Exit Sub
    End If
    
    Dim FPos As Integer, LPos As Integer, Pos As Integer, sW As String, Founds As Boolean
    
    FPos = GetLastWord(TotalBuffer)
    If FPos < LastVietOff Then FPos = LastVietOff
    LPos = KeyPushed
    Founds = False
    
    Do While FPos <= LPos
        If VowelSpecify(Mid$(TotalBuffer, FPos, 1)) <> 0 Then
            Founds = True
            Exit Do
        End If
        FPos = FPos + 1
    Loop
    
    If Not Founds Then
        PutToBuffer Ch
        Exit Sub
    End If
    
    Do While LPos >= FPos
        If VowelSpecify(Mid$(TotalBuffer, LPos, 1)) <> 0 Then
            Founds = True
            Exit Do
        End If
        LPos = LPos - 1
    Loop
    If LPos < FPos Then LPos = FPos
    
    Pos = FPos
    FPos = LPos
    Do While FPos >= Pos
        If (VowelSpecify(Mid$(TotalBuffer, FPos, 1)) = 0) Or (LPos - FPos > MAX_VOWEL_STRING_LENGTH) Then
            Exit Do
        End If
        FPos = FPos - 1
    Loop
    If FPos < Pos Then FPos = Pos
    
    sW = Mid$(TotalBuffer, FPos, LPos - FPos + 1)
    
    If LPos < KeyPushed Then
        If (InStr(1, "f,r,x", Ch, vbTextCompare) > 0) And InStr(1, "c,t,p", Mid$(TotalBuffer, LPos + 1, 1), vbTextCompare) > 0 Then
            PutToBuffer Ch
            Exit Sub
        End If
    End If
    
    Select Case Len(sW)
    
        Case 1:
            If LPos > 1 Then
                If (UCase$(Mid$(TotalBuffer, LPos - 1, 1)) = "Q" And UCase$(sW) = "U") Then
                    PutToBuffer Ch
                    Exit Sub
                End If
            End If
            Pos = LPos
            'frmMain.Caption = sW
        Case 2:
            
            If FPos > 1 Then
                If (UCase$(Mid$(TotalBuffer, FPos - 1, 1)) = "Q" And UCase$(Mid$(TotalBuffer, FPos, 1)) = "U") Or (UCase$(Mid$(TotalBuffer, FPos - 1, 1)) = "G" And UCase$(Mid$(TotalBuffer, FPos, 1)) = "I") Then
                    Pos = LPos
                End If
            End If
            If UCase$(Right$(sW, 1)) = "O" And UCase$(Left$(UniToTelex(Left$(sW, 1)), 1)) <> "A" Then
                PutToBuffer Ch
                Exit Sub
            End If
            If ((UCase$(Left$(sW, 1)) = "A" And UCase$(Right$(sW, 1)) = "Y") Or (UCase$(Left$(sW, 1)) = "A" And UCase$(Right$(sW, 1)) = "I") Or (UCase$(Left$(sW, 1)) = "O" And UCase$(Right$(sW, 1)) = "I") Or (UCase$(Left$(sW, 1)) = "U" And UCase$(Right$(sW, 1)) = "I") Or (UCase$(Left$(sW, 1)) = "U" And UCase$(Right$(sW, 1)) = "A")) Then
            ElseIf VowelSpecify(Right$(sW, 1)) > NONE_MARK_VOWEL Then
                Pos = LPos
            ElseIf VowelSpecify(Left$(sW, 1)) > NONE_MARK_VOWEL Then
                Pos = LPos - 1
            Else
                If ((UCase$(Left$(sW, 1)) = "O" And UCase$(Right$(sW, 1)) = "A") Or (UCase$(Left$(sW, 1)) = "O" And UCase$(Right$(sW, 1)) = "E") Or (UCase$(Left$(sW, 1)) = "U" And UCase$(Right$(sW, 1)) = "A") Or (UCase$(Left$(sW, 1)) = "U" And UCase$(Right$(sW, 1)) = "Y")) Then
                    If ToneMarkIsOldStyle Then
                        If LPos < KeyPushed Then
                            Pos = LPos
                        Else
                            Pos = LPos - 1
                        End If
                    Else
                        Pos = LPos
                    End If
                End If
            End If
        Case 3:
            If FPos > 1 Then
                If (UCase$(Mid$(TotalBuffer, FPos - 1, 1)) = "Q" And UCase$(Mid$(TotalBuffer, FPos, 1)) = "U") Then
                    Pos = LPos - 1
                End If
            End If
        

            Dim II As Integer
            Founds = False
            For II = Len(sW) To 1 Step -1
                If VowelSpecify(Mid$(sW, II, 1)) = TONE_AND_BREVE_MARK_VOWEL Then
                    Founds = True
                    Exit For
                End If
            Next II
                
            'frmMain.Caption = Founds
                
            If Not Founds Then
                Founds = False
                For II = Len(sW) To 1 Step -1
                    If VowelSpecify(Mid$(sW, II, 1)) = TONE_MARK_VOWEL Then
                        Founds = True
                        Exit For
                    End If
                Next II
            Else
                If II = 3 Then
                    Pos = LPos
                ElseIf II = 2 Then
                    Pos = LPos - 1
                ElseIf II = 1 Then
                    Pos = FPos
                End If
            End If
                
            'frmMain.Caption = Founds
            If Not Founds Then
                Founds = False
                For II = Len(sW) To 1 Step -1
                    If VowelSpecify(Mid$(sW, II, 1)) = BREVE_MARK_VOWEL Then
                        Founds = True
                        Exit For
                    End If
                Next II
            Else
                If II = 3 Then
                    Pos = LPos
                ElseIf II = 2 Then
                    Pos = LPos - 1
                ElseIf II = 1 Then
                    Pos = LPos - 2
                End If
            End If
            
            If Not Founds Then
                Pos = LPos - 1
            Else
                If II = 3 Then
                    Pos = LPos
                ElseIf II = 2 Then
                    Pos = LPos - 1
                ElseIf II = 1 Then
                    Pos = LPos - 2
                End If
            End If
                
    End Select
    
    UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
    BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
    
    Dim sAnsi As String
    sAnsi = UniToTelex(Mid$(TotalBuffer, Pos, 1))
    If VowelSpecify(Mid$(TotalBuffer, Pos, 1)) = NONE_MARK_VOWEL Then
        Mid$(TotalBuffer, Pos, 1) = TelexToUni(Mid$(TotalBuffer, Pos, 1) & Ch)
        UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
        Exit Sub
    ElseIf VowelSpecify(Mid$(TotalBuffer, Pos, 1)) = BREVE_MARK_VOWEL Then
        Mid$(TotalBuffer, Pos, 1) = TelexToUni(sAnsi & Ch)
        UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
        Exit Sub
    ElseIf VowelSpecify(Mid$(TotalBuffer, Pos, 1)) = TONE_MARK_VOWEL Then
        If UCase$(Right$(sAnsi, 1)) = UCase$(Ch) Then
            Mid$(TotalBuffer, Pos, 1) = Left$(sAnsi, 1)
            PutToBuffer Ch
            LastVietOff = KeyPushed
            VietKeyTempOff = True
            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
            Exit Sub
        Else
            Mid$(TotalBuffer, Pos, 1) = TelexToUni(Left$(sAnsi, 1) & Ch)
            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
            Exit Sub
        End If
    ElseIf VowelSpecify(Mid$(TotalBuffer, Pos, 1)) = TONE_AND_BREVE_MARK_VOWEL Then
        If UCase$(Right$(sAnsi, 1)) = UCase$(Ch) Then
            Mid$(TotalBuffer, Pos, 1) = TelexToUni(Left$(sAnsi, 2))
            PutToBuffer Ch
            LastVietOff = KeyPushed
            VietKeyTempOff = True
            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
            Exit Sub
        Else
            Mid$(TotalBuffer, Pos, 1) = TelexToUni(Left$(sAnsi, 2) & Ch)
            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
            Exit Sub
        End If
    End If
    
    
End Sub



Public Sub ProcessBreveMark(Ch As String)
    If LastVietOff > 0 Then
        If UCase$(Ch) <> UCase$(Mid$(TotalBuffer, LastVietOff, 1)) Then VietKeyTempOff = False
    End If

    If ((inputMethod <> TELEX_INPUT) Or (VietKeyTempOff = True) Or (Not VietNameseKeyboard)) Then
        PutToBuffer Ch
        Exit Sub
    End If

    If KeyPushed <= 0 Then
        PutToBuffer IIf(Ch = UCase$(Ch), TelexToUni("UW"), TelexToUni("uw"))
        UniBuf = Right$(TotalBuffer, 1)
        BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
        LastIsWConverted = True
        LastIsShortkeyConverted = False
        Exit Sub
    End If
    
    Dim FPos As Integer, LPos As Integer, Pos As Integer, sW As String, Founds As Boolean
    
    FPos = GetLastWord(TotalBuffer)
    If FPos < LastVietOff Then FPos = LastVietOff
    
    LPos = KeyPushed
    Founds = False
    
    Do While LPos >= FPos
        If (UCase$(Left$(UniToTelex(Mid$(TotalBuffer, LPos, 1)), 1)) = "A") Or ((UCase$(Left$(UniToTelex(Mid$(TotalBuffer, LPos, 1)), 1))) = "O") Or ((UCase$(Left$(UniToTelex(Mid$(TotalBuffer, LPos, 1)), 1))) = "U") Then
            Founds = True
            Exit Do
        End If
        LPos = LPos - 1
    Loop
    
    If Not Founds Then
        If InStr(1, "b,c,d,g,h,l,m,n,r,s,t,v,x" & STRING_RESET_TELEX, Mid$(TotalBuffer, KeyPushed, 1), vbTextCompare) > 0 Then
            UniBuf = Right$(TotalBuffer, 1)
            BackNumbers = LenX(UniBuf)

            PutToBuffer IIf(Ch = UCase$(Ch), TelexToUni("UW"), TelexToUni("uw"))
            LastIsWConverted = True
            LastIsShortkeyConverted = False
            UniBuf = Right$(TotalBuffer, 2)
        Else
            PutToBuffer Ch
        End If
        Exit Sub
    End If
    
    If Mid$(TotalBuffer, LPos, 1) = ChrW$(&H1A0) Or Mid$(TotalBuffer, LPos, 1) = ChrW$(&H1A1) Or Mid$(TotalBuffer, LPos, 1) = ChrW$(&H1AF) Or Mid$(TotalBuffer, LPos, 1) = ChrW$(&H1B0) Then
        If LastIsShortkeyConverted Then
            PutToBuffer Ch
            Exit Sub
        End If
    End If
    
    If LPos < FPos Then LPos = FPos
    Pos = LPos
    Founds = False
    Do While (Pos >= FPos) And (LPos - Pos <= MAX_VOWEL_STRING_LENGTH)
        If VowelSpecify(Mid$(TotalBuffer, Pos, 1)) = 0 Then
            Exit Do
        End If
        Pos = Pos - 1
    Loop
    
    If Pos < FPos Then Pos = FPos
    
    Do While FPos <= LPos
        If VowelSpecify(Mid$(TotalBuffer, FPos, 1)) <> 0 Then
            Exit Do
        End If
        FPos = FPos + 1
    Loop
JUMP:
    sW = Mid$(TotalBuffer, FPos, LPos - FPos + 1)
    'frmMain.Caption = sW
    Pos = LPos

    Select Case Len(sW)
        Case 1:
            If UCase$(UniToTelex(sW)) = "A" Then
                If LPos < KeyPushed Then
                    If VowelSpecify(Mid$(TotalBuffer, LPos + 1, 1)) > 0 Then
                        PutToBuffer Ch
                        Exit Sub
                    End If
                    If InStr(1, "c,m,n,p,t", Mid$(TotalBuffer, LPos + 1, 1), vbTextCompare) <= 0 Then
                        PutToBuffer Ch
                        Exit Sub
                    End If
                End If
                Pos = LPos
            End If
            If UCase$(sW) = "U" Then
                If LPos > 1 Then
                    If UCase$(Mid$(TotalBuffer, LPos - 1, 1)) = "Q" Then
                        PutToBuffer Ch
                        Exit Sub
                    Else
                        Pos = LPos
                    End If
                End If
                If LPos < KeyPushed Then
                    If VowelSpecify(Mid$(TotalBuffer, LPos + 1, 1)) <> 0 And UCase$(Left$(UniToTelex(Mid$(TotalBuffer, LPos + 1, 1)), 1)) <> "O" And UCase$(Left$(UniToTelex(Mid$(TotalBuffer, LPos + 1, 1)), 1)) <> "A" Then
                        PutToBuffer Ch
                        Exit Sub
                    End If
                End If
            End If
            LastIsShortkeyConverted = False
            Pos = LPos
        Case 2:
            If VowelSpecify(Right$(sW, 1)) = NONE_MARK_VOWEL Then
                If UCase$(Right$(sW, 1)) = "A" Then
                
                    If LPos < KeyPushed Then
                        If VowelSpecify(Mid$(TotalBuffer, LPos + 1, 1)) > 0 Then
                            PutToBuffer Ch
                            Exit Sub
                        End If
                        If InStr(1, "c,m,n,p,t", Mid$(TotalBuffer, LPos + 1, 1), vbTextCompare) <= 0 Then
                            PutToBuffer Ch
                            Exit Sub
                        End If
                    End If
                
                    If UCase$(Left$(UniToTelex(Left$(sW, 1)), 1)) = "U" Then
                        If FPos > 1 Then
                            If UCase$(Mid$(TotalBuffer, FPos - 1, 1)) = "Q" Then
                                Pos = LPos
                            Else
                                If LPos < KeyPushed Then
                                    If (VowelSpecify(Left$(sW, 1)) = TONE_AND_BREVE_MARK_VOWEL) Or (VowelSpecify(Left$(sW, 1)) = TONE_MARK_VOWEL) Or (VowelSpecify(Left$(sW, 1)) = NONE_MARK_VOWEL) Then
                                        PutToBuffer Ch
                                        Exit Sub
                                    Else
                                        Pos = LPos - 1
                                    End If
                                Else
                                    Pos = LPos - 1
                                End If
                            End If
                        Else
                            Pos = LPos - 1
                        End If
                    ElseIf UCase$(Left$(UniToTelex(Left$(sW, 1)), 1)) = "O" Then
                        If VowelSpecify(Left$(sW, 1)) = NONE_MARK_VOWEL Then
                            Pos = LPos
                        ElseIf VowelSpecify(Left$(sW, 1)) = BREVE_MARK_VOWEL Then
                            Mid$(TotalBuffer, FPos, 1) = Left$(UniToTelex(Mid$(TotalBuffer, FPos, 1)), 1)
                            Mid$(TotalBuffer, LPos, 1) = TelexToUni(Mid$(TotalBuffer, LPos, 1) & Ch)
                            Pos = LPos - 1
                            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
                            BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
                            Exit Sub
                        ElseIf VowelSpecify(Left$(sW, 1)) = TONE_MARK_VOWEL Then
                            If LPos < KeyPushed Then
                                If InStr(1, "c,m,n,p,t", Mid$(TotalBuffer, LPos + 1, 1), vbTextCompare) <= 0 Then
                                    PutToBuffer Ch
                                    Exit Sub
                                End If
                                
                                If (InStr(1, "C,P,T", Mid$(TotalBuffer, LPos + 1, 1), vbTextCompare) > 0) And (InStr(1, "F,R,X", Right$(UniToTelex(Left$(sW, 1)), 1), vbTextCompare) > 0) Then
                                    PutToBuffer Ch
                                    Exit Sub
                                End If
                            End If
                            Mid$(TotalBuffer, FPos, 1) = Left$(UniToTelex(Mid$(TotalBuffer, FPos, 1)), 1)
                            Mid$(TotalBuffer, LPos, 1) = TelexToUni(Mid$(TotalBuffer, LPos, 1) & Ch & Right$(UniToTelex(Left$(sW, 1)), 1))
                            Pos = LPos - 1
                            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
                            BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
                            Exit Sub
                        ElseIf VowelSpecify(Left$(sW, 1)) = TONE_AND_BREVE_MARK_VOWEL Then
                            PutToBuffer Ch
                            Exit Sub
                        End If
                    Else
                        PutToBuffer Ch
                        Exit Sub
                    End If
                ElseIf UCase$(Right$(sW, 1)) = "O" Then
                    If LPos < KeyPushed Then
                        If InStr(1, "i,c,m,n,p,t,u", Mid$(TotalBuffer, LPos + 1, 1), vbTextCompare) <= 0 Then
                            PutToBuffer Ch
                            Exit Sub
                        End If
                    End If
                
                    If UCase$(Left$(UniToTelex(Left$(sW, 1)), 1)) = "U" Then
                        If VowelSpecify(Left$(sW, 1)) = NONE_MARK_VOWEL Then
                            If FPos > 1 Then
                                If UCase$(Mid$(TotalBuffer, FPos - 1, 1)) = "Q" Then
                                    Pos = LPos
                                Else
                                    Mid$(TotalBuffer, LPos, 1) = TelexToUni(Mid$(TotalBuffer, LPos, 1) & Ch)
                                    Mid$(TotalBuffer, FPos, 1) = TelexToUni(Mid$(TotalBuffer, FPos, 1) & Ch)
                                    Pos = LPos - 1
                                    UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
                                    BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
                                    LastIsShortkeyConverted = False
                                    Exit Sub
                                End If
                            Else
                                Pos = LPos - 1
                            End If
                        ElseIf VowelSpecify(Left$(sW, 1)) = BREVE_MARK_VOWEL Then
                            Pos = LPos
                        Else
                            Mid$(TotalBuffer, FPos, 1) = TelexToUni(Left$(UniToTelex(Mid$(TotalBuffer, FPos, 1)), 1) & Ch)
                            Mid$(TotalBuffer, LPos, 1) = TelexToUni(Mid$(TotalBuffer, LPos, 1) & Ch & Right$(UniToTelex(Left$(sW, 1)), 1))
                            Pos = LPos - 1
                            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
                            BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
                            LastIsShortkeyConverted = False
                            Exit Sub
                        End If
                    Else
                        Pos = LPos
                    End If
                ElseIf UCase$(Right$(sW, 1)) = "U" Then
                    If LPos < KeyPushed Then
                        If (VowelSpecify(Mid$(TotalBuffer, LPos + 1, 1)) <> 0) Or (InStr(1, "c,m,n,p,t", Mid$(TotalBuffer, LPos + 1, 1), vbTextCompare) <= 0) Then
                            PutToBuffer Ch
                            Exit Sub
                        End If
                    End If
                    If UCase$(Left$(UniToTelex(Left$(sW, 1)), 1)) = "U" Then
                        If VowelSpecify(Left$(sW, 1)) = NONE_MARK_VOWEL Then
                            If LPos >= KeyPushed Then
                                Pos = LPos - 1
                            Else
                                PutToBuffer Ch
                                Exit Sub
                            End If
                        ElseIf VowelSpecify(Left$(sW, 1)) = BREVE_MARK_VOWEL Then
                            Pos = LPos - 1
                        ElseIf VowelSpecify(Left$(sW, 1)) = TONE_MARK_VOWEL Then
                            Pos = LPos - 1
                        ElseIf VowelSpecify(Left$(sW, 1)) = TONE_AND_BREVE_MARK_VOWEL Then
                            Pos = LPos - 1
                        End If
                    Else
                        Pos = LPos
                    End If
                End If
            ElseIf VowelSpecify(Right$(sW, 1)) = BREVE_MARK_VOWEL Then
                If VowelSpecify(Left$(sW, 1)) = NONE_MARK_VOWEL Then
                    If UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "A" Then
                        If UCase$(Left$(UniToTelex(Left$(sW, 1)), 1)) = "O" Then
                            Pos = LPos
                        Else
                            PutToBuffer Ch
                            Exit Sub
                        End If
                    ElseIf UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "O" Then
                        Mid$(TotalBuffer, FPos, 1) = TelexToUni(Mid$(TotalBuffer, FPos, 1) & Ch)
                        Mid$(TotalBuffer, LPos, 1) = TelexToUni(Left$(UniToTelex(Mid$(TotalBuffer, LPos, 1)), 1) & Ch)
                        Pos = LPos - 1
                        UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
                        BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
                        Exit Sub
                    ElseIf UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "U" Then
                        PutToBuffer Ch
                        Exit Sub
                    End If
                ElseIf VowelSpecify(Left$(sW, 1)) = BREVE_MARK_VOWEL Then
                    If UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "A" Then
                        If UCase$(Left$(UniToTelex(Left$(sW, 1)), 1)) = "U" Then
                            Pos = LPos - 1
                        Else
                            PutToBuffer Ch
                            Exit Sub
                        End If
                    ElseIf UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "O" Then
                        If VowelSpecify(Right$(sW, 1)) = NONE_MARK_VOWEL Then
                            If VowelSpecify(Left$(sW, 1)) = NONE_MARK_VOWEL Then
                                Mid$(TotalBuffer, FPos, 1) = TelexToUni(Mid$(TotalBuffer, FPos, 1) & Ch)
                                Mid$(TotalBuffer, LPos, 1) = TelexToUni(Left$(UniToTelex(Mid$(TotalBuffer, LPos, 1)), 1) & Ch)
                                Pos = LPos - 1
                                UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
                                BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
                                Exit Sub
                            Else
                                Pos = LPos
                            End If
                        ElseIf VowelSpecify(Right$(sW, 1)) = BREVE_MARK_VOWEL Then
                            Mid$(TotalBuffer, LPos, 1) = Left$(UniToTelex(Right$(sW, 1)), 1)
                            Pos = LPos - 1
                        Else
                            Pos = LPos
                        End If
                    ElseIf UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "U" Then
                        PutToBuffer Ch
                        Exit Sub
                    End If
                ElseIf VowelSpecify(Left$(sW, 1)) = TONE_MARK_VOWEL Then
                    If UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "A" Then
                        PutToBuffer Ch
                        Exit Sub
                    ElseIf UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "O" Then
                        Mid$(TotalBuffer, FPos, 1) = TelexToUni(Mid$(TotalBuffer, FPos, 1) & Ch)
                        Mid$(TotalBuffer, LPos, 1) = TelexToUni(Left$(UniToTelex(Mid$(TotalBuffer, LPos, 1)), 1) & Ch)
                        Pos = LPos - 1
                        UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
                        BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
                        Exit Sub
                    ElseIf UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "U" Then
                        PutToBuffer Ch
                        Exit Sub
                    End If
                ElseIf VowelSpecify(Left$(sW, 1)) = TONE_AND_BREVE_MARK_VOWEL Then
                 If UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "A" Then
                        PutToBuffer Ch
                        Exit Sub
                    ElseIf UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "O" Then
                        Mid$(TotalBuffer, FPos, 1) = TelexToUni(Mid$(TotalBuffer, FPos, 1) & Ch)
                        Mid$(TotalBuffer, LPos, 1) = TelexToUni(Left$(UniToTelex(Mid$(TotalBuffer, LPos, 1)), 1) & Ch)
                        Pos = LPos - 1
                        UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
                        BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
                        Exit Sub
                    ElseIf UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "U" Then
                        PutToBuffer Ch
                        Exit Sub
                    End If
                End If
            ElseIf VowelSpecify(Right$(sW, 1)) = TONE_MARK_VOWEL Then
                If (UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "A") Then
                    If LPos < KeyPushed Then
                        If VowelSpecify(Mid$(TotalBuffer, LPos + 1, 1)) > 0 Then
                            PutToBuffer Ch
                            Exit Sub
                        End If
                        If InStr(1, "c,m,n,p,t", Mid$(TotalBuffer, LPos + 1, 1), vbTextCompare) <= 0 Then
                            PutToBuffer Ch
                            Exit Sub
                        End If
                    End If
                    Pos = LPos
                ElseIf (UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "O") Then
                    If UCase$(Left$(UniToTelex(Left$(sW, 1)), 1)) = "U" Then
                        If VowelSpecify(Left$(sW, 1)) = NONE_MARK_VOWEL Then
                            Mid$(TotalBuffer, FPos, 1) = TelexToUni(Mid$(TotalBuffer, FPos, 1) & Ch)
                            Mid$(TotalBuffer, LPos, 1) = TelexToUni(Left$(UniToTelex(Mid$(TotalBuffer, LPos, 1)), 1) & Ch & Right$(UniToTelex(Mid$(TotalBuffer, LPos, 1)), 1))
                            Pos = LPos - 1
                            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
                            BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
                            Exit Sub
                        ElseIf VowelSpecify(Left$(sW, 1)) = BREVE_MARK_VOWEL Then
                            Mid$(TotalBuffer, FPos, 1) = Left$(UniToTelex(Mid$(TotalBuffer, FPos, 1)), 1)
                            Mid$(TotalBuffer, LPos, 1) = TelexToUni(Left$(UniToTelex(Mid$(TotalBuffer, LPos, 1)), 1) & Ch & Right$(UniToTelex(Mid$(TotalBuffer, LPos, 1)), 1))
                            Pos = LPos - 1
                            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
                            BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
                            Exit Sub
                        Else
                            Pos = LPos
                        End If
                    Else
                        PutToBuffer Ch
                        Exit Sub
                    End If
                ElseIf (UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "U") Then
                    Pos = LPos
                End If
            ElseIf VowelSpecify(Right$(sW, 1)) = BREVE_MARK_VOWEL Then
                If UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "A" Then
                ElseIf UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "O" Then
                ElseIf UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "U" Then
                End If
            ElseIf VowelSpecify(Right$(sW, 1)) = TONE_AND_BREVE_MARK_VOWEL Then
                If UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "A" Then
                    If UCase$(Left$(UniToTelex(Left$(sW, 1)), 1)) = "O" Then
                        Pos = LPos
                    Else
                        PutToBuffer Ch
                        Exit Sub
                    End If
                ElseIf UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "O" Then
                    If UCase$(Mid$(UniToTelex(Right$(sW, 1)), 2, 1)) = UCase$(Ch) Then
                        If UCase$(Left$(UniToTelex(Left$(sW, 1)), 1)) = "U" Then
                            If VowelSpecify(Left$(sW, 1)) = NONE_MARK_VOWEL Then
                                Pos = LPos
                            ElseIf VowelSpecify(Left$(sW, 1)) = BREVE_MARK_VOWEL Then
                                Mid$(TotalBuffer, LPos, 1) = Left$(UniToTelex(Mid$(TotalBuffer, LPos, 1)), 1)
                                Pos = LPos - 1
                            Else
                                Pos = LPos
                            End If
                        Else
                            PutToBuffer Ch
                            Exit Sub
                        End If
                    Else
                        If UCase$(Left$(UniToTelex(Left$(sW, 1)), 1)) = "U" Then
                            If VowelSpecify(Left$(sW, 1)) = NONE_MARK_VOWEL Then
                                Mid$(TotalBuffer, FPos, 1) = TelexToUni(Mid$(TotalBuffer, FPos, 1) & Ch)
                                Mid$(TotalBuffer, LPos, 1) = TelexToUni(Left$(UniToTelex(Mid$(TotalBuffer, LPos, 1)), 1) & Ch & Right$(UniToTelex(Mid$(TotalBuffer, LPos, 1)), 1))
                                Pos = LPos - 1
                                UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
                                BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
                                Exit Sub
                            ElseIf VowelSpecify(Left$(sW, 1)) = BREVE_MARK_VOWEL Then
                                Mid$(TotalBuffer, FPos, 1) = Left$(UniToTelex(Mid$(TotalBuffer, FPos, 1)), 1)
                                Mid$(TotalBuffer, LPos, 1) = TelexToUni(Left$(UniToTelex(Mid$(TotalBuffer, LPos, 1)), 1) & Ch & Right$(UniToTelex(Mid$(TotalBuffer, LPos, 1)), 1))
                                Pos = LPos - 1
                                UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
                                BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
                                Exit Sub
                            Else
                                Pos = LPos
                            End If
                        Else
                            PutToBuffer Ch
                            Exit Sub
                        End If
                    End If
                ElseIf UCase$(Left$(UniToTelex(Right$(sW, 1)), 1)) = "U" Then
                    Pos = LPos
                End If
                
            End If
        Case 3:
            If (UCase$(Left$(UniToTelex(Left$(sW, 1)), 10))) = "U" And (UCase$(Left$(UniToTelex(Mid$(sW, 2, 1)), 10))) = "O" Then
                If InStr(1, "c,i,m,n,p,t,u", Right$(sW, 1), vbTextCompare) <= 0 Then
                    PutToBuffer Ch
                    Exit Sub
                End If
                LPos = LPos - 1
                GoTo JUMP
            Else
                PutToBuffer Ch
                Exit Sub
            End If
        Case Else
            Pos = LPos
    End Select
    
    UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
    BackNumbers = IIf(IsDoubleCharSet(CodeTable), LenX(UniBuf), Len(UniBuf))
    
    Dim sAnsi As String
    sAnsi = UniToTelex(Mid$(TotalBuffer, Pos, 1))
    
    If VowelSpecify(Mid$(TotalBuffer, Pos, 1)) = NONE_MARK_VOWEL Then
        If UCase$(Mid$(TotalBuffer, Pos, 1)) = "U" Then LastIsWConverted = False
        Mid$(TotalBuffer, Pos, 1) = TelexToUni(Mid$(TotalBuffer, Pos, 1) & Ch)
        UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
        Exit Sub
    ElseIf VowelSpecify(Mid$(TotalBuffer, Pos, 1)) = BREVE_MARK_VOWEL Then
        If UCase$(Right$(sAnsi, 1)) = UCase$(Ch) Then
            If (UCase$(Left$(sAnsi, 1)) = "U") And LastIsWConverted Then
                Mid$(TotalBuffer, Pos, 1) = Right$(sAnsi, 1)
            Else
                Mid$(TotalBuffer, Pos, 1) = Left$(sAnsi, 1)
                PutToBuffer Ch
            End If
            LastVietOff = KeyPushed
            VietKeyTempOff = True
            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
            Exit Sub
        Else
            Mid$(TotalBuffer, Pos, 1) = TelexToUni(Left$(sAnsi, 1) & Ch)
            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
            Exit Sub
        End If
    ElseIf VowelSpecify(Mid$(TotalBuffer, Pos, 1)) = TONE_MARK_VOWEL Then
        Mid$(TotalBuffer, Pos, 1) = TelexToUni(Left$(sAnsi, 1) & Ch & Right$(sAnsi, 1))
        UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
        Exit Sub
    ElseIf VowelSpecify(Mid$(TotalBuffer, Pos, 1)) = TONE_AND_BREVE_MARK_VOWEL Then
        If UCase$(Mid$(sAnsi, 2, 1)) = UCase(Ch) Then
            Mid$(TotalBuffer, Pos, 1) = TelexToUni(Left$(sAnsi, 1) & Right$(sAnsi, 1))
            PutToBuffer Ch
            LastVietOff = KeyPushed
            VietKeyTempOff = True
            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
            Exit Sub
        Else
            Mid$(TotalBuffer, Pos, 1) = TelexToUni(Left$(sAnsi, 1) & Ch & Right$(sAnsi, 1))
            UniBuf = Mid$(TotalBuffer, Pos, KeyPushed - Pos + 1)
            Exit Sub
        End If
    End If
    
End Sub


Public Sub ProcessShortkey(Ch As String)
    If Not VietNameseKeyboard Or VietKeyTempOff Then
        PutToBuffer Ch
        Exit Sub
    End If
    
    If TempOffShortkey Then
        PutToBuffer Ch
        Exit Sub
    End If
    
    If Ch <> "[" And Ch <> "{" And Ch <> "]" And Ch <> "}" Then
        PutToBuffer Ch
        Exit Sub
    End If
    
    If KeyPushed <= 0 Then
        If Ch = "]" Then
            PutToBuffer ChrW$(&H1A1)
        ElseIf Ch = "}" Then
            PutToBuffer ChrW$(&H1A0)
        ElseIf Ch = "[" Then
            PutToBuffer ChrW$(&H1B0)
            LastIsWConverted = False
        ElseIf Ch = "{" Then
            PutToBuffer ChrW$(&H1AF)
            LastIsWConverted = False
        End If
        UniBuf = Right$(TotalBuffer, 1)
        BackNumbers = LenX(UniBuf)
        LastIsShortkeyConverted = True
        Exit Sub
    End If
    
    Dim FPos As Integer, Pos As Integer, LPos As Integer, Founds As Boolean
    
    FPos = GetLastWord(TotalBuffer)
    If FPos < LastShortkeyOff Then FPos = LastShortkeyOff
    LPos = KeyPushed
    
    Do While LPos >= FPos
        If Mid$(TotalBuffer, LPos, 1) = Ch Then
            PutToBuffer Ch
            Exit Sub
        End If
        
        If Ch = "[" Or Ch = "{" Then
            If Mid$(TotalBuffer, LPos, 1) = ChrW$(&H1AF) Or Mid$(TotalBuffer, LPos, 1) = ChrW$(&H1B0) Then
                Founds = True
                Exit Do
            End If
        ElseIf Ch = "]" Or Ch = "}" Then
            If Mid$(TotalBuffer, LPos, 1) = ChrW$(&H1A0) Or Mid$(TotalBuffer, LPos, 1) = ChrW$(&H1A1) Then
                Founds = True
                Exit Do
            End If
        End If
        LPos = LPos - 1
    Loop
    
    If LPos < FPos Then LPos = FPos
    
    If Not Founds Then
        If InStr(1, IIf(Ch = "[" Or Ch = "{", "b,c,d,g,h,l,m,n,p,r,t,v,x", "b,c,d,g,h,l,m,n,p,r,t,v,x" & ChrW$(&H1AF) & ChrW$(&H1B0)), Right$(TotalBuffer, 1), vbTextCompare) <= 0 Then
            PutToBuffer Ch
            Exit Sub
        End If
        UniBuf = Right$(TotalBuffer, 1)
        BackNumbers = LenX(UniBuf)
        
        If Ch = "[" Then
            PutToBuffer ChrW$(&H1B0)
            LastIsWConverted = False
        ElseIf Ch = "{" Then
            PutToBuffer ChrW$(&H1AF)
            LastIsWConverted = False
        ElseIf Ch = "]" Then
            PutToBuffer ChrW$(&H1A1)
        ElseIf Ch = "}" Then
            PutToBuffer ChrW$(&H1A0)
        End If
        UniBuf = Right$(TotalBuffer, 2)
        LastIsShortkeyConverted = True
        Exit Sub
    End If
    
    If Not LastIsShortkeyConverted Then
        PutToBuffer Ch
        Exit Sub
    End If
    
    UniBuf = Mid$(TotalBuffer, LPos, KeyPushed - LPos + 1)
    BackNumbers = LenX(UniBuf)
    
    If Ch = "[" Or Ch = "{" Then
        Mid$(TotalBuffer, LPos, 1) = IIf(Mid$(TotalBuffer, LPos, 1) = UCase$(Mid$(TotalBuffer, LPos, 1)), "{", "[")
    ElseIf Ch = "]" Or Ch = "}" Then
        Mid$(TotalBuffer, LPos, 1) = IIf(Mid$(TotalBuffer, LPos, 1) = UCase$(Mid$(TotalBuffer, LPos, 1)), "}", "]")
    End If
    PutToBuffer Ch
    TempOffShortkey = True
    LastShortkeyOff = KeyPushed
    UniBuf = Mid$(TotalBuffer, LPos, KeyPushed - LPos + 1)
    Exit Sub
End Sub
