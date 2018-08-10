Attribute VB_Name = "mdlStringConvert"
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



Public Function VniToUni(S As String) As String
    Dim I As Long, J As Long, sResult As String
    sResult = S
    
    For I = UBound(UNICODE_PRECOMPOSED_TABLE) To LBound(UNICODE_PRECOMPOSED_TABLE) Step -1
        For J = 1 To Len(S)
            If (LCase$(VniArr(I)) = LCase$(Mid$(S, J, 3))) And (Mid$(S, J, 1) = Left$(VniArr(I), 1)) Then sResult = Replace$(sResult, Mid$(S, J, 3), IIf(Mid$(S, J, 1) = UCase$(Mid$(S, J, 1)), UCase$(UNICODE_PRECOMPOSED_TABLE(I)), UNICODE_PRECOMPOSED_TABLE(I)))
        Next J
        
        For J = 1 To Len(S)
            If (LCase$(VniArr(I)) = LCase$(Mid$(S, J, 2))) And (Mid$(S, J, 1) = Left$(VniArr(I), 1)) Then sResult = Replace$(sResult, Mid$(S, J, 2), IIf(Mid$(S, J, 1) = UCase$(Mid$(S, J, 1)), UCase$(UNICODE_PRECOMPOSED_TABLE(I)), UNICODE_PRECOMPOSED_TABLE(I)))
        Next J
    Next I
    
    VniToUni = sResult
    
End Function

Public Function UniToVni(S As String) As String
    Dim I As Long, J As Long, sResult As String
    sResult = S
    
    For I = UBound(UNICODE_PRECOMPOSED_TABLE) To LBound(UNICODE_PRECOMPOSED_TABLE) Step -1
        For J = 1 To Len(S)
            If UNICODE_PRECOMPOSED_TABLE(I) = Mid$(S, J, 1) Then sResult = Replace$(sResult, Mid$(S, J, 1), IIf(Mid$(S, J, 1) = UCase$(Mid$(S, J, 1)), UCase$(VniArr(I)), VniArr(I)))
        Next J
    Next I
    
    UniToVni = sResult
    
End Function




Public Function ViqrToUni(S As String) As String
    Dim I As Long, J As Long, sResult As String
    sResult = S
    
    For I = UBound(UNICODE_PRECOMPOSED_TABLE) To LBound(UNICODE_PRECOMPOSED_TABLE) Step -1
        For J = 1 To Len(S)
            If (LCase$(ViqrArr(I)) = LCase$(Mid$(S, J, 3))) And (Mid$(S, J, 1) = Left$(ViqrArr(I), 1)) Then sResult = Replace$(sResult, Mid$(S, J, 3), UNICODE_PRECOMPOSED_TABLE(I))
        Next J
        
        For J = 1 To Len(S)
            If (LCase$(ViqrArr(I)) = LCase$(Mid$(S, J, 2))) And (Mid$(S, J, 1) = Left$(ViqrArr(I), 1)) Then sResult = Replace$(sResult, Mid$(S, J, 2), UNICODE_PRECOMPOSED_TABLE(I))
        Next J
    Next I
    
    ViqrToUni = sResult
    
End Function


Public Function UniToViqr(S As String) As String
    Dim I As Long, J As Long, sResult As String
    sResult = S
    
    For I = UBound(UNICODE_PRECOMPOSED_TABLE) To LBound(UNICODE_PRECOMPOSED_TABLE) Step -1
        For J = 1 To Len(S)
            If UNICODE_PRECOMPOSED_TABLE(I) = Mid$(S, J, 1) Then sResult = Replace$(sResult, Mid$(S, J, 1), ViqrArr(I))
        Next J
    Next I
    
    UniToViqr = sResult
    
End Function

