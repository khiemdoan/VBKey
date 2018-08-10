Attribute VB_Name = "mdlTools"
Option Explicit

Public Sub DelFileExecuting(Filename As String)
        On Error Resume Next
        Dim BatFile As String
        BatFile = MakeRandomFileName("bat")
        Dim Fn As Integer
        Fn = FreeFile
        Open BatFile For Output Lock Write As #Fn
            Print #1, ":try" & vbCrLf & "del " & Filename & vbCrLf & "if exist " & Filename & "  goto try" & vbCrLf & "del " & BatFile
        Close #1
        Shell BatFile, vbHide
End Sub

Private Function MakeRandomFileName(Extend As String) As String
      Dim Arr(), ArrTemp() As String
      Dim I As Long
      Arr = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "o", "p", "q", "r", "s", "t", "u", "v", "x", "y", "z", "w")
LoopAgain:
      Do
          Randomize
          ReDim ArrTemp(Int(Rnd * 100 + 100))
          For I = 0 To UBound(ArrTemp)
              ArrTemp(I) = Arr(Rnd * 24)
          Next I
          
          For I = 0 To UBound(ArrTemp)
              MakeRandomFileName = MakeRandomFileName & ArrTemp(I)
          Next I
          MakeRandomFileName = MakeRandomFileName & IIf(Left$(Extend, 1) = ".", Extend, "." & Extend)
      Loop Until MakeRandomFileName <> "" And Dir$(MakeRandomFileName) = ""
      
End Function

