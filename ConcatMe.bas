Attribute VB_Name = "ConcatMe"

Function ConcatMe(Rng As Range, startDelim As String, endDelim As String, Optional cutString As Boolean) As String

Dim cl As Range

   ConcatMe = ""
If cutString = True Then
   For Each cl In Rng
      ConcatMe = ConcatMe & startDelim & cl.Text & endDelim
   Next cl
   
    If endDelim = "" Then
        ConcatMe = ConcatMe
    Else
    ConcatMe = Left(ConcatMe, Len(ConcatMe) - 1)
    End If
Else
  For Each cl In Rng
      ConcatMe = ConcatMe & startDelim & cl.Text & endDelim
   Next cl
End If
End Function

