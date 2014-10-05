Attribute VB_Name = "ConcatMe"
Function ConcatMe(Rng As range, startDelim As String, endDelim As String) As String

Dim cl As range

   ConcatMe = ""

   For Each cl In Rng
      ConcatMe = ConcatMe & startDelim & cl.Text & endDelim
   Next cl

End Function
