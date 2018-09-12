Function Concat1(myRange As Range, Optional myDelimiter As String)
    Dim r As Range
 
    Application.Volatile
    For Each r In myRange
        Concat1 = Concat1 & r & myDelimiter
    Next r
    If Len(myDelimiter) > 0 Then
        Concat1 = Left(Concat1, Len(Concat1) - Len(myDelimiter))
    End If
End Function
 
Function Concat2(myRange As Range, Optional myDelimiter As String)
    Dim r As Range
 
    Application.Volatile
    For Each r In myRange
        If Len(r.Text) > 0 Then
            Concat2 = Concat2 & r & myDelimiter
        End If
    Next r
    If Len(myDelimiter) > 0 Then
        Concat2 = Left(Concat2, Len(Concat2) - Len(myDelimiter))
    End If
End Function