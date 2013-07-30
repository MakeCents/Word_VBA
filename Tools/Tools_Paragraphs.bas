Attribute VB_Name = "Tools_Paragraphs"
Sub showlist()
    Dim i As Integer
    Dim s As String
    Dim lst As ListParagraphs
    s = ""
    For i = 1 To ActiveDocument.Lists.Count
        Set lst = ActiveDocument.Lists(i).Range.ListParagraphs
        s = s & "List " & Format$(i) & ": " & lst.Count & _
            " paragraphs" & vbCrLf
    Next i
    MsgBox s
    
End Sub
Sub Count()
    MsgBox ActiveDocument.Lists.Count
End Sub
