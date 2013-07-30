Attribute VB_Name = "Tools_Write_xml_files"
Function cell_contents()
    
    s = Selection
    T = InStr(1, s, "…")
    If T = 0 Then T = InStr(5, s, "..")
        If T = 0 Then T = Len(s)
    cell_contents = Mid(UCase(s), 1, T - 1)
End Function
Sub RPSTL_table_tagger()
Locate
With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "What folder would you like to save the text file in?"
        .InitialFileName = ActiveDocument.Path & "\"
        .Show
        If .SelectedItems.Count = 0 Then GoTo 1
        fdlr = .SelectedItems(1)
    End With
    If StrPtr(E) = 0 Then GoTo 1
    N = ActiveDocument.Name
    N = Mid(N, 1, InStr(1, N, ".") - 1)
    file = fdlr & "\" & N & ".txt"


Open file For Output Access Write As #1
current = 1

'This is for the header
Print #1, "<thead>"
Print #1, "<row>"
Print #1, "<entry align=""center"" rowsep=""0"" valign=""top"">(1)</entry>"
Print #1, "<entry align=""center"" rowsep=""0"" valign=""top"">(2)</entry>"
Print #1, "<entry align=""center"" rowsep=""0"" valign=""top"">(3)</entry>"
Print #1, "<entry align=""center"" rowsep=""0"" valign=""top"">(4)</entry>"
Print #1, "<entry align=""center"" rowsep=""0"" valign=""top"">(5)</entry>"
Print #1, "<entry align=""center"" rowsep=""0"" valign=""top"">(6)</entry>"
Print #1, "<entry align=""center"" rowsep=""0"" valign=""top"">(7)</entry>"
Print #1, "</row>"
Print #1, "<row>"
Print #1, "<entry align=""center"">ITEM NO.</entry>"
Print #1, "<entry align=""center"">SMR CODE</entry>"
Print #1, "<entry align=""center"">NSN</entry>"
Print #1, "<entry align=""center"">CAGE CODE</entry>"
Print #1, "<entry align=""center"">PART NUMBER</entry>"
Print #1, "<entry align=""center"">DESCRIPTION AND USABLE ON CODE (UOC)</entry>"
Print #1, "<entry align=""center"">QTY</entry>"
Print #1, "</row>"
Print #1, "</thead>"


'This is for the body
Print #1, "<tbody>"

Do While InStr(1, UCase(Selection), "END OF FIGURE") = 0
    Selection.Expand wdCell
   
    If Len(Selection) <= "2" Then
        b = ""
        bc = ""
    Else:
        If Selection.font.Bold = True Then
            b = "<p><b>"
            bc = "</b></p>"
        Else:
            b = "<p>"
            bc = "</p>"
        End If
    End If
    If current = 1 Then
    
        Print #1, "<row>"
        Print #1, "<?PubTbl row rht=""0.34in""?>"
        Print #1, "<entry colsep=""0"" rowsep=""0"">" & b & cell_contents & bc & "</entry>"
    ElseIf current > 6 Then
        Print #1, "<entry rowsep=""0"">" & b & cell_contents & bc & "</entry>"
        Print #1, "</row>"
        current = 0
    Else:
        Print #1, "<entry colsep=""0"" rowsep=""0"">" & b & cell_contents & bc & "</entry>"
    End If
    current = current + 1
    Selection.MoveRight Unit:=wdCell
Loop
    Selection.Expand wdCell
    Print #1, "<entry colsep=""0"" rowsep=""0""><b>END OF FIGURE</b></entry>"
    
    Selection.MoveRight Unit:=wdCell
    Selection.Expand wdCell
    Print #1, "<entry valign=""bottom""></entry>"
    Print #1, "</row>"
    Print #1, "</tbody>"
    Print #1, "</tgroup>"
Print #1, "</table></p>"
Print #1, "</conbody>"
Print #1, "</concept>"
Close #1
1
End Sub

Sub Locate()
    x = 0
    Selection.HomeKey Unit:=wdStory
1   Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "(1)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.Expand wdCell
    If InStr(1, Selection, "ITEM") Then
        Selection.MoveDown Unit:=wdLine, Count:=1
        Exit Sub
    Else:
        x = x + 1
        If x > 10 Then
            Selection.MoveDown Unit:=wdLine, Count:=1
            Selection.Expand wdCell
            If InStr(1, Selection, "ITEM") Then
                Selection.MoveDown Unit:=wdLine, Count:=1
                Exit Sub
            End If
        End If
        GoTo 1
    End If
End Sub

