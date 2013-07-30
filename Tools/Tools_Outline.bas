Attribute VB_Name = "Tools_Outline"

Sub Outline()
Dim filelist() As String
Dim fName As String
Dim fPath As String
Dim i As Integer
Dim filetype  As String
Dname = Dimdoc(Dname)
    filetype = InputBox("What file type?" & _
    Chr(10) & "* = All", "File Type?", "doc")
    If StrPtr(filetype) = 0 Then Exit Sub
    '=======================================================
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Choose the folder that includes the documents"
            .Show
            If .SelectedItems.Count = 0 Then Exit Sub
            fPath = .SelectedItems(1) & "\"
        End With
    '=======================================================
    fName = Dir(fPath & "*." & filetype)
    While fName <> ""
        i = i + 1
        ReDim Preserve filelist(1 To i)
        filelist(i) = fName
        fName = Dir()
    Wend
    If i = 0 Then
        MsgBox "No " & T & " found in " & fPath & "       ", vbExclamation
        Exit Sub
    End If
    '=======================================================
    For i = 1 To UBound(filelist)
    ActiveWindow.ActivePane.View.Type = wdOutlineView
    On Error GoTo 1
    ChangeFileOpenDirectory _
        fPath
        If filelist(i) = Dname Then GoTo 7
        Selection.Range.Subdocuments.AddFromFile Name:=filelist(i), _
        ConfirmConversions:=True, ReadOnly:=False, PasswordDocument:="", _
        PasswordTemplate:="", Revert:=False, WritePasswordDocument:="", _
        WritePasswordTemplate:=""
7   Next
    WordBasic.ViewPageFromOutline
1   Selection.HomeKey Unit:=wdStory
End Sub
Function Dimdoc(Dname As Variant)
    On Error GoTo 1
    Dimdoc = ActiveDocument.Name
    GoTo 2
1   Documents.Add Template:="Normal", NewTemplate:=False, DocumentType:=0
2   Dimdoc = ActiveDocument.Name
End Function

