Attribute VB_Name = "Tools_NamingFiles"
Sub rename()
    Dim filelist As New Collection
    Dim fName As String
    Dim i As Integer
    Dim filetype  As String
'=======================================================================
    'Message box to ask the file type
    filetype = "doc"
'=======================================================================
    'You will get an error when you hit cancel
    On Error GoTo 3
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Show
            fPath = .SelectedItems(1) & "\"
     End With
'=======================================================================
    fName = Dir(fPath & "*." & filetype)
    While fName <> ""
        i = i + 1
        filelist.Add Item:=fName
        fName = Dir()
    Wend
    If i = 0 Then
        MsgBox "No " & filetype & "files found in " & fPath & "       ", vbExclamation
        Exit Sub
    End If
    For T = 1 To i
        namefile = fPath & filelist(T)
            Documents.Open FileName:=namefile, ConfirmConversions:=True, ReadOnly _
            :=False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate _
            :="", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="" _
            , Format:=wdOpenFormatAuto, XMLTransform:=""
        
            Selection.MoveRight Unit:=wdCharacter, Count:=4, Extend:=wdExtend
            N = Left(Selection, 4)
        newname = fPath & N & "." & filetype
        ActiveDocument.Close
        Name namefile As newname
    Next T
3
End Sub
Sub example()
    x = "and"
    MsgBox "This " & x & " That"
End Sub

'=======================================================================
'This imports the doc files specified from a single folder and renames them the the first four digits it sees.
'=======================================================================
Sub Import_File_Names()
Dim filelist() As String
Dim fName As String
Dim fPath As String
Dim i As Integer
Dim startrow As Integer
Dim filetype  As String

     'Makes sure you want to clear this sheet
    
    T = InputBox("What file type?" & _
    Chr(10) & "* = All" & _
    Chr(10) & "xls" & _
    Chr(10) & "doc" & _
    Chr(10) & "sgm" & _
    Chr(10) & "xlsx" & _
    Chr(10) & "txt", "File Type?", "doc")
    If StrPtr(T) = 0 Then
        GoTo 3
    End If
   
    '=======================================================
    'Sets sheet name
        'On Error GoTo 3
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Show
            fPath = .SelectedItems(1) & "\"
        End With
    
1
    filetype = T
    
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

    For i = 1 To UBound(filelist)
    
       namefile = fPath & filelist(i)
        Documents.Open FileName:=filelist(i), ConfirmConversions:=True, ReadOnly _
        :=False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate _
        :="", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="" _
        , Format:=wdOpenFormatAuto, XMLTransform:=""
    
        Selection.MoveRight Unit:=wdCharacter, Count:=4, Extend:=wdExtend
        N = Left(Selection, 4)
    
        oldname = filelist(i)
        newname = fPath & N & "." & filetype
        ActiveDocument.Close
    Name oldname As newname
    
    Next
    MsgBox "Done"
3
    
End Sub

