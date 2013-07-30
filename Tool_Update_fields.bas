Attribute VB_Name = "Tool_Update_fields"
Sub Flds()
    Application.ScreenUpdating = False
    Dim sec As Section
    ActiveDocument.Fields.Update
    For Each sec In ActiveDocument.Sections
        sec.Headers(wdHeaderFooterPrimary).Range.Fields.Update
        sec.Headers(wdHeaderFooterFirstPage).Range.Fields.Update
        sec.Footers(wdHeaderFooterPrimary).Range.Fields.Update
        sec.Footers(wdHeaderFooterFirstPage).Range.Fields.Update
    Next
    
    Dim oField As Field
    Dim oSection As Section
    Dim oHeader As HeaderFooter
    Dim oFooter As HeaderFooter
     For Each oSection In ActiveDocument.Sections
     For Each oHeader In oSection.Headers
         If oHeader.Exists Then
             For Each oField In oHeader.Range.Fields
                 oField.Update
             Next oField
         End If
     Next oHeader
     For Each oFooter In oSection.Footers
         If oFooter.Exists Then
              For Each oField In oFooter.Range.Fields
                  oField.Update
             Next oField
         End If
     Next oFooter
 Next oSection
    
End Sub

