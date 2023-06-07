Attribute VB_Name = "VBA_CopyAucReportSheet"
Sub CopyAucReportSheet()
    
    PostingDateSheet.UsedRange.Clear
    AUCReportSheet.UsedRange.Copy
    
    PostingDateSheet.Range("A1").PasteSpecial xlPasteAll
    
    PostingDateSheet.Cells(7, CapDateCol_AUCReportSheet).Value = "Posting Date"

End Sub
