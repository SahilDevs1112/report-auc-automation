Attribute VB_Name = "VBA_SumForEachSegment"
Sub SumForEachSegment()

'    Dim startRow, endRow As Long
    
    'initializing the starting row for Auc report Sheet to do sum
'    startRow = 7
    
    'Initializing sum variables as 0
    TotalSum = 0
'    SegmentSum = 0
'
'    While AUCReportSheet.Range("A" & Rows.Count).Row <> endRow
'
'        'Get the end row for each segment
'        endRow = AUCReportSheet.Range("A" & startRow).End(xlDown).Row
'
'        'Do the sum for each segment in Asset cost column
'        SegmentSum = Application.WorksheetFunction.Sum(AUCReportSheet.Range(ColAlphabet(AssetCostCol_AUCReportSheet) & startRow & ":" & ColAlphabet(AssetCostCol_AUCReportSheet) & endRow))
'
'        AUCReportSheet.Cells(endRow + 1, AssetCostCol_AUCReportSheet).Value = SegmentSum
'
'        'Caluculate Total Sum for all Segments
'        TotalSum = TotalSum + SegmentSum
'
'        'Set startRow as start of the next segment
'        startRow = endRow + 2
'
'    Wend
'
'    'Place the Total Sum at the end of the sheet
'    AUCReportSheet.Cells(lastRow(AUCReportSheet, 1) + 2, AssetCostCol_AUCReportSheet).Value = TotalSum

    
    If Not GrandTotal Is Nothing Then
        
        GrandTotal.EntireRow.Delete
        
    End If
    
    AUCReportSheet.Range("A6:O" & lastRow(AUCReportSheet, 1) + 1).Subtotal GroupBy:=CostCenterCol_AUCReportSheet, Function:=xlSum, TotalList:=Array(AssetCostCol_AUCReportSheet), _
    Replace:=True, PageBreaks:=False, SummaryBelowData:=True
    
    TotalSum = AUCReportSheet.Cells(GrandTotal.Row, AssetCostCol_AUCReportSheet).Value
    
    'Calculate AUC FAR sheet Sum
    FARSheetSum = Application.WorksheetFunction.Sum(FARSheet.Range(ColAlphabet(AssetCostCol_FARSheet) & 2 & ":" & ColAlphabet(AssetCostCol_FARSheet) & lastRow(FARSheet, 1)))
    
    'Cross Check Total Sum with AUC FAR sheet Sum
    AUCReportSheet.Cells(GrandTotal.Row + 2, AssetCostCol_AUCReportSheet).Value = FARSheetSum
    
    AUCReportSheet.Cells(GrandTotal.Row + 4, AssetCostCol_AUCReportSheet).Value = TotalSum - FARSheetSum

End Sub
