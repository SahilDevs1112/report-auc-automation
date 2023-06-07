Attribute VB_Name = "VBA_AssetNotFoundInReport"
Sub AssetNotFoundInReport()

    On Error Resume Next
    Set CostCenterSearch = AUCReportSheet.Range(ColAlphabet(CostCenterCol_AUCReportSheet) & 7 & ":" & ColAlphabet(CostCenterCol_AUCReportSheet) & lastRow(AUCReportSheet, CostCenterCol_AUCReportSheet)).Find(what:= _
    FARSheet.Cells(asset.Row, CostCenterCol_FARSheet).Value, LookIn:=xlValues, lookat:=xlPart)
    
    If Not CostCenterSearch Is Nothing Then
        
        Dim CostCenterFoundRow As Long
        CostCenterFoundRow = CostCenterSearch.Row
        
        CostCenterSearch.EntireRow.Insert
        
        'copy new asset data into new row
        AUCReportSheet.Cells(CostCenterFoundRow, AssetNoCol_AUCReportSheet).Value = FARSheet.Cells(asset.Row, AssetNoCol_FARSheet).Value
        AUCReportSheet.Cells(CostCenterFoundRow, AssetClassCol_AUCReportSheet).Value = FARSheet.Cells(asset.Row, AssetClassCol_FARSheet).Value
        AUCReportSheet.Cells(CostCenterFoundRow, CoCodeCol_AUCReportSheet).Value = FARSheet.Cells(asset.Row, CoCodeCol_FARSheet).Value
        AUCReportSheet.Cells(CostCenterFoundRow, CostCenterCol_AUCReportSheet).Value = FARSheet.Cells(asset.Row, CostCenterCol_FARSheet).Value
        AUCReportSheet.Cells(CostCenterFoundRow, CapDateCol_AUCReportSheet).Value = FARSheet.Cells(asset.Row, CapDateCol_FARSheet).Value
        AUCReportSheet.Cells(CostCenterFoundRow, AssetDescCol_AUCReportSheet).Value = FARSheet.Cells(asset.Row, AssetDescCol_FARSheet).Value
        AUCReportSheet.Cells(CostCenterFoundRow, WBSCol_AUCReportSheet).Value = FARSheet.Cells(asset.Row, WBSCol_FARSheet).Value
        AUCReportSheet.Cells(CostCenterFoundRow, AssetCostCol_AUCReportSheet).Value = FARSheet.Cells(asset.Row, AssetCostCol_FARSheet).Value
        
        With AUCReportSheet.Range("A" & CostCenterFoundRow & ":O" & CostCenterFoundRow)
            .Interior.ColorIndex = 15
        End With
          
    Else
    
        'If Cost center is not found then create a new Location segment at the end of the sheet
        AUCReportSheet.Range("A", lastRow(AUCReportSheet, 1) + 2).EntireRow.Insert
        AUCReportSheet.Range("A", lastRow(AUCReportSheet, 1) + 3).EntireRow.Insert
        
        Dim newRow As Long
        newRow = lastRow(AUCReportSheet, 1) + 2
        
        'Copy new Asset Data into new row
        AUCReportSheet.Cells(newRow, AssetNoCol_AUCReportSheet).Value = FARSheet.Cells(asset.Row, AssetNoCol_FARSheet).Value
        AUCReportSheet.Cells(newRow, AssetClassCol_AUCReportSheet).Value = FARSheet.Cells(asset.Row, AssetClassCol_FARSheet).Value
        AUCReportSheet.Cells(newRow, CoCodeCol_AUCReportSheet).Value = FARSheet.Cells(asset.Row, CoCodeCol_FARSheet).Value
        AUCReportSheet.Cells(newRow, CostCenterCol_AUCReportSheet).Value = FARSheet.Cells(asset.Row, CostCenterCol_FARSheet).Value
        AUCReportSheet.Cells(newRow, CapDateCol_AUCReportSheet).Value = FARSheet.Cells(asset.Row, CapDateCol_FARSheet).Value
        AUCReportSheet.Cells(newRow, AssetDescCol_AUCReportSheet).Value = FARSheet.Cells(asset.Row, AssetDescCol_FARSheet).Value
        AUCReportSheet.Cells(newRow, WBSCol_AUCReportSheet).Value = FARSheet.Cells(asset.Row, WBSCol_FARSheet).Value
        AUCReportSheet.Cells(newRow, AssetCostCol_AUCReportSheet).Value = FARSheet.Cells(asset.Row, AssetCostCol_FARSheet).Value
        
        'format Division Row after new location is added
        
        Set DivisionFormatRng = AUCReportSheet.Range("A" & newRow + 1 & ":O" & newRow + 1)
        DivisionRowFormatting (DivisionFormatRng)
    
    End If
               
End Sub
