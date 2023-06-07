Attribute VB_Name = "VBA_AssetFoundInReport"
Sub AssetFoundInReport()
    
    'if asset is found in report sheet check for the values , if new value then add the new values and mention the old value in comments
    If Not AUCReportSheet.Cells(AssetSearch.Row, AssetCostCol_AUCReportSheet).Value = FARSheet.Cells(asset.Row, AssetCostCol_FARSheet).Value Then
    
        AUCReportSheet.Range("P" & AssetSearch.Row).Value = "Asset Value is changed Previous Value is " & AUCReportSheet.Cells(AssetSearch.Row, AssetCostCol_AUCReportSheet).Value
        AUCReportSheet.Cells(AssetSearch.Row, AssetCostCol_AUCReportSheet).Value = FARSheet.Cells(asset.Row, AssetCostCol_FARSheet).Value
        
    End If

End Sub
