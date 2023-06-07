Attribute VB_Name = "VBA_UpdateLocation"
Sub UpdateLocation()
    
    For Each costCenter In AUCReportSheet.Range(ColAlphabet(CostCenterCol_AUCReportSheet) & 7 & ":" & ColAlphabet(CostCenterCol_AUCReportSheet) & GrandTotal.Row - 1)
        
        If costCenter.Value <> Empty Then
        'Give location as per Cost Center
            On Error Resume Next
            Set CostCenterSearch_HiddenSheet = HiddenSheet.Range("F:F").Find(what:=costCenter.Value, LookIn:=xlValues, lookat:=xlPart)
            
            If Not CostCenterSearch_HiddenSheet Is Nothing Then
            
                AUCReportSheet.Cells(costCenter.Row, AssetLocCol_AUCReportSheet).Value = HiddenSheet.Range("G" & CostCenterSearch_HiddenSheet.Row).Value
            
            Else
                
                'If the cost center is not found then add that cost center to the list and add a formula to get the location as and when the user gives the value in Loc column
                HiddenSheet.Range("F" & lastRow(HiddenSheet, 6) + 1).Value = AUCReportSheet.Cells(costCenter.Row, CostCenterCol_AUCReportSheet).Value
                HiddenSheet.Range("G" & lastRow(HiddenSheet, 6)).Formula = "='" & AUCReportSheet.Name & "'!" & ColAlphabet(AssetLocCol_AUCReportSheet) & costCenter.Row
            
            End If
            
        End If
        
    Next costCenter
        
End Sub
