Attribute VB_Name = "VBA_MatchFARAndReport"
Sub MatchFARAndReport()

    For Each asset In FARSheet.Range(ColAlphabet(AssetNoCol_FARSheet) & 2 & ":" & ColAlphabet(AssetNoCol_FARSheet) & lastRow(FARSheet, AssetNoCol_FARSheet))
        
        On Error Resume Next
        Set AssetSearch = AUCReportSheet.Range(ColAlphabet(AssetNoCol_AUCReportSheet) & 7 & ":" & ColAlphabet(AssetNoCol_AUCReportSheet) & lastRow(AUCReportSheet, 1)).Find(what:=asset.Value, LookIn:=xlValues, lookat:=xlPart)
        
        If Not AssetSearch Is Nothing Then
        
            Call AssetFoundInReport
        
        Else
        
            Call AssetNotFoundInReport
        
        End If
    
    Next asset
    
    For Each asset2 In AUCReportSheet.Range(ColAlphabet(AssetNoCol_AUCReportSheet) & 7 & ":" & ColAlphabet(AssetNoCol_AUCReportSheet) & lastRow(AUCReportSheet, 1))
        
        If asset2.Value <> Empty Then
        
            On Error Resume Next
            Set AssetSearch2 = FARSheet.Range(ColAlphabet(AssetNoCol_FARSheet) & 2 & ":" & ColAlphabet(AssetNoCol_FARSheet) & lastRow(FARSheet, 1)).Find(what:=asset2.Value, LookIn:=xlValues, lookat:=xlPart)
            
            If AssetSearch2 Is Nothing Then
            
                AUCReportSheet.Range("P" & aseet2.Row).Value = "This Asset is not found in FAR"
            
            End If
        
        End If
    
    Next asset2

End Sub
