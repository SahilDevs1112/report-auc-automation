Attribute VB_Name = "VBA_FinalFormatting"
Sub FinalFormatting()

    AUCReportSheet.Range("B2").Value = Date
    AUCReportSheet.Range("I5").Value = PD & " " & FY
    AUCReportSheet.Range("A3").Value = PD & " " & FY
    
    ActiveWindow.Zoom = 80
    With AUCReportSheet.UsedRange.Font
        .Name = "Calibri"
        .Size = 10
        .ThemeFont = xlThemeFontMinor
        
    End With
    
    With AUCReportSheet.UsedRange
        .Rows.AutoFit
        .Columns.AutoFit
    End With
    
    
    'MainWb.SaveAs MainWb.Path & "/USI AUC Report " & PD & " " & FY & ".xlsm"
    
End Sub
