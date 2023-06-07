Attribute VB_Name = "VBA_FetchDataFromDownloadWb"
Sub FetchDataFromDownloadWb()

    If DownloadSheet.AutoFilterMode = True Then
        DownloadSheet.AutoFilterMode = False
    End If
    
    'Delete Previous Data if any
    FARSheet.UsedRange.Clear
    
    DownloadSheet.UsedRange.AutoFilter field:=AssetCostCol_DownloadWb, Criteria1:="<>0"
    
    'copy asset data except Assets with asset cost 0
    DownloadSheet.UsedRange.SpecialCells(xlCellTypeVisible).Copy
    FARSheet.Range("A1").PasteSpecial xlPasteAll
    
    With FARSheet.UsedRange.Font
        .Name = "Calibri"
        .Size = 10
        .ThemeFont = xlThemeFontMinor
    End With
    ActiveWindow.Zoom = 80
    
End Sub
