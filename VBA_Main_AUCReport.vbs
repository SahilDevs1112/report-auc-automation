Attribute VB_Name = "VBA_Main_AUCReport"
Sub Main_AUCReport()
    
    Application.ScreenUpdating = False
    
    Call FetchDataFromDownloadWb
    
    Call MatchFARAndReport
    
    Call UpdateLocation
    
    Call SumForEachSegment
        
    Call FinalFormatting
    
    Call CopyAucReportSheet
    
    DownloadWb.Close False
    
    Application.ScreenUpdating = True

End Sub

