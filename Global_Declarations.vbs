Attribute VB_Name = "Global_Declarations"
Public AssetSearch, AssetSearch2 As Range
Public CostCenterSearch_HiddenSheet As Range
Public asset, asset2 As Range
Public CostCenterSearch As Range
Public TotalSum, SegmentSum, FARSheetSum As Double
Public selectedWb As Workbook
Public counter As Integer
Public costCenter As Range
Public DivisionFormatRng As Range

Public Property Get MainWb() As Workbook

    Set MainWb = ThisWorkbook

End Property

Public Property Get DownloadWb() As Workbook

    On Error Resume Next
    Set DownloadWb = Workbooks.Open(MainWb.Path & "\AUC FAR.XLSX")
    
    If Err Then
        
        If counter = 0 Then
        
            MsgBox "File was not found please select the Downloaded File", vbOKOnly, "Select AUC FAR"
            Dim fd As FileDialog
            Set fd = Application.FileDialog(msoFileDialogFilePicker)
            With fd
                .AllowMultiSelect = False
                .Filters.Add "Excel File", "*.xlsx"
            End With
    
            If fd.Show = -1 Then
    
                Set DownloadWb = Workbooks.Open(fd.SelectedItems(1))
                Set selectedWb = DownloadWb
            Else
    
                MsgBox "No files Selected", vbInformation, "Select a File"
                End
    
            End If
    
            counter = counter + 1
    
        Else
    
            Set DownloadWb = selectedWb
    
        End If
    
    End If

End Property

Public Property Get PD() As String

    PD = Trim(AUCReportSheet.Range("B1").Value)

End Property

Public Property Get FY() As String

    FY = Trim(AUCReportSheet.Range("D1").Value)

End Property

Public Property Get AUCReportSheet() As Worksheet

    Set AUCReportSheet = MainWb.Sheets(1)

End Property

Public Property Get DownloadSheet() As Worksheet
    
    Set DownloadSheet = DownloadWb.Sheets(1)

End Property

Public Property Get FARSheet() As Worksheet
    
    Set FARSheet = MainWb.Sheets("AUC FAR")

End Property

Public Property Get AssetCostCol_DownloadWb() As Integer

    AssetCostCol_DownloadWb = DownloadSheet.Range("A1:BP1").Find(what:="Asset Cost", LookIn:=xlValues, lookat:=xlWhole).Column

End Property

Public Property Get AssetCostCol_FARSheet() As Integer

    AssetCostCol_FARSheet = FARSheet.Range("A1:Z1").Find(what:="Asset Cost", LookIn:=xlValues, lookat:=xlWhole).Column

End Property

Public Property Get AssetCostCol_AUCReportSheet() As Integer

    AssetCostCol_AUCReportSheet = AUCReportSheet.Range("A6:Z6").Find(what:="Asset Cost", LookIn:=xlValues, lookat:=xlPart).Column

End Property

Public Property Get AssetNoCol_FARSheet() As Integer

    AssetNoCol_FARSheet = FARSheet.Range("A1:Z1").Find(what:="Asset Number", LookIn:=xlValues, lookat:=xlWhole).Column

End Property

Public Property Get AssetNoCol_AUCReportSheet() As Integer

    AssetNoCol_AUCReportSheet = AUCReportSheet.Range("A6:Z6").Find(what:="Asset No (S4)", LookIn:=xlValues, lookat:=xlWhole).Column

End Property

Public Property Get AssetClassCol_FARSheet() As Integer

    AssetClassCol_FARSheet = FARSheet.Range("A1:Z1").Find(what:="Asset Class", LookIn:=xlValues, lookat:=xlWhole).Column

End Property

Public Property Get AssetClassCol_AUCReportSheet() As Integer

    AssetClassCol_AUCReportSheet = AUCReportSheet.Range("A6:Z6").Find(what:="Class", LookIn:=xlValues, lookat:=xlWhole).Column

End Property

Public Property Get CoCodeCol_FARSheet() As Integer

    CoCodeCol_FARSheet = FARSheet.Range("A1:Z1").Find(what:="Company Code", LookIn:=xlValues, lookat:=xlWhole).Column

End Property

Public Property Get CoCodeCol_AUCReportSheet() As Integer

    CoCodeCol_AUCReportSheet = AUCReportSheet.Range("A6:Z6").Find(what:="CoCd", LookIn:=xlValues, lookat:=xlWhole).Column

End Property

Public Property Get CostCenterCol_FARSheet() As Integer

    CostCenterCol_FARSheet = FARSheet.Range("A1:Z1").Find(what:="Cost Center", LookIn:=xlValues, lookat:=xlWhole).Column

End Property

Public Property Get CostCenterCol_AUCReportSheet() As Integer

    CostCenterCol_AUCReportSheet = AUCReportSheet.Range("A6:Z6").Find(what:="Cost Ctr", LookIn:=xlValues, lookat:=xlWhole).Column

End Property

Public Property Get CapDateCol_FARSheet() As Integer

    CapDateCol_FARSheet = FARSheet.Range("A1:Z1").Find(what:="Capitalization date", LookIn:=xlValues, lookat:=xlWhole).Column

End Property

Public Property Get CapDateCol_AUCReportSheet() As Integer

    CapDateCol_AUCReportSheet = AUCReportSheet.Range("A6:Z6").Find(what:="Cap Date", LookIn:=xlValues, lookat:=xlWhole).Column

End Property

Public Property Get AssetDescCol_FARSheet() As Integer

    AssetDescCol_FARSheet = FARSheet.Range("A1:Z1").Find(what:="Description 1", LookIn:=xlValues, lookat:=xlWhole).Column

End Property

Public Property Get AssetDescCol_AUCReportSheet() As Integer

    AssetDescCol_AUCReportSheet = AUCReportSheet.Range("A6:Z6").Find(what:="Asset description", LookIn:=xlValues, lookat:=xlWhole).Column

End Property

Public Property Get WBSCol_FARSheet() As Integer

    WBSCol_FARSheet = FARSheet.Range("A1:Z1").Find(what:="Capital WBS Element", LookIn:=xlValues, lookat:=xlWhole).Column

End Property

Public Property Get WBSCol_AUCReportSheet() As Integer

    WBSCol_AUCReportSheet = AUCReportSheet.Range("A6:Z6").Find(what:="Budget based WBS", LookIn:=xlValues, lookat:=xlWhole).Column

End Property

Public Property Get AssetLocCol_AUCReportSheet() As Integer

    AssetLocCol_AUCReportSheet = AUCReportSheet.Range("A6:Z6").Find(what:="Loc", LookIn:=xlValues, lookat:=xlWhole).Column

End Property

Public Property Get HiddenSheet() As Worksheet

    Set HiddenSheet = MainWb.Sheets("Hidden")

End Property

Public Property Get GrandTotal() As Range
    
    On Error Resume Next
    Set GrandTotal = AUCReportSheet.UsedRange.Find(what:="Grand Total", LookIn:=xlValues, lookat:=xlPart)

End Property

Public Function lastRow(ws As Worksheet, col As Integer) As Long

    lastRow = ws.Cells(Rows.Count, col).End(xlUp).Row

End Function

Public Function ColAlphabet(col As Long) As String

    ColAlphabet = Chr(col + 64)

End Function

Public Property Get PostingDateSheet() As Worksheet
    
    Set PostingDateSheet = MainWb.Sheets("AUC Posting date wise Breakup")

End Property

Public Function DivisionRowFormatting(rng As Range)

    With rng.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    rng.Borders(xlInsideHorizontal).LineStyle = xlNone


End Function

