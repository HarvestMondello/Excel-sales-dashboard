Sub SaveReportAsPDF()
    Dim ws As Worksheet
    Dim filePath As String
    Dim fileName As String

    Set ws = ThisWorkbook.Sheets("Dashboard")

    ' Set the file path and file name for saving the PDF
    filePath = ThisWorkbook.Path
    fileName = filePath & "\" & ws.Name & "_Report_" & Format(Now(), "yyyymmdd_HHMMSS") & ".pdf"

    With ws
        ' Manually set the print area (adjust the range as needed)
        .PageSetup.PrintArea = "A1:Z40"  ' Change this to the range you want to print

        ' Force Landscape Orientation
        .PageSetup.Orientation = xlLandscape

        ' Ensure fitting to one page
        .PageSetup.Zoom = False
        .PageSetup.FitToPagesWide = 1
        .PageSetup.FitToPagesTall = 1

        ' Adjust margins for better fit
        .PageSetup.LeftMargin = Application.InchesToPoints(0.25)
        .PageSetup.RightMargin = Application.InchesToPoints(0.25)
        .PageSetup.TopMargin = Application.InchesToPoints(0.5)
        .PageSetup.BottomMargin = Application.InchesToPoints(0.5)

        ' Export as PDF
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName, Quality:=xlQualityStandard, _
                             IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    End With
End Sub


