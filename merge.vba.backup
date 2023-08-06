Sub 打印()

    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.75)
        .RightMargin = Application.InchesToPoints(0.75)
        .TopMargin = Application.InchesToPoints(1)
        .BottomMargin = Application.InchesToPoints(1)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintErrors = xlPrintErrorsDisplayed
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .LeftHeader = "----------------------------------"
        .CenterHeader = "华东大区工作周报汇总"
        .RightHeader = "----------------------------------"
        .CenterFooter = "第 &P 页，共 &N 页"
        .CenterHorizontally = True
    End With

    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:="520GIFTForL2023", Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True

End Sub
