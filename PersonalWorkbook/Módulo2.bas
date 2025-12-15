Attribute VB_Name = "Módulo2"
Option Explicit

Sub PudimAmassado()
Attribute PudimAmassado.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PudimAmassado Macro
'

'
    Range("B21:B29").Select
    Selection.ClearContents
    Range("D21:D29").Select
    Selection.Copy
    Range("B21").Select
    ActiveSheet.Paste
    Range("D21:D29").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("F21:F29").Select
    Selection.Copy
    Range("D21").Select
    ActiveSheet.Paste
    Range("F21:F29").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("H21:H29").Select
    Selection.Copy
    Application.CutCopyMode = False
End Sub
Sub GigaPombo()
Attribute GigaPombo.VB_ProcData.VB_Invoke_Func = " \n14"
'
' GigaPombo Macro
'

'
    Range("A1:K51").Select
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    ActiveSheet.PageSetup.PrintArea = "$A$1:$K$51"
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.511811023622047)
        .RightMargin = Application.InchesToPoints(0.511811023622047)
        .TopMargin = Application.InchesToPoints(0.196850393700787)
        .BottomMargin = Application.InchesToPoints(0.196850393700787)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    ActiveSheet.PageSetup.PrintArea = "$A$1:$K$51"
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.511811023622047)
        .RightMargin = Application.InchesToPoints(0.511811023622047)
        .TopMargin = Application.InchesToPoints(0.196850393700787)
        .BottomMargin = Application.InchesToPoints(0.196850393700787)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    ActiveWindow.SmallScroll Down:=-20
    ExecuteExcel4Macro "PRINT(1,,,1,,,,,,,,1,,,TRUE,,FALSE)"
End Sub
