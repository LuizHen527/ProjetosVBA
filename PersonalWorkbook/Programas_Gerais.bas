Attribute VB_Name = "Programas_Gerais"
Sub Find_Path()

    Dim path As String

    path = Application.StartupPath

    Range("A1").Value = path

End Sub

Private Sub EsconderPlanilha_Click()

Application.ScreenUpdating = False

For Each vplan In Sheets
    
    If vplan.Name <> "GERAL" Then
         vplan.Visible = False
         
    End If
Next

Application.ScreenUpdating = True

End Sub

Private Sub MostrarPlanilhas_Click()

Application.ScreenUpdating = False

For Each vplan In Sheets
    
    If vplan.Name <> "GERAL" Then
         vplan.Visible = True
         
    End If
Next

Application.ScreenUpdating = True

End Sub

Sub Passwordbreaker()
    'Author unknown
    'Breaks worksheet password protection.
    Dim i As Integer, j As Integer, k As Integer
    Dim l As Integer, m As Integer, n As Integer
    Dim i1 As Integer, i2 As Integer, i3 As Integer
    Dim i4 As Integer, i5 As Integer, i6 As Integer
    On Error Resume Next
    For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
    For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
    For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
    For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
    ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & _
        Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
        Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
    If ActiveSheet.ProtectContents = False Then
        Range("A1").Value = "One usable password is " & Chr(i) & Chr(j) & _
            Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
            Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
         Exit Sub
         
    End If
    Next: Next: Next: Next: Next: Next
    Next: Next: Next: Next: Next: Next
End Sub

Sub Importaçao()

Open Range("B1").Value For Output As 1

Range("A2").Select
LastRow = Cells(Cells.Rows.Count, 1).End(xlUp).Row - 1

For i = 1 To LastRow

    Print #1, ActiveCell.Value
    
    ActiveCell.Offset(1, 0).Select
    
    
Next

MsgBox "Arquivo gerado com sucesso"


Close 1


End Sub

Sub lsLigarTelaCheia()
    'Oculta todas as guias de menu
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    
    'Ocultar barra de fórmulas
    Application.DisplayFormulaBar = False
    
    'Ocultar barra de status, disposta ao final da planilha
    Application.DisplayStatusBar = False
    
    'Alterar o nome do Excel
    Application.Caption = "Controle de manutenção de veículos 3.0"
    
    With ActiveWindow
        'Ocultar barra horizontal
        .DisplayHorizontalScrollBar = False
        
        'Ocultar barra vertical
        .DisplayVerticalScrollBar = False
        
        'Ocultar guias das planilhas
        .DisplayWorkbookTabs = False
        
        'Oculta os títulos de linha e coluna
        .DisplayHeadings = False
        
        'Oculta valores zero na planilha
        .DisplayZeros = False
        
        'Oculta as linhas de grade da planilha
        .DisplayGridlines = False
    End With
End Sub

Sub lsDesligarTelaCheia()
    'Reexibe os menus
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    
    'Reexibir a barra de fórmulas
    Application.DisplayFormulaBar = True
    
    'Reexibir a barra de status, disposta ao final da planilha
    Application.DisplayStatusBar = True
    
    'Reexibir o cabeçalho da Pasta de trabalho
    ActiveWindow.DisplayHeadings = True
    
    'Retornar o nome do Excel
    Application.Caption = ""
    
    With ActiveWindow
        'Reexibir barra horizontal
        .DisplayHorizontalScrollBar = True
        
        'Reexibir barra vertical
        .DisplayVerticalScrollBar = True
        
        'Reexibir guias das planilhas
        .DisplayWorkbookTabs = True
        
        'Reexibir os títulos de linha e coluna
        .DisplayHeadings = True
        
        'Reexibir valores zero na planilha
        .DisplayZeros = True
        
        'Reexibir as linhas de grade da planilha
        .DisplayGridlines = True
    End With
End Sub

'Chama o procedimento de tela cheia ao abrir a pasta de trabalho
Private Sub Workbook_Open()
    lsLigarTelaCheia
End Sub

'Desliga o modo de tela cheia ao fechar a pasta de trabalho
Private Sub Workbook_Close()
    lsDesligarTelaCheia
End Sub

Sub TecSerp_GerencialBI()

Application.WindowState = xlMinimized
Application.SendKeys ("^%a")
Application.Wait DateTime.Now + DateTime.TimeValue("00:00:05")
Application.SendKeys ("lucas1234")
Application.Wait DateTime.Now + DateTime.TimeValue("00:00:02")
Application.SendKeys ("~")
Application.SendKeys ("%")
Application.SendKeys ("{RIGHT}")
Application.SendKeys ("{DOWN}")
Application.SendKeys ("{DOWN}")
Application.SendKeys ("{RIGHT}")
Application.SendKeys ("{DOWN}")
Application.SendKeys ("~")

End Sub

Sub TecSerp_Pedidos()

Application.WindowState = xlMinimized
Application.SendKeys ("^%a")
Application.Wait DateTime.Now + DateTime.TimeValue("00:00:05")
Application.SendKeys ("lucas1234")
Application.Wait DateTime.Now + DateTime.TimeValue("00:00:02")
Application.SendKeys ("~")
Application.SendKeys ("%")
Application.SendKeys ("{RIGHT}")
Application.SendKeys ("{DOWN}")
Application.SendKeys ("{DOWN}")
Application.SendKeys ("{RIGHT}")
Application.SendKeys ("~")
Application.SendKeys ("~")

End Sub

Sub lsLigarTelaCheia_2()
    'Oculta todas as guias de menu
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    
    'Ocultar barra de fórmulas
    Application.DisplayFormulaBar = False
    
    'Ocultar barra de status, disposta ao final da planilha
    Application.DisplayStatusBar = False
    
    With ActiveWindow
        'Ocultar barra horizontal
        '.DisplayHorizontalScrollBar = False
        
        'Ocultar barra vertical
        '.DisplayVerticalScrollBar = False
        
        'Ocultar guias das planilhas
        .DisplayWorkbookTabs = False
        
        'Oculta os títulos de linha e coluna
        '.DisplayHeadings = False
        
        'Oculta valores zero na planilha
        .DisplayZeros = False
        
        'Oculta as linhas de grade da planilha
        .DisplayGridlines = False
    End With
End Sub

Sub Calculoautomatico()

Application.Calculation = xlCalculationAutomatic

End Sub

Sub CalendarMaker()

'       ' Unprotect sheet if had previous calendar to prevent error.
'       ActiveSheet.Protect DrawingObjects:=False, Contents:=False, _
'          Scenarios:=False
       ' Prevent screen flashing while drawing calendar.
       Application.ScreenUpdating = False
       ' Set up error trapping.
       On Error GoTo MyErrorTrap
       ' Clear area a1:g14 including any previous calendar.
       Range("a1:g14").Clear
       ' Use InputBox to get desired month and year and set variable
       ' MyInput.
       MyInput = InputBox("Type in Month and year for Calendar ")
       ' Allow user to end macro with Cancel in InputBox.
       If MyInput = "" Then Exit Sub
       ' Get the date value of the beginning of inputted month.
       StartDay = DateValue(MyInput)
       ' Check if valid date but not the first of the month
       ' -- if so, reset StartDay to first day of month.
       If Day(StartDay) <> 1 Then
           StartDay = DateValue(Month(StartDay) & "/1/" & _
               Year(StartDay))
       End If
       ' Prepare cell for Month and Year as fully spelled out.
       Range("a1").NumberFormat = "mmmm yyyy"
       ' Center the Month and Year label across a1:g1 with appropriate
       ' size, height and bolding.
       With Range("a1:g1")
           .HorizontalAlignment = xlCenterAcrossSelection
           .VerticalAlignment = xlCenter
           .Font.Size = 18
           .Font.Bold = True
           .RowHeight = 35
       End With
       ' Prepare a2:g2 for day of week labels with centering, size,
       ' height and bolding.
       With Range("a2:g2")
           .ColumnWidth = 11
           .VerticalAlignment = xlCenter
           .HorizontalAlignment = xlCenter
           .VerticalAlignment = xlCenter
           .Orientation = xlHorizontal
           .Font.Size = 12
           .Font.Bold = True
           .RowHeight = 20
       End With
       ' Put days of week in a2:g2.
       Range("a2") = "Sunday"
       Range("b2") = "Monday"
       Range("c2") = "Tuesday"
       Range("d2") = "Wednesday"
       Range("e2") = "Thursday"
       Range("f2") = "Friday"
       Range("g2") = "Saturday"
       ' Prepare a3:g7 for dates with left/top alignment, size, height
       ' and bolding.
       With Range("a3:g8")
           .HorizontalAlignment = xlRight
           .VerticalAlignment = xlTop
           .Font.Size = 18
           .Font.Bold = True
           .RowHeight = 21
       End With
       ' Put inputted month and year fully spelling out into "a1".
       Range("a1").Value = Application.Text(MyInput, "mmmm yyyy")
       ' Set variable and get which day of the week the month starts.
       DayofWeek = Weekday(StartDay)
       ' Set variables to identify the year and month as separate
       ' variables.
       CurYear = Year(StartDay)
       CurMonth = Month(StartDay)
       ' Set variable and calculate the first day of the next month.
       FinalDay = DateSerial(CurYear, CurMonth + 1, 1)
       ' Place a "1" in cell position of the first day of the chosen
       ' month based on DayofWeek.
       Select Case DayofWeek
           Case 1
               Range("a3").Value = 1
           Case 2
               Range("b3").Value = 1
           Case 3
               Range("c3").Value = 1
           Case 4
               Range("d3").Value = 1
           Case 5
               Range("e3").Value = 1
           Case 6
               Range("f3").Value = 1
           Case 7
               Range("g3").Value = 1
       End Select
       ' Loop through range a3:g8 incrementing each cell after the "1"
       ' cell.
       For Each cell In Range("a3:g8")
           RowCell = cell.Row
           ColCell = cell.Column
           ' Do if "1" is in first column.
           If cell.Column = 1 And cell.Row = 3 Then
           ' Do if current cell is not in 1st column.
           ElseIf cell.Column <> 1 Then
               If cell.Offset(0, -1).Value >= 1 Then
                   cell.Value = cell.Offset(0, -1).Value + 1
                   ' Stop when the last day of the month has been
                   ' entered.
                   If cell.Value > (FinalDay - StartDay) Then
                       cell.Value = ""
                       ' Exit loop when calendar has correct number of
                       ' days shown.
                       Exit For
                   End If
               End If
           ' Do only if current cell is not in Row 3 and is in Column 1.
           ElseIf cell.Row > 3 And cell.Column = 1 Then
               cell.Value = cell.Offset(-1, 6).Value + 1
               ' Stop when the last day of the month has been entered.
               If cell.Value > (FinalDay - StartDay) Then
                   cell.Value = ""
                   ' Exit loop when calendar has correct number of days
                   ' shown.
                   Exit For
               End If
           End If
       Next

       ' Create Entry cells, format them centered, wrap text, and border
       ' around days.
       For x = 0 To 5
           Range("A4").Offset(x * 2, 0).EntireRow.Insert
           With Range("A4:G4").Offset(x * 2, 0)
               .RowHeight = 65
               .HorizontalAlignment = xlCenter
               .VerticalAlignment = xlTop
               .WrapText = True
               .Font.Size = 10
               .Font.Bold = False
               ' Unlock these cells to be able to enter text later after
               ' sheet is protected.
               .Locked = False
           End With
           ' Put border around the block of dates.
           With Range("A3").Offset(x * 2, 0).Resize(2, _
           7).Borders(xlLeft)
               .Weight = xlThick
               .ColorIndex = xlAutomatic
           End With

           With Range("A3").Offset(x * 2, 0).Resize(2, _
           7).Borders(xlRight)
               .Weight = xlThick
               .ColorIndex = xlAutomatic
           End With
           Range("A3").Offset(x * 2, 0).Resize(2, 7).BorderAround _
              Weight:=xlThick, ColorIndex:=xlAutomatic
       Next
       If Range("A13").Value = "" Then Range("A13").Offset(0, 0) _
          .Resize(2, 8).EntireRow.Delete
       ' Turn off gridlines.
       ActiveWindow.DisplayGridlines = False
'       ' Protect sheet to prevent overwriting the dates.
'       ActiveSheet.Protect DrawingObjects:=True, Contents:=True, _
'          Scenarios:=True

       ' Resize window to show all of calendar (may have to be adjusted
       ' for video configuration).
       ActiveWindow.WindowState = xlMaximized
       ActiveWindow.ScrollRow = 1

       ' Allow screen to redraw with calendar showing.
       Application.ScreenUpdating = True
       ' Prevent going to error trap unless error found by exiting Sub
       ' here.
       Exit Sub
   ' Error causes msgbox to indicate the problem, provides new input box,
   ' and resumes at the line that caused the error.
MyErrorTrap:
       MsgBox "You may not have entered your Month and Year correctly." _
           & Chr(13) & "Spell the Month correctly" _
           & " (or use 3 letter abbreviation)" _
           & Chr(13) & "and 4 digits for the Year"
       MyInput = InputBox("Type in Month and year for Calendar")
       If MyInput = "" Then Exit Sub
       Resume
   End Sub


