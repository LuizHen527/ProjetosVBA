Attribute VB_Name = "C02_FILTROS"
Sub Pedidos()

'Esconde bot�es do menu
ActiveSheet.Shapes.Range(BarShapes).Visible = Not ActiveSheet.Shapes.Range(BarShapes).Visible
ActiveSheet.Shapes.Range(ButtonShapes).Visible = Not ActiveSheet.Shapes.Range(ButtonShapes).Visible

Application.ScreenUpdating = False

Sheets("BASE").Select

'Filtra os pedidos conclu�dos e cancelados
ActiveSheet.Range("$A$2:$AD$6000").AutoFilter
ActiveSheet.Range("$A$2:$AD$6000").AutoFilter Field:=27, Criteria1:=Array( _
    "Aguardando aprova��o da compra", "Aguardando entrega", "Aguardando retirada", _
    "Cotando", "Pesquisa de Mercado", "="), Operator:=xlFilterValues

'Classica em ordem alfab�tica
ActiveWorkbook.Worksheets("BASE").AutoFilter.Sort.SortFields.Clear
ActiveWorkbook.Worksheets("BASE").AutoFilter.Sort.SortFields.Add Key:=Range( _
    "AA2:AA5500"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortTextAsNumbers
        With ActiveWorkbook.Worksheets("BASE").AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

Range("A2").Select
    
'Call C03_BOTOES.Delete

End Sub

Sub Clear()

'Limpa filtros
Application.ScreenUpdating = False

ActiveWorkbook.Worksheets("BASE").AutoFilter.Sort.SortFields.Clear

On Error GoTo AlreadyFiltered
Worksheets("BASE").ShowAllData
On Error GoTo 0

'Call C03_BOTOES.Delete
    
'Esconde bot�es do menu
ActiveSheet.Shapes.Range(BarShapes).Visible = Not ActiveSheet.Shapes.Range(BarShapes).Visible
ActiveSheet.Shapes.Range(ButtonShapes).Visible = Not ActiveSheet.Shapes.Range(ButtonShapes).Visible

Application.ScreenUpdating = True

Exit Sub

AlreadyFiltered:
    MsgBox "Os pedidos j� foram filtrados.", vbInformation, "Dados j� filtrados"
 
    'Esconde bot�es do menu
    ActiveSheet.Shapes.Range(BarShapes).Visible = Not ActiveSheet.Shapes.Range(BarShapes).Visible
    ActiveSheet.Shapes.Range(ButtonShapes).Visible = Not ActiveSheet.Shapes.Range(ButtonShapes).Visible
    
    Application.ScreenUpdating = True
End Sub

Sub Solicitante()

Range("B2").Select
'Application.DoubleClick
'Call C03_BOTOES.Delete

'Esconde bot�es do menu
ActiveSheet.Shapes.Range(BarShapes).Visible = Not ActiveSheet.Shapes.Range(BarShapes).Visible
ActiveSheet.Shapes.Range(ButtonShapes).Visible = Not ActiveSheet.Shapes.Range(ButtonShapes).Visible

End Sub


Sub Financeiro()

Range("L2").Select
'Application.DoubleClick
'Call C03_BOTOES.Delete

'Esconde bot�es do menu
ActiveSheet.Shapes.Range(BarShapes).Visible = Not ActiveSheet.Shapes.Range(BarShapes).Visible
ActiveSheet.Shapes.Range(ButtonShapes).Visible = Not ActiveSheet.Shapes.Range(ButtonShapes).Visible

End Sub

Sub Classifica��o()

Range("U2").Select
'Application.DoubleClick
'Call C03_BOTOES.Delete

'Esconde bot�es do menu
ActiveSheet.Shapes.Range(BarShapes).Visible = Not ActiveSheet.Shapes.Range(BarShapes).Visible
ActiveSheet.Shapes.Range(ButtonShapes).Visible = Not ActiveSheet.Shapes.Range(ButtonShapes).Visible

End Sub

Sub Acompanhamento()

Range("Y2").Select
'Application.DoubleClick
'Call C03_BOTOES.Delete

'Esconde bot�es do menu
ActiveSheet.Shapes.Range(BarShapes).Visible = Not ActiveSheet.Shapes.Range(BarShapes).Visible
ActiveSheet.Shapes.Range(ButtonShapes).Visible = Not ActiveSheet.Shapes.Range(ButtonShapes).Visible

End Sub







