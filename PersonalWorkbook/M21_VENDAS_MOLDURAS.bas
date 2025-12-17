Attribute VB_Name = "M21_VENDAS_MOLDURAS"
Sub MOLDURAS_VENDAS() 'Geração da aba referente as familias

Application.DisplayAlerts = False

Dim TabRange As Range
Dim TabCache As PivotCache
Dim TabDin As PivotTable

Sheets("Macro").Select

plan_name = "Vendas_Molduras"
Table = "Molduras_1"
Id_Kit = "Molduras"

'Seleciona os dados
Set TabRange = Cells(1, 1).CurrentRegion


'Define a fonte de dados da Tabela Dinâmica (que ficará em cache)
Set TabCache = ActiveWorkbook.PivotCaches _
.Create(SourceType:=xlDatabase, SourceData:=TabRange)

ActiveWorkbook.Sheets.Add 'Adiciona uma nova planilha, que se torna ativa
    ActiveSheet.Name = plan_name

'Inserir a Tabela Dinâmica na planilha
Set TabDin = TabCache.CreatePivotTable _
(TableDestination:=Cells(1, 1), TableName:=Table)

'Impede auto formatação das células
With ActiveSheet.PivotTables(Table)
        .HasAutoFormat = False
        .MergeLabels = True
End With

'Filtros de análise
With ActiveSheet.PivotTables(Table).PivotFields("5.Familia")
        .Orientation = xlPageField
        .Position = 1
End With
    ActiveSheet.PivotTables(Table).PivotFields("5.Familia").CurrentPage = "MOLDURAS"

'Inclui campo de data
With ActiveSheet.PivotTables("Molduras_1").PivotFields("Data")
        .Orientation = xlRowField
        .Position = 1
End With

'Adicion Campos a serem analisados
ActiveSheet.PivotTables("Molduras_1").AddDataField ActiveSheet.PivotTables( _
    "Molduras_1").PivotFields("Pedido"), "Contar de Pedido", xlCount
ActiveSheet.PivotTables("Molduras_1").AddDataField ActiveSheet.PivotTables( _
    "Molduras_1").PivotFields("21.ConvQtd"), "Soma de 21.ConvQtd", xlSum
    

End Sub

Sub MOLDURAS_VENDAS_2()

last_line = Cells(Cells.Rows.Count, 1).End(xlUp).Row
select_line = last_line - 1
    Range("A" & select_line & ":" & "C" & select_line).Copy
        
'Atualiza planilha Histórico de vendas
Workbooks.Open fileName:="\\121.137.1.5\manutencao1\Lucas\12_Relatórios\2023\01_Relatórios Diários\07_Histórico Vendas de Molduras.xlsx"
    Sheets("BASE").Select
    last_line_2 = Cells(Cells.Rows.Count, 1).End(xlUp).Row
        Range("A" & last_line_2).Select
            Selection.End(xlUp).Select
        ActiveCell.Offset(1, 0).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        ActiveWorkbook.RefreshAll
        ActiveWorkbook.Save

End Sub








