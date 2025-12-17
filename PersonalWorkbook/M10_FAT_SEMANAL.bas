Attribute VB_Name = "M10_FAT_SEMANAL"
Sub Menu()

Application.DisplayAlerts = False

On Error GoTo panda

    If Sheets("R_RAUL").Activate = True Then
        Sheets("R_RAUL").Delete
    End If

panda:

Call R1_RAUL
Call R2_RAUL

Application.DisplayAlerts = True

End Sub

Sub R1_RAUL() 'Geração da aba referente as familias

Application.DisplayAlerts = False

Dim TabRange As Range
Dim TabCache As PivotCache
Dim TabDin As PivotTable

Sheets("Macro").Select

plan_name = "R_RAUL"
Table = "R1_RAUL"

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

'Insere campo de data
With ActiveSheet.PivotTables("R1_RAUL").PivotFields("Data")
        .Orientation = xlRowField
        .Position = 1
End With
   
'Insere valor totalR$
ActiveSheet.PivotTables("R1_RAUL").AddDataField ActiveSheet.PivotTables( _
    "R1_RAUL").PivotFields("Total"), "Soma de Total", xlSum

'Altera formato da tabela
ActiveSheet.PivotTables("R1_RAUL").RowAxisLayout xlTabularRow

'Desativa Totais
ActiveSheet.PivotTables("R1_RAUL").PivotSelect "Data[All]", xlLabelOnly, True
    With ActiveSheet.PivotTables("R1_RAUL")
        .ColumnGrand = False
        .RowGrand = False
    End With
    
End Sub

Sub R2_RAUL() 'Geração da aba referente as familias

panda = Cells(Cells.Rows.Count, 1).End(xlUp).Row

'Copia os dados e insere semanas
Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
Range("D2").Select
    ActiveSheet.Paste
Range("F2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=WEEKNUM(RC[-2],1)"
Range("F2").Select
    Selection.Copy
    
    
Range("F2").Select
    
    Selection.AutoFill Destination:=Range("F2:F" & panda), Type:=xlFillDefault
    Range("F2:F" & panda).Select
Range("F2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Range("G2:G" & panda).Select
    Application.CutCopyMode = False
    ActiveSheet.Range("$G$2:$G$" & panda).RemoveDuplicates Columns:=1, Header:=xlNo
Range("G2:G" & panda).Select
    Selection.Copy
Range("I2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    Application.CutCopyMode = False

'Cálculo por semana
Range("I3").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(C6,R[-1]C,C5)"
Range("I3").Select
    Selection.AutoFill Destination:=Range("I3:M3"), Type:=xlFillDefault
Range("I3:M3").Select

End Sub


Sub R3_RAUL() 'Geração da aba referente as familias

'Transforma o nome da planilha base em variavel
plan_name = ActiveWorkbook.Name

Range("I3:M3").Select
    Selection.Copy


Workbooks.Open fileName:="\\121.137.1.5\manutencao1\Lucas\12_Relatorios\2025\02_Relatorios Semanais\09_Relatórios Semanais Compilados - Raul_25.xlsx"
    Windows("09_Relatórios Semanais Compilados - Raul_25.xlsx").Activate
        Sheets("Faturado x Venda").Select
        
        
        
            Range("E1").Select
                Range(Selection, Selection.End(xlToRight)).Select
                    Selection.Copy
    ActiveWorkbook.Save
    ActiveWindow.Close


End Sub
