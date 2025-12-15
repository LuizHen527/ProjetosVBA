Attribute VB_Name = "M06_GIOVANI"
Sub Menu()

Application.DisplayAlerts = False

On Error GoTo panda

    If Sheets("A_GIOVANI").Activate = True Then
        Sheets("A_GIOVANI").Delete
    End If

panda:

Call A1_Giovani
Call A2_Giovani
Call A3_Giovani
Call A4_Giovani

Application.DisplayAlerts = True

End Sub

Sub A1_Giovani() 'Geração da aba referente as familias

Application.DisplayAlerts = False

Dim TabRange As Range
Dim TabCache As PivotCache
Dim TabDin As PivotTable

Sheets("Macro").Select

plan_name = "A_GIOVANI"
Table = "A1_GIOVANI"

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
With ActiveSheet.PivotTables("A1_GIOVANI").PivotFields("Data")
        .Orientation = xlRowField
        .Position = 1
End With
   
'Insere campo de família
With ActiveSheet.PivotTables("A1_GIOVANI").PivotFields("5.Familia")
        .Orientation = xlColumnField
        .Position = 1
End With

'Insere valor totalR$
ActiveSheet.PivotTables("A1_GIOVANI").AddDataField ActiveSheet.PivotTables( _
    "A1_GIOVANI").PivotFields("Total"), "Soma de Total", xlSum

'Altera formato da tabela
ActiveSheet.PivotTables("A1_GIOVANI").RowAxisLayout xlTabularRow

End Sub
Sub A2_Giovani() 'Geração da aba referente as familias

Application.DisplayAlerts = False

Dim TabRange As Range
Dim TabCache As PivotCache
Dim TabDin As PivotTable

Sheets("Macro").Select

plan_name = "A_GIOVANI"
Table = "A2_GIOVANI"

'Seleciona os dados
Set TabRange = Cells(1, 1).CurrentRegion

'Define a fonte de dados da Tabela Dinâmica (que ficará em cache)
Set TabCache = ActiveWorkbook.PivotCaches _
.Create(SourceType:=xlDatabase, SourceData:=TabRange)

Sheets(plan_name).Select

'Inserir a Tabela Dinâmica na planilha
Set TabDin = TabCache.CreatePivotTable _
(TableDestination:=Cells(1, 20), TableName:=Table)

'Impede auto formatação das células
With ActiveSheet.PivotTables(Table)
        .HasAutoFormat = False
        .MergeLabels = True
End With

'Insere campo de data
With ActiveSheet.PivotTables(Table).PivotFields("Data")
        .Orientation = xlRowField
        .Position = 1
End With
   
ActiveSheet.PivotTables(Table).PivotSelect "Data[All]", xlLabelOnly, _
        True
    With ActiveSheet.PivotTables(Table).PivotFields("5.Familia")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables(Table).PivotFields("5.Familia").ClearAllFilters
    ActiveSheet.PivotTables(Table).PivotFields("5.Familia").CurrentPage = _
        "KITS"
    With ActiveSheet.PivotTables(Table).PivotFields("6.Identificaçao")
        .Orientation = xlColumnField
        .Position = 1
    End With

'Insere valor totalR$
ActiveSheet.PivotTables(Table).AddDataField ActiveSheet.PivotTables( _
    Table).PivotFields("Total"), "Soma de Total", xlSum
    
'Kits
Range("T4").Select
    Selection.CurrentRegion.Select
    Selection.Copy
Range("A80").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'Perfis
Range("U1").Value = "PERFIS"

Range("T4").Select
    Selection.CurrentRegion.Select
    Selection.Copy
Range("A110").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'Botoes
Range("U1").Value = "ACESSORIOS"

Range("T4").Select
    Selection.CurrentRegion.Select
    Selection.Copy
Range("A140").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
End Sub
Sub A3_Giovani()

'Geral
Range("A1").Select
    Selection.CurrentRegion.Select
    Selection.Copy
Range("A50").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   
Range("A51").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
Range("T51").Select
    ActiveSheet.Paste
Range("U51").Select
    ActiveCell.FormulaR1C1 = "Total geral"
Range("V51").Select
    ActiveCell.FormulaR1C1 = "BLINDEX"
Range("W51").Select
    ActiveCell.FormulaR1C1 = "BOX"
Range("X51").Select
    ActiveCell.FormulaR1C1 = "ROAPLAS"
Range("Y51").Select
    ActiveCell.FormulaR1C1 = "KITS"
Range("Z51").Select
    ActiveCell.FormulaR1C1 = "MOLDURAS"
Range("AA51").Select
    ActiveCell.FormulaR1C1 = "BOX"
Range("AB51").Select
    ActiveCell.FormulaR1C1 = "ENGENHARIA"
Range("AC51").Select
    ActiveCell.FormulaR1C1 = "PERFIS"
Range("AD51").Select
    ActiveCell.FormulaR1C1 = "BOTOES"
Range("AE51").Select
    ActiveCell.FormulaR1C1 = "OUTROS"
    
End Sub

Sub A4_Giovani()

'Total Geral
Range("U52").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(INDEX(R52C2:R79C18,MATCH(RC[-1],R52C1:R79C1,0),MATCH(R51C21,R51C2:R51C18,0)),"""")"
    Selection.AutoFill Destination:=Range("U52:U80"), Type:=xlFillDefault
    Range("U52:U80").Select
    
''Kits Blindex, Box e Roaplas
'Range("V52").Select
'    Selection.FormulaArray = _
'         "=IFERROR(INDEX(R82C2:R109C18,MATCH(RC20,R82C1:R109C1,0),MATCH(R51C,R81C2:R81C18,0))+INDEX(R82C2:R109C18,MATCH(RC20,R82C1:R109C1,0),MATCH(""COMBATE"",R81C2:R81C18,0)),"""")"
'    Selection.AutoFill Destination:=Range("V52:V80"), Type:=xlFillDefault
'        Selection.AutoFill Destination:=Range("V52:V80"), Type:=xlFillDefault
'    Range("V52:V80").Select
'    Selection.AutoFill Destination:=Range("V52:X80"), Type:=xlFillDefault
'    Range("V52:X80").Select

'Kits Blindex
    Range("V52").Select
    Selection.FormulaArray = _
        "=IFERROR(INDEX(R82C2:R109C18,MATCH(RC20,R82C1:R109C1,0),MATCH(R51C,R81C2:R81C18,0)),0)+IFERROR(INDEX(R82C2:R109C18,MATCH(RC20,R82C1:R109C1,0),MATCH(""COMBATE"",R81C2:R81C18,0)),0)"
    Selection.AutoFill Destination:=Range("V52:V80"), Type:=xlFillDefault
    Range("V52:V80").Select

'Kits Box
    Range("W52").Select
    Selection.FormulaArray = _
        "=IFERROR(INDEX(R82C2:R109C18,MATCH(RC20,R82C1:R109C1,0),MATCH(R51C,R81C2:R81C18,0)),0)"
    Selection.AutoFill Destination:=Range("W52:W80"), Type:=xlFillDefault
    Range("W52:W80").Select
    
'Kits Roaplas
    Range("X52").Select
    Selection.FormulaArray = _
        "=IFERROR(INDEX(R82C2:R109C18,MATCH(RC20,R82C1:R109C1,0),MATCH(R51C,R81C2:R81C18,0)),0)"
    Selection.AutoFill Destination:=Range("X52:X80"), Type:=xlFillDefault
    Range("X52:X80").Select
  
'Kits Outros
    Range("Y52").Select
    Selection.FormulaArray = _
        "=IFERROR(INDEX(R52C2:R79C18,MATCH(RC[-5],R52C1:R79C1,0),MATCH(R51C25,R51C2:R51C18,0))-SUM(RC[-3]:RC[-1]),"""")"
    Selection.AutoFill Destination:=Range("Y52:Y80"), Type:=xlFillDefault
    Range("Y52:Y80").Select
    
    
'Molduras
Range("Z52").Select
    Selection.FormulaArray = _
        "=IFERROR(INDEX(R52C2:R79C18,MATCH(RC[-6],R52C1:R79C1,0),MATCH(R51C26,R51C2:R51C18,0)),"""")"
    Selection.AutoFill Destination:=Range("Z52:Z79"), Type:=xlFillDefault
    Range("Z52:Z80").Select
    
'Perfis Box e Engenharia
Range("AA52").Select
    Selection.FormulaArray = _
        "=IFERROR(INDEX(R112C2:R139C18,MATCH(RC20,R112C1:R139C1,0),MATCH(R51C,R111C2:R111C18,0)),"""")"
    Selection.AutoFill Destination:=Range("AA52:AA80"), Type:=xlFillDefault
    Range("AA52:AA80").Select
    Selection.AutoFill Destination:=Range("AA52:AB80"), Type:=xlFillDefault
    Range("AA52:AB80").Select


'Perfis Outros
    Range("AC52").Select
    Selection.FormulaArray = _
        "=IFERROR(INDEX(R52C2:R79C18,MATCH(RC[-9],R52C1:R79C1,0),MATCH(R51C29,R51C2:R51C18,0))-SUM(RC[-2]:RC[-1]),"""")"
    Selection.AutoFill Destination:=Range("AC52:AC78"), Type:=xlFillDefault
    Range("AC52:AC80").Select

'Botões
    Range("AD52").Select
    Selection.FormulaArray = _
        "=IFERROR(INDEX(R142C2:R170C18,MATCH(RC[-10],R142C1:R170C1,0),MATCH(R51C30,R141C2:R141C18,0)),"""")"
    Selection.AutoFill Destination:=Range("AD52:AD80"), Type:=xlFillDefault
    Range("AD52:AD80").Select
    
'Outros
    Range("AE52").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-10]<>"""",RC[-10]-SUM(RC[-9]:RC[-1]),"""")"
    Selection.AutoFill Destination:=Range("AE52:AE77"), Type:=xlFillDefault
    Range("AE52:AE80").Select

End Sub
