Attribute VB_Name = "M23_BD_MOLDURAS"

Option Base 1
Sub BD_MOLDURAS()
''As mudanças do programa não aparecem passo-a-passo na tela
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim TabRange As Range
Dim TabCache As PivotCache
Dim TabDin As PivotTable

Sheets("Macro").Select

plan_name = "BD_MOLDURAS"
Table = "MOLDURAS_1"
Id_Kit = "MOLDURAS"

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
        .MergeLabels = False
End With

'Filtros de análise
With ActiveSheet.PivotTables(Table).PivotFields("5.Familia")
        .Orientation = xlPageField
        .Position = 1
End With
    ActiveSheet.PivotTables(Table).PivotFields("5.Familia").CurrentPage = "MOLDURAS"
    


With ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("Ano")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("Mes")
        .Orientation = xlRowField
        .Position = 2
    End With
    
    With ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("6.Identificaçao")
        .Orientation = xlRowField
        .Position = 3
    End With
    
    
    With ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("10.Acabamentos")
        .Orientation = xlRowField
        .Position = 4
    End With
    With ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("7.SubIdentificaçao")
        .Orientation = xlRowField
        .Position = 5
    End With
   

    With ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("Conv. Unid")
        .Orientation = xlRowField
        .Position = 6
    End With
    ActiveWindow.SmallScroll Down:=15
    Sheets("BD_MOLDURAS").Select
    With ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("12.Comprimento")
        .Orientation = xlRowField
        .Position = 7
    End With
    
    With ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("17.Peso Total")
        .Orientation = xlRowField
        .Position = 8
    End With
    Sheets("BD_MOLDURAS").Select
    ActiveSheet.PivotTables("MOLDURAS_1").AddDataField ActiveSheet.PivotTables( _
        "MOLDURAS_1").PivotFields("21.ConvQtd"), "Soma de 21.ConvQtd", xlSum
    ActiveSheet.PivotTables("MOLDURAS_1").AddDataField ActiveSheet.PivotTables( _
        "MOLDURAS_1").PivotFields("23.Peso total"), "Soma de 23.Peso total", xlSum
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("Data").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("Ano").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("Mes").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("Hora").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("Pedido").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("Cliente").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("UF").Subtotals = Array(False _
        , False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("Pessoa").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("Cadastrado").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("Código").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("Referencia").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("Produto").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("Total").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("Qtde").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("Unidade").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("5.Familia").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("6.Identificaçao").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("7.SubIdentificaçao"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("8.Formato").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("9.Caracteristica"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("10.Acabamentos").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("11.Material").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("12.Comprimento").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("13.Altura").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("14.Unidade").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("15.SubUnidade").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("16.Peso/m").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("17.Peso Total").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("18. Custo").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("19.Conv. Metros").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("20.VML").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("Conv. Qtde").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("Conv. Unid").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("21.ConvQtd").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("22.ConvUnid").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").PivotFields("23.Peso total").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("MOLDURAS_1").RowAxisLayout xlTabularRow

'End Sub
'
'
'Sub molduras()

Range("A1").Select

'range_row inicial da Base
inicial_range_mold = Cells(Cells.Rows.Count, 1).End(xlUp).Row

'Transforma o nome da planilha base em variavel
last_line_mold = Cells(Cells.Rows.Count, 1).End(xlUp).Row
    
    Range("A3").Select
        Selection.CurrentRegion.Select
        Selection.Copy
    Range("L3").Select
        ActiveSheet.Paste

'---------------------------------------------------------------------------------------------------------------------

células = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U")

'Range das células

For a = 11 To 20
Range(células(a) & "1").Select
    
    'Nº linhas para fazer o preenchimento
    For i = 1 To last_line_mold - 1
    
        If ActiveCell.Offset(1, 0).Value = "" Then
            ActiveCell.Copy
            ActiveCell.Offset(1, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Else
            ActiveCell.Offset(1, 0).Select
        End If
    Next

Next
End Sub



Sub Preencher_Dados()

'Application.ScreenUpdating = False
'Application.DisplayAlerts = False

'Sheets("Base").Select

células = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U")

'Range das células

For a = 11 To 20
Range(células(a) & "1").Select
    
    'Nº linhas para fazer o preenchimento
    For i = 1 To 210
    
        If ActiveCell.Offset(1, 0).Value = "" Then
            ActiveCell.Copy
            ActiveCell.Offset(1, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Else
            ActiveCell.Offset(1, 0).Select
        End If
    Next

Next
End Sub

