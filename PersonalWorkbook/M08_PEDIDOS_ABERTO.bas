Attribute VB_Name = "M08_PEDIDOS_ABERTO"
Sub Menu()

Application.DisplayAlerts = False

On Error GoTo panda

    If Sheets("Pedidos_Aberto").Activate = True Then
        Sheets("Pedidos_Aberto").Delete
    End If

panda:

Call Pedidos

Application.DisplayAlerts = True

End Sub

Sub Pedidos() 'Geração da aba referente as familias

Application.DisplayAlerts = False

Dim TabRange As Range
Dim TabCache As PivotCache
Dim TabDin As PivotTable

Sheets("Macro").Select

plan_name = "Pedidos_Aberto"
Table = "Pedidos"

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

'Dados de coluna
With ActiveSheet.PivotTables("Pedidos").PivotFields("Ano")
        .Orientation = xlRowField
        .Position = 1
End With

With ActiveSheet.PivotTables("Pedidos").PivotFields("Mes")
        .Orientation = xlRowField
        .Position = 2
End With

  ActiveSheet.PivotTables("Pedidos").AddDataField ActiveSheet.PivotTables( _
        "Pedidos").PivotFields("Total"), "Soma de Total", xlSum
    ActiveSheet.PivotTables("Pedidos").AddDataField ActiveSheet.PivotTables( _
        "Pedidos").PivotFields("Pedido"), "Contar de Pedido", xlCount
        
    Range("C6").Select
    ActiveSheet.PivotTables("Pedidos").PivotFields("Data").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("Ano").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("Mes").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("Hora").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("Pedido").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("Cliente").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("UF").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("Pessoa").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("Cadastrado").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("Código").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("Referencia").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("Produto").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("Total").Subtotals = Array(False _
        , False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("Qtde").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("Unidade").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("5.Familia").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("6.Identificaçao").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("7.SubIdentificaçao").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("Pedidos").PivotFields("8.Formato").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("9.Caracteristica").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("10.Acabamentos").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("11.Material").Subtotals = Array _
        (False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("12.Comprimento").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("13.Altura").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("14.Unidade").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("15.SubUnidade").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("16.Peso/m").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("17.Peso Total").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("18. Custo").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("19.Conv. Metros").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("20.VML").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("Conv. Qtde").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("Conv. Unid").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("21.ConvQtd").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("22.ConvUnid").Subtotals = Array _
        (False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").PivotFields("23.Peso total").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Pedidos").RowAxisLayout xlTabularRow

End Sub
