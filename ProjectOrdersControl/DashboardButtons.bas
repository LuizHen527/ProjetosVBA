Attribute VB_Name = "DashboardButtons"
'@Folder("VBAProject")
Option Explicit

Sub TodosPedidosEmAberto()
    Dim pedidosEmAberto() As String
    
    
    pedidosEmAberto = PegaPedidosEmAberto
    
    'Contar numero de pedidos
    ContaNumeroDePedidos pedidosEmAberto
    
End Sub

'Os numeros de pedidos precisam estar na segunda "coluna" do array

'Conta numero de pedidos
Function ContaNumeroDePedidos(pedidos() As String) As Integer
    Dim i As Integer, contPedidos As Integer, iVer As Integer
    
    For i = 0 To UBound(pedidos)
        
        For iVer = 0 To i - 1
            If pedidos(i, 2) = pedidos(iVer, 2) Then
                GoTo ProximoPedido
            End If
        Next iVer
        
        
        contPedidos = contPedidos + 1
ProximoPedido:
    Next i
    
End Function

Function PegaPedidosEmAberto() As String()
    Dim rng As Range
    Dim pedidos() As String
    Dim i As Integer

    
    i = 0
    
    ThisWorkbook.Sheets("base").Select
    
    Range("A2").Select
    
    'tirar filtros
    If ActiveSheet.FilterMode Then
        ActiveSheet.ShowAllData
    End If
    
    ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=10, Criteria1:= _
        "EM ABERTO"
    ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=11, Criteria1:= _
        "SIM"
        
    ReDim pedidos(Worksheets("base").AutoFilter.Range.Offset(1, 0).Columns("B").SpecialCells(xlCellTypeVisible).Count, 13)
        
    For Each rng In Worksheets("base").AutoFilter.Range.Offset(1, 0).Columns("B").SpecialCells(xlCellTypeVisible)
        'Pegar todas as colunas
        
        'DATA PEDIDO
        pedidos(i, 0) = rng.Offset(0, -1).Value
        
        'PEDIDO
        pedidos(i, 1) = rng.Value
        
        'CLIENTE
        pedidos(i, 2) = rng.Offset(0, 1).Value
        
        'VENDEDOR
        pedidos(i, 3) = rng.Offset(0, 2).Value
        
        'CADASTRADO
        pedidos(i, 4) = rng.Offset(0, 3).Value
        
        'PRODUTO
        pedidos(i, 5) = rng.Offset(0, 4).Value
        
        'QUANTIDADE
        pedidos(i, 6) = rng.Offset(0, 5).Value
        
        'UNID.
        pedidos(i, 7) = rng.Offset(0, 6).Value
        
        'R$
        pedidos(i, 8) = rng.Offset(0, 7).Value
        
        'SITUA플O
        pedidos(i, 9) = rng.Offset(0, 8).Value
        
        'PEDIDO ATEN플O
        pedidos(i, 10) = rng.Offset(0, 9).Value
        
        'OBSERVA플O
        pedidos(i, 11) = rng.Offset(0, 10).Value
        
        'DATA ATUALIZA플O
        pedidos(i, 12) = rng.Offset(0, 11).Value
        
        i = i + 1
        
    Next rng
    
    PegaPedidosEmAberto = pedidos
    
End Function
