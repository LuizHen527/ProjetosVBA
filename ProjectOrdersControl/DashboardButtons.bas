Attribute VB_Name = "DashboardButtons"
'@Folder("VBAProject")
Option Explicit

Sub TodosPedidosEmAberto()
    ActiveSheet.Shapes.Range("PedidoMenu").Visible = Not ActiveSheet.Shapes.Range("PedidoMenu").Visible
    
    Application.ScreenUpdating = False
    
    Dim pedidosEmAberto() As String
    Dim numeroPedidos As Integer
    
    pedidosEmAberto = PegaPedidosEmAberto
    
    'Contar numero de pedidos
    numeroPedidos = ContaNumeroDePedidos(pedidosEmAberto)
    
    MostrarDadosTodosItensPedidos pedidosEmAberto, numeroPedidos

    Application.ScreenUpdating = True
    
End Sub

Sub OpenCloseMenu()

    ActiveSheet.Shapes.Range("PedidoMenu").Visible = Not ActiveSheet.Shapes.Range("PedidoMenu").Visible
    
End Sub

Sub ResumoDeCadaPedido()
    ActiveSheet.Shapes.Range("PedidoMenu").Visible = Not ActiveSheet.Shapes.Range("PedidoMenu").Visible
    
    Application.ScreenUpdating = False
    
    Dim pedidosEmAberto() As String
    Dim resumoPedidos() As String
    
    'Pegar dados de todos os pedidos
    pedidosEmAberto = PegaPedidosEmAberto
    
    resumoPedidos = CriarResumoPedidos(pedidosEmAberto)
    
    MostrarResumoPedidos resumoPedidos
    
    Application.ScreenUpdating = True
End Sub

Sub MateriaisParaProduzir()
    Dim materiaisProduzir() As String
    
    materiaisProduzir = PegaMateriaisProduzir
    
    MostrarMateriaisProduzir materiaisProduzir
    
End Sub


'-------------- FUNÇÕES DE PROCESSAMENTO --------------

Function PegaMateriaisProduzir() As String()
    Dim materiaisProduzir() As String, pedidosEmAberto() As String
    Dim iterator As Integer, i As Integer
    Dim rng As Range
    
    'Pegar dados de todos os pedidos
    pedidosEmAberto = PegaPedidosEmAberto
    
    ThisWorkbook.Sheets("perfis_pedido").Select
    
    ReDim materiaisProduzir(300, 5)
    
    For Each rng In Range("A3", "A" & Cells(Rows.Count, 1).End(xlUp).row)
        If rng.Offset(0, 4).Value = "PRODUZIR" Then
            
            'Salvar todos os dados em array
            
            'NUMERO
            materiaisProduzir(iterator, 1) = rng.Value
            
            'PERFIL
            materiaisProduzir(iterator, 2) = rng.Offset(0, 1).Value
            
            'COR
            materiaisProduzir(iterator, 3) = rng.Offset(0, 2).Value
            
            'QUANTIDADE
            materiaisProduzir(iterator, 4) = rng.Offset(0, 3).Value
            
            'ULTIMA ATUALIZAÇÃO
            materiaisProduzir(iterator, 5) = rng.Offset(0, 5).Value
            
            'Procurar numero do pedido nos pedidos em aberto e salvar a data
            For i = 0 To UBound(pedidosEmAberto)
                
                If rng.Value = pedidosEmAberto(i, 1) Then
                
                    'DATA PEDIDO
                    materiaisProduzir(iterator, 0) = pedidosEmAberto(i, 0)
                    
                    Exit For
                End If
                
            Next i
            
            iterator = iterator + 1
            
        End If
    Next rng
    
    PegaMateriaisProduzir = materiaisProduzir
End Function

'Retorna array com cada pedido sendo apenas uma "linha" do array
Function CriarResumoPedidos(pedidos() As String) As String()
    
    Dim i As Integer, iterator As Integer, numeroPedidos As Integer
    Dim dataPedido As String, numero As String, cliente As String, dataAtualização As String, observacao As String
    Dim resumoPedidos() As String
    Dim valor As Double
    
    numeroPedidos = ContaNumeroDePedidos(pedidos)
    
    dataPedido = pedidos(0, 0)
    numero = pedidos(0, 1)
    cliente = pedidos(0, 2)
    valor = CDbl(pedidos(0, 8))
    observacao = pedidos(0, 11)
    dataAtualização = pedidos(0, 12)
    
    ReDim resumoPedidos(numeroPedidos - 1, 5)
    
    iterator = 0
    
    'Processar pedido de forma que cada pedido fique em uma linha
    For i = 1 To UBound(pedidos)
        
        'Quando vim um pedido diferente, salva o anterior no array e pega o proximo
        If numero <> pedidos(i, 1) Then
    
            'Cadastrar pedido no array
            resumoPedidos(iterator, 0) = dataPedido
            resumoPedidos(iterator, 1) = numero
            resumoPedidos(iterator, 2) = cliente
            resumoPedidos(iterator, 3) = CStr(valor)
            resumoPedidos(iterator, 4) = observacao
            resumoPedidos(iterator, 5) = dataAtualização
            
            iterator = iterator + 1
            
            'Coloca os dados do proximo pedido
            dataPedido = pedidos(i, 0)
            numero = pedidos(i, 1)
            cliente = pedidos(i, 2)
            If pedidos(i, 8) = "" Then valor = 0 Else valor = CDbl(pedidos(i, 8))
            observacao = pedidos(i, 11)
            dataAtualização = pedidos(i, 12)
            
            GoTo NextIteration
            
        End If
        
        If pedidos(i, 8) = "" Then valor = valor + 0 Else valor = valor + CDbl(pedidos(i, 8))
        
        
NextIteration:
    Next i
    
    CriarResumoPedidos = resumoPedidos
    
End Function


'Retorna numero de pedidos no array. Não conta duplicadas
Function ContaNumeroDePedidos(pedidos() As String) As Integer
    Dim i As Integer, contPedidos As Integer, iVer As Integer
    
    For i = 0 To UBound(pedidos)
        
        For iVer = 0 To i - 1
            If pedidos(i, 1) = pedidos(iVer, 1) Then
                GoTo ProximoPedido
            End If
        Next iVer
        
        If pedidos(i, 5) <> "" Then
            contPedidos = contPedidos + 1
        End If
        
ProximoPedido:
    Next i
    
    ContaNumeroDePedidos = contPedidos
    
End Function

'Retorna array com todos os pedidos em aberto e todas as colunas
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
        
        'SITUAÇÃO
        pedidos(i, 9) = rng.Offset(0, 8).Value
        
        'PEDIDO ATENÇÃO
        pedidos(i, 10) = rng.Offset(0, 9).Value
        
        'OBSERVAÇÃO
        pedidos(i, 11) = rng.Offset(0, 10).Value
        
        'DATA ATUALIZAÇÃO
        pedidos(i, 12) = rng.Offset(0, 11).Value
        
        i = i + 1
        
    Next rng
    
    PegaPedidosEmAberto = pedidos
    
End Function



'-------------- FUNÇÕES DE MONTAR TABELAS --------------

Function MostrarMateriaisProduzir(materiaisArr() As String)
    Dim i As Integer
    Dim colorLine As String
    
    'Deletar planilha antiga
    
    'Fazer header
    ThisWorkbook.Sheets("dashboard").Select
    
    Range("A3", "K" & Cells(Rows.Count, 1).End(xlUp).row).Delete
    
    Range("A1").Value = "DASHBOARD - PERFIS PARA PRODUZIR"
    
    Range("A6").Value = "DATA PEDIDO"
    Range("B6").Value = "NUMERO"
    Range("C6").Value = "PERFIL"
    Range("D6").Value = "COR"
    Range("E6").Value = "QUANTIDADE"
    Range("F6").Value = "DATA ATUALIZAÇÃO"
    
    'Estilo HEADER
    With Range("A6:F6")
        .Interior.Color = RGB(97, 183, 241)
        .Font.Bold = True
        .RowHeight = 30
        .Borders.Color = vbWhite
        .Borders(xlInsideVertical).Weight = xlThin
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Font.Size = 11
    End With
    
    colorLine = "blue"
    
    'Incerir dados
    For i = 0 To UBound(materiaisArr)
        
        If materiaisArr(i, 1) <> "" Then
            If colorLine = "white" Then colorLine = "blue" Else colorLine = "white"
            
            'DATA PEDIDO
            Range("A" & 7 + i).Value = CDate(materiaisArr(i, 0))
            
            'NUMERO PEDIDO
            Range("B" & 7 + i).Value = CDbl(materiaisArr(i, 1))
            
            'PERFIL
            Range("C" & 7 + i).Value = materiaisArr(i, 2)
            
            'COR
            Range("D" & 7 + i).Value = materiaisArr(i, 3)
            
            'QUANTIDADE
            Range("E" & 7 + i).Value = materiaisArr(i, 4)
            
            'ATUALIZAÇÃO
            Range("F" & 7 + i).Value = materiaisArr(i, 5)
            
            If colorLine = "blue" Then
                With Range("A" & 7 + i & ":F" & 7 + i)
                    .Interior.Color = RGB(215, 245, 239)
                    .VerticalAlignment = xlCenter
                    .HorizontalAlignment = xlCenter
                    .Borders(xlEdgeTop).Weight = xlThin
                    .Borders(xlEdgeBottom).Weight = xlThin
                    .Borders(xlInsideVertical).Weight = xlThin
                    .Borders(xlInsideVertical).Color = vbWhite
                    .Borders(xlEdgeBottom).Color = vbWhite
                    .Borders(xlEdgeTop).Color = vbWhite
                    .Font.Size = 10
                End With
            Else
                With Range("A" & 7 + i & ":F" & 7 + i)
                    .VerticalAlignment = xlCenter
                    .HorizontalAlignment = xlCenter
                    .Font.Size = 10
                End With
            End If
        End If
        
        
        
    Next i
    
    'Transformar em tabela
    
    'Arruma tamanho da coluna
    Range("A:A").ColumnWidth = 16
    Range("B:B").ColumnWidth = 14
    Columns("C:C").EntireColumn.AutoFit
    Range("D:D").ColumnWidth = 11
    Range("F:F").ColumnWidth = 19.5
    
    Columns("E:E").EntireColumn.AutoFit
    Range("F:F").ColumnWidth = 19.5
    
    'Transformar dados em tabela
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A6:F" & Cells(Rows.Count, 1).End(xlUp).row), , xlYes).Name = "DashBoardResumoPedidos"
    ActiveSheet.ListObjects("DashBoardResumoPedidos").TableStyle = ""
    
End Function

'Monta tabela com o array de pedidos resumidos
Function MostrarResumoPedidos(resumoPedidos() As String)
    Dim i As Integer, iterator As Integer, numeroPedidos As Integer
    Dim colorLine As String
    
    'Apagar tabela que estava antes
    'Formatar tabela de resumo pedidos
    
    numeroPedidos = ContaNumeroDePedidos(resumoPedidos)
    
    ThisWorkbook.Sheets("dashboard").Select
    
    Range("A3", "K" & Cells(Rows.Count, 1).End(xlUp).row).Delete
    
    Range("A1").Value = "DASHBOARD - RESUMO DOS PEDIDOS EM ABERTO"
    
    Range("A3").Value = "TOTAL PEDIDOS"
    Range("B3").Value = "VALOR TOTAL"
    
    Range("A6").Value = "DATA PEDIDO"
    Range("B6").Value = "NUMERO"
    Range("C6").Value = "CLIENTE"
    Range("D6").Value = "VALOR"
    Range("E6").Value = "OBSERVAÇÃO"
    Range("F6").Value = "DATA ATUALIZAÇÃO"
    
    'Estilo TOTAL PEDIDO e VALOR TOTAL
    With Range("A3:B3")
        .Interior.Color = RGB(97, 183, 241)
        .Font.Bold = True
        .RowHeight = 18
        .Borders.Color = vbWhite
        .Borders(xlInsideVertical).Weight = xlThin
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Font.Size = 11
    End With
    
    'Estilo HEADER
    With Range("A6:F6")
        .Interior.Color = RGB(97, 183, 241)
        .Font.Bold = True
        .RowHeight = 30
        .Borders.Color = vbWhite
        .Borders(xlInsideVertical).Weight = xlThin
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Font.Size = 11
    End With
    
    'Como vou montar a tabela?
    'Colocar as linhas do meu array
    'Cada linha precisa ser de uma cor
    
    colorLine = "blue"
    
    For i = 0 To UBound(resumoPedidos)
        
        If colorLine = "white" Then colorLine = "blue" Else colorLine = "white"
        
        'DATA PEDIDO
        Range("A" & 7 + i).Value = CDate(resumoPedidos(i, 0))
        
        'NUMERO
        Range("B" & 7 + i).Value = CDbl(resumoPedidos(i, 1))
        
        'CLIENTE
        Range("C" & 7 + i).Value = resumoPedidos(i, 2)
        
        'VALOR
        With Range("D" & 7 + i)
            .Value = CDbl(resumoPedidos(i, 3))
            .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        End With
        
        'OBSERVAÇÃO
        Range("E" & 7 + i).Value = resumoPedidos(i, 4)
        
        'DATA ATUALIZAÇÃO
        Range("F" & 7 + i).Value = CDate(resumoPedidos(i, 5))
        
        If colorLine = "blue" Then
            With Range("A" & 7 + i & ":F" & 7 + i)
                .Interior.Color = RGB(215, 245, 239)
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
                .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeBottom).Weight = xlThin
                .Borders(xlInsideVertical).Weight = xlThin
                .Borders(xlInsideVertical).Color = vbWhite
                .Borders(xlEdgeBottom).Color = vbWhite
                .Borders(xlEdgeTop).Color = vbWhite
                .Font.Size = 10
            End With
        Else
            With Range("A" & 7 + i & ":F" & 7 + i)
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
                .Font.Size = 10
            End With
        End If
        
    Next i
    
    'Arruma tamanho da coluna
    Range("A:A").ColumnWidth = 16
    Range("B:B").ColumnWidth = 14
    Range("C:C").ColumnWidth = 30
    Columns("D:D").EntireColumn.AutoFit
    Columns("E:E").EntireColumn.AutoFit
    Range("F:F").ColumnWidth = 19.5
    
    'Transformar dados em tabela
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A6:F" & Cells(Rows.Count, 1).End(xlUp).row), , xlYes).Name = "DashBoardResumoPedidos"
    ActiveSheet.ListObjects("DashBoardResumoPedidos").TableStyle = ""
    
    'Insere TOTAL PEDIDOS e VALOR TOTAL
    With Range("A4")
        .Value = numeroPedidos
        .HorizontalAlignment = xlCenter
    End With
    Range("B4").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Range("B4").Formula = "=SUM(DashBoardResumoPedidos[VALOR])"
    
End Function

'Monta tabela com todos os itens dos pedidos
Function MostrarDadosTodosItensPedidos(pedidosEmAberto() As String, numeroDePedidos As Integer)
    Dim i As Integer
    Dim numeroPedido As String, colorLine As String

    'Antes de colocar os dados, converter eles para o formato certo
    'Estilizar cada pedido com uma cor
    
    ThisWorkbook.Sheets("dashboard").Select
    
    Range("A3", "K" & Cells(Rows.Count, 1).End(xlUp).row).Delete
    
    Range("A1").Value = "DASHBOARD - TODOS OS ITENS DOS PEDIDOS EM ABERTO"
    
    Range("A3").Value = "TOTAL PEDIDOS"
    Range("B3").Value = "VALOR TOTAL"
    
    Range("A6").Value = "DATA PEDIDO"
    Range("B6").Value = "NUMERO"
    Range("C6").Value = "PRODUTO"
    Range("D6").Value = "QUANTIDADE"
    Range("E6").Value = "VALOR"
    Range("F6").Value = "OBSERVAÇÃO"
    Range("G6").Value = "DATA ATUALIZAÇÃO"
    
    'Estilo TOTAL PEDIDO e VALOR TOTAL
    With Range("A3:B3")
        .Interior.Color = RGB(97, 183, 241)
        .Font.Bold = True
        .RowHeight = 18
        .Borders.Color = vbWhite
        .Borders(xlInsideVertical).Weight = xlThin
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Font.Size = 11
    End With
    
    'Estilo HEADER
    With Range("A6:G6")
        .Interior.Color = RGB(97, 183, 241)
        .Font.Bold = True
        .RowHeight = 30
        .Borders.Color = vbWhite
        .Borders(xlInsideVertical).Weight = xlThin
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Font.Size = 11
    End With
     
    colorLine = "blue"
    
    For i = 0 To UBound(pedidosEmAberto)
    
        If pedidosEmAberto(i, 1) <> "" Then
        
            If numeroPedido <> pedidosEmAberto(i, 1) Then
            
                If colorLine = "white" Then colorLine = "blue" Else colorLine = "white"
                
                numeroPedido = pedidosEmAberto(i, 1)
                
            End If
        
            'DATA PEDIDO
            Range("A" & 7 + i).Value = CDate(pedidosEmAberto(i, 0))
            
            'PEDIDO
            Range("B" & 7 + i).Value = pedidosEmAberto(i, 1)
            
            'PRODUTO
            Range("C" & 7 + i).Value = pedidosEmAberto(i, 5)
            
            'QUANTIDADE
            Range("D" & 7 + i).Value = pedidosEmAberto(i, 6)
                
            'VALOR
            With Range("E" & 7 + i)
                .Value = CDbl(pedidosEmAberto(i, 8))
                .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            End With
            
            'OBSERVAÇÃO
            Range("F" & 7 + i).Value = pedidosEmAberto(i, 11)
            
            'DATA ATUALIZAÇÃO
            Range("G" & 7 + i).Value = CDate(pedidosEmAberto(i, 12))
            
            
            If colorLine = "blue" Then
                With Range("A" & 7 + i & ":G" & 7 + i)
                    .Interior.Color = RGB(215, 245, 239)
                    .VerticalAlignment = xlCenter
                    .HorizontalAlignment = xlCenter
                    .Borders(xlEdgeTop).Weight = xlThin
                    .Borders(xlEdgeBottom).Weight = xlThin
                    .Borders(xlInsideVertical).Weight = xlThin
                    .Borders(xlInsideVertical).Color = vbWhite
                    .Borders(xlEdgeBottom).Color = vbWhite
                    .Borders(xlEdgeTop).Color = vbWhite
                    .Font.Size = 10
                End With
            Else
                With Range("A" & 7 + i & ":G" & 7 + i)
                    .VerticalAlignment = xlCenter
                    .HorizontalAlignment = xlCenter
                    .Font.Size = 10
                End With
            End If
        
        End If
    Next i
    
    'Arruma tamanho da coluna
    Range("A:A").ColumnWidth = 16
    Range("B:B").ColumnWidth = 14
    Columns("C:C").EntireColumn.AutoFit
    Range("D:D").ColumnWidth = 15.5
    Columns("E:E").EntireColumn.AutoFit
    Columns("F:F").EntireColumn.AutoFit
    Range("G:G").ColumnWidth = 19.5
    
    'Transformar dados em tabela
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A6:G" & Cells(Rows.Count, 7).End(xlUp).row), , xlYes).Name = "DashBoardTodosPedidos"
    ActiveSheet.ListObjects("DashBoardTodosPedidos").TableStyle = ""
    
    'Insere TOTAL PEDIDOS e VALOR TOTAL
    With Range("A4")
        .Value = numeroDePedidos
        .HorizontalAlignment = xlCenter
    End With
    Range("B4").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Range("B4").Formula = "=SUM(DashBoardTodosPedidos[VALOR])"
    
End Function
