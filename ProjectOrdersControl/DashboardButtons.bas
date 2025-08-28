Attribute VB_Name = "DashboardButtons"
'@Folder("VBAProject")
Option Explicit

Sub TodosPedidosEmAberto()
    OpenCloseMenu
    
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

    ActiveSheet.Shapes.Range("pedido_menu").Visible = Not ActiveSheet.Shapes.Range("pedido_menu").Visible
    
End Sub

Sub ResumoDeCadaPedido()
    OpenCloseMenu
    
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
    
    OpenCloseMenu
    Application.ScreenUpdating = False
    
    Dim materiaisProduzir() As String
    
    materiaisProduzir = PegaMateriaisProduzir
    
    MostrarMateriaisProduzir materiaisProduzir
    
    Application.ScreenUpdating = True
    
End Sub

Sub PesquisarPedido()
    
    Application.ScreenUpdating = False
    
    Dim pedidos() As String, pedidosEncontrados() As String
    Dim numeroPedido As Double
    
    'Como quero que funcione?
        'Ao selecionar a opção, aparece uma caixa de dialogo.
        'O numero do pedido é digitado
        'Todos os dados do pedido aparece
        'Permitir pesquisa de todos os pedidos ou apenas pedidos em aberto?
        'Faz mais sentido pesquisar por todos os tipos
        
    'Como vou desenvolver essa funcionalidade?
        'Pegar dados com o mesmo numero de pedido
            'Pegar pedidos em aberto
            'Comparar com o numero
            'Salvar pedidos com o mesmo numero
            'Retornar pedidos com o mesmo numero
        
    numeroPedido = PegaNumeroPedidoInput
    
    If numeroPedido <> 0 Then
            
        pedidos = PegaTodosPedidos
        
        pedidosEncontrados = ProcurarPedido(pedidos, numeroPedido)
        
        If UBound(pedidosEncontrados) = 0 Then
            
            ThisWorkbook.Sheets("dashboard").Select
            
            MsgBox "O numero pode estar errado ou o pedido pode não estar na base.", vbOKOnly + vbInformation, "Pedido não encontrado"
            
            Exit Sub
        End If
        
        MostrarPedidoEncontrado pedidosEncontrados
        
    End If
    
    Application.ScreenUpdating = True
    
End Sub

'-------------- FUNÇÕES DE PROCESSAMENTO --------------

Function ProcurarPedido(pedidosArray() As String, numeroPedido As Double) As String()
    Dim pedidosEncontrados() As String
    Dim i As Integer, iterator As Integer
    
    'Comparar com o numero
        'Salvar pedidos com o mesmo numero
        'Retornar pedidos com o mesmo numero
        
    ReDim pedidosEncontrados(200, 12)
    
    iterator = 0
    
    For i = 0 To UBound(pedidosArray)
        
        'Se achar o numero procurado nos pedidos
        If pedidosArray(i, 1) = CStr(numeroPedido) Then
            
            'DATA PEDIDO
            pedidosEncontrados(iterator, 0) = pedidosArray(i, 0)
            
            'PEDIDO
            pedidosEncontrados(iterator, 1) = pedidosArray(i, 1)
            
            'CLIENTE
            pedidosEncontrados(iterator, 2) = pedidosArray(i, 2)
            
            'VENDEDOR
            pedidosEncontrados(iterator, 3) = pedidosArray(i, 3)
            
            'CADASTRADO
            pedidosEncontrados(iterator, 4) = pedidosArray(i, 4)
            
            'PRODUTO
            pedidosEncontrados(iterator, 5) = pedidosArray(i, 5)
            
            'QUANTIDADE
            pedidosEncontrados(iterator, 6) = pedidosArray(i, 6)
            
            'UNID.
            pedidosEncontrados(iterator, 7) = pedidosArray(i, 7)
            
            'R$
            pedidosEncontrados(iterator, 8) = pedidosArray(i, 8)
            
            'SITUAÇÃO
            pedidosEncontrados(iterator, 9) = pedidosArray(i, 9)
            
            'PEDIDO ATENÇÃO
            pedidosEncontrados(iterator, 10) = pedidosArray(i, 10)
            
            'OBSERVAÇÃO
            pedidosEncontrados(iterator, 11) = pedidosArray(i, 11)
            
            'DATA ATUALIZAÇÃO
            pedidosEncontrados(iterator, 12) = pedidosArray(i, 12)
            
            iterator = iterator + 1
        End If
        
    Next i
    
    If iterator = 0 Then
        ReDim pedidosEncontrados(0)
    End If
    
    ProcurarPedido = pedidosEncontrados
    
End Function

Function PegaNumeroPedidoInput() As Double
    Dim numeroPedido As Double
    
    
    numeroPedido = Application.InputBox("Digite o numero do pedido", "Pesquisar pedido", , , , , , 1)
    
    PegaNumeroPedidoInput = numeroPedido
End Function

Function PegaTodosPedidos() As String()
    Dim pedidos() As String
    Dim i As Integer
    Dim rng As Range
    
    ThisWorkbook.Sheets("base").Select
    
    Range("A2").Select
    
    'tirar filtros
    If ActiveSheet.FilterMode Then
        ActiveSheet.ShowAllData
    End If
    
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
    
    PegaTodosPedidos = pedidos
    
End Function

Function PegaMateriaisProduzir() As String()
    Dim materiaisProduzir() As String, pedidosEmAberto() As String
    Dim iterator As Integer, i As Integer
    Dim rng As Range
    
    'Pegar dados de todos os pedidos
    pedidosEmAberto = PegaPedidosEmAberto
    
    ThisWorkbook.Sheets("perfis_pedido").Select
    
    ReDim materiaisProduzir(300, 6)
    
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
                    
                    'CLIENTE
                    materiaisProduzir(iterator, 6) = pedidosEmAberto(i, 2)
                    
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

Function MostrarPedidoEncontrado(pedidoArr() As String)
    Dim i As Integer, iterator As Integer
    'Como montar a tabela de pedido pesquisado
    'Como?
        'Apagar tabela anterior
        'Montar Header
        'Loopar pelos dados e colocar cada linha
        'Transformar tudo em tabela
    
    ThisWorkbook.Sheets("dashboard").Select
    
    
    If Range("A6").Value > 2000 Then

        Range("A3", "M" & Cells(Rows.Count, 1).End(xlUp).row).Delete
        
    ElseIf IsObject(ActiveSheet.ListObjects("DashBoardTable")) Then
    
        Range("A6").Select
        
        If ActiveSheet.FilterMode Then
            ActiveSheet.ShowAllData
        End If
        
        Range("A3", "M" & Cells(Rows.Count, 1).End(xlUp).row).Delete
        
    End If
    
    Range("A3:A50").RowHeight = 15
    
    Range("A1").Value = "DASHBOARD - PEDIDO PESQUISADO " & pedidoArr(0, 1)
    
    Range("A6").Value = "DATA PEDIDO"
    Range("B6").Value = "NUMERO"
    Range("C6").Value = "VENDEDOR"
    Range("D6").Value = "CADASTRADO"
    Range("E6").Value = "CLIENTE"
    Range("F6").Value = "PRODUTO"
    Range("G6").Value = "QUANTIDADE"
    Range("H6").Value = "UNID."
    Range("I6").Value = "VALOR"
    Range("J6").Value = "SITUAÇÃO"
    Range("K6").Value = "PEDIDO ATENÇÃO"
    Range("L6").Value = "OBSERVAÇÃO"
    Range("M6").Value = "DATA ATUALIZAÇÃO"
    
    'Estilo HEADER
    With Range("A6:M6")
        .Interior.Color = RGB(97, 183, 241)
        .Font.Bold = True
        .RowHeight = 30
        .Borders.Color = vbWhite
        .Borders(xlInsideVertical).Weight = xlThin
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Font.Size = 11
    End With
    
    iterator = 0
    
    'Incerir dados
    For i = 0 To UBound(pedidoArr)
        
        If pedidoArr(i, 1) <> "" Then
            
            'DATA PEDIDO
            Range("A" & 7 + i).Value = CDate(pedidoArr(i, 0))
            
            'NUMERO PEDIDO
            Range("B" & 7 + i).Value = CDbl(pedidoArr(i, 1))
            
            'VENDEDOR
            Range("C" & 7 + i).Value = pedidoArr(i, 3)
            
            'CADASTRADO
            Range("D" & 7 + i).Value = pedidoArr(i, 4)
            
            'CLIENTE
            Range("E" & 7 + i).Value = pedidoArr(i, 2)
            
            'PRODUTO
            Range("F" & 7 + i).Value = pedidoArr(i, 5)
            
            'QUANTIDADE
            Range("G" & 7 + i).Value = pedidoArr(i, 6)
            
            'UNID.
            Range("H" & 7 + i).Value = pedidoArr(i, 7)
            
            'VALOR
            With Range("I" & 7 + i)
                .Value = CDbl(pedidoArr(i, 8))
                .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            End With
            
            'SITUAÇÃO
            If pedidoArr(i, 9) = "EM ABERTO" Then
                With Range("J" & 7 + i)
                    .Value = pedidoArr(i, 9)
                    .Interior.Color = RGB(255, 189, 189)
                    .Font.Color = RGB(208, 0, 0)
                End With
            Else
                With Range("J" & 7 + i)
                    .Value = pedidoArr(i, 9)
                    .Interior.Color = RGB(169, 233, 169)
                    .Font.Color = RGB(25, 101, 25)
                End With
            End If
            
            If pedidoArr(i, 10) = "SIM" Then
                'PEDIDO ATENÇÃO
                With Range("K" & 7 + i)
                    .Value = pedidoArr(i, 10)
                    .Interior.Color = RGB(255, 189, 189)
                    .Font.Color = RGB(208, 0, 0)
                End With
            Else
                'PEDIDO ATENÇÃO
                With Range("K" & 7 + i)
                    .Value = pedidoArr(i, 10)
                    .Interior.Color = RGB(169, 233, 169)
                    .Font.Color = RGB(25, 101, 25)
                End With
            End If
            
            'OBSERVAÇÃO
            Range("L" & 7 + i).Value = pedidoArr(i, 11)
            
            'DATA ATUALIZAÇÃO
            Range("M" & 7 + i).Value = CDate(pedidoArr(i, 12))
            

            With Range("A" & 7 + i & ":M" & 7 + i)
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
                .Font.Size = 10
            End With
            
            iterator = iterator + 1
            
        End If
    Next i
    

    Columns("C:C").ColumnWidth = 18
    Range("D:D").ColumnWidth = 17
    Columns("E:E").EntireColumn.AutoFit
    Columns("F:F").EntireColumn.AutoFit
    Columns("G:G").ColumnWidth = 12
    Columns("H:H").EntireColumn.AutoFit
    Columns("I:I").EntireColumn.AutoFit
    Columns("J:J").ColumnWidth = 13
    Columns("K:K").EntireColumn.AutoFit
    Columns("L:L").EntireColumn.AutoFit
    Columns("M:M").EntireColumn.AutoFit
    
    'Range ("J:J").FormatConditions.Add(xlcellvalue, xlequal, ""
    
    'Transformar dados em tabela
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A6:M" & Cells(Rows.Count, 1).End(xlUp).row), , xlYes).Name = "DashBoardTable"
    ActiveSheet.ListObjects("DashBoardTable").TableStyle = ""
    
    
    
End Function

Function MostrarMateriaisProduzir(materiaisArr() As String)
    Dim i As Integer
    Dim colorLine As String
    
    'Deletar planilha antiga
    
    'Fazer header
    ThisWorkbook.Sheets("dashboard").Select
    
    If Range("A6").Value > 2000 Then

        Range("A3", "M" & Cells(Rows.Count, 1).End(xlUp).row).Delete
        
    ElseIf IsObject(ActiveSheet.ListObjects("DashBoardTable")) Then
    
        Range("A6").Select
        
        If ActiveSheet.FilterMode Then
            ActiveSheet.ShowAllData
        End If
        
        Range("A3", "M" & Cells(Rows.Count, 1).End(xlUp).row).Delete
        
    End If
    
    Range("A3:A50").RowHeight = 15
    
    Range("A1").Value = "DASHBOARD - PERFIS PARA PRODUZIR"
    
    Range("A6").Value = "DATA PEDIDO"
    Range("B6").Value = "NUMERO"
    Range("C6").Value = "CLIENTE"
    Range("D6").Value = "PERFIL"
    Range("E6").Value = "COR"
    Range("F6").Value = "QUANTIDADE"
    Range("G6").Value = "DATA ATUALIZAÇÃO"
    
    
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
    
    'Incerir dados
    For i = 0 To UBound(materiaisArr)
        
        If materiaisArr(i, 1) <> "" Then
            If colorLine = "white" Then colorLine = "blue" Else colorLine = "white"
            
            'DATA PEDIDO
            Range("A" & 7 + i).Value = CDate(materiaisArr(i, 0))
            
            'NUMERO PEDIDO
            Range("B" & 7 + i).Value = CDbl(materiaisArr(i, 1))
            
            'CLIENTE
            Range("C" & 7 + i).Value = materiaisArr(i, 6)
            
            'PERFIL
            Range("D" & 7 + i).Value = materiaisArr(i, 2)
            
            'COR
            Range("E" & 7 + i).Value = materiaisArr(i, 3)
            
            'QUANTIDADE
            Range("F" & 7 + i).Value = materiaisArr(i, 4)
            
            'ATUALIZAÇÃO
            Range("G" & 7 + i).Value = materiaisArr(i, 5)
            
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
    
    'Transformar em tabela
    
    'Arruma tamanho da coluna
    Range("A:A").ColumnWidth = 16
    Range("B:B").ColumnWidth = 14
    Columns("C:C").EntireColumn.AutoFit
    Columns("D:D").EntireColumn.AutoFit
    Range("E:E").ColumnWidth = 11
    Range("G:G").ColumnWidth = 19.5
    
    Columns("F:F").EntireColumn.AutoFit
    Range("G:G").ColumnWidth = 19.5
    
    'Transformar dados em tabela
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A6:G" & Cells(Rows.Count, 1).End(xlUp).row), , xlYes).Name = "DashBoardTable"
    ActiveSheet.ListObjects("DashBoardTable").TableStyle = ""
    
End Function

'Monta tabela com o array de pedidos resumidos
Function MostrarResumoPedidos(resumoPedidos() As String)
    Dim i As Integer, iterator As Integer, numeroPedidos As Integer
    Dim colorLine As String
    
    'Apagar tabela que estava antes
    'Formatar tabela de resumo pedidos
    
    numeroPedidos = ContaNumeroDePedidos(resumoPedidos)
    
    ThisWorkbook.Sheets("dashboard").Select
    
    If Range("A6").Value > 2000 Then

        Range("A3", "M" & Cells(Rows.Count, 1).End(xlUp).row).Delete
        
    ElseIf IsObject(ActiveSheet.ListObjects("DashBoardTable")) Then
    
        Range("A6").Select
        
        If ActiveSheet.FilterMode Then
            ActiveSheet.ShowAllData
        End If
        
        Range("A3", "M" & Cells(Rows.Count, 1).End(xlUp).row).Delete
        
    End If
    
    Range("A3:A50").RowHeight = 15
    
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
    Range("B:B").ColumnWidth = 16
    Range("C:C").ColumnWidth = 30
    Columns("D:D").EntireColumn.AutoFit
    Columns("E:E").EntireColumn.AutoFit
    Range("F:F").ColumnWidth = 19.5
    
    'Transformar dados em tabela
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A6:F" & Cells(Rows.Count, 1).End(xlUp).row), , xlYes).Name = "DashBoardTable"
    ActiveSheet.ListObjects("DashBoardTable").TableStyle = ""
    
    'Insere TOTAL PEDIDOS e VALOR TOTAL
    With Range("A4")
        .Value = numeroPedidos
        .HorizontalAlignment = xlCenter
    End With
    Range("B4").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Range("B4").Formula = "=SUM(DashBoardTable[VALOR])"
    
End Function

'Monta tabela com todos os itens dos pedidos
Function MostrarDadosTodosItensPedidos(pedidosEmAberto() As String, numeroDePedidos As Integer)
    Dim i As Integer
    Dim numeroPedido As String, colorLine As String

    'Antes de colocar os dados, converter eles para o formato certo
    'Estilizar cada pedido com uma cor
    
    ThisWorkbook.Sheets("dashboard").Select
    If Range("A6").Value > 2000 Then

        Range("A3", "M" & Cells(Rows.Count, 1).End(xlUp).row).Delete
        
    ElseIf IsObject(ActiveSheet.ListObjects("DashBoardTable")) Then
    
        Range("A6").Select
        
        If ActiveSheet.FilterMode Then
            ActiveSheet.ShowAllData
        End If
        
        Range("A3", "M" & Cells(Rows.Count, 1).End(xlUp).row).Delete
        
    End If

    Range("A3:A50").RowHeight = 15

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
    Range("B:B").ColumnWidth = 16
    Columns("C:C").EntireColumn.AutoFit
    Range("D:D").ColumnWidth = 15.5
    Columns("E:E").EntireColumn.AutoFit
    Columns("F:F").EntireColumn.AutoFit
    Range("G:G").ColumnWidth = 19.5
    
    'Transformar dados em tabela
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A6:G" & Cells(Rows.Count, 7).End(xlUp).row), , xlYes).Name = "DashBoardTable"
    ActiveSheet.ListObjects("DashBoardTable").TableStyle = ""
    
    'Insere TOTAL PEDIDOS e VALOR TOTAL
    With Range("A4")
        .Value = numeroDePedidos
        .HorizontalAlignment = xlCenter
    End With
    Range("B4").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Range("B4").Formula = "=SUM(DashBoardTable[VALOR])"
    
End Function
