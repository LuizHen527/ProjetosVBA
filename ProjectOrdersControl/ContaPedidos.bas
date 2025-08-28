Attribute VB_Name = "ContaPedidos"
'@Folder("VBAProject")
Option Explicit

Sub ContaPedidosEmAberto()

    Application.ScreenUpdating = False
    
    ActiveSheet.Shapes.Range("pedido_menu").Visible = Not ActiveSheet.Shapes.Range("pedido_menu").Visible
    
    Dim pedidosEmAbertoo() As String
    Dim i As Integer

    pedidosEmAbertoo = GetPedidosEmAberto
    
    MontaTabela pedidosEmAbertoo
    
    Application.ScreenUpdating = True
    
    'Debug.Print GetPedidosEmAberto
    
End Sub


'---------- FUNCTIONS ----------

Function MontaTabela(pedidosArray() As String)
    Dim i As Integer, iMes As Integer, iCount As Integer, ano As Byte, iterator As Byte, row As Byte
    Dim pedidoSemValor As Integer, pedidoComValor As Integer, totalPedidos As Integer, totalValor As Double
    Dim initialRange As String
    Dim dataArr() As String, dataArrComp() As String, anosContados() As String
    
    initialRange = Range("A6").Address

    Sheets("dashboard").Select
    
    If IsObject("DashBoardTable") Then
        Range("A6").Select
        
        If ActiveSheet.FilterMode Then
            ActiveSheet.ShowAllData
        End If
    End If
    
    Range("A3", "M" & Cells(Rows.Count, 1).End(xlUp).row).Delete
        
    iterator = 0
    
    ReDim anosContados(10)
    
    'Loop que define o ano
    For i = 0 To UBound(pedidosArray)
        If pedidosArray(i, 0) = "" Then GoTo NextYear
        
        'A cada ano, descer uma certa quantidade de celulas
        
        dataArr = Split(pedidosArray(i, 0), "/")
        
        
        'Se o ano já foi analizado, pular.
        For ano = 0 To UBound(anosContados)
            If anosContados(ano) = dataArr(2) Then GoTo NextYear
        Next ano
        
        'Montar cabeçalho
    
        Range(initialRange).Value = dataArr(2)
        
        Range(initialRange).Offset(0, 1).Value = "COM VALOR"
        
        Range(initialRange).Offset(0, 2).Value = "SEM VALOR"
        
        Range(initialRange).Offset(0, 3).Value = "TOTAL PEDIDOS"
        
        Range(initialRange).Offset(0, 4).Value = "VALOR TOTAL"
        
        With Range(Range(initialRange).Address, Range(initialRange).Offset(0, 4).Address)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Interior.Color = RGB(120, 193, 243)
            .RowHeight = 19
            .ColumnWidth = 14.2
            .Borders.Color = vbWhite
            .Borders.Weight = xlThin
        End With
        
        row = 1
        
        'Loop que define o mes
        For iMes = 1 To 12
            pedidoSemValor = 0
            pedidoComValor = 0
            totalPedidos = 0
            totalValor = 0
            
            
            
            'Loop procurando todos com esse ano e esse mes
            For iCount = 0 To UBound(pedidosArray)
                
                If pedidosArray(iCount, 0) = "" Then GoTo NextIteration
                
                dataArrComp = Split(pedidosArray(iCount, 0), "/")
            
                'Se for do mesmo ano e do mesmo mes. Conta
                If dataArr(2) = dataArrComp(2) And _
                iMes = dataArrComp(1) Then
                    
                    'Contar se é com valor ou sem valor. Um dos dois apenas
                    If pedidosArray(iCount, 2) = 0 Then
                        pedidoSemValor = pedidoSemValor + 1
                    Else
                        pedidoComValor = pedidoComValor + 1
                    End If
                        
                    'Somar ao total de pedidos do mes
                    totalPedidos = totalPedidos + 1
                    
                    'Somar ao total de dinheiro do mes
                    totalValor = totalValor + pedidosArray(iCount, 2)
                    
                End If
            
NextIteration:
            Next iCount
            
            If pedidoComValor > 0 Or _
            pedidoSemValor > 0 Or _
            totalPedidos > 0 Or _
            totalValor > 0 Then
            
                Range(initialRange).Offset(row, 0).Value = UCase(CStr(MonthName(iMes)))
                
                With Range(initialRange).Offset(row, 0)
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .Font.Bold = True
                    .Interior.Color = RGB(155, 232, 216)
                    .RowHeight = 19
                    .Borders.Color = vbWhite
                    .Borders.Weight = xlThin
                End With
                
                Range(initialRange).Offset(row, 1).Value = pedidoComValor
                
                Range(initialRange).Offset(row, 2).Value = pedidoSemValor
                
                Range(initialRange).Offset(row, 3).Value = totalPedidos
                
                Range(initialRange).Offset(row, 4).Value = totalValor
                Range(initialRange).Offset(row, 4).NumberFormat = "_-$ * #,##0.00_-;-$ * #,##0.00_-;_-$ * ""-""??_-;_-@_-"
                '= "_-$ * #,##0.00_-;-$ * #,##0.00_-;_-$ * " - "??_-;_-@_-"
                
                With Range(Range(initialRange).Offset(row, 1), Range(initialRange).Offset(row, 3))
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With
                
                row = row + 1
            End If
            
        Next iMes
        
        anosContados(iterator) = dataArr(2)
        
        initialRange = Range(initialRange).Offset(2 + row, 0).Address
        
        iterator = iterator + 1
        
NextYear:
    Next i
    
    'Contagem por mes
    'Separar string de data
    'Conta
    
    
End Function

Function GetPedidosEmAberto() As String()
    Dim rng As Range
    Dim pedidosEmAberto() As String
    Dim i As Integer, iterador As Integer
    
    
    'Retornar array
    
    Sheets("base").Activate
    
    Range("A3").Select
    
    If ActiveSheet.FilterMode Then
        ActiveSheet.ShowAllData
    End If
    
    ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=10, Criteria1:= _
        "EM ABERTO"
    
    ReDim pedidosEmAberto(Range("A1:A" & Cells(Rows.Count, 1).End(xlUp).row).SpecialCells(xlCellTypeVisible).Count, 4)
    
    i = 0
    
    For Each rng In Worksheets("base").AutoFilter.Range.Offset(1, 0).Columns("B").SpecialCells(xlCellTypeVisible)
        
        'Salvar colunas: DATA, NUMERO, SITUAÇÂO, FINALIZADO
        
        'Verifica se não é duplicado
        
        If i > 0 Then
            For iterador = 0 To i
                If pedidosEmAberto(iterador, 1) = rng.Value Then
                    
                    'Soma valor
                    If pedidosEmAberto(iterador, 2) = "" Then
                        pedidosEmAberto(iterador, 2) = Range(rng.Address).Offset(0, 7).Value
                    Else
                        pedidosEmAberto(iterador, 2) = CStr(CDbl(pedidosEmAberto(iterador, 2)) + CDbl(Range(rng.Address).Offset(0, 7).Value))
                    End If
                    
                    GoTo NextIteration
                End If
            Next iterador
        End If
        
        'DATA
        pedidosEmAberto(i, 0) = Range(rng.Address).Offset(0, -1).Value
        
        'NUMERO PEDIDO
        pedidosEmAberto(i, 1) = rng.Value
        
        'VALOR
        pedidosEmAberto(i, 2) = Range(rng.Address).Offset(0, 7).Value
        
        'SITUAÇÂO
        pedidosEmAberto(i, 3) = Range(rng.Address).Offset(0, 8).Value
        
        'FINALIZADO
        pedidosEmAberto(i, 4) = Range(rng.Address).Offset(0, 9).Value
        
        'Debug.Print pedidosEmAberto(i, 0) & ";" & pedidosEmAberto(i, 1) & ";" & pedidosEmAberto(i, 2) & ";" & pedidosEmAberto(i, 3)
        
        i = i + 1
        
NextIteration:
    Next rng
    
    GetPedidosEmAberto = pedidosEmAberto
    
    
End Function
