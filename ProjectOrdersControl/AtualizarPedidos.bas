Attribute VB_Name = "AtualizarPedidos"
'@Folder("VBAProject")
Option Explicit

Sub AtualizarPedidos()
    Dim pedidosAberto() As String, novosPedidosArr() As String, pedidosFinalizadosArr() As String
    Dim pastaPedidos As Object
    Dim systemDate() As String, inputDateResult As String
    
    inputDateResult = InputDate
    
    If inputDateResult = "End" Then
        Exit Sub
    End If
    
    Set pastaPedidos = CreateObject("Scripting.FileSystemObject").getfolder("\\121.137.1.5\manutencao1\Lucas\12_Relatorios\2025\01_Relatorios Diarios\01_Relatorios TecSerp")
    
    pedidosAberto = PedidosEmAberto
    
    ReDim systemDate(2)
    
    systemDate = Split(Date, "/")

    On Error GoTo FolderNotFound
        AbrePlanilhaTecSerp systemDate, pastaPedidos, inputDateResult
    On Error GoTo 0
    
    novosPedidosArr = NovosPedidos(pedidosAberto())
    
    pedidosFinalizadosArr = pedidosFinalizados(pedidosAberto())
    
    FechaPlanilhaTecSerp
    
    If novosPedidosArr(0, 0) = "Vazio" And pedidosFinalizadosArr(0) = "Vazio" Then
        msgBox "A planilha já está atualizada.", vbInformation, "Sem novos dados"
        Exit Sub
    End If
    
    If MsgItensAtualizados(novosPedidosArr, pedidosFinalizadosArr) = 2 Then
        Exit Sub
    End If
    
    If novosPedidosArr(0, 0) <> "Vazio" Then
        
        AtualizaPedidosNovos novosPedidosArr
        
    End If
    
    If pedidosFinalizadosArr(0) <> "Vazio" Then
        AtualizaPedidosFinalizados pedidosFinalizadosArr()
    End If
    
    Exit Sub
FolderNotFound:
    msgBox "Verifique se a planilha de pedidos a faturar de hoje (" & systemDate(0) & "/" & systemDate(1) & "/" & systemDate(2) & ") foi gerada." & vbNewLine & vbNewLine & "Verifique a pasta em: " & pastaPedidos, _
    vbExclamation, "Planilha do TecSerp não encontrada"

End Sub




'------------ FUNCTIONS ------------

Function InputDate() As String
    Dim inputBoxAnswer As Variant
    Dim msgBoxAnswer As VbMsgBoxResult
    Dim ultimaData As String
    Dim dateInput As Date
    
    Sheets("base").Select
    
    Range("A3").Select
    
    If ActiveSheet.FilterMode Then
        ActiveSheet.ShowAllData
    End If
    
    'Perguntar se ele quer data limite ser igual a ultima data inserida
    ultimaData = Range("A" & Cells(Rows.Count, 1).End(xlUp).Row).Value
    
    msgBoxAnswer = msgBox("Quer pegar os pedidos até essa data: " & ultimaData & "?", vbYesNoCancel + vbQuestion, "Data de procura")
    
    If msgBoxAnswer = vbYes Then
        'Retornar a ultima data
        InputDate = ultimaData
        
        Exit Function
        
    ElseIf msgBoxAnswer = vbCancel Then
        InputDate = "End"
            
        Exit Function
    End If
    
    
    While True
        inputBoxAnswer = Application.InputBox("Colocar data limite?", "Data limite", , , , , , 2 + 4 + 16)
        
        If inputBoxAnswer = False Then
            InputDate = "End"
            
            Exit Function
        End If
        
        'Converter string pra data. Se não for, perguntar denovo
        On Error GoTo ErrorDate
            dateInput = CDate(inputBoxAnswer)
        On Error GoTo 0
        
        If CDbl(dateInput) < 45297 Then GoTo ErrorDate
        
        InputDate = inputBoxAnswer
        
        Exit Function
  
ErrorDate:
        msgBox "Digite uma data valida. Ex: 14/05/2025", vbOKOnly + vbExclamation, "Data incorreta"
    Wend
    
    
End Function

Function MsgItensAtualizados(pedidosNovos() As String, pedidosFinalizados() As String) As VbMsgBoxResult
    Dim item As Variant
    Dim msgItensParaAtualizar As String
    Dim msgBoxResult As VbMsgBoxResult
    
    msgItensParaAtualizar = "Os seguintes itens foram passados pra planilha:" & vbNewLine

    If pedidosNovos(0, 0) <> "Vazio" Then
        For item = 0 To UBound(pedidosNovos)
            If InStr(1, msgItensParaAtualizar, pedidosNovos(item, 1), vbTextCompare) = 0 Then
                msgItensParaAtualizar = msgItensParaAtualizar & vbNewLine & "PEDIDO NOVO:               " & pedidosNovos(item, 1)
            End If
        Next item
    End If
    
    msgItensParaAtualizar = msgItensParaAtualizar & vbNewLine
    
    If pedidosFinalizados(0) <> "Vazio" Then
        For Each item In pedidosFinalizados
            If InStr(1, msgItensParaAtualizar, item, vbTextCompare) = 0 Then
                msgItensParaAtualizar = msgItensParaAtualizar & vbNewLine & "PEDIDO FINALIZADO:     " & item
            End If
        Next item
    End If
    
    msgBoxResult = msgBox(msgItensParaAtualizar, vbInformation + vbOKCancel, "Itens para atualizar na planilha")
    
    MsgItensAtualizados = msgBoxResult
    
End Function


Function AtualizaPedidosNovos(pedidosNovos() As String)
    Dim ultimaLinha As String
    Dim i As Integer
    
    Range("Tabela3[[#Headers],[PRODUTO]]").Select
    ActiveSheet.ShowAllData
    
    For i = 0 To UBound(pedidosNovos)
    
        ultimaLinha = Range("A" & Cells(Rows.Count, 1).End(xlUp).Row).Offset(1, 0).Address
        
        'Data
        Range(ultimaLinha).Value = pedidosNovos(i, 0)
        
        'Numero do pedido
        Range(ultimaLinha).Offset(0, 1).Value = pedidosNovos(i, 1)
        
        'Cliente
        Range(ultimaLinha).Offset(0, 2).Value = pedidosNovos(i, 2)
        
        'Vendedor
        Range(ultimaLinha).Offset(0, 3).Value = pedidosNovos(i, 4)
        
        'Cadastrado
        Range(ultimaLinha).Offset(0, 4).Value = pedidosNovos(i, 5)
        
        'Produto
        Range(ultimaLinha).Offset(0, 5).Value = pedidosNovos(i, 3)
        
        'Quantidade
        Range(ultimaLinha).Offset(0, 6).Value = pedidosNovos(i, 6)
        
        'Unidade
        Range(ultimaLinha).Offset(0, 7).Value = pedidosNovos(i, 7)
        
        'Valor
        Range(ultimaLinha).Offset(0, 8).Value = CDbl(pedidosNovos(i, 8))
        
        'Situação
        Range(ultimaLinha).Offset(0, 9).Value = "EM ABERTO"
        
        If pedidosNovos(i, 3) = "DESPESA DE CORREIO" Or CDbl(pedidosNovos(i, 8)) = 0 Then
            'Pedido atenção
            Range(ultimaLinha).Offset(0, 10).Value = "NÃO"
            
            'Motivo
            Range(ultimaLinha).Offset(0, 11).Value = "Pedido sem valor."
        Else
            'Pedido atenção
            Range(ultimaLinha).Offset(0, 10).Value = "SIM"
            
            'Motivo
            Range(ultimaLinha).Offset(0, 11).Value = "Perguntar para vendedoras."
        End If
        
        
        
        'Data atualização
        Range(ultimaLinha).Offset(0, 12).Value = Date
 
    Next i

End Function

Function AtualizaPedidosFinalizados(pedidosFinalizados() As String)
    Dim rng As Variant
    Dim pedido As Variant

    For Each pedido In pedidosFinalizados
        For Each rng In ActiveSheet.AutoFilter.Range.Offset(1, 0).Columns("B").SpecialCells(xlCellTypeVisible)
            If CStr(rng) = pedido Then
            
                'SITUAÇAO
                Range(rng.Address).Offset(0, 8).Value = "FINALIZADO"
                
                'ATENÇÃO
                Range(rng.Address).Offset(0, 9).Value = "NÃO"
                
                'MOTIVO
                Range(rng.Address).Offset(0, 10).Value = "Pedido sumiu do sistema."
                
                'DATA ATUALIZAÇÃO
                Range(rng.Address).Offset(0, 11).Value = Date
                        
            End If
        Next rng
    Next pedido
    
End Function

Function pedidosFinalizados(meusPedidos() As String) As String()
    Dim item As Variant, rng As Variant
    Dim arrReturn() As String
    Dim i As Integer
    
    i = 0

    For Each item In meusPedidos
    
            For Each rng In ActiveSheet.AutoFilter.Range.Offset(1, 0).Columns("E").SpecialCells(xlCellTypeVisible)
                If CStr(rng) <> "" Then
                
                    If CStr(rng) = item Then
                    
                        GoTo NextI
                        
                    End If
                    
                End If
            Next rng
            
            If item <> "" Then
                ReDim Preserve arrReturn(i)
            
                arrReturn(i) = item
                
                i = i + 1
            End If
NextI:
        
    Next item
    
    If i = 0 Then
        ReDim arrReturn(1)
        arrReturn(0) = "Vazio"
        pedidosFinalizados = arrReturn
        
        Exit Function
    End If
    
    pedidosFinalizados = arrReturn
End Function

Function NovosPedidos(meusPedidos() As String) As String()
    Dim returnArray() As String, item As Variant, novoPedido As Variant
    Dim rng As Range
    Dim i As Integer, arrSize As Integer
    
    i = 0
    
    ReDim returnArray(i, 8)
    
    'Verifica a quantidade de pedidos novos
    For Each rng In ActiveSheet.AutoFilter.Range.Offset(1, 0).Columns("E").SpecialCells(xlCellTypeVisible)
        
        If rng.Value <> "" Then
            
            For Each item In meusPedidos
                
                If rng = item Then
                    
                    GoTo NextIteration
                    
                End If
                
            Next item
            
            'Loopar pelo range, cadastrando os dados em um array
            arrSize = arrSize + Range(rng.Address, rng.End(xlUp).Offset(1, 0).Address).Count
        End If
        
NextIteration:
        
    Next rng
    
    If arrSize = 0 Then
        ReDim returnArray(1, 1)
        returnArray(0, 0) = "Vazio"
        NovosPedidos = returnArray
        
        Exit Function
    End If
    
    ReDim returnArray(arrSize - 1, 8)
    
    
    For Each rng In ActiveSheet.AutoFilter.Range.Offset(1, 0).Columns("E").SpecialCells(xlCellTypeVisible)
        
        If rng.Value <> "" Then
            
            'Se o numero do pedido(rng) estiver no array meusPedidos, vai pra proxima iteracao
            For Each item In meusPedidos
                
                If rng = item Then
                    
                    GoTo NextIt
                    
                End If
                
            Next item



            For Each novoPedido In Range(rng.Address, rng.End(xlUp).Offset(1, 0).Address)

                'Data do pedido
                returnArray(i, 0) = Range(novoPedido.Offset(0, -4).Address).Value
                
                'Numero do pedido
                returnArray(i, 1) = rng
                
                'Cliente
                returnArray(i, 2) = Range(novoPedido.Offset(0, 1).Address).Value
                
                'Produto
                returnArray(i, 3) = Range(novoPedido.Offset(0, 7).Address).Value
                
                'Vendedor
                returnArray(i, 4) = Range(novoPedido.Offset(0, 3).Address).Value
                
                'Cadastrado
                returnArray(i, 5) = Range(novoPedido.Offset(0, 4).Address).Value
                
                'Quantidade
                returnArray(i, 6) = Range(novoPedido.Offset(0, 9).Address).Value
                
                'Unidade
                returnArray(i, 7) = Range(novoPedido.Offset(0, 10).Address).Value
                
                'Valor
                returnArray(i, 8) = Range(novoPedido.Offset(0, 8).Address).Value
                
                i = i + 1
            Next novoPedido

        End If
        
NextIt:
        
    Next rng
    
    NovosPedidos = returnArray
    
End Function


Function AbrePlanilhaTecSerp(systemDate() As String, pastaRaiz As Object, filterDate As String)
    Dim pasta As String, arquivo As String
    
    
    'Abre a planilha do tecserp de acordo com a data do sistema
    pasta = Dir(pastaRaiz & "\" & Right(systemDate(2), 2) & "_" & systemDate(1) & "_*", vbDirectory)
    
    arquivo = Dir(pastaRaiz & "\" & pasta & "\" & Right(systemDate(2), 2) & "_" & systemDate(1) & "_" & systemDate(0) & "_Molducolor A FATURAR" & "*" & ".xlsx")
    
    Workbooks.Open fileName:=pastaRaiz & "\" & pasta & "\" & arquivo
    
    ActiveWorkbook.Sheets("Macro").Select
    
    'Criar uma tabela e filtra essa tabela com uma data
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1:" & "AJ" & Cells(Rows.Count, 1).End(xlUp).Row), , xlYes).Name _
        = "Tabela1"
    
    ActiveSheet.ListObjects("Tabela1").TableStyle = ""
    
        
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=1, Criteria1:= _
        "<=" & CLng(CDate(filterDate)), Operator:=xlFilterValues
    
End Function

Function FechaPlanilhaTecSerp()

    ActiveWorkbook.Close SaveChanges:=False
    
End Function

Function PedidosEmAberto() As String()
    Dim rangeNumPedidos() As String, rng As Range
    Dim arrPedidos() As String
    Dim iterator As Integer, i As Integer, numPedidos As Integer
    Dim item As Variant
    
    numPedidos = 0
    iterator = 0
    
    Sheets("base").Select
    
    Range("A3").Select
    
    If ActiveSheet.FilterMode Then
        ActiveSheet.ShowAllData
    End If
    
    ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=10, Criteria1:= _
        "EM ABERTO"

    i = 0
    
    ReDim arrPedidos(i)

    For Each rng In Worksheets("base").AutoFilter.Range.Offset(1, 0).Columns("B").SpecialCells(xlCellTypeVisible)
        
        
        For iterator = 0 To i - 1
            
            If arrPedidos(iterator) = CStr(rng) Then
                
                GoTo nxt
            End If
            
            
        Next iterator
        
        ReDim Preserve arrPedidos(i)
        arrPedidos(i) = CStr(rng)
        
        i = i + 1
        
nxt:
    Next rng
    
    
    
    PedidosEmAberto = arrPedidos()
    
End Function

