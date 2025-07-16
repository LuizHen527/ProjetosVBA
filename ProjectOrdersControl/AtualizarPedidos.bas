Attribute VB_Name = "AtualizarPedidos"
'@Folder("VBAProject")
Option Explicit

Sub AtualizarPedidos()
    Dim inputBoxAnswer As Variant
    Dim pedidosAberto() As String, pedidosSistema() As String, novosPedidosArr() As String, pedidosFinalizadosArr() As String
    Dim pastaPedidos As Object
    Dim systemDate() As String
    
    Dim item As Variant
    
    Set pastaPedidos = CreateObject("Scripting.FileSystemObject").getfolder("C:\Users\Molducolor7\Desktop\TesteMacroPedidos")
    
    'inputBoxAnswer = Application.InputBox("Colocar data limite? Digite 0 para não colocar data limite", "Data limite", , , , , , 2 + 4 + 16)
    
    pedidosAberto = PedidosEmAberto
    
    
    ReDim systemDate(2)
    systemDate = Split(Date, "/")

    On Error GoTo FolderNotFound
        AbrePlanilhaTecSerp systemDate, pastaPedidos
    On Error GoTo 0
    
    pedidosSistema = PedidosTecserp()
    
    novosPedidosArr = NovosPedidos(pedidosAberto())
    
    pedidosFinalizadosArr = PedidosFinalizado(pedidosAberto())
    
    
    For item = 0 To UBound(novosPedidosArr)
        
        Debug.Print "PEDIDO NOVO:" & novosPedidosArr(item, 1)
    Next item
    
    For Each item In pedidosFinalizadosArr
        Debug.Print "PEDIDO FINALIZADO:" & item
    Next item

    'Função pra encontrar pedidos novos
    'Função pra encontrar pedidos apagados
    
    
    'Comparar pedidos da minha planilha com os do sistema
        'FEITO: Loopar pelo range de numeros de pedidos. Salvar numeros em array
        'FEITO: Compara os numeros pra pegar pedidos novos.
        'FEITO: Se tiver pedido novo, salva dados dele em um array
        
        'Fecha planilha tecserp
        'Compara numeros pra saber os finalizados
        'O que for finalizado, atualizar o status do pedido na minha planilha
        'Colar os pedidos novos
        
        'Colocar retorno na função de novos pedidos
    
    Exit Sub
FolderNotFound:
    MsgBox "Verifique se a planilha de pedidos a faturar de hoje (" & systemDate(0) & "/" & systemDate(1) & "/" & systemDate(2) & ") foi gerada." & vbNewLine & vbNewLine & "Verifique a pasta em: " & pastaPedidos, _
    vbExclamation, "Planilha do TecSerp não encontrada"

End Sub

Function PedidosFinalizado(meusPedidos() As String) As String()
    Dim item As Variant, rng As Variant
    Dim arrReturn() As String
    Dim i As Integer
    
    i = 0

    'Pegar pedidos finalizados e colocar em um array
    'Loopar por pedidos do sistema e procurar os meus pedidos com os que estão la
    'O que não tiver na planilha, é pedido finalizado
    'Salvar numero do pedido em um array e retornar
    For Each item In meusPedidos
    
            For Each rng In ActiveSheet.AutoFilter.Range.Offset(1, 0).Columns("E").SpecialCells(xlCellTypeVisible)
                If rng.Value <> "" Then
                
                    If rng = item Then
                    
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
    
    PedidosFinalizado = arrReturn
End Function

Function NovosPedidos(meusPedidos() As String) As String()
    Dim returnArray() As String, item As Variant, novoPedido As Variant
    Dim rng As Range
    Dim i As Integer, arrSize As Integer
    
    i = 0
    
    ReDim returnArray(i, 8)
    
    For Each rng In ActiveSheet.AutoFilter.Range.Offset(1, 0).Columns("E").SpecialCells(xlCellTypeVisible)
        
        If rng.Value <> "" Then
            
            For Each item In meusPedidos
                'Debug.Print rng.Value & "/" & item
                
                If rng = item Then
                
                    'Debug.Print "Encontrado: " & item
                    
                    GoTo NextIteration
                    
                End If
                
            Next item
            'Loopar pelo range, cadastrando os dados em um array
            
            'Debug.Print "Pedido Novo: " & rng & " " & Range(rng.Address, rng.End(xlUp).Offset(1, 0).Address).Address
            arrSize = arrSize + Range(rng.Address, rng.End(xlUp).Offset(1, 0).Address).Count
        End If
        
NextIteration:
        
    Next rng
    
    
    ReDim returnArray(arrSize - 1, 8)
    
    
    For Each rng In ActiveSheet.AutoFilter.Range.Offset(1, 0).Columns("E").SpecialCells(xlCellTypeVisible)
        
        If rng.Value <> "" Then
            
            For Each item In meusPedidos
                'Debug.Print rng.Value & "/" & item
                
                If rng = item Then
                
                    'Debug.Print "Encontrado: " & item
                    
                    GoTo NextIt
                    
                End If
                
            Next item
            'Loopar pelo range, cadastrando os dados em um array
            
            'Debug.Print "Pedido Novo: " & rng & " " & Range(rng.Address, rng.End(xlUp).Offset(1, 0).Address).Address
            
            
            
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
                
'                Debug.Print "DATA:      " & returnArray(i, 0)
'                Debug.Print "NUMERO:    " & returnArray(i, 1)
'                Debug.Print "CLIENTE:   " & returnArray(i, 2)
'                Debug.Print "PRODUTO:   " & returnArray(i, 3)
'                Debug.Print "VENDEDOR:  " & returnArray(i, 4)
'                Debug.Print "CADAST:    " & returnArray(i, 5)
'                Debug.Print "QUANT:     " & returnArray(i, 6)
'                Debug.Print "UNID:      " & returnArray(i, 7)
'                Debug.Print "VALOR:     " & returnArray(i, 8)
                
                i = i + 1
            Next novoPedido
            

            
            
            
        End If
        
NextIt:
        
    Next rng
    
    NovosPedidos = returnArray
    
End Function



Function PedidosTecserp() As String()
    Dim returnArray() As String
    Dim rng As Range
    Dim i As Integer
    
    i = 0
    
    'Criar uma tabela e filtra essa tabela com uma data
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1:" & "AJ" & Cells(Rows.Count, 1).End(xlUp).Row), , xlYes).Name _
        = "Tabela1"
    
    ActiveSheet.ListObjects("Tabela1").TableStyle = ""
        
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=1, Criteria1:= _
        "<=" & CLng(CDate("22/04/2025")), Operator:=xlFilterValues
        
        

    'Pega os numeros de pedidos e salva no array que a função retorna
    For Each rng In ActiveSheet.AutoFilter.Range.Offset(1, 0).Columns("E").SpecialCells(xlCellTypeVisible)
        
        If rng.Value <> "" Then
            ReDim returnArray(i)
            returnArray(i) = rng.Value
            
            'Redim no array
            
            i = i + 1
        End If
        
    Next rng
    
    PedidosTecserp = returnArray
    
End Function



Function AbrePlanilhaTecSerp(systemDate() As String, pastaRaiz As Object)
    Dim pasta As String, arquivo As String
    
    
    'Abre a planilha do tecserp de acordo com a data do sistema
    pasta = Dir(pastaRaiz & "\" & Right(systemDate(2), 2) & "_" & systemDate(1) & "_*", vbDirectory)
    
    arquivo = Dir(pastaRaiz & "\" & pasta & "\" & Right(systemDate(2), 2) & "_" & systemDate(1) & "_" & systemDate(0) & "_Molducolor A FATURAR" & "*" & ".xlsx")
    
    Workbooks.Open fileName:=pastaRaiz & "\" & pasta & "\" & arquivo
    
    ActiveWorkbook.Sheets("Macro").Select
    
End Function



Function PedidosEmAberto() As String()
    Dim rangeNumPedidos() As String, rng As Range
    Dim arrPedidos() As String
    Dim iterator As Integer, i As Integer, numPedidos As Integer
    Dim item As Variant
    
    numPedidos = 0
    iterator = 0
    
    Sheets("base").Select
    
    'Filtra pedidos em aberto
    ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=10, Criteria1:= _
        "EM ABERTO"

    i = 0
    
    ReDim arrPedidos(i)
    
    'Loopar pelo range conta quantos pedidos tem
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

