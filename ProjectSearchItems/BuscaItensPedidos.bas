Attribute VB_Name = "BuscaItensPedidos"
'@Folder("VBAProject")
Option Explicit

Sub Main()
    Dim pedidosAberto() As String, novosPedidosArr() As String, pedidosFinalizadosArr() As String, systemDate() As String
    Dim pastaPedidos As Object
    Dim numerosDePedidos() As String, numeroPedido As Range, numerosRange As Range, iterator As Integer
    Dim informacoesPedidos() As String, i As Variant
    Dim countPrintedInfo As Integer
    
    countPrintedInfo = 2
    
    Application.ScreenUpdating = False
    
    '-------- Salva numeros de pedidos em um array --------
    
    Set numerosRange = Range("A2:A" & Cells(Rows.Count, 1).End(xlUp).row)
    
    ReDim numerosDePedidos(numerosRange.Count)
    
    For iterator = 1 To numerosRange.Count
        numerosDePedidos(iterator) = CStr(Cells(iterator + 1, 1).Value)
    Next iterator
    
    
    '-------- Abre planilha TecSerp --------
    
    Set pastaPedidos = CreateObject("Scripting.FileSystemObject").getfolder("\\121.137.1.5\manutencao1\Lucas\12_Relatorios\2025\01_Relatorios Diarios\01_Relatorios TecSerp")
    
    ReDim systemDate(2)
    
    systemDate = Split(Date, "/")
    
    On Error GoTo FolderNotFound
        AbrePlanilhaTecSerp systemDate, pastaPedidos
    On Error GoTo 0
    
    
    '-------- Procura e salva pedidos --------
    informacoesPedidos = ProcurarPedidos(numerosDePedidos)
    
    ActiveWorkbook.Close SaveChanges:=False
    
    ThisWorkbook.Activate
    
    Sheets.Add
    
    'Titulos das colunas
    Cells(1, 1).Value = "Numero"
    Cells(1, 2).Value = "Identificacao"
    Cells(1, 3).Value = "Cor"
    Cells(1, 4).Value = "Quantidade"
    Cells(1, 5).Value = "Comprimento"
    Cells(1, 6).Value = "Altura"
    
    With Range(Cells(1, 1), Cells(1, 6))
        .Interior.Color = RGB(153, 204, 255)
        .Columns.AutoFit
    End With
    
    Range(Cells(1, 3).Address).Columns.ColumnWidth = 16
    
    For i = 0 To UBound(informacoesPedidos)
        If Not informacoesPedidos(i, 0) = "" Then
            'Debug.Print informacoesPedidos(i, 0) & " "
            Cells(countPrintedInfo, 1).Value = informacoesPedidos(i, 0)
            
            Cells(countPrintedInfo, 2).Value = informacoesPedidos(i, 1)
            
            Cells(countPrintedInfo, 3).Value = informacoesPedidos(i, 2)
            
            Cells(countPrintedInfo, 4).Value = informacoesPedidos(i, 3)
            
            Cells(countPrintedInfo, 5).Value = informacoesPedidos(i, 4)
            
            Cells(countPrintedInfo, 6).Value = informacoesPedidos(i, 5)
            
            countPrintedInfo = countPrintedInfo + 1
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    Exit Sub
FolderNotFound:
    MsgBox "Verifique se a planilha de pedidos a faturar de hoje (" & systemDate(0) & "/" & systemDate(1) & "/" & systemDate(2) & ") foi gerada." & vbNewLine & vbNewLine & "Verifique a pasta em: " & pastaPedidos, _
    vbExclamation, "Planilha do TecSerp não  encontrada"
    
    Application.ScreenUpdating = True
End Sub


Function AbrePlanilhaTecSerp(systemDate() As String, pastaRaiz As Object)
    Dim pasta As String, arquivo As String, horarioMockado As String, horarioQualquer As String
    
    horarioMockado = " 14.28"
    horarioQualquer = "*"
    
    'Abre a planilha do tecserp de acordo com a data do sistema
    pasta = Dir(pastaRaiz & "\" & Right(systemDate(2), 2) & "_" & systemDate(1) & "_*", vbDirectory)
    
    arquivo = Dir(pastaRaiz & "\" & pasta & "\" & Right(systemDate(2), 2) & "_" & systemDate(1) & "_" & systemDate(0) & "_Molducolor A FATURAR" & horarioQualquer & ".xlsx")
    
    Workbooks.Open Filename:=pastaRaiz & "\" & pasta & "\" & arquivo
    
    ActiveWorkbook.Sheets("Macro").Select
    
End Function


Function ProcurarPedidos(numerosParaProcurar() As String) As String()
    Dim numerosTecSerp As Range, iterator As Integer, numeroPedidoTecSerp As Range
    Dim achouPedido As Boolean, itensPedido() As String
    Dim returnArray() As String, rowReturnArray As Integer, returnArrayIterator As Integer
    
    Set numerosTecSerp = Range("E2:E" & Cells(Rows.Count, 5).End(xlUp).row)
    
    rowReturnArray = 0
    
    ReDim returnArray(5000, 5)
    
    'Loop por cada numero
    
    For iterator = 0 To numerosTecSerp.Count
        Set numeroPedidoTecSerp = Cells(iterator + 2, 5)
    
        If Not numeroPedidoTecSerp.Value = "" Then
            'Chama função que verifica pedido
            achouPedido = VerificarPedido(numeroPedidoTecSerp.Value, numerosParaProcurar)
            
            If achouPedido = True Then
                'Fazer função que retorna informações do pedido
                itensPedido = CapturarDadosPedido(numeroPedidoTecSerp)
                
                'Guardar itens retornados no array de todos os pedidos
                For returnArrayIterator = 0 To UBound(itensPedido) - 1
                    'Numero pedido
                    returnArray(rowReturnArray, 0) = itensPedido(returnArrayIterator, 0)
                    
                    'Identificacao
                    returnArray(rowReturnArray, 1) = itensPedido(returnArrayIterator, 1)
                    
                    'Cor
                    returnArray(rowReturnArray, 2) = itensPedido(returnArrayIterator, 2)
                    
                    'Quantidade
                    returnArray(rowReturnArray, 3) = itensPedido(returnArrayIterator, 3)
                    
                    'Comprimento
                    returnArray(rowReturnArray, 4) = itensPedido(returnArrayIterator, 4)
                    
                    'Altura
                    returnArray(rowReturnArray, 5) = itensPedido(returnArrayIterator, 5)
                    
                    rowReturnArray = rowReturnArray + 1
                Next
            End If
            
        End If
    Next iterator
    
    ProcurarPedidos = returnArray
End Function

'Verifica se um numero de pedido esta na lista
Function VerificarPedido(numeroParaProcurar As String, listaPedidos() As String) As Boolean
    Dim numero As Variant
    
    For Each numero In listaPedidos
    
        If numero = numeroParaProcurar Then
            VerificarPedido = True
            Exit Function
        End If
        
    Next numero
    
    VerificarPedido = False
End Function

'Pega todos os itens do pedido e retorna em forma de array
Function CapturarDadosPedido(celulaPedido As Range) As String()
    Dim item As Variant, itensPedido As Range, row As Integer, returnArray() As String
    Dim celula As Range
    
    'Essa função vai retornar todas as informacoes do pedido em um array 2D
    'Na outra funcao, eu vou ter que loopar por esse array e colocar ele no array de retorno da outra funcao
    
    'CODIGO UTIL: esse codigo seleciona as celulas vazias acima da celula do primeiro parametro.
    'Range(celulaPedido, celulaPedido.End(xlUp).Offset(1, 0)).Select
    
    Set celula = celulaPedido.End(xlUp).Offset(1, 0)

    
    
    While Not celula = "" And Not celula = celulaPedido
    
         Set celula = celula.Offset(1, 0)
    
    Wend
    
    'Set itensPedido = Range(celulaPedido, celulaPedido.End(xlUp).Offset(1, 0))
    
    Set itensPedido = Range(celulaPedido, celula)
    
    ReDim returnArray(itensPedido.Count, 5)
    
    row = 0
    
    For Each item In itensPedido
    
        'Numero do pedido
        'Debug.Print celulaPedido
        
        returnArray(row, 0) = CStr(celulaPedido.Value)
        
        'Identificação
        'Debug.Print item.Offset(0, 12).Value
        
        returnArray(row, 1) = CStr(item.Offset(0, 12).Value)
        
        'Cor
        'Debug.Print item.Offset(0, 16).Value
        
        returnArray(row, 2) = CStr(item.Offset(0, 16).Value)
        
        'Quantidade
        'Debug.Print item.Offset(0, 9).Value
        
        returnArray(row, 3) = CStr(item.Offset(0, 9).Value)
        
        'Comprimento
        'Debug.Print item.Offset(0, 18).Value
        
        returnArray(row, 4) = CStr(item.Offset(0, 18).Value)
        
        'Altura
        'Debug.Print item.Offset(0, 19).Value
        
        returnArray(row, 5) = CStr(item.Offset(0, 19).Value)
        
        row = row + 1
    Next item
    
    CapturarDadosPedido = returnArray
End Function
