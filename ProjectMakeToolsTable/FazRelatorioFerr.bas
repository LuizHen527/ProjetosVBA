Attribute VB_Name = "FazRelatorioFerr"
Option Explicit

Sub RelatorioFerramentas()
    'Ao usar a macro:
        'deixe a planilha que você quer copiar selecionada
        'a planilha precisa estar com esse formato de nome:
        'Mes_Numero do mes_Utimos dois digitos do ano
        'Deve ficar assim: Mar_3_25
    
        'Antes precisa corrigir os nomes
        
        'Capturar erros:
            'Erro de rodar sem tem o historico aberto
            'Erro de rodar sem ter colocado nome certo na planilha
    
    Dim data() As Variant, processedData As Variant, perfil As Variant
    Dim fileName As String, arrDate() As String, inicialRange() As String, rowAddress As String, empresas() As String, nome As Variant, empresa As Variant
    Dim numRows As Integer, colInt As Integer, rowInt As Integer, copyInt As Integer, x As Integer, numRowsArray As Integer, columnIcr As Integer, numRowsNames As Integer
    
    '---- Inicializando variaveis ----
    
    columnIcr = 3
    numRowsArray = 0
    x = 2
    fileName = ThisWorkbook.Name
    arrDate = Split(ActiveSheet.Name, "_")
    numRowsNames = Range("C4", "C" & Cells(Rows.Count, 1).End(xlUp).Row).Rows.Count
    ReDim empresas(5)
    empresas() = Split("MOLDUCOLOR,ALUMITEC,POLLUX,ALHENA,EXTERNO", ",")
    
    
    Workbooks("HISTÓRICO PRODUÇÃO 2022-2024_V5.xlsm").Activate
    
    Worksheets("02_Correção Nomes").Select
    numRowsNames = Range("C4", "C" & Cells(Rows.Count, 1).End(xlUp).Row).Rows.Count
    
    Worksheets("01_Base").Select
    
    'Tira filtros aplicados
    ActiveWorkbook.Worksheets("01_Base").AutoFilter.Sort.SortFields.Clear
 
    'Filtra os dados da base pela data, de acordo com o nome da planilha(Ex:Mar_1_25)
    ActiveSheet.Range("$A$3:$BA$4805").AutoFilter Field:=1, Operator:= _
    xlFilterValues, Criteria2:=Array(1, arrDate(1) & "/10/20" & arrDate(2))
    
'------------------------- SALVANDO DADOS NO ARRAY -------------------------
    
    'Conta linha visiveis
    numRows = Range("A3", "A" & Cells(Rows.Count, 1).End(xlUp).Row).Rows.SpecialCells(xlCellTypeVisible).Count - 1
    
    'Endereço da ultima celula da coluna A
    inicialRange = Split(Range("A" & Cells(Rows.Count, 1).End(xlUp).Row).Address, "$")
    
    ReDim data(numRows, 7) As Variant
    
    'Salva dados da coluna DATA
    For rowInt = 1 To numRows
        rowAddress = (inicialRange(2) - numRows) + rowInt
        
        data(rowInt, 0) = Range("A" & rowAddress).Value
    Next rowInt
    
    'Salva dados da coluna NOME CORRIGIDO
    
    
    For rowInt = 1 To numRows
        rowAddress = (inicialRange(2) - numRows) + rowInt
        
        data(rowInt, 1) = Range("C" & rowAddress).Value
        
        'Fazer loop que procura o nome corrigido na planilha de ferramentas
        'Quando achar ele salva o nome do setor
        Worksheets("02_Correção Nomes").Select
        
        'Salva nome da empresa
        For Each nome In Range("C4", "C" & numRowsNames)
            If data(rowInt, 1) = nome.Value Then
                data(rowInt, 7) = Cells(nome.Row, 4).Value
                
                Debug.Print data(rowInt, 1) & " " & data(rowInt, 7)
                
                GoTo nextName
            End If
        Next nome
        
nextName:
        Worksheets("01_Base").Select
        
    Next rowInt
    
    'Salva dados da coluna Numero da peça(N)
    For rowInt = 1 To numRows
        rowAddress = (inicialRange(2) - numRows) + rowInt
        
        data(rowInt, 2) = Range("D" & rowAddress).Value
    Next rowInt
    
    'Salva dados da coluna Peso do perfil
    For rowInt = 1 To numRows
        rowAddress = (inicialRange(2) - numRows) + rowInt
        
        data(rowInt, 3) = Range("E" & rowAddress).Value
    Next rowInt
    
    'Salva dados da coluna Produção Bruta
    For rowInt = 1 To numRows
        rowAddress = (inicialRange(2) - numRows) + rowInt
        
        data(rowInt, 4) = Range("Z" & rowAddress).Value
    Next rowInt
    
    'Salva dados da coluna Talão
    For rowInt = 1 To numRows
        rowAddress = (inicialRange(2) - numRows) + rowInt
        
        data(rowInt, 5) = Range("X" & rowAddress).Value
    Next rowInt
    
    'Salva dados da coluna Ponta
    For rowInt = 1 To numRows
        rowAddress = (inicialRange(2) - numRows) + rowInt
        
        data(rowInt, 6) = Range("Y" & rowAddress).Value
    Next rowInt
    
    'Fazer for each que compara o nome corrigido com
    
    '---- Fazer aqui um loop pra colocar nome da empresa no array tb ----
    
'--------------------- MONTANDO COLUNA NOMES DE PERFIL ---------------------
           
    Workbooks(fileName).Activate
    
    'Loop que passa por todos os nomes de perfis
    For Each empresa In empresas
        For rowInt = 1 To numRows
        
             For copyInt = 1 To rowInt - 1
            
                'Verifica se o nome já foi copiado
                If (data(rowInt, 1) = data(copyInt, 1)) And (data(rowInt, 2) = data(copyInt, 2)) Then
                    
                    'Pula o nome
                    GoTo NextIteration
                End If
            Next copyInt
        
        'Verifica se é da empresa
        If Not data(rowInt, 7) = empresa Then
            GoTo NextIteration
        End If
        
        Range("A" & x) = data(rowInt, 1)
        Range("B" & x) = data(rowInt, 2)
        Range("C" & x) = empresa
        
        'Incrementa 1
        x = x + 1
           
NextIteration:
        Next rowInt
    Next empresa

    
    'Estilo das colunas
    Range("A1") = "PERFIL"
    Range("B1") = "Nº"
    Range("C1") = "EMPRESA"
    
    With Range("A1:A250")
        .Columns.AutoFit
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    With Range("B1:B250")
        .ColumnWidth = 5.29
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    With Range("C1:C250")
        .ColumnWidth = 17
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    Range("A1").Font.Size = 14
    Range("B1").Font.Size = 14
    Range("C1").Font.Size = 14
    
    '---- Fazer com que os nomes sejam colados seguindo a ordem da empresa que ele pertence ----
    
'------------------------- PROCESSA OS DADOS -------------------------
    
    'Verifica se nos perfis do mesmo dia tem algum perfil com o mesmo nome
    'Se tiver ele soma valores de Prod., Talão e Ponta
    'No fim desse loop tenho todos os necessarios para fazer a planilha
    ReDim processedData(250, 6)
    For rowInt = 1 To numRows
        
        
        For copyInt = 1 To rowInt - 1
        
            'Verifica se o nome já foi copiado
            If (data(rowInt, 1) = processedData(copyInt, 1)) And _
                (data(rowInt, 0) = processedData(copyInt, 0)) And _
                (data(rowInt, 2) = processedData(copyInt, 2)) _
            Then
            
                'Soma a produção bruta
                processedData(copyInt, 4) = processedData(copyInt, 4) + data(rowInt, 4)
                
                'Soma Talão
                processedData(copyInt, 5) = processedData(copyInt, 5) + data(rowInt, 5)
                
                'Soma Ponta
                processedData(copyInt, 6) = processedData(copyInt, 6) + data(rowInt, 6)
                
                'Debug.Print "Somou " & data(rowInt, 1) & " " & data(rowInt, 4) & " " & processedData(copyInt, 1) & " "; processedData(copyInt, 4)
                'Debug.Print "Somou " & processedData(copyInt, 0) & " " & processedData(copyInt, 1) & " " & processedData(copyInt, 4)
                
                GoTo NextIt
            End If
        Next copyInt
        
    
        'Salva Data
        processedData(numRowsArray, 0) = data(rowInt, 0)
        
        'Salva Nome Perfil
        processedData(numRowsArray, 1) = data(rowInt, 1)
        
        'Salva Numero de peça
        processedData(numRowsArray, 2) = data(rowInt, 2)
        
        'Salva Peso do perfil
        processedData(numRowsArray, 3) = data(rowInt, 3)
        
        'Salva Prod. Bruta
        processedData(numRowsArray, 4) = data(rowInt, 4)
        
        'Salva Talão
        processedData(numRowsArray, 5) = data(rowInt, 5)
        
        'Salva Ponta
        processedData(numRowsArray, 6) = data(rowInt, 6)
        
        'Salva numero da linha
        processedData(numRowsArray, 6) = data(rowInt, 6)
        
        numRowsArray = numRowsArray + 1
        
NextIt:
        
    Next rowInt
    
'------------------------- COLOCANDO DADOS DE CADA DIA -------------------------

    'Loop percorre cada data do array
    For rowInt = 0 To numRowsArray
        
        If Not IsEmpty(processedData(rowInt, 0)) Then
        
            'Loop que compara a data com as que já foram copiadas
            For copyInt = 0 To rowInt - 1
                
                'Se a data for igual a uma que ja foi copiada
                If (processedData(rowInt, 0) = processedData(copyInt, 0)) Then
                    
                    'Percorre os nomes de perfis
                    For Each perfil In Range("A2", "A" & numRows)
                    
                        'Os valores são colocados na linha com o mesmo nome e numero
                        If (perfil.Value = processedData(rowInt, 1)) And _
                        (Cells(perfil.Row, perfil.Column + 1) = processedData(rowInt, 2)) Then
                        
                            Debug.Print processedData(rowInt, 1) & perfil.Row
                            
                            Cells(perfil.Row, columnIcr + 1) = processedData(rowInt, 4)
                            
                            GoTo NextDate
                        End If
                    Next perfil
                    
                    
                End If
                
            Next copyInt
            
            If Not rowInt = 0 Then
                If Not (processedData(rowInt, 0) = processedData(rowInt - 1, 0)) Then
                    columnIcr = columnIcr + 3
                End If
            End If
            
            
            
            'Incere coluna e dados
            Cells(1, columnIcr) = "Furos"
            Cells(1, columnIcr + 1) = Format(processedData(rowInt, 0), "dd/mmm")
            Cells(1, columnIcr + 2) = "Grs/MT"
            
            '-------- STYLE --------
            
            'Estiliza coluna de furos
            With Cells(1, columnIcr)
                .ColumnWidth = 5.43
                .Font.Size = 10
                .Font.Color = RGB(0, 112, 192)
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            
            'Estiliza coluna de data
            With Cells(1, columnIcr + 1)
                .ColumnWidth = 5.43
                .Font.Size = 9
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            
            'Estiliza coluna de Grs/MT
            With Cells(1, columnIcr + 2)
                .ColumnWidth = 5.43
                .Font.Size = 10
                .Font.Color = RGB(255, 0, 0)
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            
            '-------- INCERIR DADOS --------
            
            'Percorre os nomes de perfis
            For Each perfil In Range("A2", "A" & numRows)
            
                'Os valores são colocados na linha com o mesmo nome e numero
                If (perfil.Value = processedData(rowInt, 1)) And _
                (Cells(perfil.Row, perfil.Column + 1) = processedData(rowInt, 2)) Then
                
                    'Debug.Print processedData(rowInt, 1) & perfil.Row
                    
                    Cells(perfil.Row, columnIcr + 1) = processedData(rowInt, 4)
                    
                    GoTo NextDate
                End If
            Next perfil
            
            
NextDate:
            
        End If
        
    Next rowInt
    
End Sub
