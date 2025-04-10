Attribute VB_Name = "FazRelatorioFerr"
Option Explicit

Sub RelatorioFerramentas()
    'Ao usar a macro:
        'deixe a planilha que voc� quer copiar selecionada
        'a planilha precisa estar com esse formato de nome:
        'Mes_Numero do mes_Utimos dois digitos do ano
        'Deve ficar assim: Mar_3_25
    
        'Antes precisa corrigir os nomes
        
        'Capturar erros:
            'Erro de rodar sem tem o historico aberto
            'Erro de rodar sem ter colocado nome certo na planilha
    
    Dim data() As Variant, processedData As Variant, perfil As Variant
    Dim fileName As String, arrDate() As String, inicialRange() As String, rowAddress As String
    Dim numRows As Integer, colInt As Integer, rowInt As Integer, copyInt As Integer, x As Integer, numRowsArray As Integer, columnIcr As Integer
    
    columnIcr = 3
    numRowsArray = 0
    x = 2
    fileName = ThisWorkbook.Name
    arrDate = Split(ActiveSheet.Name, "_")
    
    Workbooks("HIST�RICO PRODU��O 2022-2024_V5.xlsm").Activate
    Worksheets("01_Base").Select
    
    'Tira filtros aplicados
    ActiveWorkbook.Worksheets("01_Base").AutoFilter.Sort.SortFields.Clear
 
    'Filtra os dados da base pela data, de acordo com o nome da planilha(Ex:Mar_1_25)
    ActiveSheet.Range("$A$3:$BA$4805").AutoFilter Field:=1, Operator:= _
    xlFilterValues, Criteria2:=Array(1, arrDate(1) & "/10/20" & arrDate(2))
    
'------------------------- SALVANDO DADOS NO ARRAY -------------------------
    
    'Conta linha visiveis
    numRows = Range("A3", "A" & Cells(Rows.Count, 1).End(xlUp).Row).Rows.SpecialCells(xlCellTypeVisible).Count - 1
    
    'Endere�o da ultima celula da coluna A
    inicialRange = Split(Range("A" & Cells(Rows.Count, 1).End(xlUp).Row).Address, "$")
    
    ReDim data(numRows, 6) As Variant
    
    'Salva dados da coluna DATA
    For rowInt = 1 To numRows
        rowAddress = (inicialRange(2) - numRows) + rowInt
        
        data(rowInt, 0) = Range("A" & rowAddress).Value
    Next rowInt
    
    'Salva dados da coluna NOME CORRIGIDO
    For rowInt = 1 To numRows
        rowAddress = (inicialRange(2) - numRows) + rowInt
        
        data(rowInt, 1) = Range("C" & rowAddress).Value
    Next rowInt
    
    'Salva dados da coluna Numero da pe�a(N)
    For rowInt = 1 To numRows
        rowAddress = (inicialRange(2) - numRows) + rowInt
        
        data(rowInt, 2) = Range("D" & rowAddress).Value
    Next rowInt
    
    'Salva dados da coluna Peso do perfil
    For rowInt = 1 To numRows
        rowAddress = (inicialRange(2) - numRows) + rowInt
        
        data(rowInt, 3) = Range("E" & rowAddress).Value
    Next rowInt
    
    'Salva dados da coluna Produ��o Bruta
    For rowInt = 1 To numRows
        rowAddress = (inicialRange(2) - numRows) + rowInt
        
        data(rowInt, 4) = Range("Z" & rowAddress).Value
    Next rowInt
    
    'Salva dados da coluna Tal�o
    For rowInt = 1 To numRows
        rowAddress = (inicialRange(2) - numRows) + rowInt
        
        data(rowInt, 5) = Range("X" & rowAddress).Value
    Next rowInt
    
    'Salva dados da coluna Ponta
    For rowInt = 1 To numRows
        rowAddress = (inicialRange(2) - numRows) + rowInt
        
        data(rowInt, 6) = Range("Y" & rowAddress).Value
    Next rowInt
    
    '---- Fazer aqui um loop pra colocar nome da empresa no array tb ----
    
'--------------------- MONTANDO COLUNA NOMES DE PERFIL ---------------------
           
    Workbooks(fileName).Activate
    
    'Loop que passa por todos os nomes de perfis
    For rowInt = 1 To numRows
    
         For copyInt = 1 To rowInt - 1
        
            'Verifica se o nome j� foi copiado
            If (data(rowInt, 1) = data(copyInt, 1)) And (data(rowInt, 2) = data(copyInt, 2)) Then

                GoTo NextIteration
            End If
        Next copyInt
    
        
    Range("A" & x) = data(rowInt, 1)
    Range("B" & x) = data(rowInt, 2)
    
    'Incrementa 1
    x = x + 1
       
NextIteration:
    Next rowInt
    
    'Estilo das colunas
    Range("A1") = "PERFIL"
    Range("B1") = "N�"
    
    With Range("A1:B250")
        .Columns.AutoFit
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    Range("B1:B250").ColumnWidth = 5.29
    Range("A1").Font.Size = 16
    Range("B1").Font.Size = 10
    
    '---- Fazer com que os nomes sejam colados seguindo a ordem da empresa que ele pertence ----
    
'------------------------- PROCESSA OS DADOS -------------------------
    
    'Verifica se nos perfis do mesmo dia tem algum perfil com o mesmo nome
    'Se tiver ele soma valores de Prod., Tal�o e Ponta
    'No fim desse loop tenho todos os necessarios para fazer a planilha
    ReDim processedData(250, 6)
    For rowInt = 1 To numRows
        
        
        For copyInt = 1 To rowInt - 1
        
            'Verifica se o nome j� foi copiado
            If (data(rowInt, 1) = data(copyInt, 1)) And _
                (data(rowInt, 0) = data(copyInt, 0)) And _
                (data(rowInt, 2) = data(copyInt, 2)) _
            Then
                'Soma a produ��o bruta
                data(copyInt, 4) = data(copyInt, 4) + data(rowInt, 4)
                
                'Soma Tal�o
                data(copyInt, 5) = data(copyInt, 5) + data(rowInt, 5)
                
                'Soma Ponta
                data(copyInt, 6) = data(copyInt, 6) + data(rowInt, 6)
                
                'Debug.Print "Somou " & data(rowInt, 0) & " " & data(rowInt, 1) & " " & data(copyInt, 4)
                
                GoTo NextIt
            End If
        Next copyInt
        
        'Debug.Print "Copiou " & data(rowInt, 0) & " " & data(rowInt, 1) & " " & data(copyInt, 4)
    
        'Salva Data
        processedData(numRowsArray, 0) = data(rowInt, 0)
        
        'Salva Nome Perfil
        processedData(numRowsArray, 1) = data(rowInt, 1)
        
        'Salva Numero de pe�a
        processedData(numRowsArray, 2) = data(rowInt, 2)
        
        'Salva Peso do perfil
        processedData(numRowsArray, 3) = data(rowInt, 3)
        
        'Salva Prod. Bruta
        processedData(numRowsArray, 4) = data(rowInt, 4)
        
        'Salva Tal�o
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
        
            'Loop que compara a data com as que j� foram copiadas
            For copyInt = 0 To rowInt - 1
                
                'Se a data for igual a uma que ja foi copiada
                If (processedData(rowInt, 0) = processedData(copyInt, 0)) Then
                    
                    'Percorre os nomes de perfis
                    For Each perfil In Range("A2", "A" & numRows)
                    
                        'Os valores s�o colocados na linha com o mesmo nome e numero
                        If (perfil.Value = processedData(rowInt, 1)) And _
                        (Cells(perfil.Row, perfil.Column + 1) = processedData(rowInt, 2)) Then
                        
                            Debug.Print processedData(rowInt, 1) & perfil.Row
                            
                            Cells(perfil.Row, columnIcr + 1) = Cells(perfil.Row, columnIcr + 1) + processedData(rowInt, 4)
                            
                            GoTo NextDate
                        End If
                    Next perfil
                    
                    
                End If
                
                columnIcr = columnIcr + 3
            Next copyInt
            
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
            
                'Os valores s�o colocados na linha com o mesmo nome e numero
                If (perfil.Value = processedData(rowInt, 1)) And _
                (Cells(perfil.Row, perfil.Column + 1) = processedData(rowInt, 2)) Then
                
                    Debug.Print processedData(rowInt, 1) & perfil.Row
                    
                    Cells(perfil.Row, columnIcr + 1) = Cells(perfil.Row, columnIcr + 1) + processedData(rowInt, 4)
                    
                    GoTo NextDate
                End If
            Next perfil
            
         
NextDate:
            
        End If
        
    Next rowInt
    
End Sub
