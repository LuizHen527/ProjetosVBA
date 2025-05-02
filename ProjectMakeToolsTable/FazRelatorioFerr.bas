Attribute VB_Name = "FazRelatorioFerr"
Option Explicit

Sub RelatorioFerramentas()
    'Ao usar a macro:
        'C planilha que voc� quer copiar selecionada
         
        'a planilha precisa estar com esse formato de nome:
        'ano_mes = 24_03
        
        'ano: Deve ser os ultimos dois digitos. Exemplo para 2022: 22
        'mes: Deve ser o numero do mes. Para abril: 04.
        'Coloque o underline separando os dois
    
        'Se os nomes no Hist�rico de produ��o n�o estiver corrigido, ele
        'vai ignorar aquela linha.
    
    Dim data() As Variant, processedData As Variant, perfil As Variant, somTalaoPonta() As Variant, rng As Range
    Dim fileName As String, arrDate() As String, inicialRange() As String, rowAddress As String, databaseName As String, strArray() As String, lastColTotais() As String, nome As Variant, empresa As Variant
    Dim numRows As Integer, colInt As Integer, rowInt As Integer, copyInt As Integer, x As Integer, numRowsArray As Integer, columnIcr As Integer, numRowsNames As Integer, _
    numPerfis As Integer, lastRowPerfis As Integer, iterador As Integer, moldRowSum As Integer, alumRowSum As Integer, pollRowSum As Integer, extRowSum As Integer, alhRowSum As Integer
    
    Application.ScreenUpdating = False
    
    '---- Inicializando variaveis ----
    
    columnIcr = 4
    numRowsArray = 0
    x = 2
    
    databaseName = "HIST�RICO PRODU��O 2022-2024_V5.xlsm"
    
    fileName = ActiveWorkbook.Name
    arrDate = Split(ActiveSheet.Name, "_")
    
    'Nomes corrigidos no hist�rico
    numRowsNames = Range("C4", "C" & Cells(Rows.Count, 1).End(xlUp).row).Rows.Count
    
    ReDim strArray(5)
    strArray() = Split("MOLDUCOLOR,ALUMITEC,POLLUX,ALHENA,EXTERNO", ",")
    
    On Error GoTo msgAbrirHistorico
    Workbooks(databaseName).Activate
    On Error GoTo 0
    
    Worksheets("02_Corre��o Nomes").Select
    numRowsNames = Range("C4", "C" & Cells(Rows.Count, 1).End(xlUp).row).Rows.Count
    
    Worksheets("01_Base").Select
    
    'Tira filtros aplicados
    ActiveWorkbook.Worksheets("01_Base").AutoFilter.Sort.SortFields.Clear
 
    'Filtra os dados da base pela data, de acordo com o nome da planilha(Ex:Mar_1_25)
    On Error GoTo msgPlanilhaNomeErrado
    
    ActiveSheet.Range("$A$3:$BA$4805").AutoFilter Field:=1, Operator:= _
    xlFilterValues, Criteria2:=Array(1, arrDate(1) & "/10/20" & arrDate(0))
    
    ActiveWorkbook.Worksheets("01_Base").AutoFilter.Sort.SortFields.Add Key:= _
        Range("A3:A4805"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("01_Base").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    On Error GoTo 0

'------------------------- SALVANDO DADOS NO ARRAY -------------------------
    
    'Conta linha visiveis
    numRows = Range("A3", "A" & Cells(Rows.Count, 1).End(xlUp).row).Rows.SpecialCells(xlCellTypeVisible).Count - 1
    
    'Endere�o da ultima celula da coluna A
    inicialRange = Split(Range("A" & Cells(Rows.Count, 1).End(xlUp).row).Address, "$")
    
    ReDim data(numRows, 8) As Variant
    
    'Salva dados da coluna DATA
    For rowInt = 1 To numRows
        rowAddress = (inicialRange(2) - numRows) + rowInt
        
        data(rowInt, 0) = Range("A" & rowAddress).Value
    Next rowInt
    
    'Salva dados da coluna NOME CORRIGIDO
    
    On Error GoTo corrigirNome
    For rowInt = 1 To numRows
        rowAddress = (inicialRange(2) - numRows) + rowInt

        data(rowInt, 1) = Range("C" & rowAddress).Value
        
        
        'Fazer loop que procura o nome corrigido na planilha de ferramentas
        'Quando achar ele salva o nome do setor
        Worksheets("02_Corre��o Nomes").Select
        
        'Salva nome da empresa
        For Each nome In Range("C4", "C" & numRowsNames)
            If data(rowInt, 1) = nome.Value Then
                data(rowInt, 7) = Cells(nome.row, 4).Value
                
                'Debug.Print data(rowInt, 1) & " " & data(rowInt, 7)
                
                GoTo nextName
            End If
        Next nome
        On Error GoTo 0
        
nextName:
        Worksheets("01_Base").Select
        
    Next rowInt
    
    On Error GoTo 0
    
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
        
        If IsNumeric(Range("Z" & rowAddress).Value) Then
            data(rowInt, 4) = Range("Z" & rowAddress).Value
        Else
            data(rowInt, 4) = 1
        End If
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
    
    'Salva dados do numero de furos
    For rowInt = 1 To numRows
        rowAddress = (inicialRange(2) - numRows) + rowInt
        
        data(rowInt, 8) = Range("F" & rowAddress).Value
    Next rowInt
    
    'Fazer for each que compara o nome corrigido com
    
    '---- Fazer aqui um loop pra colocar nome da empresa no array tb ----
    
'--------------------- MONTANDO COLUNA NOMES DE PERFIL ---------------------
           
    Workbooks(fileName).Activate
    
    'Loop que passa por todos os nomes de perfis
    For Each empresa In strArray
        For rowInt = 1 To numRows
        
             For copyInt = 1 To rowInt - 1
            
                'Verifica se o nome j� foi copiado. Confere nome e numero do perfil
                If (data(rowInt, 1) = data(copyInt, 1)) And (data(rowInt, 2) = data(copyInt, 2)) Then
                    
                    'Pula o nome
                    GoTo NextIteration
                End If
            Next copyInt
        
        'Verifica se � da empresa
        If Not data(rowInt, 7) = empresa Then
            GoTo NextIteration
        End If
        
        'PERFIL
        Range("A" & x + 1) = data(rowInt, 1)
        
        'Numero (N)
        Range("B" & x + 1) = data(rowInt, 2)
        
        'Empresa
        Range("C" & x + 1) = empresa
        
        'Incrementa 1
        x = x + 1
           
NextIteration:
        Next rowInt
    Next empresa
    
    'Salva o numero de nomes de perfil
    numPerfis = Range("A1", "A" & Cells(Rows.Count, 1).End(xlUp).row).Count
    
    'Estilo das colunas
    Range("A2") = "PERFIL"
    Range("B2") = "N�"
    Range("C2") = "EMPRESA"
    
    With Range("A2:A" & numPerfis)
        .ColumnWidth = 42
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    With Range("B2:B" & numPerfis)
        .ColumnWidth = 5.29
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    With Range("C2:C" & numPerfis)
        .ColumnWidth = 17
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    Range("A2").Font.Size = 14
    Range("B2").Font.Size = 14
    Range("C2").Font.Size = 14
    
    '---- Fazer com que os nomes sejam colados seguindo a ordem da empresa que ele pertence ----
    
'------------------------- PROCESSA OS DADOS -------------------------
    
    'Verifica se nos perfis do mesmo dia tem algum perfil com o mesmo nome
    'Se tiver ele soma valores de Prod., Tal�o e Ponta
    'No fim desse loop tenho todos os necessarios para fazer a planilha
    ReDim processedData(numRows, 7)
    ReDim somTalaoPonta(numRows, 2)
    
    For rowInt = 1 To numRows
        
        For copyInt = 1 To rowInt - 1
        
            'Verifica se o nome j� foi copiado
            If (data(rowInt, 1) = processedData(copyInt, 1)) And _
                (data(rowInt, 0) = processedData(copyInt, 0)) And _
                (data(rowInt, 2) = processedData(copyInt, 2)) _
            Then
            
                'Soma a produ��o bruta
                If IsNumeric(data(rowInt, 4)) Then
                processedData(copyInt, 4) = processedData(copyInt, 4) + data(rowInt, 4)
                End If
                
                'Soma Tal�o
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
        
        'Salva Numero de furos
        processedData(numRowsArray, 7) = data(rowInt, 8)
        
        'Salva numero da linha
        'processedData(numRowsArray, 6) = data(rowInt, 6)
        
        'Debug.Print numRowsArray & "" & processedData(numRowsArray, 0) & " Tal�o: " & processedData(numRowsArray, 5) & " Ponta: " & processedData(numRowsArray, 6)
        
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
                    For Each perfil In Range("A3", "A" & numRows)
                    
                        'Os valores s�o colocados na linha com o mesmo nome e numero
                        If (perfil.Value = processedData(rowInt, 1)) And _
                        (Cells(perfil.row, perfil.Column + 1) = processedData(rowInt, 2)) Then
                        
                            'Debug.Print processedData(rowInt, 1) & perfil.Row
                            
                            'Produ��o bruta
                            If IsEmpty(processedData(rowInt, 4)) Or processedData(rowInt, 4) = 0 Then
                                Cells(perfil.row, columnIcr + 1) = 0
                                
                                With Cells(perfil.row, columnIcr + 1)
                                .Font.Size = 12
                                .Font.Bold = True
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter
                                .NumberFormat = 0
                            End With
                            Else
                                Cells(perfil.row, columnIcr + 1) = processedData(rowInt, 4)
                                
                                With Cells(perfil.row, columnIcr + 1)
                                    .Font.Size = 12
                                    .Font.Bold = True
                                    .HorizontalAlignment = xlCenter
                                    .VerticalAlignment = xlCenter
                                    .NumberFormat = "#,###"
                                End With
                            End If

                            'Grs/MT
                            If IsEmpty(processedData(rowInt, 3)) Then
                                Cells(perfil.row, columnIcr + 2) = "Vazio"
                            Else
                                Cells(perfil.row, columnIcr + 2) = processedData(rowInt, 3)
                            End If
                            
                            With Cells(perfil.row, columnIcr + 2)
                                .Font.Size = 12
                                .Font.Bold = True
                                .Font.Color = RGB(255, 0, 0)
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter
                                .NumberFormat = "0.000"
                            End With
                            
                            'Furos
                            If IsEmpty(processedData(rowInt, 7)) Then
                                Cells(perfil.row, columnIcr) = "Vazio"
                            Else
                                Cells(perfil.row, columnIcr) = processedData(rowInt, 7)
                            End If
                            
                            With Cells(perfil.row, columnIcr)
                                .Font.Size = 12
                                .Font.Bold = True
                                .Font.Color = RGB(255, 0, 0)
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter
                            End With
                            
                            GoTo NextDate
                        End If
                    Next perfil
                    
                    
                End If
                
            Next copyInt
            
            'Se for o primeiro item ele n�o entra no if
            If Not rowInt = 0 Then
                'Se a data for diferente da anterior ele incrementa
                If Not (processedData(rowInt, 0) = processedData(rowInt - 1, 0)) Then
                    columnIcr = columnIcr + 3
                End If
            End If
            
            
            
            'Incere coluna e dados
            Cells(2, columnIcr) = "Furos"
            Cells(2, columnIcr + 1) = Format(processedData(rowInt, 0), "dd/mmm")
            Cells(2, columnIcr + 2) = "Grs/MT"
            
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
            
            '-------- LINHAS DAS BORDAS DA CELULA --------
            With Range(Col_Letter(columnIcr + 0) & 2, Col_Letter(columnIcr + 2) & x - 1)
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlThick
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                .Borders(xlInsideHorizontal).Weight = xlThin
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideVertical).Weight = xlThin
            End With
            
            '-------- INCERIR DADOS --------
            
            'Percorre os nomes de perfis
            For Each perfil In Range("A3", "A" & numRows)
            
                'Os valores s�o colocados na linha com o mesmo nome e numero
                If (perfil.Value = processedData(rowInt, 1)) And _
                (Cells(perfil.row, perfil.Column + 1) = processedData(rowInt, 2)) Then
                
                    'Debug.Print processedData(rowInt, 1) & perfil.Row
                    
                    'Produ��o bruta
                    Cells(perfil.row, columnIcr + 1) = processedData(rowInt, 4)
                    
                    With Cells(perfil.row, columnIcr + 1)
                        .Font.Size = 12
                        .Font.Bold = True
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .NumberFormat = "#,###"
                    End With
                    
                    'Grs/MT
                    Cells(perfil.row, columnIcr + 2) = processedData(rowInt, 3)
                    
                    If IsEmpty(processedData(rowInt, 3)) Then
                        Cells(perfil.row, columnIcr + 2) = "Vazio"
                    End If
                            
                    With Cells(perfil.row, columnIcr + 2)
                        .Font.Size = 12
                        .Font.Bold = True
                        .Font.Color = RGB(255, 0, 0)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .NumberFormat = "0.000"
                    End With
                    
                    'Furos
                    Cells(perfil.row, columnIcr) = processedData(rowInt, 7)
                            
                    With Cells(perfil.row, columnIcr)
                        .Font.Size = 12
                        .Font.Bold = True
                        .Font.Color = RGB(255, 0, 0)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                    End With
                    
                    GoTo NextDate
                End If
            Next perfil
            
            
NextDate:
            
        End If
        
    Next rowInt
    
'------------------------- COLUNA DA SOMA DA PRODU��O DE CADA DIA -------------------------
    
    Cells(2, Columns.Count).End(xlToLeft).Offset(0, 1) = "TOTAIS"
    
    Cells(2, Columns.Count).End(xlToLeft).Offset(1, 0).Formula = "=E3+H3+K3+N3+Q3+T3+W3+Z3+AC3+AF3+AI3+AL3+AO3+AR3+AU3+AX3+BA3+BD3+BG3+BJ3+BM3+BP3+BS3+BV3+BY3+CB3+CE3+CH3+CK3+CN3+CQ3+CT3"
    
    Cells(2, Columns.Count).End(xlToLeft).Offset(1, 0).Select
    
    Selection.AutoFill Destination:=Range(Cells(2, Columns.Count).End(xlToLeft).Offset(1, 0).Address, Col_Letter(Cells(2, Columns.Count).End(xlToLeft).Offset(1, 0).Column) & numPerfis), Type:=xlFillDefault
    
    
    '-------- STYLE --------
    With Range(Cells(2, Columns.Count).End(xlToLeft), Col_Letter(Cells(2, Columns.Count).End(xlToLeft).Column) & Cells(Rows.Count, 1).End(xlUp).row)
        .Font.Size = 12
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ColumnWidth = 11.43
        .NumberFormat = "#,###"
    End With
    
'------------------------- FAZER LINHAS DE TOTAIS -------------------------
    lastRowPerfis = Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).row
    ReDim strArray(5)
    strArray() = Split("TOTAL BRUTO [Kg],TOTAL TAL�O [Kg],TOTAL PONTA [Kg],TOTAL L�QUIDO [Kg],PERDA TAL�O [%],PERDA PONTA [%]", ",")
    
    For x = 0 To 5
        
        With Range("A" & lastRowPerfis + x, "C" & lastRowPerfis + x)
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            If strArray(x) = "TOTAL BRUTO [Kg]" Or strArray(x) = "TOTAL L�QUIDO [Kg]" Then
                .Interior.Color = RGB(255, 255, 102)
            End If
        End With
        
        Cells(Rows.Count, 1).End(xlUp).Offset(1, 0) = strArray(x)
    
    Next x
    
    
'------------------------- FAZER FORMULAS DE TOTAIS -------------------------

    ReDim lastColTotais(2)
    lastColTotais() = Split(Cells(lastRowPerfis - 1, Cells(lastRowPerfis - 1, Columns.Count).End(xlToLeft).Column).Offset(1, -1).Address, "$")
    
    '------ TOTAL DO DIA ------
    
    'Aplica estilo e formula na primeira celula
    With Range(Cells(Rows.Count, 1).End(xlUp).Offset(-5, 1), _
    Cells(Rows.Count, 1).End(xlUp).Offset(-5, 3))
    
        .Merge
        .Interior.Color = RGB(255, 255, 102)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "#,###"
        .Font.Size = 12
        .Formula = "=SUM(E3:E" & lastRowPerfis - 1 & ")"
        
    End With
    
    Cells(Rows.Count, 1).End(xlUp).Offset(-5, 1).Select
    
    Selection.AutoFill Destination:=Range(Cells(Rows.Count, 1).End(xlUp).Offset(-5, 1), Range(lastColTotais(1) & lastColTotais(2))), Type:=xlFillDefault
    
    '------ TOTAL TAL�O E PONTA ------
    
    iterador = 0
    
    'mescla celulas na linha do tal�o
    With Range(Cells(Rows.Count, 1).End(xlUp).Offset(-4, 1), _
    Cells(Rows.Count, 1).End(xlUp).Offset(-4, 3))
    
        .Merge
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "#,###"
        .Font.Size = 12
        
    End With
    
    Cells(Rows.Count, 1).End(xlUp).Offset(-4, 1).Select
    
    Selection.AutoFill Destination:=Range(Cells(Rows.Count, 1).End(xlUp).Offset(-4, 1), Range(lastColTotais(1) & lastColTotais(2) + 1)), Type:=xlFillDefault
    
    'mescla celulas na linha da ponta
    
    With Range(Cells(Rows.Count, 1).End(xlUp).Offset(-3, 1), _
    Cells(Rows.Count, 1).End(xlUp).Offset(-3, 3))
    
        .Merge
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "#,###"
        .Font.Size = 12
        
    End With
    
    Cells(Rows.Count, 1).End(xlUp).Offset(-3, 1).Select
    
    Selection.AutoFill Destination:=Range(Cells(Rows.Count, 1).End(xlUp).Offset(-3, 1), Range(lastColTotais(1) & lastColTotais(2) + 2)), Type:=xlFillDefault
    
    
    'Percorre o processedData() somando os tal�es e pontas com mesma data e colocando em um array
    For x = 0 To numRowsArray
        
        'percorre dados que j� foram salvos
        For copyInt = 0 To x - 1
            If processedData(x, 0) = somTalaoPonta(copyInt, 0) Then
                'soma tal�o
                somTalaoPonta(copyInt, 1) = somTalaoPonta(copyInt, 1) + processedData(x, 5)
                
                'soma ponta
                somTalaoPonta(copyInt, 2) = somTalaoPonta(copyInt, 2) + processedData(x, 6)
                
                'Debug.Print somTalaoPonta(copyInt, 0) & vbTab & somTalaoPonta(copyInt, 1) & vbTab & somTalaoPonta(copyInt, 2) & vbTab & "SOMOU"
                
                GoTo NextX
            End If
        Next copyInt
        
        'data
        somTalaoPonta(iterador, 0) = processedData(x, 0)
        'tal�o
        somTalaoPonta(iterador, 1) = processedData(x, 5)
        'ponta
        somTalaoPonta(iterador, 2) = processedData(x, 6)
        
        'Debug.Print somTalaoPonta(iterador, 0) & vbTab & somTalaoPonta(iterador, 1) & vbTab & somTalaoPonta(iterador, 2)
        
        iterador = iterador + 1
NextX:
    Next x
    
    colInt = 1
    
    'Loop que percorre as datas e cola
    For x = 0 To iterador
        'ponta
        Cells(Rows.Count, 1).End(xlUp).Offset(-3, colInt) = somTalaoPonta(x, 2)
        'talao
        Cells(Rows.Count, 1).End(xlUp).Offset(-4, colInt) = somTalaoPonta(x, 1)
        
        colInt = colInt + 3
    Next x
    
    '------ TOTAL L�QUIDO ------
    
    'lastColTotais() = Split(Cells(lastRowPerfis - 1, Cells(lastRowPerfis - 1, Columns.Count).End(xlToLeft).Column).Offset(1, 0).Address, "$")
    
    With Range(Cells(Rows.Count, 1).End(xlUp).Offset(-2, 1), _
    Cells(Rows.Count, 1).End(xlUp).Offset(-2, 3))
    
        .Merge
        .Interior.Color = RGB(255, 255, 102)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "#,###"
        .Font.Size = 12
        .Formula = "=D" & Cells(Rows.Count, 1).End(xlUp).Offset(-5, 1).row _
                    & "-D" & Cells(Rows.Count, 1).End(xlUp).Offset(-4, 1).row _
                    & "-D" & Cells(Rows.Count, 1).End(xlUp).Offset(-3, 1).row
        
    End With
    
    Cells(Rows.Count, 1).End(xlUp).Offset(-2, 1).Select
    
    Selection.AutoFill Destination:=Range(Cells(Rows.Count, 1).End(xlUp).Offset(-2, 1), Range(lastColTotais(1) & lastColTotais(2) + 3)), Type:=xlFillDefault
    
    
    '------ PERDA TAL�O(%) ------
    
    With Range(Cells(Rows.Count, 1).End(xlUp).Offset(-1, 1), _
    Cells(Rows.Count, 1).End(xlUp).Offset(-1, 3))
    
        .Merge
        '.Interior.Color = RGB(255, 255, 102)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "0%"
        .Font.Size = 12
        .Formula = "=(D" & Cells(Rows.Count, 1).End(xlUp).Offset(-4, 1).row _
                    & ")/D" & Cells(Rows.Count, 1).End(xlUp).Offset(-5, 1).row
        
    End With
    
    Cells(Rows.Count, 1).End(xlUp).Offset(-1, 1).Select
    
    Selection.AutoFill Destination:=Range(Cells(Rows.Count, 1).End(xlUp).Offset(-1, 1), Range(lastColTotais(1) & lastColTotais(2) + 4)), Type:=xlFillDefault
    
    '------ PERDA PONTA(%) ------
    
    With Range(Cells(Rows.Count, 1).End(xlUp).Offset(0, 1), _
    Cells(Rows.Count, 1).End(xlUp).Offset(0, 3))
    
        .Merge
        '.Interior.Color = RGB(255, 255, 102)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "0%"
        .Font.Size = 12
        .Formula = "=(D" & Cells(Rows.Count, 1).End(xlUp).Offset(-3, 1).row _
                    & ")/D" & Cells(Rows.Count, 1).End(xlUp).Offset(-5, 1).row
        
    End With
    
    Cells(Rows.Count, 1).End(xlUp).Offset(0, 1).Select
    
    Selection.AutoFill Destination:=Range(Cells(Rows.Count, 1).End(xlUp).Offset(0, 1), Range(lastColTotais(1) & lastColTotais(2) + 5)), Type:=xlFillDefault
    
'------------------------- FAZER FORMULAS DE TOTAIS -------------------------
    
    '------ SOMA DE TOTAIS DO DIA ------
    With Cells(lastRowPerfis - 1, Cells(lastRowPerfis - 1, Columns.Count).End(xlToLeft).Column).Offset(1, 0)
        .Interior.Color = RGB(255, 255, 102)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "#,###"
        .Font.Size = 12
        .Formula = "=SUM(D" & lastRowPerfis & ":" & Cells(lastRowPerfis - 1, Cells(lastRowPerfis, Columns.Count).End(xlToLeft).Column).Offset(1, 2).Address & ")"
    End With
    
    '------ SOMA DE TOTAIS DE TAL�O ------
    With Cells(lastRowPerfis + 1, Cells(lastRowPerfis + 1, Columns.Count).End(xlToLeft).Column)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "#,###"
        .Font.Size = 12
        .Formula = "=SUM(D" & lastRowPerfis + 1 & ":" & Cells(lastRowPerfis + 1, Cells(lastRowPerfis + 1, Columns.Count).End(xlToLeft).Column).Offset(0, -1).Address & ")"
    End With
    
    '------ SOMA DE TOTAIS DE PONTA ------
    With Cells(lastRowPerfis + 2, Cells(lastRowPerfis + 2, Columns.Count).End(xlToLeft).Column)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "#,###"
        .Font.Size = 12
        .Formula = "=SUM(D" & lastRowPerfis + 2 & ":" & Cells(lastRowPerfis + 2, Cells(lastRowPerfis + 2, Columns.Count).End(xlToLeft).Column).Offset(0, -1).Address & ")"
    End With
    
    '------ SOMA DO TOTAL L�QUIDO ------
    With Cells(lastRowPerfis + 3, Cells(lastRowPerfis + 3, Columns.Count).End(xlToLeft).Column).Offset(0, 1)
        .Interior.Color = RGB(255, 255, 102)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "#,###"
        .Font.Size = 12
        .Formula = "=" & Cells(lastRowPerfis, Cells(lastRowPerfis, Columns.Count).End(xlToLeft).Column).Address _
                        & "-" & Cells(lastRowPerfis + 1, Cells(lastRowPerfis + 1, Columns.Count).End(xlToLeft).Column).Address _
                        & "-" & Cells(lastRowPerfis + 2, Cells(lastRowPerfis + 2, Columns.Count).End(xlToLeft).Column).Address
    End With
    
    '------ % PERDA TAL�O ------
    With Cells(lastRowPerfis + 4, Cells(lastRowPerfis + 4, Columns.Count).End(xlToLeft).Column).Offset(0, 1)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "0%"
        .Font.Size = 12
        .Formula = "=" & Cells(lastRowPerfis + 1, Cells(lastRowPerfis + 1, Columns.Count).End(xlToLeft).Column).Address _
                    & "/" & Cells(lastRowPerfis, Cells(lastRowPerfis, Columns.Count).End(xlToLeft).Column).Address
    End With
    
    '------ % PERDA PONTA ------
    With Cells(lastRowPerfis + 5, Cells(lastRowPerfis + 5, Columns.Count).End(xlToLeft).Column).Offset(0, 1)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "0%"
        .Font.Size = 12
        .Formula = "=" & Cells(lastRowPerfis + 2, Cells(lastRowPerfis + 2, Columns.Count).End(xlToLeft).Column).Address _
                    & "/" & Cells(lastRowPerfis, Cells(lastRowPerfis, Columns.Count).End(xlToLeft).Column).Address
    End With
    
'------------------------- FAZENDO BORDAS E CORES -------------------------
    
    
    '------ Colocando cores nas linhas de acordo com empresa ------
    
    For Each rng In Range("C2", Cells(lastRowPerfis - 1, 3))
        Select Case rng.Value
        
            Case "MOLDUCOLOR"
                Range(Cells(rng.row, 1), Cells(rng.row, Cells(rng.row, Columns.Count).End(xlToLeft).Column)) _
                .Interior.Color = RGB(199, 211, 227)
                moldRowSum = moldRowSum + 1
                
            Case "ALUMITEC"
                Range(Cells(rng.row, 1), Cells(rng.row, Cells(rng.row, Columns.Count).End(xlToLeft).Column)) _
                .Interior.Color = RGB(236, 197, 243)
                alumRowSum = alumRowSum + 1
                
            Case "POLLUX"
                Range(Cells(rng.row, 1), Cells(rng.row, Cells(rng.row, Columns.Count).End(xlToLeft).Column)) _
                .Interior.Color = RGB(205, 222, 172)
                pollRowSum = pollRowSum + 1
                
            Case "EXTERNO"
                Range(Cells(rng.row, 1), Cells(rng.row, Cells(rng.row, Columns.Count).End(xlToLeft).Column)) _
                .Interior.Color = RGB(212, 211, 198)
                extRowSum = extRowSum + 1
                
            Case "ALHENA"
                Range(Cells(rng.row, 1), Cells(rng.row, Cells(rng.row, Columns.Count).End(xlToLeft).Column)) _
                .Interior.Color = RGB(252, 213, 180)
                alhRowSum = alhRowSum + 1
                
            Case Else
                
        End Select
    Next rng
    
    '------ Colocando bordas nas colunas perfil, N e Empresa ------
    
    With Range("A1", "C" & lastRowPerfis - 1)
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThick
        .Borders(xlEdgeLeft).Weight = xlThick
    End With
    
    '------ Colocando bordas na coluna de totais ------
    
    With Range(Cells(1, Columns.Count).End(xlToLeft), Cells(lastRowPerfis - 1, Columns.Count).End(xlToLeft))
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThick
    End With
    
    '------ Colocando bordas no rodap�. Linhas de total. ------
    
    With Range("A" & lastRowPerfis, Cells(lastRowPerfis + 5, Columns.Count).End(xlToLeft))
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlInsideVertical).Weight = xlThick
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlThick
        .Borders(xlEdgeLeft).Weight = xlThick
    End With
    
    '------ Colocando bordas no cabe�alho ------
    
    With Range("a1", Cells(1, Columns.Count).End(xlToLeft))
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    '------ Filtrando perfis de A a Z pra cada empresa ------
    
    'Molducolor
    If Not moldRowSum = 0 Then
        With ActiveWorkbook.ActiveSheet.Sort
            .SortFields.Clear
            .SortFields.Add Key:=Range("A2"), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Range("A2", Cells(moldRowSum + 1, Columns.Count).End(xlToLeft))
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
    
    'Alumitec
    If Not alumRowSum = 0 Then
        With ActiveWorkbook.ActiveSheet.Sort
            .SortFields.Clear
            .SortFields.Add Key:=Range("A" & moldRowSum + 2), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Range("A" & moldRowSum + 2, Cells(moldRowSum + alumRowSum + 1, Columns.Count).End(xlToLeft))
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
    
    'Pollux
    If Not pollRowSum = 0 Then
        With ActiveWorkbook.ActiveSheet.Sort
            .SortFields.Clear
            .SortFields.Add Key:=Range("A" & moldRowSum + alumRowSum + 2), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Range("A" & moldRowSum + alumRowSum + 2, Cells(moldRowSum + alumRowSum + pollRowSum + 1, Columns.Count).End(xlToLeft))
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
    
    'Alhena
    If Not alhRowSum = 0 Then
        With ActiveWorkbook.ActiveSheet.Sort
            .SortFields.Clear
            .SortFields.Add Key:=Range("A" & moldRowSum + alumRowSum + pollRowSum + 2), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Range("A" & moldRowSum + alumRowSum + pollRowSum + 2, Cells(moldRowSum + alumRowSum + pollRowSum + alhRowSum + 1, Columns.Count).End(xlToLeft))
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
    
    'Externo
    If Not extRowSum = 0 Then
        With ActiveWorkbook.ActiveSheet.Sort
            .SortFields.Clear
            .SortFields.Add Key:=Range("A" & moldRowSum + alumRowSum + pollRowSum + alhRowSum + 2), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Range("A" & moldRowSum + alumRowSum + pollRowSum + alhRowSum + 2, Cells(moldRowSum + alumRowSum + pollRowSum + alhRowSum + extRowSum + 1, Columns.Count).End(xlToLeft))
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
    
    
    Application.ScreenUpdating = True
    
'------------------------- ERROR HANDLING -------------------------

    Exit Sub
    
msgAbrirHistorico:
    MsgBox "Abra a planilha: HIST�RICO PRODU��O 2022-2024_V5.xlsm", vbOKOnly + vbExclamation, "Sem base de dados"
    
    Exit Sub
    
msgPlanilhaNomeErrado:
    MsgBox "Coloque o nome da tabela dessa forma:" & vbNewLine & vbNewLine & "Mes_Dia do mes_Ultimos dois digitos do ano" _
     & vbNewLine & vbNewLine & "Deve ficar assim:" & vbNewLine & "Fev_2_24" & vbNewLine & vbNewLine & _
     "� importante colocar os underlines(_).", vbOKOnly + vbExclamation, "Nome da tabela incorreto"
    
    Exit Sub
    
corrigirNome:
    MsgBox "Corrija os nomes na data que est� querendo fazer a tabela ante de executar o programa." & vbNewLine & vbNewLine & _
    "Na planilha HIST�RICO PRODU��O 2022-2024_V5, preencha todos os campos da coluna NOME CORRIGIDO", vbOKOnly + vbExclamation, "Nomes a ser corrigidos"
End Sub

Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function

Sub RelatorioEmVariasPlanilhas()

    Dim arquivo As String
    Dim i As Integer
    
    arquivo = "2022_Produ��o Por Ferramenta.xlsx"
    
    Workbooks(arquivo).Activate
    
    For i = 1 To 9
        ActiveWorkbook.Worksheets("22_0" & i).Activate
        
        RelatorioFerramentas
    Next i
    
    For i = 10 To 12
        ActiveWorkbook.Worksheets("22_" & i).Activate
        
        RelatorioFerramentas
    Next i
    
End Sub
