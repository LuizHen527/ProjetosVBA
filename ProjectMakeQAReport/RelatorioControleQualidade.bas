Attribute VB_Name = "RelatorioControleQualidade"
'@Folder("VBAProject")
Option Explicit
Public baseData() As String
Public productionYes As Byte, productionNo As Byte, productionProblem As Byte
Public selectedDate As Variant

'Bug:
    'RESOLVIDO: Ao gerar relatorio de abril de 2025 a data está vindo errado
    'O excel traduz tudo que vem do VBA pro formato da lingua que tá no excel (Portugues)
    'Quando eu mandei a data no formato brasileiro (dd/mm/yyyy) ele pensou que tava em ingles(mm/dd/yyyy)
    'Entao ele inverteu o dia e o mes na hora de passar a data pra planilha com a intenção de formatar pra portugues,
    'mas já estava em portugues
    
Sub CapturarDados()

    If ActiveSheet.Shapes("btnCancel").Visible = True Then
        MsgBox "Confirme ou cancele antes de gerar outro relatorio.", vbExclamation, "Botão desativado"
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    

    '--------------- VARIAVEIS ---------------
    Dim targetMonth As String, baseName As String
    Dim selectedDateResponse As VbMsgBoxResult
    Dim productionFolder As Object
    Dim iterator As Integer, i As Integer
    Dim convertedDate As Date
    
    productionYes = 0
    productionProblem = 0
    productionNo = 0
    
    '.getfolder("\\121.137.1.5\alumitec9\COMPRAS\25_Compras")
    Set productionFolder = CreateObject("Scripting.FileSystemObject").getfolder("\\121.137.1.5\alumitec9\PRODUÇÃO")

    '--------------- SELECIONANDO A DATA ---------------
    
    'pega data do ultimo relatorio
    selectedDate = Split(Range("J5").Value, "_")
    
    If LCase(selectedDate(0)) = "dezembro" Then
        selectedDate(0) = "janeiro"
        selectedDate(1) = CStr(CInt(selectedDate(1)) + 1)
        
        'Retorna o numero do mês
        targetMonth = month(DateValue("01 " & selectedDate(0) & " 2025"))
        
    Else
        targetMonth = month(DateValue("01 " & selectedDate(0) & " 2025")) + 1
    End If
    

    
    'Pergunta se usuario quer mes previsto
    selectedDateResponse = MsgBox("Quer pegar os dados da data abaixo?" & vbNewLine & vbNewLine & MonthName(targetMonth) & " de " & "20" & selectedDate(1) _
    , vbQuestion + vbYesNoCancel, "Selecionar data")
    
    If selectedDateResponse = vbNo Then
        'Chama função que mostra a caixa de input
        selectedDate = InputBoxDialog()
        
        'Caso a caixa de dialogo seja cancelada
        If Not IsArray(selectedDate) Then
            Exit Sub
        End If

    ElseIf selectedDateResponse = vbCancel Then
        'Executa se o usuario cancelar
        Exit Sub
        
    ElseIf selectedDateResponse = vbYes Then
        selectedDate(0) = MonthName(targetMonth)
    End If

    'tranforma mes em numero
    targetMonth = month(DateValue("01 " & selectedDate(0) & " 2025"))
    
    If targetMonth < 10 Then targetMonth = "0" & targetMonth
    
    '--------------- CAPTURANDO DADOS ---------------
    On Error GoTo FolderNotFound
    Workbooks.Open Filename:=productionFolder & "\" & "\20" & selectedDate(1) & " Extrusão e Produção\02_PRODUÇÃO DIÁRIA\" & targetMonth & " - PROD. DIÁRIA " & UCase(selectedDate(0)) & " 20" & selectedDate(1) & ".xlsm"
    On Error GoTo 0

    baseName = ActiveWorkbook.Name
    
    ActiveWorkbook.Worksheets("Base").Select
    
    ReDim baseData(Range("A5", "A" & Cells(Rows.count, 1).End(xlUp).Row).count - 1, 6)
    
    'Array que salva dados da base da producao diaria e conta os tipos de producao
    For iterator = 0 To Range("A5", "A" & Cells(Rows.count, 1).End(xlUp).Row).count - 1
        
        'Fiz essa conversao por causa de um bug bem estranho do excel
        'Vou fazer um video sobre esse bug mais tarde
        
        'Converter pra date type
        convertedDate = CDate(Range("A" & iterator + 5))
        
        'Mudar data pra formato americano
        convertedDate = Format(convertedDate, "mm/dd/yyyy")
        
        'Converter pra string denovo
        baseData(iterator, 0) = CStr(convertedDate)
        
        'Salva nome
        baseData(iterator, 1) = Range("E" & iterator + 5)
        
        'Salva produção
        baseData(iterator, 2) = Range("AM" & iterator + 5)
        
        'Salva problema
        baseData(iterator, 3) = Range("AN" & iterator + 5)
        
        'Salva observação
        baseData(iterator, 4) = Range("AO" & iterator + 5)
        
        'Salva numero
        baseData(iterator, 5) = Range("F" & iterator + 5)
        
        'Salva corte. Pra saber o que é parada de producao
        baseData(iterator, 6) = Range("I" & iterator + 5)
        
        'Conta produção = sim
        If Range("AM" & iterator + 5) = "SIM" And Not Range("AN" & iterator + 5) = "TESTE" Then productionYes = productionYes + 1
        
        'Conta produção = nao
        If Range("AM" & iterator + 5) = "NÃO" And Not Range("AN" & iterator + 5) = "TESTE" Then productionNo = productionNo + 1
        
        'Conta produção = problema
        If Range("AM" & iterator + 5) = "PROBLEMA" And Not Range("AN" & iterator + 5) = "TESTE" Then productionProblem = productionProblem + 1
        
    Next iterator
    
    ThisWorkbook.Worksheets("Relatório").Activate
    
    'filtrando dados para um array que tenha apenas ferramentas com problema
    For iterator = 0 To UBound(baseData)
        
        If (baseData(iterator, 3) = "RISCO" Or baseData(iterator, 3) = "ACABAMENTO" Or baseData(iterator, 3) = "") _
        And Not baseData(iterator, 1) = "PARADA PRODUÇÃO" _
        And Not baseData(iterator, 2) = "SIM" _
        And Not baseData(iterator, 6) = "" Then
            
            'data
            Range("P" & 3 + i) = baseData(iterator, 0)
            
            'nome
            Range("Q" & 3 + i).Value = baseData(iterator, 1)
            
            'produção
            Range("R" & 3 + i).Value = baseData(iterator, 2)
            
            'problema
            Range("S" & 3 + i).Value = baseData(iterator, 3)
            
            'observacao
            Range("T" & 3 + i).Value = baseData(iterator, 4)
            
            'index
            Range("U" & 3 + i).Value = iterator
            
            i = i + 1
        End If
    Next iterator
    
    Workbooks(targetMonth & " - PROD. DIÁRIA " & UCase(selectedDate(0)) & " 20" & selectedDate(1) & ".xlsm").Close False
    
    ActiveSheet.Shapes("btnCancel").Visible = True
    ActiveSheet.Shapes("btnConfirm").Visible = True
    ActiveSheet.Shapes("btnStart").Fill.ForeColor.RGB = RGB(115, 147, 179)
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Exit Sub
    
    '--------------- ERROR HANDLING ---------------

    'Caso não encontre o arquivo da produção diaria
FolderNotFound:
    MsgBox "Verifique se arquivo existe ou esta com o nome errado.", vbExclamation + vbOKOnly, "Arquivo não encontrado"
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub



Sub ConstruirTabelas()
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim iterator As Byte, countCorrida As Byte, countEsquadro As Byte, countOvalizacao As Byte, countDimensional As Byte, ferramenta As Byte, ite As Byte
    Dim countRisco As Byte, countQuebra As Byte, countOutros As Byte, countVazios As Byte, countBolha As Byte, countProblema As Byte
    Dim i As Integer, problemas As Integer, index As Integer
    Dim itensToUpdate As String, ferrProblemaQuant() As String, ferrTotais() As String, topFerr() As String, probTotais() As String, topProb() As String
    Dim bestItemName As String, bestItemProblem As String, bestItemQuant As String, initialRange As String, colorLine As String, formulaString As String
    Dim msgBoxAnswer As VbMsgBoxResult
    Dim rng As Range, compRng As Range
    
    
    itensToUpdate = "Quer mudar o problema dos itens abaixo?" & vbNewLine & vbNewLine & "NOME" & vbTab & vbTab & "ANTES" & vbTab & vbTab & "DEPOIS" & vbNewLine & vbNewLine
    
    For iterator = 0 To Range("p3", "p" & Cells(Rows.count, 16).End(xlUp).Row).count
        
        'Se problema do array não for igual problema da tabela, executa código
        If Not baseData(Range("U" & 3 + iterator).Value, 3) = Range("S" & 3 + iterator).Value And Not IsEmpty(Range("S" & 3 + iterator).Value) Then
        
            'Esse if é só pra deixar o numero de espaços na mensagem certos
            If Len(Range("S" & 3 + iterator).Value) >= 7 Then
                'String com mensagem de itens que serao atualizados
                itensToUpdate = itensToUpdate & Range("Q" & 3 + iterator).Value & vbTab & vbTab & baseData(Range("U" & 3 + iterator).Value, 3) & vbTab & Range("S" & 3 + iterator).Value & vbNewLine
            Else
                'String com mensagem de itens que serao atualizados
                itensToUpdate = itensToUpdate & Range("Q" & 3 + iterator).Value & vbTab & vbTab & baseData(Range("U" & 3 + iterator).Value, 3) & vbTab & vbTab & Range("S" & 3 + iterator).Value & vbNewLine
            End If

        End If
    Next iterator
    
    'Se ele atualizou algum item, pergunta se tem certeza
    If Len(itensToUpdate) > 67 Then
        msgBoxAnswer = MsgBox(itensToUpdate, vbQuestion + vbYesNo, "Confirmar mudanças")
        
        If msgBoxAnswer = vbNo Then Exit Sub
    End If
    
    
    'Atualiza dados pela planilha de verificação
    For iterator = 0 To Range("p3", "p" & Cells(Rows.count, 16).End(xlUp).Row).count
        
        'Se problema do array não for igual problema da tabela, executa código
        If Not baseData(Range("U" & 3 + iterator).Value, 3) = Range("S" & 3 + iterator).Value And Not IsEmpty(Range("S" & 3 + iterator).Value) Then
        
            baseData(Range("U" & 3 + iterator).Value, 3) = Range("S" & 3 + iterator).Value
            
        End If
    Next iterator
    
    '-------------- BASE DO RANKING --------------
    
    'Monta array com PERFIL, ERRO e QUANTIDADE pra fazer ranking
    ferramenta = 0
    ReDim ferrProblemaQuant(50, 2)
    
    For iterator = 0 To UBound(baseData)
        countProblema = 0
        
        If baseData(iterator, 3) = "" Or baseData(iterator, 3) = "TESTE" Then GoTo ProximaFerr
        
        'Verificar se a ferramenta e Nome já não esta no array
        If ferramenta > 0 Then
            For i = 0 To ferramenta - 1
                If ferrProblemaQuant(i, 0) = baseData(iterator, 1) And _
                ferrProblemaQuant(i, 1) = baseData(iterator, 3) Then
                
                    GoTo ProximaFerr
                    
                End If
            Next i
        End If
        
        'Contar todas as ferramentas com o mesmo NOME e PROBLEMA
        For i = 0 To UBound(baseData)
        
            If baseData(i, 1) = baseData(iterator, 1) And _
            baseData(i, 3) = baseData(iterator, 3) Then
                
                countProblema = countProblema + 1
            
            End If
        Next i
        
        
        'NOME
        ferrProblemaQuant(ferramenta, 0) = baseData(iterator, 1)
        
        'PROBLEMA
        ferrProblemaQuant(ferramenta, 1) = baseData(iterator, 3)
        
        'QUANTIDADE
        ferrProblemaQuant(ferramenta, 2) = CStr(countProblema)
        
        'Debug.Print ferrProblemaQuant(ferramenta, 0) & "," & ferrProblemaQuant(ferramenta, 1) & "," & ferrProblemaQuant(ferramenta, 2)
        
        ferramenta = ferramenta + 1
        
ProximaFerr:
        
    Next iterator
    
    
    
    '-------------- TOTAIS DE FERRAMENTAS --------------
    
    ReDim ferrTotais(50, 1)
    ferramenta = 0
    'Monta um array com as ferramentas e os totais de problemas
    For iterator = 0 To UBound(ferrProblemaQuant)
        problemas = 0
        
        If ferrProblemaQuant(iterator, 0) = "" Then GoTo GoNext
        
        
        'Se a ferramenta já esta no array, proxima iteracao
        If ferramenta > 0 Then
            For i = 0 To ferramenta - 1
                If ferrProblemaQuant(iterator, 0) = ferrTotais(i, 0) Then
                
                    GoTo GoNext
                    
                End If
            Next i
        End If
              
              
        'Procura por todas as ocorrencias da ferramenta e soma problemas
        For i = 0 To UBound(ferrProblemaQuant)
            If ferrProblemaQuant(iterator, 0) = ferrProblemaQuant(i, 0) Then
            
                'Debug.Print ferrProblemaQuant(i, 0) & "," & ferrProblemaQuant(iterator, 1) & "," & ferrProblemaQuant(iterator, 2)
            
                problemas = problemas + ferrProblemaQuant(i, 2)
                
            End If
        Next i
        
        'Debug.Print ferrProblemaQuant(iterator, 0) & "," & problemas
        
        'NOME
        ferrTotais(ferramenta, 0) = ferrProblemaQuant(iterator, 0)
        
        'QUANTIDADE
        ferrTotais(ferramenta, 1) = problemas
        
        'Debug.Print ferrTotais(ferramenta, 0) & "," & ferrTotais(ferramenta, 1)
        
        ferramenta = ferramenta + 1
            
GoNext:
    Next iterator

    '-------------- FERRAMENTAS COM MAIS ERROS --------------
    
    Range("A21", "D" & Cells(Rows.count, 1).End(xlUp).Row + 1).Delete (xlShiftUp)
        
    'Loopar 5 vezes pegando 5 ferramentas com mais problemas
    ReDim topFerr(4, 1)
    
    For iterator = 0 To 4
    
        
        
        'Procurar ferramenta com maior quantidade de problemas
        For i = 0 To UBound(ferrTotais)
        
            'Se a ferramenta já estiver no array, pula
            If iterator > 0 Then
                For ite = 0 To iterator
                    If ferrTotais(i, 0) = topFerr(ite, 0) Then
                    
                        GoTo GoNextFerr
                        
                    End If
                Next ite
            End If
        
        
            'Se a quantidade de erro for maior, salva essa ferramenta
            If ferrTotais(i, 1) > topFerr(iterator, 1) Then
                'NOME
                topFerr(iterator, 0) = ferrTotais(i, 0)
        
                'QUANTIDADE
                topFerr(iterator, 1) = ferrTotais(i, 1)
            End If
            
GoNextFerr:
        Next i
        
        'Debug.Print topFerr(iterator, 0) & "," & topFerr(iterator, 1)
        
        '-------------- MONTANDO TABELA DO RANKING DE FERRAMENTAS --------------
        
        'Pega todos os erros e quantidades da ferramenta que foi colocada array topFerr
        colorLine = "azul"
        For i = 0 To UBound(ferrProblemaQuant)
            If topFerr(iterator, 0) = ferrProblemaQuant(i, 0) Then
                'Debug.Print ferrProblemaQuant(i, 0) & "," & ferrProblemaQuant(i, 1) & "," & ferrProblemaQuant(i, 2)
                
                
                'Colocar ferramenta na tabela
                
                'PERFIL
                Range("A" & Cells(Rows.count, 1).End(xlUp).Row).Offset(1, 0).Value = ferrProblemaQuant(i, 0)
                
                'PROBLEMA
                Range("B" & Cells(Rows.count, 2).End(xlUp).Row).Offset(1, 0).Value = ferrProblemaQuant(i, 1)
                
                'QUANTIDADE
                Range("D" & Cells(Rows.count, 4).End(xlUp).Row).Offset(1, 0).Value = ferrProblemaQuant(i, 2)
                
                'Mescla celulas de problema
                
                Range("B" & Cells(Rows.count, 1).End(xlUp).Row, "C" & Cells(Rows.count, 1).End(xlUp).Row).Merge
                
                Range("B" & Cells(Rows.count, 2).End(xlUp).Row).HorizontalAlignment = xlCenter
                Range("B" & Cells(Rows.count, 2).End(xlUp).Row).VerticalAlignment = xlCenter
                
                
            End If
            
            
        Next i
        
        
        
    Next iterator
    
    
    '-------------- ORGANIZA RANKING --------------
    
    For Each rng In Range("A21:" & Cells(Rows.count, 1).End(xlUp).Address)
        bestItemQuant = 0
          
        'Verifica colocados no ranking. Salvando o que tem maior numero de erros
        For Each compRng In Range(rng.Offset(1, 0).Address & ":" & Cells(Rows.count, 1).End(xlUp).Address)
            
            'Se nome for igual, defeito diferente e quan
            If rng.Value = compRng.Value And _
            rng.Offset(0, 1).Value <> compRng.Offset(0, 1).Value And _
            rng.Offset(0, 3).Value < compRng.Offset(0, 3).Value And _
            compRng.Offset(0, 3) > bestItemQuant Then
                
                bestItemName = compRng.Value
                bestItemProblem = compRng.Offset(0, 1).Value
                bestItemQuant = compRng.Offset(0, 3).Value
            End If
        Next compRng
        
        'Se for encontrado um melhor colocado
        If bestItemQuant > 0 Then
            
            'Colocar o menor colocado em baixo
            For Each compRng In Range(rng.Offset(1, 0).Address & ":" & Cells(Rows.count, 1).End(xlUp).Address)
                
                If compRng.Value = bestItemName And _
                compRng.Offset(0, 1).Value = bestItemProblem Then
                    
                    'NOME
                    compRng.Value = rng.Value
                    
                    'PROBLEMA
                    compRng.Offset(0, 1).Value = rng.Offset(0, 1).Value
                    
                    'QUANTIDADE
                    compRng.Offset(0, 3).Value = rng.Offset(0, 3).Value
                    
                    Exit For
                End If
            Next compRng
            
            'Coloca melhor colocado em cima
            
            'NOME
            rng.Value = bestItemName
            
            'PROBLEMA
            rng.Offset(0, 1).Value = bestItemProblem
            
            'QUANTIDADE
            rng.Offset(0, 3).Value = bestItemQuant
            
            
        End If
        
        
    Next rng
    
    '-------------- ESTILIZANDO O CONTEUDO DA TABELA DE RANKING --------------
    
    
    With Range("A21:" & Cells(Rows.count, 4).End(xlUp).Address)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.Color = vbWhite
        .Borders.Weight = xlThin
    End With
    
    initialRange = Range("A21").Address
    
    For Each rng In Range("A21:" & Cells(Rows.count, 1).End(xlUp).Address).Offset(1, 0)
        
        If rng.Value <> Range(initialRange).Value Then
            With Range(initialRange & ":" & rng.Offset(-1, 0).Address)
                .Merge
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Borders.Color = vbWhite
            End With
            
            If colorLine = "azul" Then
                With Range(initialRange & ":" & rng.Offset(-1, 3).Address)
                    .Interior.Color = RGB(184, 204, 228)
                    .Borders(xlInsideVertical).Weight = xlThin
                    .Borders(xlInsideVertical).Color = vbWhite
                    .Borders(xlEdgeRight).Color = vbWhite
                    .Borders(xlEdgeLeft).Color = vbWhite
                End With
                
                colorLine = "branco"
            Else
                With Range(initialRange & ":" & rng.Offset(-1, 3).Address)
                    .Interior.Color = vbWhite
                    .Borders(xlInsideVertical).Weight = xlThin
                    .Borders(xlInsideVertical).Color = RGB(184, 204, 228)
                    .Borders(xlEdgeRight).Color = vbWhite
                    .Borders(xlEdgeLeft).Color = vbWhite
                End With
                
                colorLine = "azul"
            End If
            
            
            
            
            initialRange = rng.Address
        End If
    Next rng
    
    
    With Range("A" & Cells(Rows.count, 2).End(xlUp).Row, "C" & Cells(Rows.count, 2).End(xlUp).Row).Offset(1, 0)
        .Merge
        .Value = "SOMA DE PARADAS"
        .Interior.Color = RGB(247, 150, 70)
        .Font.Color = vbWhite
        .Borders.Weight = xlThin
        .Borders.Color = vbWhite
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
    End With
    
    Range("A" & Cells(Rows.count, 1).End(xlUp).Row, "C" & Cells(Rows.count, 1).End(xlUp).Row).Copy
    
    Range("A" & Cells(Rows.count, 1).End(xlUp).Row, "C" & Cells(Rows.count, 1).End(xlUp).Row).Offset(1, 0).PasteSpecial (xlAll)
    
    With Range("A" & Cells(Rows.count, 1).End(xlUp).Row, "C" & Cells(Rows.count, 1).End(xlUp).Row)
        .Value = "%"
        .Interior.Color = RGB(151, 71, 7)
    End With
    
    formulaString = "=SUM(D21:" & Range("D" & Cells(Rows.count, 1).End(xlUp).Row - 1).Offset(-1, 0).Address & ")"
    
    With Range("D" & Cells(Rows.count, 1).End(xlUp).Row - 1)
        .Formula = formulaString
        .Interior.Color = RGB(247, 150, 70)
        .Font.Color = vbWhite
        .Borders.Weight = xlThin
        .Borders.Color = vbWhite
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
    End With
    
    formulaString = "=" & Range("D" & Cells(Rows.count, 1).End(xlUp).Row - 1).Address & "/(" & "J7+J8)"
    
    With Range("D" & Cells(Rows.count, 1).End(xlUp).Row)
        .Formula = formulaString
        .Interior.Color = RGB(151, 71, 7)
        .Font.Color = vbWhite
        .Borders.Weight = xlThin
        .Borders.Color = vbWhite
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .NumberFormat = "0%"
    End With

    
    'Rankear 4 maiores problemas
    
    
    '-------------- TOTAIS DE PROBLEMAS --------------
    ReDim probTotais(20, 1)
    index = 0
    problemas = 0
    
    For iterator = 0 To UBound(ferrProblemaQuant)
    
        'Verificar se o nome já não esta no array
        If iterator > 0 Then
            For i = 0 To UBound(probTotais)
                If ferrProblemaQuant(iterator, 1) = probTotais(i, 0) Then
                    GoTo NextProb
                End If
            Next i
        End If
        
        'Nome
        probTotais(index, 0) = ferrProblemaQuant(iterator, 1)
    
        'Somar todos os problemas com o mesmo nome
        For i = 0 To UBound(ferrProblemaQuant)
        
            If ferrProblemaQuant(iterator, 1) = ferrProblemaQuant(i, 1) Then
                
                'Soma
                problemas = problemas + CInt(ferrProblemaQuant(i, 2))
                
            End If
        Next i
        
        probTotais(index, 1) = problemas
        
        Debug.Print probTotais(index, 0) & "," & probTotais(index, 1)
        
        problemas = 0
        
        index = index + 1
NextProb:
        
    Next iterator
    
    
    
    '-------------- BUSCANDO ERROS COM MAIS OCORRENCIAS --------------
    
    'iterator
    'ite
    'GoNextProb
    ReDim topProb(3, 1)
    
    Debug.Print "----------- MELHORES -----------"
    
    'Procurar ferramenta com maior quantidade de problemas
    For iterator = 0 To 3
    
        For i = 0 To UBound(probTotais)
        
            'Debug.Print probTotais(i, 1) & "(" & probTotais(i, 0) & "), " & topProb(iterator, 1) & "(" & topProb(iterator, 0) & ")"
            
            If probTotais(i, 1) = "" Then GoTo GoNextProb
            
            'Se a ferramenta já estiver no array, pula
            If iterator > 0 Then
                For ite = 0 To iterator
                    If probTotais(i, 0) = topProb(ite, 0) Then
                    
                        GoTo GoNextProb
                        
                    End If
                Next ite
            End If
        
            If topProb(iterator, 1) = "" Then topProb(iterator, 1) = "0"
            
            
            If CInt(probTotais(i, 1)) > CInt(topProb(iterator, 1)) Then
                'NOME
                topProb(iterator, 0) = probTotais(i, 0)
        
                'QUANTIDADE
                topProb(iterator, 1) = probTotais(i, 1)
                
                'Se a quantidade de erro for maior, salva essa ferramenta
                'Debug.Print topProb(iterator, 0) & "," & topProb(iterator, 1)
            End If
            
GoNextProb:
        Next i
        
        
        Debug.Print topProb(iterator, 0) & "," & topProb(iterator, 1)
        
    Next iterator
    
    
    '-------------- MOTANDO TABELA DE ERROS MAIS RECORRENTES --------------
    
    colorLine = "azul"
    For iterator = 0 To 3
        Range("F" & 21 + iterator, "G" & 21 + iterator).Merge
        Range("F" & 21 + iterator).Value = topProb(iterator, 0)
        Range("H" & 21 + iterator).Value = topProb(iterator, 1)
        Range("I" & 21 + iterator).Formula = "=H" & 21 + iterator & "/J6"
        Range("I" & 21 + iterator).NumberFormat = "0%"
        With Range("F" & 21 + iterator, "I" & 21 + iterator)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            
        End With
        
        If colorLine = "azul" Then
        
            With Range("F" & 21 + iterator, "I" & 21 + iterator)
                .Interior.Color = RGB(184, 204, 228)
                .Borders(xlInsideVertical).Weight = xlThin
                .Borders(xlInsideVertical).Color = vbWhite
                .Borders(xlEdgeRight).Color = vbWhite
                .Borders(xlEdgeLeft).Color = vbWhite
            End With
            
            colorLine = "branco"
        Else
        
            With Range("F" & 21 + iterator, "I" & 21 + iterator)
                .Interior.Color = vbWhite
                .Borders(xlInsideVertical).Weight = xlThin
                .Borders(xlInsideVertical).Color = RGB(184, 204, 228)
                .Borders(xlEdgeRight).Color = vbWhite
                .Borders(xlEdgeLeft).Color = vbWhite
            End With
            
            colorLine = "azul"
        End If
    Next iterator
    
    'Footer
    Range("F25", "G25").Merge
    Range("F25").Value = "SOMA"
    Range("H25").Formula = "=SUM(H21:H24)"
    Range("I25").Formula = "=H25/SUM(J7:J8)"
    Range("I25").NumberFormat = "0%"
    
    With Range("F25", "I25")
        .Interior.Color = RGB(151, 72, 7)
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideVertical).Color = vbWhite
    End With
    
    
    'Se tiver dando bug na contagem, ative o codigo abaixo pra debugar
'    Debug.Print "BOLHA: " & countBolha
'    Debug.Print "CORRIDA: " & countCorrida
'    Debug.Print "DIMENSIONAL: " & countDimensional
'    Debug.Print "ESQUADRO: " & countEsquadro
'    Debug.Print "OVAL: " & countOvalizacao
'    Debug.Print "QUEBRA: " & countQuebra
'    Debug.Print "RISCO: " & countRisco
'    Debug.Print "VAZIOS: " & countVazios
'    Debug.Print "OUTROS: " & countOutros
    
    
    'Apaga primeira coluna e passa dados das outras colunas pra esquerda
    Application.CutCopyMode = False
    Range("B6:B8").ClearContents
    Range("D6:D8").Copy
    Range("B6").Select
    ActiveSheet.Paste
    Range("D6:D8").ClearContents
    Range("F6:F8").Copy
    Range("D6").Select
    ActiveSheet.Paste
    Range("F6:F8").ClearContents
    Range("H6:H8").Copy
    Range("F6").Select
    ActiveSheet.Paste
    Range("H6:H8").ClearContents
    Range("J6:J8").Copy
    Range("H6").Select
    ActiveSheet.Paste
    Range("J6:J8").ClearContents
    Range("B5:C5").ClearContents
    Range("D5:E5").Copy
    Range("B5:C5").Select
    ActiveSheet.Paste
    Range("D5:E5").ClearContents
    Range("F5:G5").Copy
    Range("D5:E5").Select
    ActiveSheet.Paste
    Range("F5:G5").ClearContents
    Range("H5:I5").Copy
    Range("F5:G5").Select
    ActiveSheet.Paste
    Range("H5:I5").ClearContents
    Range("J5:K5").Copy
    Range("H5:I5").Select
    ActiveSheet.Paste
    Range("J5:K5").ClearContents
    
    'Colocar os dados de quantidade do mes
    Range("J5").Value = CStr(UCase(selectedDate(0)) & "_" & selectedDate(1))
    Range("J6").Value = productionYes
    Range("J7").Value = productionNo
    Range("J8").Value = productionProblem
    

    'Apagar array de dados
    Erase baseData

    'Apaga lista de verificação
    If Not Range("p3").Value = "" Then
        Range("p3:u" & Cells(Rows.count, 21).End(xlUp).Row).Delete
    End If

    ActiveSheet.Shapes("btnStart").Visible = True
    ActiveSheet.Shapes("btnConfirm").Visible = False
    ActiveSheet.Shapes("btnCancel").Visible = False
    ActiveSheet.Shapes("btnStart").Fill.ForeColor.RGB = RGB(11, 29, 81)
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub

Sub btnCancel()

    Application.ScreenUpdating = False
    
    'Apagar array de dados
    Erase baseData
    
    'Apaga lista de verificação
    If Not Range("p3").Value = "" Then
        Range("p3:u" & Cells(Rows.count, 21).End(xlUp).Row).Delete
    End If

    'Esconde botões
    ActiveSheet.Shapes("btnStart").Visible = True
    ActiveSheet.Shapes("btnConfirm").Visible = False
    ActiveSheet.Shapes("btnCancel").Visible = False
    ActiveSheet.Shapes("btnStart").Fill.ForeColor.RGB = RGB(11, 29, 81)
    
    Application.ScreenUpdating = True
End Sub

Sub GerarPDF()
    
    'Definir caminho que o pdf sera salvo
    'Mudar orientação pra retrato
    'Definir pagina como retrato
    'Definir area de impressao
    'Exportar como pdf
    
    Dim folderPath As Object
    Dim lastDate As String
    Dim filePath As Variant
    
    lastDate = ActiveSheet.Range("J5").Value
    
    Set folderPath = CreateObject("Scripting.FileSystemObject").getfolder("\\121.137.1.5\manutencao1\Lucas\21_Luiz\QA")
    filePath = folderPath & "\Rel_QA" & lastDate & ".pdf"
    
    Debug.Print filePath
    
    With ActiveSheet.PageSetup
        .Orientation = xlPortrait
        .PrintArea = "$A$1:$K$51"
    End With
    
    'ActiveWorkbook.sheet().ExportAsFixedFormat Type:=xlTypePDF, Filename:="aaaaaaaa.pdf"
    ' Export the active sheet as PDF
    'ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=filePath, Quality:=xlQualityStandard
    
    ThisWorkbook.Sheets("Relatório").ExportAsFixedFormat Type:=xlTypePDF, Filename:=folderPath, Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    
End Sub





'----------------------------- FUNCTIONS -----------------------------


'Funtion que mostra a caixa e lida com o tratamento de excessao.
'Retorna false se o usuario cancelar ou clicar no "X". Retorna o mes e o ano(ABRIL, 25) caso o usuario digite corretamente
Function InputBoxDialog() As Variant

    '--------------- VARIAVEIS ---------------
    Dim inputBoxAnswer As Variant
    Dim returnValue() As String
    Dim verifyMonth As Boolean
    
    InputBoxDialog = False

    '--------------- INPUTBOX ---------------
InputBoxError:

    inputBoxAnswer = Application.InputBox("Escreva a data que deseja:" & vbNewLine & vbNewLine & "Siga o seguinte padrão: abril_24", "Selecione uma data", , , , , , 2 + 4 + 16)
    
    'Finaliza macro caso ele clique em cancelar ou no X
    If inputBoxAnswer = False Then
        Exit Function
    End If
    
    'Separa mes e ano
    returnValue() = Split(inputBoxAnswer, "_")
    
    'Verifica se o mes existe
    verifyMonth = VerificaMes(LCase(returnValue(0)))
    
    'Tratamento de excessoes
    If verifyMonth = False Then
        MsgBox "Digite um mês valido.", vbExclamation, "Aviso"
        GoTo InputBoxError:
        
    ElseIf UBound(returnValue, 1) < 1 Then
        MsgBox "Digite um mês e um ano. Separe eles com um underline (_). Dessa forma: " _
        & " abril_25", vbExclamation, "Aviso"
        GoTo InputBoxError:
        
    ElseIf returnValue(1) = "" Then
        MsgBox "Digite um mês e um ano. Separe eles com um underline (_). Dessa forma: " _
        & " abril_25", vbExclamation, "Aviso"
        GoTo InputBoxError:
        
    ElseIf returnValue(1) < 24 Or returnValue(1) > 40 Or returnValue(1) = "" Then
        MsgBox "Digite um ano valido.(De 2023 pra frente) ", vbExclamation, "Aviso"
        GoTo InputBoxError:
    
    End If

    InputBoxDialog = returnValue
    
End Function

'Function que valida se o mes digitado pode ser usado ou nao.
'Recebe o nome do mes como parametro
'Retorna false caso nao seja um mes valido. True se for um mes valido.
Function VerificaMes(mes As String) As Boolean

    Dim meses As Variant
    Dim n As Integer
    
    LCase (mes)
    
    meses = Array("janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro")

    For n = 0 To 11
        If mes = meses(n) Then
            'Mes valido
            VerificaMes = True
            
            Exit Function
        End If
    Next n
    
    'Mes nao valido
    VerificaMes = False
    
End Function


