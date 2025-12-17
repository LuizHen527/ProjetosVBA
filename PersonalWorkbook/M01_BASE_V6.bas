Attribute VB_Name = "M01_BASE_V6"
Public inicial_range As Integer
Public last_line As Integer
'Macro para organizar os dados da planilha à faturar da MOLDUCOLOR de acordo com a BD TecSerp

Sub M1_BASE_V3()

Dim objCelulaAtual As Range
Dim varCelula As Variant
                                                         
'Mensagem de inícido do código
'MsgBox "Start!"

''As mudanças do programa não aparecem passo-a-passo na tela
'Application.ScreenUpdating = False
Application.DisplayAlerts = False

Application.Calculation = xlCalculationAutomatic

'range_row inicial da Base
inicial_range = Cells(Cells.Rows.Count, 1).End(xlUp).Row

'Transforma o nome da planilha base em variavel
plan_name = ActiveWorkbook.Name
last_line = Cells(Cells.Rows.Count, 1).End(xlUp).Row


'Configura as abas iniciais: cria abas, exclui abas, altera nomes e oculta aba
Sheets(1).Select
    Sheets(1).Name = "Base"
Sheets(Array(2, 3)).Select
    ActiveWindow.SelectedSheets.Delete
Sheets("Base").Select
    Sheets("Base").Copy After:=Sheets(1)
Sheets("Base (2)").Select
    Sheets("Base (2)").Name = "Macro"
        Sheets("Base").Select
            ActiveWindow.SelectedSheets.Visible = False

'Apagas as colunas não desejadas
Range("B:E, G:I, K:K, R:R, V:V, X:X, Z:Z, AC:AC, AE:BL").Select
    Selection.Delete Shift:=xlToLeft

'Cria coluna "Conv. Unid"

Range("R1").Value = "Conv. Unid"

Set objCelulaAtual = Range(Cells(Rows.Count, 17).End(xlUp).Address)

objCelulaAtual.Select

While Not (objCelulaAtual.Address = "$Q$1")

    If objCelulaAtual.Value > 0 Then
        If objCelulaAtual.Offset(-1, 0) > 0 Then
            'Pega o range e coloca "PEÇA" do lado de todas as cells
            For Each varCelula In Range(objCelulaAtual.Address, objCelulaAtual.End(xlUp).Address)
                varCelula.Offset(0, 1).Value = "PEÇA"
            Next
            
            'Pula pra proxima celula
            ActiveCell.End(xlUp).Select
            ActiveCell.End(xlUp).Select
            
        Else
            objCelulaAtual.Offset(0, 1).Value = "PEÇA"
            'Pula pra proxima celula
            ActiveCell.End(xlUp).Select
            
        End If
        
    Else
        'Pula pra proxima celula
        ActiveCell.End(xlUp).Select
    
    End If
    
    Set objCelulaAtual = Range(ActiveCell.Address)
    
Wend


'Transforma data quebrada em data inteira               '
Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Columns("A:A").Select
    Selection.NumberFormat = "0.00"
Range("B2").Select
    ActiveCell.FormulaR1C1 = "=INT(RC[-1])"
        Selection.AutoFill Destination:=Range("B2:B" & last_line)

'Converte a hora
Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("C2").Select
            ActiveCell.FormulaR1C1 = "=RC[-2]-RC[-1]"
            Selection.AutoFill Destination:=Range("C2:C" & last_line)
        Range("C2").Select
            Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
        Range("C2").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False

Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Columns("A:A").Select
    Selection.NumberFormat = "dd/mm/yy;@"
Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft

Range("B1").Value = "Hora"
    Columns("B:B").Select
        Selection.NumberFormat = "h:mm;@"

'Tranforma txt em número
Set rngPedido = Range("H2" & ":" & "H" & last_line)
    With rngPedido
        .NumberFormat = "General"
        .FormulaLocal = rngPedido.Value
    End With

'---------------------------------- Processo de preenchimento de dados -----------------------------------------------------

 'Processo para reorganizar as colunas, altera a ordem das colunas (total,Qtde e Unidade)
Columns("K:K").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Columns("R:T").Select
    Selection.Cut
Range("K1").Select
    ActiveSheet.Paste

'Processo para reorganizar as colunas, adiciona o espaco para as colunas unidade, subunida., peso/m e peso
'E organiza a ordem das colunas caracte. e acabament.
Columns("Q:Q").Select
    Application.CutCopyMode = False
        Selection.Cut
Range("S1").Select
    ActiveSheet.Paste
Columns("P:P").Select
    Selection.Cut
Range("Q1").Select
    ActiveSheet.Paste
Columns("T:T").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

'Desabilita os calculos automatico
Application.Calculation = xlCalculationManual

'Alocando COnv qtd e unid
Columns("X:Y").Select
    Selection.Copy
    Range("AD1").Select
        ActiveSheet.Paste

'Processo para arrumar os titulos do intervalo E1:V1
'Abri a planilha da calculadora conferir o local do arquivo
'MsgBox "Organizando..."
Workbooks.Open fileName:="\\121.137.1.5\manutencao1\Lucas\09_Banco Dados\BD TecSerp.xlsm"
    Windows("BD TecSerp.xlsm").Activate
        Sheets("Análise").Select
            Range("E1").Select
                Range(Selection, Selection.End(xlToRight)).Select
                    Selection.Copy
    Windows(plan_name).Activate
    Sheets("Macro").Select
        Range("N1").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False

'Procv referente a coluna N, 5.familia
Range("N2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-6],'[BD TecSerp.xlsm]Análise'!R1C1:R90000C13,5,0),"""")"
        With ActiveSheet
            LastRow = Range("A" & .Rows.Count).End(xlUp).Row
            Set RngAutopreencher = Range("N2:N" & LastRow)
        End With
    Selection.AutoFill Destination:=RngAutopreencher

'Procv referente a coluna O, 6.Identificacao
Range("O2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-7],'[BD TecSerp.xlsm]Análise'!R1C1:R90000C13,6,0),"""")"
        With ActiveSheet
            LastRow = Range("A" & .Rows.Count).End(xlUp).Row
            Set RngAutopreencher = Range("O2:O" & LastRow)
        End With
    Selection.AutoFill Destination:=RngAutopreencher

'Procv referente a coluna P, 7.Subidentificacao
Range("P2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-8],'[BD TecSerp.xlsm]Análise'!R1C1:R90000C13,7,0),"""")"
        With ActiveSheet
            LastRow = Range("A" & .Rows.Count).End(xlUp).Row
            Set RngAutopreencher = Range("P2:P" & LastRow)
        End With
    Selection.AutoFill Destination:=RngAutopreencher

'Procv referente a coluna Q, 8.Formato
Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-9],'[BD TecSerp.xlsm]Análise'!R1C1:R90000C13,8,0),"""")"
        With ActiveSheet
            LastRow = Range("A" & .Rows.Count).End(xlUp).Row
            Set RngAutopreencher = Range("Q2:Q" & LastRow)
        End With
    Selection.AutoFill Destination:=RngAutopreencher

'Procv referente a coluna R, 9.Caracteristica
Range("R2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-10],'[BD TecSerp.xlsm]Análise'!R1C1:R90000C13,9,0),"""")"
        With ActiveSheet
            LastRow = Range("A" & .Rows.Count).End(xlUp).Row
            Set RngAutopreencher = Range("R2:R" & LastRow)
        End With
    Selection.AutoFill Destination:=RngAutopreencher

'Procv referente a coluna S, 10.Acabamento
Range("S2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-11],'[BD TecSerp.xlsm]Análise'!R1C1:R90000C13,10,0),"""")"
        With ActiveSheet
            LastRow = Range("A" & .Rows.Count).End(xlUp).Row
            Set RngAutopreencher = Range("S2:S" & LastRow)
        End With
    Selection.AutoFill Destination:=RngAutopreencher

'Procv referente a coluna T, 11.Material
Range("T2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-12],'[BD TecSerp.xlsm]Análise'!R1C1:R90000C13,11,0),"""")"
        With ActiveSheet
            LastRow = Range("A" & .Rows.Count).End(xlUp).Row
            Set RngAutopreencher = Range("T2:T" & LastRow)
        End With
    Selection.AutoFill Destination:=RngAutopreencher

'Procv referente a coluna U, 12.Comprimento
Range("U2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-13],'[BD TecSerp.xlsm]Análise'!R1C1:R90000C13,12,0),"""")"
        With ActiveSheet
            LastRow = Range("A" & .Rows.Count).End(xlUp).Row
            Set RngAutopreencher = Range("U2:U" & LastRow)
        End With
     Selection.AutoFill Destination:=RngAutopreencher

'Procv referente a coluna V, 13.Altura
Range("V2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-14],'[BD TecSerp.xlsm]Análise'!R1C1:R90000C13,13,0),"""")"
        With ActiveSheet
            LastRow = Range("A" & .Rows.Count).End(xlUp).Row
            Set RngAutopreencher = Range("V2:V" & LastRow)
        End With
     Selection.AutoFill Destination:=RngAutopreencher

'Procv referente a coluna W, 14.Unidade
Range("W2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-15],'[BD TecSerp.xlsm]Análise'!R1C1:R90000C14,14,0),"""")"
        With ActiveSheet
            LastRow = Range("A" & .Rows.Count).End(xlUp).Row
            Set RngAutopreencher = Range("W2:W" & LastRow)
        End With
     Selection.AutoFill Destination:=RngAutopreencher


'Procv referente a coluna X, 15.SubUnidade
Range("X2").Select
    ActiveCell.FormulaR1C1 = _
    "=IFERROR(VLOOKUP(RC[-16],'[BD TecSerp.xlsm]Análise'!R1C1:R90000C15,15,0),"""")"
        With ActiveSheet
            LastRow = Range("A" & .Rows.Count).End(xlUp).Row
            Set RngAutopreencher = Range("X2:X" & LastRow)
        End With
     Selection.AutoFill Destination:=RngAutopreencher

'Procv referente a coluna Y, 16.Peso/m
Range("Y2").Select
    ActiveCell.FormulaR1C1 = _
    "=IFERROR(VLOOKUP(RC[-17],'[BD TecSerp.xlsm]Análise'!R1C1:R90000C16,16,0),"""")"
        With ActiveSheet
            LastRow = Range("A" & .Rows.Count).End(xlUp).Row
            Set RngAutopreencher = Range("Y2:Y" & LastRow)
        End With
     Selection.AutoFill Destination:=RngAutopreencher

'Procv referente a coluna Z, 17.Peso
Range("Z2").Select
    ActiveCell.FormulaR1C1 = _
    "=IFERROR(VLOOKUP(RC[-18],'[BD TecSerp.xlsm]Análise'!R1C1:R90000C17,17,0),"""")"
        With ActiveSheet
            LastRow = Range("A" & .Rows.Count).End(xlUp).Row
            Set RngAutopreencher = Range("Z2:Z" & LastRow)
        End With
     Selection.AutoFill Destination:=RngAutopreencher

'Procv referente a coluna AA, 18.Custo
Range("AA2").Select
    ActiveCell.FormulaR1C1 = _
    "=IFERROR(VLOOKUP(RC[-19],'[BD TecSerp.xlsm]Análise'!R1C1:R90000C18,18,0),"""")"
        With ActiveSheet
            LastRow = Range("A" & .Rows.Count).End(xlUp).Row
            Set RngAutopreencher = Range("AA2:AA" & LastRow)
        End With
     Selection.AutoFill Destination:=RngAutopreencher

'Procv referente a coluna AB, 19.?
Range("AB2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-20],'[BD TecSerp.xlsm]Análise'!R1C1:R90000C19,19,0),"""")"
        With ActiveSheet
            LastRow = Range("A" & .Rows.Count).End(xlUp).Row
            Set RngAutopreencher = Range("AB2:AB" & LastRow)
        End With
     Selection.AutoFill Destination:=RngAutopreencher

'Procv referente a coluna AB, 20.X
Range("AC1").Select
    ActiveCell.FormulaR1C1 = "20.VML"
Range("AC2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-26],Base!C6:C7,2,0)"
        With ActiveSheet
            LastRow = Range("A" & .Rows.Count).End(xlUp).Row
            Set RngAutopreencher = Range("AC2:AC" & LastRow)
        End With
     Selection.AutoFill Destination:=RngAutopreencher

Application.Calculation = xlCalculationAutomatic

Columns("AC:AC").Select
    Selection.Copy
Range("AC1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

Application.Calculation = xlCalculationManual

'Seleciona a célula A2 e armazena o número da última linha em last_line
Range("A2").Select
last_line = Cells(Cells.Rows.Count, 1).End(xlUp).Row


'Classifica as celulas por data -> cliente -> Pedido -> Código -> Quantidade
    Range("A1").Select
    Selection.CurrentRegion.Select
    ActiveWorkbook.Worksheets("Macro").Sort.SortFields.Clear
    
    'Data
    ActiveWorkbook.Worksheets("Macro").Sort.SortFields.Add Key:= _
        Range("A2:A" & last_line), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    'Cliente
    ActiveWorkbook.Worksheets("Macro").Sort.SortFields.Add Key:= _
        Range("D2:D" & last_line), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    'Hora
    ActiveWorkbook.Worksheets("Macro").Sort.SortFields.Add Key:= _
        Range("B2:B" & last_line), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    'Pedido
    ActiveWorkbook.Worksheets("Macro").Sort.SortFields.Add Key:= _
        Range("C2:C" & last_line), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    'Código
    ActiveWorkbook.Worksheets("Macro").Sort.SortFields.Add Key:= _
        Range("H1:H" & last_line), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    'Quantidade
    ActiveWorkbook.Worksheets("Macro").Sort.SortFields.Add Key:= _
        Range("L1:L" & last_line), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal

    With ActiveWorkbook.Worksheets("Macro").Sort
        .SetRange Range("A1:AJ" & last_line)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'Remove valores duplicados referentes aos pedidos
Range("C1").Select
    For PEDIDO = 1 To last_line
        aa = ActiveCell.Value
        ab = ActiveCell.Offset(1, 0).Value
            If aa = ab Then
                Selection.ClearContents
            End If
        ActiveCell.Offset(1, 0).Select
    Next


'Copia e cola os valores das colunas para análise de duplicadas
'Total
Range("K2:K" & last_line).Copy
Range("AH2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'Código
Range("H2:H" & last_line).Copy
Range("AI2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'Qtd
Range("L2:L" & last_line).Copy
Range("AJ2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'Insere 50 linhas vazias para evitar bug
Rows("2:70").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

'Início remoção duplicados
Range("C1").Select
    For contador = 1 To 10000
        'Saída do looping
        If ActiveCell.Offset(1, 0).Value = "" Then
            Selection.End(xlDown).Select
            If ActiveCell.Value = "" Then
                Exit For
            End If
        Else
            ActiveCell.Offset(1, 0).Select
        End If

        'Análise pedido duplicado
        ActiveCell.Offset(0, 31).Select
            If ActiveCell.Value = "" Then
                ActiveCell.Offset(0, -31).Select
            Else
                first_row = ActiveCell.Row
                last_row = ActiveCell.Row
                
                'Verificação duplicada +1 linha
                If ActiveCell.Offset(-1, 0) <> "" Then
                    Selection.End(xlUp).Select
                        first_row = ActiveCell.Row
                        range_row = last_row - first_row + 1
                    Range("AH" & first_row & ":" & "AJ" & last_row).Select
                        Selection.Cut
                    ActiveCell.Offset(-range_row, 0).Select
                        ActiveSheet.Paste

                    'Variáveis
                    duplicate_row = first_row - range_row
                    range_duplicate_row = duplicate_row
                    k = 0

                    Do While duplicate_row < first_row:
                        If Range("AI" & duplicate_row).Value = Range("H" & duplicate_row).Value And Range("AJ" & duplicate_row).Value = Range("L" & duplicate_row).Value Then
                                k = k + 1
                            Range("AK" & duplicate_row).Value = Range("AH" & duplicate_row).Value + Range("K" & duplicate_row).Value
                            'processo de substituicao dos valores
                            If k = range_row Then
                                Range("AK" & duplicate_row & ":" & "AK" & range_duplicate_row).Select
                                    Selection.Copy
                                Range("K" & range_duplicate_row).Select
                                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                                    :=False, Transpose:=False
                                ActiveCell.Offset(0, -8).Select
                                    Selection.End(xlDown).Select
                                ActiveCell.Offset(1, 0).Select
                                    Selection.End(xlToLeft).Select
                                        Selection.End(xlToLeft).Select
                                            Selection.End(xlToLeft).Select
                                range_duplicate_row_delete = first_row & ":" & last_row
                                    Rows(range_duplicate_row_delete).Select
                                        Range(Selection, Selection.End(xlToRight)).EntireRow.Delete
                                    ActiveCell.Offset(0, 1).Select
                            End If
                        End If
                        duplicate_row = duplicate_row + 1
                    Loop
                    Range("AH" & first_row - range_row & ":" & "AK" & duplicate_row - 1).Select
                            Selection.Clear
                            
                'Verificação duplicada 1 linha
                Else
                    Range("AH" & last_row & ":" & "AJ" & last_row).Select
                        Selection.Cut
                    ActiveCell.Offset(-1, 0).Select
                        ActiveSheet.Paste
                    ActiveCell.Offset(2, -31).Select
                        duplicate_row = last_row - 1
                    'processo de substituicao dos valores
                    If Range("AI" & duplicate_row).Value = Range("H" & duplicate_row).Value And Range("AJ" & duplicate_row).Value = Range("L" & duplicate_row).Value Then
                            k = k + 1
                        Range("AK" & duplicate_row).Value = Range("AH" & duplicate_row).Value + Range("K" & duplicate_row).Value
                            Range("AK" & duplicate_row & ":" & "AK" & duplicate_row).Select
                                Selection.Copy
                            Range("K" & duplicate_row).Select
                                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                                :=False, Transpose:=False
                        ActiveCell.Offset(1, 0).Select
                            ActiveCell.Offset(0, -8).Select
                                Selection.Copy
                            ActiveCell.Offset(-1, 0).Select
                                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                                :=False, Transpose:=False
                            ActiveCell.Offset(1, 0).Select
                                Selection.End(xlToLeft).Select
                                    Range(Selection, Selection.End(xlToRight)).EntireRow.Delete
                    End If
                    Range("AH" & duplicate_row & ":" & "AK" & duplicate_row).Select
                        Selection.Clear
                End If
            Range("C" & first_row - 1).Select
            End If
    Next

'Deletar linhas
Range("A2").Select
    delete_first_line = ActiveCell.Row
Selection.End(xlDown).Select
ActiveCell.Offset(-1, 0).Select
    delete_last_line = ActiveCell.Row
Rows(delete_first_line & ":" & delete_last_line).Select
    Selection.Delete Shift:=xlUp
Columns("AF:AG").Select
    Selection.Delete Shift:=xlToLeft
    
'Contagem de linhas
Range("A2").Select
last_line = Cells(Cells.Rows.Count, 1).End(xlUp).Row

Range("AF1").Select
    ActiveCell.FormulaR1C1 = "21.ConvQtd"
Range("AG1").Select
    ActiveCell.FormulaR1C1 = "22.ConvUnid"

'Preenche a coluna AC, ConvQtd
Range("AF2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-19]<>""METROS"",RC[-20],ROUNDUP(RC[-20]/RC[-11],0))"
        Range("AF2").Select
            Selection.AutoFill Destination:=Range("AF2:AF" & last_line)
        Range("AF2:AF" & last_line).Select

'Preenche a coluna AB, ConvUnid
Range("AG2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-20]<>""METROS"",RC[-20],RC[-2])"
        Range("AG2").Select
            Selection.AutoFill Destination:=Range("AG2:AG" & last_line)
        Range("AG2:AG" & last_line).Select

'Preenche a coluna AC, Peso Total
Range("AH1").Value = "23.Peso total"
    Range("AH2").Select
        ActiveCell.FormulaR1C1 = "=RC[-2]*RC[-8]"
    Range("AH2").Select
        Selection.AutoFill Destination:=Range("AH2:AH" & last_line)
    Range("AH2:AH" & last_line).Select
        Selection.NumberFormat = "0.00"

Range("A1:AH1").Select
    Selection.AutoFilter

Application.Calculation = xlCalculationAutomatic

'Copia e cola valores
Range("A1").Select
    Selection.CurrentRegion.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Dim range_row_final As Integer
Range("A1").Select
    range_row_final = Cells(Cells.Rows.Count, 1).End(xlUp).Row
    linhas_excluidas = inicial_range - range_row_final
        MsgBox ("Número de linhas excluidas: " & linhas_excluidas)

'Fecha planilha BD Tecserp
Windows("BD TecSerp.xlsm").Activate
    ActiveWindow.Close
    
'Adiciona Ano e Mês
Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Range("B1").Value = "Ano"
    Range("B2").Select
        ActiveCell.FormulaR1C1 = "=YEAR(RC[-1])"
        Selection.AutoFill Destination:=Range("B2:B" & last_line)
    Columns("B:B").Select
        Selection.NumberFormat = "General"

    Columns("C:C").Select
    Selection.TextToColumns Destination:=Range("C1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C2").Select
        ActiveCell.FormulaR1C1 = "=TEXT(MONTH(RC[-2])*29,""mmm"")"
    Range("C2").Select
        Selection.AutoFill Destination:=Range("C2:C" & last_line)

    Range("C1").Value = "Mes"


'Procv referente a coluna AB, 19.?
Range("AD1").Value = "19.Conv. Metros"
Range("AD2").Select
    ActiveCell.FormulaR1C1 = "=RC[-7]*RC[4]"
        With ActiveSheet
            LastRow = Range("A" & .Rows.Count).End(xlUp).Row
            Set RngAutopreencher = Range("AD2:AD" & LastRow)
        End With
     Selection.AutoFill Destination:=RngAutopreencher

Range("A1").Select

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub




