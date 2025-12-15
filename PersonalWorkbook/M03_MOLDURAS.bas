Attribute VB_Name = "M03_MOLDURAS"
Sub molduras() 'Geração da aba referente as familias

Sheets("Macro").Select
    plan_name = "3.Molduras"

ActiveWorkbook.Sheets.Add 'Adiciona uma nova planilha, que se torna ativa
    ActiveSheet.Name = plan_name

'Nomes Molduras
Range("A1").FormulaR1C1 = "MOLDURAS"
Range("A2").FormulaR1C1 = "AF01"
Range("A3").FormulaR1C1 = "AF01"
Range("A4").FormulaR1C1 = "AF06"
Range("A5").FormulaR1C1 = "AF07"
Range("A6").FormulaR1C1 = "AF12"
Range("A7").FormulaR1C1 = "AF13"
Range("A8").FormulaR1C1 = "AF13"
Range("A9").FormulaR1C1 = "AF14"
Range("A10").FormulaR1C1 = "AF14"
Range("A11").FormulaR1C1 = "AF15"
Range("A12").FormulaR1C1 = "AF15"
Range("A13").FormulaR1C1 = "AF16"
Range("A14").FormulaR1C1 = "AF16"
Range("A15").FormulaR1C1 = "AF18"
Range("A16").FormulaR1C1 = "AF18"
Range("A17").FormulaR1C1 = "AF20"
Range("A18").FormulaR1C1 = "AF21"
Range("A19").FormulaR1C1 = "AF22"
Range("A20").FormulaR1C1 = "AF30"
Range("A21").FormulaR1C1 = "AF40"
Range("A22").FormulaR1C1 = "OVAL"
Range("A23").FormulaR1C1 = "OVAL"
Range("A24").FormulaR1C1 = "OVAL"
Range("A25").FormulaR1C1 = "OVAL"
Range("A26").FormulaR1C1 = "OVAL"
Range("A27").FormulaR1C1 = "OVAL"
Range("A28").FormulaR1C1 = "TOTAL"

'Medidas Molduras
Range("B1").FormulaR1C1 = "MEDIDAS"
Range("B2").FormulaR1C1 = "2.2"
Range("B3").FormulaR1C1 = "2.5"
Range("B4").FormulaR1C1 = "2.5"
Range("B5").FormulaR1C1 = "2.5"
Range("B6").FormulaR1C1 = "2.5"
Range("B7").FormulaR1C1 = "2.2"
Range("B8").FormulaR1C1 = "2.5"
Range("B9").FormulaR1C1 = "2.2"
Range("B10").FormulaR1C1 = "2.5"
Range("B11").FormulaR1C1 = "2.2"
Range("B12").FormulaR1C1 = "2.5"
Range("B13").FormulaR1C1 = "2.2"
Range("B14").FormulaR1C1 = "2.5"
Range("B15").FormulaR1C1 = "2.2"
Range("B16").FormulaR1C1 = "2.5"
Range("B17").FormulaR1C1 = "2.5"
Range("B18").FormulaR1C1 = "2.5"
Range("B19").FormulaR1C1 = "2.5"
Range("B20").FormulaR1C1 = "2.5"
Range("B21").FormulaR1C1 = "2.5"
Range("B22").FormulaR1C1 = "1.6"
Range("B23").FormulaR1C1 = "1.8"
Range("B24").FormulaR1C1 = "2.1"
Range("B25").FormulaR1C1 = "3.4"
Range("B26").FormulaR1C1 = "4.3"
Range("B27").FormulaR1C1 = "3.5"

'Acabamentos
Range("C1").FormulaR1C1 = "AZUL"
Range("D1").FormulaR1C1 = "BRANCO"
Range("E1").FormulaR1C1 = "BRONZE B."
Range("F1").FormulaR1C1 = "BRONZE F."
Range("G1").FormulaR1C1 = "DOURADO B."
Range("H1").FormulaR1C1 = "DOURADO F."
Range("I1").FormulaR1C1 = "FUME B."
Range("J1").FormulaR1C1 = "FUME F."
Range("K1").FormulaR1C1 = "INCOLOR B."
Range("L1").FormulaR1C1 = "INCOLOR F."
Range("M1").FormulaR1C1 = "PRETO B."
Range("N1").FormulaR1C1 = "PRETO F."
Range("O1").FormulaR1C1 = "VERDE"
Range("P1").FormulaR1C1 = "VINHO"
Range("Q1").FormulaR1C1 = "TOTAL"

'Quantidade de Molduras em aberto
Range("C2").Select
    ActiveCell.FormulaR1C1 = _
    "=SUMIFS(Macro!C34,Macro!C16,""MOLDURAS"",Macro!C17,RC1,Macro!C23,RC2,Macro!C21,R1C)"
Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:P2"), Type:=xlFillDefault
Range("C2:P2").Select
    Selection.AutoFill Destination:=Range("C2:P27"), Type:=xlFillDefault
    Range("C2:P27").Select

'Totais por acabamento
Range("C28").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-26]C:R[-1]C)"
        Range("C28").Select
    Selection.AutoFill Destination:=Range("C28:P28"), Type:=xlFillDefault
        Range("C28:P28").Select

'Totais por tipo de moldura
Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-14]:RC[-1])"
Range("Q2").Select
    Selection.AutoFill Destination:=Range("Q2:Q27"), Type:=xlFillDefault
        Range("Q2:Q27").Select

'Totais das molduras
Range("Q28").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-26]C:R[-1]C)"
Range("C28:P28").Select
    Range("P28").Activate



'Resumo Acabamentos
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "FOSCO"
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "BRILHO"
    Range("T1").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    Range("R2").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIFS(Macro!C34,Macro!C16,""MOLDURAS"",Macro!C17,RC1,Macro!C23,RC2,Macro!C18,""COLORIDO FOSCO"")+SUMIFS(Macro!C34,Macro!C16,""MOLDURAS"",Macro!C17,RC1,Macro!C23,RC2,Macro!C18,""INCOLOR FOSCO"")+SUMIFS(Macro!C34,Macro!C16,""MOLDURAS"",Macro!C17,RC1,Macro!C23,RC2,Macro!C18,""MACSO FOSCO"")"
    Range("R2").Select
        Selection.AutoFill Destination:=Range("R2:R27"), Type:=xlFillDefault
    Range("R2:R27").Select
    Range("S2").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIFS(Macro!C34,Macro!C16,""MOLDURAS"",Macro!C17,RC1,Macro!C23,RC2,Macro!C18,""COLORIDO BRILHO"")+SUMIFS(Macro!C34,Macro!C16,""MOLDURAS"",Macro!C17,RC1,Macro!C23,RC2,Macro!C18,""INCOLOR BRILHO"")+SUMIFS(Macro!C34,Macro!C16,""MOLDURAS"",Macro!C17,RC1,Macro!C23,RC2,Macro!C18,""MACSO BRILHO"")"
    Range("S2").Select
        Selection.AutoFill Destination:=Range("S2:S27"), Type:=xlFillDefault
    Range("S2:S27").Select
    Range("T2").Select
        ActiveCell.FormulaR1C1 = "=SUM(RC[-2]:RC[-1])"
    Range("T2").Select
        Selection.AutoFill Destination:=Range("T2:T27"), Type:=xlFillDefault
    Range("T2:T27").Select
    Range("R28").Select
        ActiveCell.FormulaR1C1 = "=SUM(R[-26]C:R[-1]C)"
    Range("R28").Select
        Selection.AutoFill Destination:=Range("R28:T28"), Type:=xlFillDefault
    Range("R28:T28").Select

'Formatação Condicional
Sheets("3.Molduras").Select
    Range("Q2:Q27").Select
    Selection.FormatConditions.AddColorScale ColorScaleType:=3
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueLowestValue
    With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = 8109667
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValuePercentile
    Selection.FormatConditions(1).ColorScaleCriteria(2).Value = 50
    With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .Color = 8711167
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(3).Type = _
        xlConditionValueHighestValue
    With Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor
        .Color = 7039480
        .TintAndShade = 0
    End With
End Sub

    
Sub Molduras_2()
Application.DisplayAlerts = False

'Transforma o nome da planilha base em variavel
plan_base = ActiveWorkbook.Name

'Variáveis de data
dia = Format(Date, "dd")
mes = Format(Date, "mm")
ano = Format(Date, "yy")

'Acessa a planilha de molduras em aberto e cria as abas
Workbooks.Open fileName:="\\121.137.1.5\alumitec9\PRODUÇÃO\2024 Extrusão e Produção\07_MOLDUCOLOR\1. Pedidos Molduras em Aberto.xlsx"
    plan_molduras = ActiveWorkbook.Name
    On Error GoTo panda
        If Sheets(ano & "_" & mes & "_" & dia).Activate = True Then
            Sheets(ano & "_" & mes & "_" & dia).Delete
        End If
            
panda:
num_abas = Sheets.Count
Sheets("BASE").Select
Sheets("BASE").Copy After:=Sheets(num_abas)
    ActiveSheet.Name = ano & "_" & mes & "_" & dia
    Range("A1").Value = "PEDIDOS DE MOLDURAS EM ABERTO - " & dia & "/" & mes & "/" & ano
    
'Acessa planilha base para copiar os dados
Windows(plan_base).Activate
    Sheets("3.Molduras").Select
        Range("C2:P27").Select
        Selection.Copy
        
'Atualiza os dados da planilha de molduras em aberto
Windows(plan_molduras).Activate
    Sheets(ano & "_" & mes & "_" & dia).Select
        Range("C3").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
  ActiveWorkbook.Save

Windows(plan_base).Activate

End Sub

