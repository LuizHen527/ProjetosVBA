Attribute VB_Name = "M24_ANALISE_VENDAS_MOLDURAS"
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"

'Application.ScreenUpdating = False
'Application.DisplayAlerts = False


'Replica os pedidos
Sheets("Macro").Select
    If Range("F1").Value <> "Pedido_2" Then
        Call M20_REPLICA_PEDIDOS.Replicar_Pedidos
    End If

'---------------------------------------------
'Converte os pedidos de texto para número
Sheets("Macro").Select
    Range("F2").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.TextToColumns Destination:=Range("F2"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True

'Adiciona a aba
Sheets.Add After:=Sheets(Sheets.Count)
        ActiveSheet.Name = "Análise Pedidos - Molduras"
        
'Títulos
Sheets("Análise Pedidos - Molduras").Select
    Range("A1").FormulaR1C1 = "ANÁLISE DE PEDIDOS - SETOR MOLDURAS"
    Range("A1").Select
        Selection.Font.Bold = True
   
    Range("A2").FormulaR1C1 = "PEDIDO"
    Range("B2").FormulaR1C1 = "CLIENTE"
    Range("C2").FormulaR1C1 = "MOLDURAS"
    Range("D2").FormulaR1C1 = "DATA FATURAMENTO"
   
    Range("A2:D2").Select
        Selection.Font.Bold = True
    
'Fórmulas
    Range("B3").FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],Macro!C6:C36,2,0),"""")"
    Range("C3").FormulaR1C1 = "=IF(SUMIFS(Macro!C35,Macro!C6,RC1,Macro!C17,R2C)=0,"""",SUMIFS(Macro!C35,Macro!C6,RC1,Macro!C17,R2C))"
    Range("D3").FormulaR1C1 = "=IFERROR(INDEX(Macro!C1:C6,MATCH(RC[-3],Macro!C6,0),1),"""")"
    Range("B3").Select
    Range(Selection, Selection.End(xlToRight)).Select
        Selection.AutoFill Destination:=Range("B3:D200"), Type:=xlFillDefault
    Range("B3:D200").Select

'Corrige colunas
Columns("A:D").Select
    Selection.ColumnWidth = 12
Columns("D:D").Select
    Selection.NumberFormat = "m/d/yyyy"

End Sub

