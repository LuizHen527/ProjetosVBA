Attribute VB_Name = "M02_CONV_PERFIS"
Option Base 1

Sub Menu()

Application.DisplayAlerts = False


On Error GoTo panda

    If Sheets("Conversão Perfis").Activate = True Then
        Sheets("Conversão Perfis").Delete
    End If

panda:
Call Nomes
Call Fórmulas
Call Perfis

End Sub


Sub Nomes()

'Application.ScreenUpdating = False
Application.DisplayAlerts = False

Sheets.Add After:=Sheets(Sheets.Count)
        ActiveSheet.Name = "Conversão Perfis"

'Títulos
Range("A1").Select
    ActiveCell.FormulaR1C1 = "QUANTIDADE VENDIDO"
Range("A2").Select
    ActiveCell.FormulaR1C1 = "KITS"
Range("B2").Select
    ActiveCell.FormulaR1C1 = "FOSCO"
Range("C2").Select
    ActiveCell.FormulaR1C1 = "BRANCO"
Range("D2").Select
    ActiveCell.FormulaR1C1 = "BRILHO"
Range("E2").Select
    ActiveCell.FormulaR1C1 = "PRETO"
Range("F2").Select
    ActiveCell.FormulaR1C1 = "BRONZE"
Range("G2").Select
    ActiveCell.FormulaR1C1 = "DOURADO"
Range("H2").Select
    ActiveCell.FormulaR1C1 = "ROSE"
Range("I2").Select
    ActiveCell.FormulaR1C1 = "INOX"
Range("J2").Select
    ActiveCell.FormulaR1C1 = "POLIDO"
Range("K2").Select
    ActiveCell.FormulaR1C1 = "DOURADO FOSCO"
    

'Kits e Perfis
kits = Array("BF1", "BF2", "BF3", "BC1", "RETO BF1", "RETO BF2", "RETO BF3", "RETO BC1", "KF2P", "KF3P", "KF4P", "KC4P", "RETO KF2P", _
            "RETO KF3P", "RETO KF4P", "RETO KC4P", "OPEN COM TRANSPASSE", "OPEN SEM TRANSPASSE", "2F", "4F", "KIT ENG", "M2000", "PIA", "PACIFIC F1", "PACIFIC F2", "PACIFIC F3", _
            "PACIFIC C1", "FLEX", "ATLANTIC F1", "DUALDOOR", "BOX TRILHO SUPERIOR - TEC072", "BOX TRILHO SUPERIOR RETO - TEC073", "BOX U08", _
            "BOX U10", "BXB083 - CLICK", "BXB095 - CADEIRINHA BF", "BXB141 - CADEIRINHA BC", "BXB219 - CAPA", "BXB238 - U10", "BXB239 - U08", _
            "BXB241 - TRILHO SUPERIOR ABAULADO", "BXB247 - GUIA INFERIOR ABAULADO", "BXB248 - GUIA INFERIOR RETO", "BXB249 - TRILHO SUPERIOR RETO", _
            "GUIA INFERIOR OPEN", "ENG CAPA - AL101", "ENG GUIA INFERIOR", "ENG TRILHO SUPERIOR - AL100", "ENG U08", "ENG U10", "ENG VEDAPO", "VITRINE H", "VITRINE U", _
            "VITRINE W", "KMA01 - SUPERIOR", "KMA02 - INFERIOR", "KMA03 - LATERAL", "KMA04 - LATERAL", "KP2308 - INFERIOR", "KP3314 - LEITO", "KP4119 - SUPERIOR", _
            "KP4822 - LATERAL", "M2000 CAPA - BX61", "M2000 INFERIOR DUPLO - BX500", "M2000 SUPERIOR - BX100", "STL TRILHO SUPERIOR - AL68", _
            "STL U15 CAVALAO", "LMK1", "LMK2", "LMK3", "R1212 - SEGURANÇA", "R2212 - INFERIOR", "R2414 - SUPERIOR", "DD01 - SUPERIOR", "DD03 - LATERAL", "KF1611", "R2106", _
            "R3125", "R3523", "R3830", "R5931", "R6126", "LIGHT F1", "LIGHT F2", "LIGHT F3", "LIGHT C1", "MULTIUSO", "CLASSIC F1", "CLASSIC F2", "CLASSIC F3", "CLASSIC C1", "KS01", "KS01 + F", "KS01 CANTO", "ABRIR", "MULTI 2F", "MULTI 3F")
            

Range("A3").Select
For i = 1 To UBound(kits)
    ActiveCell.Value = kits(i)
        ActiveCell.Offset(1, 0).Select
Next
    
Application.ScreenUpdating = True
Application.DisplayAlerts = True
    
End Sub
Sub Fórmulas()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

last_line = Cells(Cells.Rows.Count, 1).End(xlUp).Row

Range("B3").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIFS(Macro!R2C34:R1048575C34,Macro!R2C18:R1048575C18,RC1,Macro!R2C21:R1048575C21,R2C)"
Range("B3").Select
    Selection.AutoFill Destination:=Range("B3:B" & last_line), Type:=xlFillDefault
Range("B3:B" & last_line).Select
    Selection.AutoFill Destination:=Range("B3:K" & last_line), Type:=xlFillDefault

  
Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Sub Perfis()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

nome_da_planilha = ActiveWorkbook.Name
last_line = Cells(Cells.Rows.Count, 1).End(xlUp).Row

    Range("B3:K" & last_line).Select
        Selection.Copy
    
Workbooks.Open fileName:="\\121.137.1.5\alumitec9\PRODUÇÃO\2025 Extrusão e Produção\07_MOLDUCOLOR\4. Calculadora Perfis_V3.xlsx"

    Windows("4. Calculadora Perfis_V3.xlsx").Activate
        Range("C4").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
        Range("O2:Z80").Select
            Selection.Copy
        Windows(nome_da_planilha).Activate
            Range("M1").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Columns("M:M").Select
                Selection.ColumnWidth = 50
        Range("M1").Select
            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
     Workbooks("4. Calculadora Perfis_V3.xlsx").Close SaveChanges:=False
     
End Sub

