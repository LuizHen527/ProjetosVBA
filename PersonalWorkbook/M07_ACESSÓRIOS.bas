Attribute VB_Name = "M07_ACESSÓRIOS"
Option Base 1

Sub Menu()

Application.DisplayAlerts = False


On Error GoTo panda

    If Sheets("Acessórios Roaplas").Activate = True Then
        Sheets("Acessórios Roaplas").Delete
    End If

panda:
Call Nome_Acessórios
Call Fórmulas_Acessórios
Call Dobradiças
Call PuxadorBC

End Sub


Sub Nome_Acessórios()

'Application.ScreenUpdating = False
Application.DisplayAlerts = False

Sheets.Add After:=Sheets(Sheets.Count)
        ActiveSheet.Name = "Acessórios Roaplas"

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
    ActiveCell.FormulaR1C1 = "ROSE"
Range("H2").Select
    ActiveCell.FormulaR1C1 = "DOURADO"
Range("I2").Select
    ActiveCell.FormulaR1C1 = "INOX"

'Kits e Perfis
kits = Array("KF2P", "KF3P", "KF4P", "KC4P", "RETO KF2P", "RETO KF3P", "RETO KF4P", "RETO KC4P", _
            "BF1", "BF2", "BF3", "BC1", "RETO BF1", "RETO BF2", "RETO BF3", "RETO BC1", _
            "2F", "4F", "PIA", "MULTIUSO")

Range("A3").Select
For i = 1 To UBound(kits)
    ActiveCell.Value = kits(i)
        ActiveCell.Offset(1, 0).Select
Next
    
Application.ScreenUpdating = True
Application.DisplayAlerts = True
    
End Sub

Sub Fórmulas_Acessórios()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Range("B3").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIFS(Macro!R2C34:R1048575C34,Macro!R2C18:R1048575C18,RC1,Macro!R2C21:R1048575C21,R2C)"
Range("B3").Select
    Selection.AutoFill Destination:=Range("B3:B22"), Type:=xlFillDefault
Range("B3:B22").Select
    Selection.AutoFill Destination:=Range("B3:I22"), Type:=xlFillDefault

  
Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub



Sub Dobradiças()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

'Dobradiças Open e Clean
    Range("A24").FormulaR1C1 = "KIT"
    Range("A25").FormulaR1C1 = "OPEN SEM TRANSPASSE"
    Range("A26").FormulaR1C1 = "OPEN SEM TRANSPASSE"
    Range("A27").FormulaR1C1 = "OPEN COM TRANSPASSE"
    Range("A28").FormulaR1C1 = "OPEN COM TRANSPASSE"
    Range("A29").FormulaR1C1 = "CLEAN"
    Range("A30").FormulaR1C1 = "CLEAN"
    Range("B24").FormulaR1C1 = "MATERIAL"
    Range("B25").FormulaR1C1 = "ZAMACK"
    Range("B26").FormulaR1C1 = "LATAO"
    Range("B27").FormulaR1C1 = "ZAMACK"
    Range("B28").FormulaR1C1 = "LATAO"
    Range("B29").FormulaR1C1 = "ZAMACK"
    Range("B30").FormulaR1C1 = "LATAO"
    Range("C24").FormulaR1C1 = "FOSCO"
    Range("D24").FormulaR1C1 = "BRANCO"
    Range("E24").FormulaR1C1 = "BRILHO"
    Range("F24").FormulaR1C1 = "PRETO"
    Range("G24").FormulaR1C1 = "BRONZE"
    Range("H24").FormulaR1C1 = "DOURADO"
    
    Range("C25").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIFS(Macro!R2C34:R1048575C34,Macro!R2C18:R1048575C18,RC1,Macro!R2C21:R1048575C21,R24C,Macro!R2C19:R1048575C19,RC2)"
    Range("C25").Select
    Selection.AutoFill Destination:=Range("C25:H25"), Type:=xlFillDefault
    Range("C25:H25").Select
    Selection.AutoFill Destination:=Range("C25:H30"), Type:=xlFillDefault
    Range("C25:H30").Select

'Soma total acessórios
    Range("A33").FormulaR1C1 = "OPEN"
    Range("A34").FormulaR1C1 = "OPEN"
    Range("A35").FormulaR1C1 = "CLEAN"
    Range("A36").FormulaR1C1 = "CLEAN"
    Range("B33").FormulaR1C1 = "ZAMACK"
    Range("B34").FormulaR1C1 = "LATAO"
    Range("B35").FormulaR1C1 = "ZAMACK"
    Range("B36").FormulaR1C1 = "LATAO"
    Range("C33").FormulaR1C1 = "=SUM(R[-8]C,R[-6]C)"
    Range("C34").FormulaR1C1 = "=SUM(R[-8]C,R[-6]C)"
    Range("C35").FormulaR1C1 = "=R[-6]C"
    Range("C36").FormulaR1C1 = "=R[-6]C"
    Range("C33:C36").Select
    Selection.AutoFill Destination:=Range("C33:H36"), Type:=xlFillDefault
    Range("C33:H36").Select

End Sub

Sub PuxadorBC()

    Range("A39").FormulaR1C1 = "KIT"
    Range("A40").FormulaR1C1 = "BARRA CHATA H"
    Range("A41").FormulaR1C1 = "BARRA CHATA H"
    Range("B39").FormulaR1C1 = "MATERIAL"
    Range("B40").FormulaR1C1 = "0,30 X 0,20"
    Range("B41").FormulaR1C1 = "0,40 X 0,30"
    Range("C39").FormulaR1C1 = "FOSCO"
    Range("D39").FormulaR1C1 = "BRANCO"
    Range("E39").FormulaR1C1 = "BRILHO"
    Range("F39").FormulaR1C1 = "PRETO"
    Range("G39").FormulaR1C1 = "BRONZE"
    Range("H39").FormulaR1C1 = "DOURADO"
    Range("I39").FormulaR1C1 = "DOURADO FOSCO"
    Range("J39").FormulaR1C1 = "ROSE"
    Range("K39").FormulaR1C1 = "INOX"
    
        ActiveCell.FormulaR1C1 = _
        "=SUMIFS(Macro!R2C34:R1048575C34,Macro!R2C18:R1048575C18,RC1,Macro!R2C21:R1048575C21,R39C,Macro!R2C23:R1048575C23,RC2)"
    Range("C40").Select
    Selection.AutoFill Destination:=Range("C40:K40"), Type:=xlFillDefault
    Range("C40:K40").Select
    Selection.AutoFill Destination:=Range("C40:K41"), Type:=xlFillDefault
    Range("C40:K41").Select
    
    
End Sub
