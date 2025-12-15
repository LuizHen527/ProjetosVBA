Attribute VB_Name = "M09_LUCRATIVIDADE"
Sub CALL_LUCRATIVIDADE()

Call L01_LUCRO_MOLDURAS
Call L02_LUCRO_KITS
Call L03_LUCRO_ROAPLAS

End Sub

Sub L01_LUCRO_MOLDURAS()

'As mudanças do programa não aparecem passo-a-passo na tela
Application.ScreenUpdating = False
Application.DisplayAlerts = False

modelos_molduras = Array("MOLDURAS", "AF01", "AF01", "AF13", "AF13", "AF14", "AF14", "AF15", "AF15", "AF16", "AF16", "AF18", "AF18", "OVAL", "OVAL", "OVAL")
medidas_molduras = Array("MEDIDAS", 2.2, 2.5, 2.2, 2.5, 2.2, 2.5, 2.2, 2.5, 2.2, 2.5, 2.2, 2.5, 1.6, 1.8, 2.1)
acabamentos_molduras = Array("AZUL", "BRANCO", "BRONZE B.", "BRONZE F.", "DOURADO B.", "DOURADO F.", "FUME B.", "FUME F.", "INCOLOR B.", "INCOLOR F.", "PRETO B.", "PRETO F.", "VERDE", "VINHO")

ActiveWorkbook.Sheets.Add 'Adiciona uma nova planilha, que se torna ativa
    ActiveSheet.Name = "Lucratividade_Molduras"

Range("A1").FormulaR1C1 = "ANÁLISE LUCRATIVIDADE DE MOLDURAS"
    
'QUANTIDADE [PEÇAS] DE MOLDURAS FATURADAS
Range("A2").FormulaR1C1 = "QUANTIDADE [PEÇAS] DE MOLDURAS FATURADAS"
    Range("A3").Select
        For i = 0 To 15
            ActiveCell.Value = modelos_molduras(i)
                ActiveCell.Offset(1, 0).Select
        Next
    Range("B3").Select
        For i = 0 To 15
            ActiveCell.Value = medidas_molduras(i)
                ActiveCell.Offset(1, 0).Select
        Next
    Range("C3").Select
        For i = 0 To 13
            ActiveCell.Value = acabamentos_molduras(i)
                ActiveCell.Offset(0, 1).Select
        Next

    Range("C4").FormulaR1C1 = "=ROUND((SUMIFS(Macro!C14,Macro!C17,RC1,Macro!C23,RC2,Macro!C21,R3C))/RC2,0)"
        Range("C4").Select
            Selection.AutoFill Destination:=Range("C4:P4"), Type:=xlFillDefault
        Range("C4:P4").Select
            Selection.AutoFill Destination:=Range("C4:P15"), Type:=xlFillDefault
        Range("C4:P15").Select
    Range("C16").FormulaR1C1 = "=ROUND((SUMIFS(Macro!C14,Macro!C17,RC1,Macro!C23,RC2,Macro!C21,R3C)),0)"
        Range("C16").Select
            Selection.AutoFill Destination:=Range("C16:C18"), Type:=xlFillDefault
        Range("C16:C18").Select
            Selection.AutoFill Destination:=Range("C16:P18"), Type:=xlFillDefault
        Range("C16:P18").Select

'QUANTIDADE. [KG] DE MOLDURAS FATURADAS
Range("A20").FormulaR1C1 = "QUANTIDADE [KG] DE MOLDURAS FATURADAS"
    Range("A21").Select
        For i = 0 To 15
            ActiveCell.Value = modelos_molduras(i)
                ActiveCell.Offset(1, 0).Select
        Next
    Range("B21").Select
        For i = 0 To 15
            ActiveCell.Value = medidas_molduras(i)
                ActiveCell.Offset(1, 0).Select
        Next
    Range("C21").Select
        For i = 0 To 13
            ActiveCell.Value = acabamentos_molduras(i)
                ActiveCell.Offset(0, 1).Select
        Next
    Range("C22").FormulaR1C1 = "=ROUND((SUMIFS(Macro!C36,Macro!C17,RC1,Macro!C23,RC2,Macro!C21,R21C))/RC2,1)"
    Range("C22").Select
        Selection.AutoFill Destination:=Range("C22:P22"), Type:=xlFillDefault
    Range("C22:P22").Select
        Selection.AutoFill Destination:=Range("C22:P36"), Type:=xlFillDefault

'VALOR [R$] DE MOLDURAS FATURADAS
Range("A38").FormulaR1C1 = "VALOR [R$] DE MOLDURAS FATURADAS"
    Range("A39").Select
        For i = 0 To 15
            ActiveCell.Value = modelos_molduras(i)
                ActiveCell.Offset(1, 0).Select
        Next
    Range("B39").Select
        For i = 0 To 15
            ActiveCell.Value = medidas_molduras(i)
                ActiveCell.Offset(1, 0).Select
        Next
    Range("C39").Select
        For i = 0 To 13
            ActiveCell.Value = acabamentos_molduras(i)
                ActiveCell.Offset(0, 1).Select
        Next
        
    Range("C40").FormulaR1C1 = "=ROUND((SUMIFS(Macro!C13,Macro!C17,RC1,Macro!C23,RC2,Macro!C21,R39C)),1)"
    Range("C40").Select
        Selection.AutoFill Destination:=Range("C40:P40"), Type:=xlFillDefault
    Range("C40:P40").Select
        Selection.AutoFill Destination:=Range("C40:P54"), Type:=xlFillDefault

End Sub

Sub L02_LUCRO_KITS()

familia_kits = Array("FAMILIA", "BLINDEX", "BLINDEX", "BLINDEX", "BLINDEX", "BLINDEX", "BLINDEX", "BLINDEX", "BLINDEX", "BOX", "BOX", "BOX", "BOX", "BOX", "BOX", "BOX", "BOX")
modelos_kits = Array("KITS", "KF2P", "KF3P", "KF4P", "KC4P", "RETO KF2P", "RETO KF3P", "RETO KF4P", "RETO KC4P", "BF1", "BF2", "BF3", "BC1", "RETO BF1", "RETO BF2", "RETO BF3", "RETO BC1")
acabamentos_kits = Array("FOSCO", "BRANCO", "BRILHO", "PRETO", "BRONZE", "DOURADO", "ROSE", "INOX")

'As mudanças do programa não aparecem passo-a-passo na tela
Application.ScreenUpdating = False
Application.DisplayAlerts = False

ActiveWorkbook.Sheets.Add
    ActiveSheet.Name = "Lucratividade_Kits"

Range("A1").FormulaR1C1 = "ANÁLISE LUCRATIVIDADE KITS"

'QUANTIDADE [PEÇAS] DE KITS FATURADOS
Range("A2").FormulaR1C1 = "QUANTIDADE [PEÇAS] DE KITS FATURADOS"
Range("A3").Select
    For i = 0 To 16
        ActiveCell.Value = familia_kits(i)
            ActiveCell.Offset(1, 0).Select
    Next
Range("B3").Select
    For i = 0 To 16
        ActiveCell.Value = modelos_kits(i)
            ActiveCell.Offset(1, 0).Select
    Next
Range("C3").Select
    For i = 0 To 7
        ActiveCell.Value = acabamentos_kits(i)
            ActiveCell.Offset(0, 1).Select
    Next
Range("C4").FormulaR1C1 = "=ROUND((SUMIFS(Macro!C14,Macro!C17,RC1,Macro!C18,RC2,Macro!C21,R3C,Macro!C16,""KITS"")),0)"
Range("C4").Select
    Selection.AutoFill Destination:=Range("C4:J4"), Type:=xlFillDefault
Range("C4:J4").Select
    Selection.AutoFill Destination:=Range("C4:J19"), Type:=xlFillDefault

'VALOR [R$] DE KITS FATURADOS
Range("A21").FormulaR1C1 = "VALOR [R$] DE KITS FATURADOS"
Range("A22").Select
    For i = 0 To 16
        ActiveCell.Value = familia_kits(i)
            ActiveCell.Offset(1, 0).Select
    Next
Range("B22").Select
    For i = 0 To 16
        ActiveCell.Value = modelos_kits(i)
            ActiveCell.Offset(1, 0).Select
    Next
Range("C22").Select
    For i = 0 To 7
        ActiveCell.Value = acabamentos_kits(i)
            ActiveCell.Offset(0, 1).Select
    Next
Range("C23").FormulaR1C1 = "=ROUND((SUMIFS(Macro!C13,Macro!C17,RC1,Macro!C18,RC2,Macro!C21,R22C,Macro!C16,""KITS"")),1)"
Range("C23").Select
    Selection.AutoFill Destination:=Range("C23:J23"), Type:=xlFillDefault
Range("C23:J23").Select
    Selection.AutoFill Destination:=Range("C23:J38"), Type:=xlFillDefault

'MEDIDAS [m] KITS
Range("A40").FormulaR1C1 = "MEDIDAS [m] MÉDIAS KITS"
Range("A41").FormulaR1C1 = "FAMILIA"
    Range("A42").FormulaR1C1 = "BLINDEX"
    Range("A43").FormulaR1C1 = "BLINDEX"
    Range("A44").FormulaR1C1 = "BLINDEX"
    Range("A45").FormulaR1C1 = "BLINDEX"
    Range("A46").FormulaR1C1 = "BOX"
    Range("A47").FormulaR1C1 = "BOX"
    Range("A48").FormulaR1C1 = "BOX"
    Range("A49").FormulaR1C1 = "BOX"
    Range("B41").FormulaR1C1 = "KITS"
    Range("B42").FormulaR1C1 = "KF2P"
    Range("B43").FormulaR1C1 = "KC4P"
    Range("B44").FormulaR1C1 = "RETO KF2P"
    Range("B45").FormulaR1C1 = "RETO KC4P"
    Range("B46").FormulaR1C1 = "BF1"
    Range("B47").FormulaR1C1 = "BC1"
    Range("B48").FormulaR1C1 = "RETO BF1"
    Range("B49").FormulaR1C1 = "RETO BC1"
    Range("C41").FormulaR1C1 = "FOSCO"
    Range("D41").FormulaR1C1 = "BRANCO"
    Range("E41").FormulaR1C1 = "BRILHO"
    Range("F41").FormulaR1C1 = "PRETO"
    Range("G41").FormulaR1C1 = "BRONZE"
    Range("H41").FormulaR1C1 = "DOURADO"
    Range("I41").FormulaR1C1 = "ROSE"
    Range("J41").FormulaR1C1 = "INOX"
    
    Range("C42").Select
        ActiveCell.FormulaR1C1 = "=ROUND((SUMIFS(Macro!C30,Macro!C17,RC1,Macro!C18,RC2,Macro!C21,R41C))/(SUMIFS(Macro!C34,Macro!C17,RC1,Macro!C18,RC2,Macro!C21,R41C,Macro!C16,""KITS"")),2)"
    Range("C42").Select
        Selection.AutoFill Destination:=Range("C42:J42"), Type:=xlFillDefault
    Range("C42:J42").Select
        Selection.AutoFill Destination:=Range("C42:J49"), Type:=xlFillDefault

End Sub

Sub L03_LUCRO_ROAPLAS()

'As mudanças do programa não aparecem passo-a-passo na tela
Application.ScreenUpdating = False
Application.DisplayAlerts = False

ActiveWorkbook.Sheets.Add
    ActiveSheet.Name = "Lucratividade_Roaplas"

'TÍTULOS
    Range("A1").Value = "ANÁLISE LUCRATIVIDADE ROAPLAS"
    Range("A2").FormulaR1C1 = "QUANTIDADE [PEÇAS] DE KITS FATURADOS"
    Range("A3").FormulaR1C1 = "KITS"
    Range("A4").FormulaR1C1 = "PACIFIC F1"
    Range("A5").FormulaR1C1 = "PACIFIC F2"
    Range("A6").FormulaR1C1 = "PACIFIC F3"
    Range("A7").FormulaR1C1 = "PACIFIC C1"
    Range("B3").FormulaR1C1 = "BRILHO"
    Range("C3").FormulaR1C1 = "PRETO"
    Range("D3").FormulaR1C1 = "DOURADO"
    
    Range("A2:D7").Select
        Selection.Copy
    Range("A9").Select
        ActiveSheet.Paste
    Range("A16").Select
        ActiveSheet.Paste
    Range("A9").Value = "VALOR [R$] DE KITS FATURADOS"
    Range("A16").Value = "MEDIDAS [m] MÉDIAS KITS"
    

'QUANTIDADE [PEÇAS] DE KITS FATURADOS
Range("B4").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIFS(Macro!C34,Macro!C18,RC1,Macro!C21,R3C,Macro!C6,""RGKIT COMERCIO & INDUSTRIA EIR"")"
Range("B4").Select
    Selection.AutoFill Destination:=Range("B4:D4"), Type:=xlFillDefault
Range("B4:D4").Select
    Selection.AutoFill Destination:=Range("B4:D7"), Type:=xlFillDefault

'VALOR [R$] DE KITS FATURADOS
Range("B11").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIFS(Macro!C13,Macro!C18,R[-7]C1,Macro!C21,R3C,Macro!C6,""RGKIT COMERCIO & INDUSTRIA EIR"")"
Range("B11").Select
    Selection.AutoFill Destination:=Range("B11:D11"), Type:=xlFillDefault
Range("B11:D11").Select
    Selection.AutoFill Destination:=Range("B11:D14"), Type:=xlFillDefault

'MEDIDAS [m] MÉDIAS KITS
Range("B18").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIFS(Macro!C30,Macro!C18,R[-14]C1,Macro!C21,R3C,Macro!C6,""RGKIT COMERCIO & INDUSTRIA EIR"")/SUMIFS(Macro!C34,Macro!C18,R[-14]C1,Macro!C21,R3C,Macro!C6,""RGKIT COMERCIO & INDUSTRIA EIR"")"
Range("B18").Select
    Selection.AutoFill Destination:=Range("B18:D18"), Type:=xlFillDefault
Range("B18:D18").Select
    Selection.AutoFill Destination:=Range("B18:D21"), Type:=xlFillDefault
    


End Sub

