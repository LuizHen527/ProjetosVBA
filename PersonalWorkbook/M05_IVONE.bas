Attribute VB_Name = "M05_IVONE"
Sub Executar_Ivone()

Application.DisplayAlerts = False

On Error GoTo panda

    If Sheets("Fechamento_Ivone").Activate = True Then
        Sheets("Fechamento_Ivone").Delete
    End If

panda:

Call M11_Ivone

Application.DisplayAlerts = True

End Sub


Sub M11_Ivone()
ActiveWorkbook.Sheets.Add 'Adiciona uma nova planilha, que se torna ativa
ActiveSheet.Name = "Fechamento_Ivone"
    
'Nome famílias
    Range("A2").FormulaR1C1 = "Revenda"
    Range("A3").FormulaR1C1 = "Kit box"
    Range("A4").FormulaR1C1 = "Kit Blindex"
    Range("A5").FormulaR1C1 = "Kit Engenharia"
    Range("A6").FormulaR1C1 = "Kit Pia"
    Range("A7").FormulaR1C1 = "Kit Slim"
    Range("A8").FormulaR1C1 = "Molduras"
    Range("A9").FormulaR1C1 = "Perfis"
    Range("A10").FormulaR1C1 = "Suportes Alcoa"
    Range("A11").FormulaR1C1 = "Botao frances"
    Range("A12").FormulaR1C1 = "Cremalheira e suporte"
    Range("A13").FormulaR1C1 = "Prolongadores"
    Range("A14").FormulaR1C1 = "Suporte S"
    Range("A15").FormulaR1C1 = "Vitrines"
    Range("A16").FormulaR1C1 = "Roaplas"
    Range("A17").FormulaR1C1 = "Dualdoor"
    Range("A18").FormulaR1C1 = "Outros"

    
    Range("B1").FormulaR1C1 = "R$ VENDA"
    Range("B2").FormulaR1C1 = "=SUMIF(Macro!C[14],""ACESSORIOS"",Macro!C[11])-R[9]C"
    Range("B3").FormulaR1C1 = "=SUMIFS(Macro!C13,Macro!C16,""KITS"",Macro!C17,""BOX"")"
    Range("B4").FormulaR1C1 = "=SUMIFS(Macro!C13,Macro!C16,""KITS"",Macro!C17,""BLINDEX"")+SUMIFS(Macro!C13,Macro!C16,""KITS"",Macro!C17,""COMBATE"")"
    Range("B5").FormulaR1C1 = "=SUMIFS(Macro!C13,Macro!C16,""KITS"",Macro!C17,""ENGENHARIA"")"
    Range("B6").FormulaR1C1 = "=SUMIFS(Macro!C13,Macro!C16,""KITS"",Macro!C17,""PIA"")"
    Range("B7").FormulaR1C1 = "=SUMIFS(Macro!C13,Macro!C16,""KITS"",Macro!C17,""SLIM"")"
    Range("B8").FormulaR1C1 = "=SUMIFS(Macro!C13,Macro!C16,""MOLDURAS"")"
    Range("B9").FormulaR1C1 = "=SUMIFS(Macro!C13,Macro!C16,""PERFIS"")-SUMIFS(Macro!C13,Macro!C16,""PERFIS"",Macro!C17,""ROAPLAS"")-SUMIFS(Macro!C13,Macro!C16,""PERFIS"",Macro!C17,""VITRINE"")"
    Range("B10").FormulaR1C1 = "=SUMIFS(Macro!C13,Macro!C16,""SUPORTES"",Macro!C17,""PRATELEIRA"")"
    Range("B11").FormulaR1C1 = "=SUMIFS(Macro!C13,Macro!C16,""ACESSORIOS"",Macro!C17,""BOTOES"",Macro!C18,""FRANCES"")"
    Range("B12").FormulaR1C1 = "=SUMIFS(Macro!C13,Macro!C16,""SUPORTES"",Macro!C17,""CREMALHEIRA"")+SUMIFS(Macro!C13,Macro!C16,""SUPORTES"",Macro!C17,""BARRA CREMALHEIRA"")"
    Range("B13").FormulaR1C1 = "=SUMIFS(Macro!C13,Macro!C16,""PROLONGADOR"")"
    Range("B14").FormulaR1C1 = "=SUMIFS(Macro!C13,Macro!C16,""SUPORTES"",Macro!C17,""S"")"
    Range("B15").FormulaR1C1 = "=SUMIFS(Macro!C13,Macro!C16,""PERFIS"",Macro!C17,""VITRINE"")"
    Range("B16").FormulaR1C1 = "=SUMIFS(Macro!C13,Macro!C16,""KITS"",Macro!C17,""ROAPLAS"")+SUMIFS(Macro!C13,Macro!C16,""PERFIS"",Macro!C17,""ROAPLAS"")-SUMIFS(Macro!C13,Macro!C16,""KITS"",Macro!C17,""ROAPLAS"",Macro!C18,""DUALDOOR"")"
    Range("B17").FormulaR1C1 = "=SUMIFS(Macro!C13,Macro!C16,""KITS"",Macro!C17,""ROAPLAS"",Macro!C18,""DUALDOOR"")"
    Range("B18").FormulaR1C1 = "=SUM(Macro!C[11])-SUM(R[-16]C:R[-1]C)"
    Range("B19").FormulaR1C1 = "=SUM(R[-17]C:R[-1]C)"

    
    Range("C1").FormulaR1C1 = "QTD. VENDA"
    Range("C2").FormulaR1C1 = "=SUMIF(Macro!C[13],""ACESSORIOS"",Macro!C[31])-R[9]C"
    Range("C3").FormulaR1C1 = "=SUMIFS(Macro!C34,Macro!C16,""KITS"",Macro!C17,""BOX"")"
    Range("C4").FormulaR1C1 = "=SUMIFS(Macro!C34,Macro!C16,""KITS"",Macro!C17,""BLINDEX"")+SUMIFS(Macro!C34,Macro!C16,""KITS"",Macro!C17,""COMBATE"")"
    Range("C5").FormulaR1C1 = "=SUMIFS(Macro!C34,Macro!C16,""KITS"",Macro!C17,""ENGENHARIA"")"
    Range("C6").FormulaR1C1 = "=SUMIFS(Macro!C34,Macro!C16,""KITS"",Macro!C17,""PIA"")"
    Range("C7").FormulaR1C1 = "=SUMIFS(Macro!C34,Macro!C16,""KITS"",Macro!C17,""SLIM"")"
    Range("C8").FormulaR1C1 = "=SUMIFS(Macro!C34,Macro!C16,""MOLDURAS"")"
    Range("C9").FormulaR1C1 = "=SUMIFS(Macro!C34,Macro!C16,""PERFIS"")-SUMIFS(Macro!C34,Macro!C16,""PERFIS"",Macro!C17,""ROAPLAS"")-SUMIFS(Macro!C34,Macro!C16,""PERFIS"",Macro!C17,""VITRINE"")"
    Range("C10").FormulaR1C1 = "=SUMIFS(Macro!C34,Macro!C16,""SUPORTES"",Macro!C17,""PRATELEIRA"")"
    Range("C11").FormulaR1C1 = "=SUMIFS(Macro!C34,Macro!C16,""ACESSORIOS"",Macro!C17,""BOTOES"",Macro!C18,""FRANCES"")"
    Range("C12").FormulaR1C1 = "=SUMIFS(Macro!C34,Macro!C16,""SUPORTES"",Macro!C17,""CREMALHEIRA"")+SUMIFS(Macro!C34,Macro!C16,""SUPORTES"",Macro!C17,""BARRA CREMALHEIRA"")"
    Range("C13").FormulaR1C1 = "=SUMIFS(Macro!C34,Macro!C16,""PROLONGADOR"")"
    Range("C14").FormulaR1C1 = "=SUMIFS(Macro!C34,Macro!C16,""SUPORTES"",Macro!C17,""S"")"
    Range("C15").FormulaR1C1 = "=SUMIFS(Macro!C34,Macro!C16,""PERFIS"",Macro!C17,""VITRINE"")"
    Range("C16").FormulaR1C1 = "=SUMIFS(Macro!C34,Macro!C16,""KITS"",Macro!C17,""ROAPLAS"")+SUMIFS(Macro!C34,Macro!C16,""PERFIS"",Macro!C17,""ROAPLAS"")-SUMIFS(Macro!C34,Macro!C16,""KITS"",Macro!C17,""ROAPLAS"",Macro!C18,""DUALDOOR"")"
    Range("C17").FormulaR1C1 = "=SUMIFS(Macro!C34,Macro!C16,""KITS"",Macro!C17,""ROAPLAS"",Macro!C18,""DUALDOOR"")"
    Range("C18").FormulaR1C1 = "=SUM(Macro!C[31])-SUM(R[-16]C:R[-1]C)"
    Range("C19").FormulaR1C1 = "=SUM(R[-17]C:R[-1]C)"



    Range("D1").FormulaR1C1 = "PESO VENDA"
    Range("D2").FormulaR1C1 = "=SUMIF(Macro!C[12],""ACESSORIOS"",Macro!C[32])-R[9]C"
    Range("D3").FormulaR1C1 = "=SUMIFS(Macro!C36,Macro!C16,""KITS"",Macro!C17,""BOX"")"
    Range("D4").FormulaR1C1 = "=SUMIFS(Macro!C36,Macro!C16,""KITS"",Macro!C17,""BLINDEX"")+SUMIFS(Macro!C36,Macro!C16,""KITS"",Macro!C17,""COMBATE"")"
    Range("D5").FormulaR1C1 = "=SUMIFS(Macro!C36,Macro!C16,""KITS"",Macro!C17,""ENGENHARIA"")"
    Range("D6").FormulaR1C1 = "=SUMIFS(Macro!C36,Macro!C16,""KITS"",Macro!C17,""PIA"")"
    Range("D7").FormulaR1C1 = "=SUMIFS(Macro!C36,Macro!C16,""KITS"",Macro!C17,""SLIM"")"
    Range("D8").FormulaR1C1 = "=SUMIFS(Macro!C36,Macro!C16,""MOLDURAS"")"
    Range("D9").FormulaR1C1 = "=SUMIFS(Macro!C36,Macro!C16,""PERFIS"")-SUMIFS(Macro!C36,Macro!C16,""PERFIS"",Macro!C17,""ROAPLAS"")-SUMIFS(Macro!C36,Macro!C16,""PERFIS"",Macro!C17,""VITRINE"")"
    Range("D10").FormulaR1C1 = "=SUMIFS(Macro!C36,Macro!C16,""SUPORTES"",Macro!C17,""PRATELEIRA"")"
    Range("D11").FormulaR1C1 = "=SUMIFS(Macro!C36,Macro!C16,""ACESSORIOS"",Macro!C17,""BOTOES"",Macro!C18,""FRANCES"")"
    Range("D12").FormulaR1C1 = "=SUMIFS(Macro!C36,Macro!C16,""SUPORTES"",Macro!C17,""CREMALHEIRA"")+SUMIFS(Macro!C36,Macro!C16,""SUPORTES"",Macro!C17,""BARRA CREMALHEIRA"")"
    Range("D13").FormulaR1C1 = "=SUMIFS(Macro!C36,Macro!C16,""PROLONGADOR"")"
    Range("D14").FormulaR1C1 = "=SUMIFS(Macro!C36,Macro!C16,""SUPORTES"",Macro!C17,""S"")"
    Range("D15").FormulaR1C1 = "=SUMIFS(Macro!C36,Macro!C16,""PERFIS"",Macro!C17,""VITRINE"")"
    Range("D16").FormulaR1C1 = "=SUMIFS(Macro!C36,Macro!C16,""KITS"",Macro!C17,""ROAPLAS"")+SUMIFS(Macro!C36,Macro!C16,""PERFIS"",Macro!C17,""ROAPLAS"")-SUMIFS(Macro!C36,Macro!C16,""KITS"",Macro!C17,""ROAPLAS"",Macro!C18,""DUALDOOR"")"
    Range("D17").FormulaR1C1 = "=SUMIFS(Macro!C36,Macro!C16,""KITS"",Macro!C17,""ROAPLAS"",Macro!C18,""DUALDOOR"")"
    Range("D18").FormulaR1C1 = "=SUM(Macro!C[32])-SUM(R[-16]C:R[-1]C)"
    Range("D19").FormulaR1C1 = "=SUM(R[-17]C:R[-1]C)"

    
    ActiveCell.FormulaR1C1 = "=SUM(R[-15]C:R[-1]C)"
    Range("F1").FormulaR1C1 = ""
    Range("F2").FormulaR1C1 = "AF01"
    Range("F3").FormulaR1C1 = "AF06"
    Range("F4").FormulaR1C1 = "AF07"
    Range("F5").FormulaR1C1 = "AF12"
    Range("F6").FormulaR1C1 = "AF13"
    Range("F7").FormulaR1C1 = "AF14"
    Range("F8").FormulaR1C1 = "AF15"
    Range("F09").FormulaR1C1 = "AF16"
    Range("F10").FormulaR1C1 = "AF18"
    Range("F11").FormulaR1C1 = "AF20"
    Range("F12").FormulaR1C1 = "AF21"
    Range("F13").FormulaR1C1 = "AF22"
    Range("F14").FormulaR1C1 = "AF30"
    Range("F15").FormulaR1C1 = "AF40"
    Range("F16").FormulaR1C1 = "OVAL"
    Range("F17").FormulaR1C1 = "ESQ-01"
    Range("G1").FormulaR1C1 = "TOTAL VENDA"
    Range("G2").Select
        ActiveCell.FormulaR1C1 = "=SUMIFS(Macro!C13,Macro!C16,""MOLDURAS"",Macro!C17,RC6)"
            Selection.AutoFill Destination:=Range("G2:G17"), Type:=xlFillDefault
            Range("G2:G17").Select
    Range("G18").Select
        ActiveCell.FormulaR1C1 = "=SUM(R[-16]C:R[-1]C)"
    Range("H1").Select
        ActiveCell.FormulaR1C1 = "QUANTIDADE [PÇ]"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIFS(Macro!C34,Macro!C16,""MOLDURAS"",Macro!C17,RC6)"
            Selection.AutoFill Destination:=Range("H2:H17"), Type:=xlFillDefault
            Range("H2:H17").Select
    Range("H18").Select
        ActiveCell.FormulaR1C1 = "=SUM(R[-16]C:R[-1]C)"
    Range("I1").Select
        ActiveCell.FormulaR1C1 = "PESO"
    Range("I2").Select
        ActiveCell.FormulaR1C1 = "=SUMIFS(Macro!C36,Macro!C16,""MOLDURAS"",Macro!C17,RC6)"
            Selection.AutoFill Destination:=Range("I2:I17"), Type:=xlFillDefault
            Range("I2:I17").Select
    Range("I18").Select
        ActiveCell.FormulaR1C1 = "=SUM(R[-16]C:R[-1]C)"
    
    Range("K1").FormulaR1C1 = "ACABAMENTO"
    Range("K2").FormulaR1C1 = "AZUL"
    Range("K3").FormulaR1C1 = "BRANCO"
    Range("K4").FormulaR1C1 = "BRONZE B."
    Range("K5").FormulaR1C1 = "BRONZE F."
    Range("K6").FormulaR1C1 = "DOURADO B."
    Range("K7").FormulaR1C1 = "DOURADO F."
    Range("K8").FormulaR1C1 = "FUME B."
    Range("K9").FormulaR1C1 = "FUME F."
    Range("K10").FormulaR1C1 = "INCOLOR B."
    Range("K11").FormulaR1C1 = "INCOLOR F."
    Range("K12").FormulaR1C1 = "PRETO B."
    Range("K13").FormulaR1C1 = "PRETO F."
    Range("K14").FormulaR1C1 = "VERDE"
    Range("K15").FormulaR1C1 = "VINHO"
    Range("K16").FormulaR1C1 = "BRUTO"
    Range("L1").FormulaR1C1 = "TOTAL R$"
    Range("L2").FormulaR1C1 = _
        "=SUMIFS(Macro!C13,Macro!C16,""MOLDURAS"",Macro!C21,RC11)"
    Range("L2").Select
        Selection.AutoFill Destination:=Range("L2:L16"), Type:=xlFillDefault
        Range("L2:L16").Select
    Range("L17").Select
        ActiveCell.FormulaR1C1 = "=SUM(R[-15]C:R[-1]C)"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "TOTAL PÇS"
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=SUMIFS(Macro!C34,Macro!C16,""MOLDURAS"",Macro!C21,RC11)"
    Range("M2").Select
        Selection.AutoFill Destination:=Range("M2:M16"), Type:=xlFillDefault
        Range("M2:M16").Select
    Range("M17").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-15]C:R[-1]C)"
    Range("M18").Select
End Sub

