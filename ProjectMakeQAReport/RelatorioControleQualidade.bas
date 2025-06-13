Attribute VB_Name = "RelatorioControleQualidade"
'@Folder("VBAProject")
Option Explicit


Sub CapturarDados()

    
    '--------------- VARIAVEIS ---------------
    Dim selectedDate() As String, inputBoxAnswer As Variant
    Dim targetMonth As Integer
    Dim selectedDateResponse As VbMsgBoxResult
    Dim verifyMonth As Boolean
    
    '--------------- SELECIONANDO A DATA ---------------
    
    selectedDate = Split(Range("J5").Value, "_")
    
    'Retorna o numero do mês
    targetMonth = month(DateValue("01 " & selectedDate(0) & " 2025"))
    
    targetMonth = targetMonth + 1
    
    'Pergunta se usuario quer mes previsto
    selectedDateResponse = MsgBox("Quer pegar os dados da data abaixo?" & vbNewLine & vbNewLine & MonthName(targetMonth) & " de " & "20" & selectedDate(1) _
    , vbQuestion + vbYesNoCancel, "Selecionar data")
    
    If selectedDateResponse = vbNo Then
        '---- INPUT BOX ----
InputBoxError:

        inputBoxAnswer = Application.InputBox("Escreva a data que deseja:" & vbNewLine & vbNewLine & "Siga o seguinte padrão: abril_24", "Selecione uma data", , , , , , 2 + 4 + 16)
        
        'Finaliza macro caso ele clique em cancelar ou no X
        If inputBoxAnswer = False Then
            Exit Sub
        End If
        
        'Separa mes e ano
        selectedDate = Split(inputBoxAnswer, "_")
        
        'Verifica se o mes existe
        verifyMonth = VerificaMes(selectedDate(0))
        
        'Tratamento de excessoes
        If verifyMonth = False Then
            MsgBox "Digite um mês valido.", vbExclamation, "Aviso"
            GoTo InputBoxError:
            
        ElseIf UBound(selectedDate, 1) < 1 Or selectedDate(1) = "" Then
            MsgBox "Digite um mês e um ano. Separe eles com um underline (_). Dessa forma: " _
            & " abril_25", vbExclamation, "Aviso"
            GoTo InputBoxError:
            
        ElseIf selectedDate(1) < 24 Or selectedDate(1) > 40 Or selectedDate(1) = "" Then
            MsgBox "Digite um ano valido.", vbExclamation, "Aviso"
            GoTo InputBoxError:
        
        End If
        
        
    ElseIf selectedDateResponse = vbCancel Then
        'Executa se o usuario cancelar
        Exit Sub
    End If
    
    'tranforma mes em numero
    targetMonth = month(DateValue("01 " & selectedDate(0) & " 2025"))
    
End Sub

Function VerificaMes(mes As String) As Boolean

    'pegar mes e verificar se ele é igual ao nome de um mes
    'se for um mes que existe ele retorna 1
    'se não for valido ele retorna 0
    Dim meses As Variant
    Dim n As Integer
    
    LCase (mes)
    
    meses = Array("janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro")

    For n = 0 To 11
        If mes = meses(n) Then
            'Mes é valido
            VerificaMes = True
            
            Exit Function
        End If
    Next n
    
    VerificaMes = False
    
End Function


