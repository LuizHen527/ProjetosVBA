Attribute VB_Name = "RelatorioControleQualidade"
'@Folder("VBAProject")
Option Explicit

'A fazer:
    'Pegar dados da coluna PRODUÇÃO na planilha
    'Quantidade PRODUÇAO = Sim (Produziu)
    'PRODUÇÃO

'Catch:

'Bug:


    
Sub CapturarDados()
    Application.DisplayAlerts = False

    
    '--------------- VARIAVEIS ---------------
    Dim selectedDate As Variant
    Dim targetMonth As String
    Dim selectedDateResponse As VbMsgBoxResult
    Dim productionFolder As Object
    Dim a As Object
    
    '.getfolder("\\121.137.1.5\alumitec9\COMPRAS\25_Compras")
    Set productionFolder = CreateObject("Scripting.FileSystemObject").getfolder("\\121.137.1.5\alumitec9\PRODUÇÃO")

    '--------------- SELECIONANDO A DATA ---------------
    
    'pega data do ultimo relatorio
    selectedDate = Split(Range("J5").Value, "_")
    
    If LCase(selectedDate(0)) = "dezembro" Then
        selectedDate(0) = "janeiro"
    End If
    
    'Retorna o numero do mês
    targetMonth = month(DateValue("01 " & selectedDate(0) & " 2025"))
    
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
    End If

    'tranforma mes em numero
    targetMonth = month(DateValue("01 " & selectedDate(0) & " 2025"))
    
    If targetMonth < 10 Then targetMonth = "0" & targetMonth
    
    Debug.Print "\20" & selectedDate(1) & " Extrusão e Produção\02_PRODUÇÃO DIÁRIA\" & targetMonth & " - PROD. DIÁRIA " & UCase(selectedDate(0)) & " 20" & selectedDate(1) & ".xlsm"
    
    On Error GoTo FolderNotFound
    Workbooks.Open Filename:=productionFolder & "\" & "\20" & selectedDate(1) & " Extrusão e Produção\02_PRODUÇÃO DIÁRIA\" & targetMonth & " - PROD. DIÁRIA " & UCase(selectedDate(0)) & " 20" & selectedDate(1) & ".xlsm"
    
    
    Exit Sub
    
    '--------------- ERROR HANDLING ---------------

    'Caso não encontre o arquivo da produção diaria
FolderNotFound:
    MsgBox "Verifique se arquivo existe ou esta com o nome errado.", vbExclamation + vbOKOnly, "Arquivo não encontrado"
    
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


