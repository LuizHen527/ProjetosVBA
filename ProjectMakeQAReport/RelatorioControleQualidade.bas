Attribute VB_Name = "RelatorioControleQualidade"
'@Folder("VBAProject")
Option Explicit
Public baseData() As String

'Fazer o botao cancelar vai ser mais facil. Easy win
'Na segunda parte do programa (Confirmar)

'A fazer:
    'Part2: Mensagem das ferramentas que vao ser atualizadas
    'Part2: Verificar se alguma ferramenta foi editada. Se foi, salvar alterações
    'Part2: Fazer logica de contagem de problemas

'Catch:

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
    Dim selectedDate As Variant
    Dim targetMonth As String, baseName As String
    Dim selectedDateResponse As VbMsgBoxResult
    Dim productionFolder As Object
    Dim productionYes As Byte, productionNo As Byte, productionProblem As Byte
    Dim iterator As Integer, i As Integer
    Dim convertedDate As Date
    
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
    
    '--------------- CAPTURANDO DADOS ---------------
    On Error GoTo FolderNotFound
    Workbooks.Open Filename:=productionFolder & "\" & "\20" & selectedDate(1) & " Extrusão e Produção\02_PRODUÇÃO DIÁRIA\" & targetMonth & " - PROD. DIÁRIA " & UCase(selectedDate(0)) & " 20" & selectedDate(1) & ".xlsm"
    On Error GoTo 0

    baseName = ActiveWorkbook.Name
    
    ActiveWorkbook.Worksheets("Base").Select
    
    ReDim baseData(Range("A5", "A" & Cells(Rows.Count, 1).End(xlUp).Row).Count - 1, 5)
    
    
    'Array que salva dados da base da producao diaria e conta os tipos de producao
    For iterator = 0 To Range("A5", "A" & Cells(Rows.Count, 1).End(xlUp).Row).Count - 1
        
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
        
        'Conta produção = sim
        If Range("AM" & iterator + 5) = "SIM" Then productionYes = productionYes + 1
        
        'Conta produção = nao
        If Range("AM" & iterator + 5) = "NÃO" And Not Range("AM" & iterator + 5) = "TESTE" Then productionNo = productionNo + 1
        
        'Conta produção = problema
        If Range("AM" & iterator + 5) = "PROBLEMA" Then productionProblem = productionProblem + 1
        
    Next iterator
    
    ThisWorkbook.Worksheets("Relatório").Activate
    
    'filtrando dados para um array que tenha apenas ferramentas com problema
    For iterator = 0 To UBound(baseData)
        
        If baseData(iterator, 3) = "RISCO" Or baseData(iterator, 3) = "ACABAMENTO" Then
            
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
    Dim iterator As Byte
    Dim itensToUpdate As String
    
    'Verificar se alguma ferramenta foi editada. Como?
        'Loopar pela tabela e comparar o que está no array com o que esta na tabela
        'Se algum for diferente, salva a mudança no array.
        
        'Pra fazer o aviso de que vão ser alteradas: Loop pelos itens atualizados e vai montando uma string.
        'Depois joga a sring no msgbox
    
    For iterator = 0 To Range("p3", "p" & Cells(Rows.Count, 16).End(xlUp).Row).Count
        
        
        If Not baseData(Range("U" & 3 + iterator).Value, 3) = Range("S" & 3 + iterator).Value Then
            
            itensToUpdate = itensToUpdate & Range("Q" & 3 + iterator).Value & vbTab & baseData(Range("U" & 3 + iterator).Value, 3) & vbTab & Range("S" & 3 + iterator).Value & vbNewLine
            
        End If
    Next iterator
    
    MsgBox itensToUpdate





    ActiveSheet.Shapes("btnStart").Visible = True
    ActiveSheet.Shapes("btnConfirm").Visible = False
    ActiveSheet.Shapes("btnCancel").Visible = False
    ActiveSheet.Shapes("btnStart").Fill.ForeColor.RGB = RGB(11, 29, 81)
    
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



'----------- TESTES -----------

Sub btnStart()

    ActiveSheet.Shapes("btnCancel").Visible = True
    ActiveSheet.Shapes("btnConfirm").Visible = True
    ActiveSheet.Shapes("btnStart").Fill.ForeColor.RGB = RGB(115, 147, 179)
End Sub


Sub btnConfirm()

    ActiveSheet.Shapes("btnStart").Visible = True
    ActiveSheet.Shapes("btnConfirm").Visible = False
    ActiveSheet.Shapes("btnCancel").Visible = False
    ActiveSheet.Shapes("btnStart").Fill.ForeColor.RGB = RGB(11, 29, 81)
End Sub

Sub btnCancel()

    ActiveSheet.Shapes("btnStart").Visible = True
    ActiveSheet.Shapes("btnConfirm").Visible = False
    ActiveSheet.Shapes("btnCancel").Visible = False
    ActiveSheet.Shapes("btnStart").Fill.ForeColor.RGB = RGB(11, 29, 81)
End Sub
