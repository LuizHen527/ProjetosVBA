Attribute VB_Name = "CopyPaste"
Option Explicit

Sub CopyPaste()

'--------------------------- LEIA ---------------------------
'Esse programa passa dados da tabela de produção diaria para
'a tabela de produção mensal

'Como usar?
    'Abra os arquivos:
        '"HISTÓRICO PRODUÇÃO 2022-2024_V5.xlsm"
        'E o arquivo da produção diaria do mes
        
    'IMPORTANTE: Tire os filtros da tabelas PROD. DIARIA (Base) e HISTORICO PRODUÇÂO (01_Base)
 
    'Rode o programa(tecla F5)

'Como adicionar mais colunas?
    'Coloque o mesmo nome da colunas nas duas tabelas.
    'Assim ele vai reconhecer a coluna sozinho.
    
    'Se o nome for diferente:
    'Tem um If gigante mais pra baixo com um comentario em cima.
    'É só colocar mais uma condição com o nome das colunas
    'Nome coluna1 = Nome da coluna na tabela de Produção
    'Nome coluna2 = Nome da coluna no Historico de produção
    'columnPaste = "Nome coluna1" And columnName = "Nome coluna2"
    
'------------------------------------------------------------

'Set de variaveis
Dim columnName As String, names As Long, x As Long
Dim columnPaste As Range, rngCopy As Long
Dim numRows As Integer, numRowsCopy As Integer, rowsSum As Integer, rngComumn As Long
Dim strFileName As String
Dim i As Integer, rowVerify As Integer, numRowShift As Integer

'Guarda o nome do arquivo da planilha de produção
strFileName = ThisWorkbook.Name

'Set do numero de linhas
Workbooks("HISTÓRICO PRODUÇÃO 2022-2024_V5.xlsm").Activate
Worksheets("01_Base").Select

'Desliga a atualização da tela
Application.ScreenUpdating = False

'Linha que os dados serão colados. Primeira linha vazia depois dos dados. Resolve problema da coluna OBS e Problema que os dados são copiados na linha errada
numRows = Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Row

Workbooks(strFileName).Activate
Worksheets("Base").Select

'Pega numero de linhas que estão sendo copiadas
numRowsCopy = Range("A5", "A" & Cells(Rows.Count, 1).End(xlUp).Row).Count

rowsSum = (numRows + numRowsCopy) - 1

'-------------------------------------------------------------------------------------------

'Loop que passa por cada coluna da produção diaria
For rngCopy = 1 To 47

    Workbooks(strFileName).Activate
    
    'Se tiver formula ele pula pra proxima coluna
    If ActiveWorkbook.Worksheets("Base").Cells(5, rngCopy).HasFormula = True Then
        GoTo NextIteration
    End If
    
    'Salva nome da coluna na produção diaria
    columnName = ActiveWorkbook.Worksheets("Base").Cells(4, rngCopy).Value
    
    'Copia os dados da coluna na produção diaria
    ActiveWorkbook.Worksheets("Base") _
    .Range(Cells(5, rngCopy).Address, Col_Letter(rngCopy) & Cells(Rows.Count, 1).End(xlUp).Row).Copy
    
'-------------------------------------------------------------------------------------------
    
    Workbooks("HISTÓRICO PRODUÇÃO 2022-2024_V5.xlsm").Activate
    Worksheets("01_Base").Select

    'Localiza coluna comparando os nomes das colunas
    For Each columnPaste In Range("A3", "BC3")
        'Verifica se o nome das duas colunas são iguais
        If LCase(columnPaste) = LCase(columnName) Then
        
            'Cola a coluna de dados na ultima linha
            Range(Col_Letter(columnPaste.Column) & numRows).PasteSpecial (xlPasteValues)
            Exit For
            
GoTo NextIteration
            Exit For
        End If
    Next
    
'-------------------------------------------------------------------------------------------

    'Localiza nomes que são diferentes
    For Each columnPaste In Range("A3", "BC3")
    
        '--------- ADICIONE COLUNAS COM NOMES DIFERENTES AQUI ---------
        
        'Verifica Nomes de colunas que não são iguais
        If columnPaste = "H. INICIO" And columnName = "HORA INICIAL" Or _
           columnPaste = "H. FINAL" And columnName = "HORA FINAL" Or _
           columnPaste = "QTD.1" And columnName = "QUANTIDADE TARUGO 1" Or _
           columnPaste = "COMP.1 [mm]" And columnName = "COMPRIMENTO 1 [MM]" Or _
           columnPaste = "COMP.2 [mm]" And columnName = "COMPRIMENTO 2 [MM]" Or _
           columnPaste = "QTD.2" And columnName = "QUANTIDADE TARUGO 2" Or _
           columnPaste = "PONTA [kg]" And columnName = "PONTAS [KG]" Or _
           columnPaste = "PROBLEMA" And columnName = "PROBLEMA2" Or _
           columnPaste = "OBS" And columnName = "OBSERVAÇÃO" Or _
           columnPaste = "T FERRAMENTA[°C]" And columnName = "TEMPERATURA FERRAMENTA [ºC]" Or _
           columnPaste = "T TARUGO [°C]" And columnName = "TEMPERATURA TARUGO [ºC]" Or _
           columnPaste = "T EMERGENTE [°C]" And columnName = "TEMPERATURA EMERGENTE [°C]" Or _
           columnPaste = "T CONTENEDOR [°C]" And columnName = "TEMPERATURA CONTENEDOR [°C]" Or _
           columnPaste = "V EXTRUSÃO [m/min]" And columnName = "VELOCIDADE EXTRUSÃO [M/MIN]" Or _
           columnPaste = "V PULLER [m/min]" And columnName = "VELOCIDADE DO PULLER [M/MIN]" _
        Then
            'Cola a coluna de dados na ultima linha
            Range(Col_Letter(columnPaste.Column) & numRows).PasteSpecial (xlPasteValues)
            Exit For
        End If
        
    Next
    
NextIteration:

Next

'-------------------------------------------------------------------------------------------

Workbooks("HISTÓRICO PRODUÇÃO 2022-2024_V5.xlsm").Activate

'Copia formulas do historico pra baixo
For rngComumn = 1 To 55

    'Se tiver formula ele pula pra proxima coluna
    If Worksheets("01_Base").Cells(5, rngComumn).HasFormula = True Then
        
        Range(Col_Letter(rngComumn) & Rows.Count).End(xlUp).Copy
        
        Range(Col_Letter(rngComumn) & numRows, Col_Letter(rngComumn) & rowsSum).PasteSpecial
        
    End If
    
Next

'-------------------------------------------------------------------------------------------

Workbooks("HISTÓRICO PRODUÇÃO 2022-2024_V5.xlsm").Activate

'Exclui todas as linhas que o nome do perfil que o nome é "PARADA DE PRODUÇÃO"
For i = numRowsCopy + 1 To 0 Step -1
    rowVerify = (rowsSum - i) - numRowShift

 If Worksheets("01_Base").Range("B" & rowVerify).Value = "PARADA PRODUÇÃO" Then
    Range("A" & rowVerify, "BC" & rowVerify).Delete (xlShiftUp)
    numRowShift = numRowShift + 1
 End If
Next

Application.ScreenUpdating = True

MsgBox "Dados da planilha " & strFileName & " foram copiados", vbInformation, "Sucesso"

End Sub

'Troca um numero por uma letra equivalente a uma coluna
Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function




