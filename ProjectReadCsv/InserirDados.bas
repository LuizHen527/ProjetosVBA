Attribute VB_Name = "InserirDados"
'@Folder("VBAProject")
Option Explicit

'Macro coloca todos os endereços a partir de um arquivo csv.
'Fiz um programa em Python pra ler o arquivo PDF e fazer um csv

'Coisas pra adicionar ao projeto
    'Padrão de nomes de arquivos: Os file names e file paths estão todos mockados
    'uma melhoria pro projeto seria fazer nomes dinamicos de arquivos.
    
'Colocar codigo no github e colocar aqui


'Ler arquivo
'Fazer header -> O que vai mudar é só o bloco. Quando for outro bloco, mudar de aba.
'Conteudo -> Colocar os dados e ir estilizando as linhas

Sub InserirDados()
    Dim csvLine As String
    Dim caminho As String
    Dim lineData() As String
    Dim data As Variant
    Dim initialBloco As String
    Dim initialRange As Range
    
    initialBloco = "Bloco A"
    Set initialRange = Range("A5")
    
    DoHeader (initialBloco)
    

    Open "C:\Users\Molducolor7\Desktop\pdf_to_csv\csv_files\condominio_residencial_scs3.csv" For Input As #1

        Do While Not EOF(1)
            
            Line Input #1, csvLine

            lineData = Split(csvLine, ";")
            
            'Se mudar de bloco, cria nova aba
            If Not Left(lineData(2), 7) = initialBloco Then
                Set initialRange = Range("A5")
                
                'cria nova aba
                Sheets.Add
                
                initialBloco = Left(lineData(2), 7)
                
                ActiveSheet.Name = initialBloco
                
                DoHeader (initialBloco)
                
            End If

            DoAddress lineData, initialRange
            
            Set initialRange = initialRange.Offset(7, 0)
        Loop

    Close #1
End Sub

'Faz a celula com o titulo
Function DoHeader(bloco As String)
    
    'Estilizando celula
    With Range("A3:B3")
        .Interior.Color = RGB(165, 165, 165)
        .Merge
        .ColumnWidth = 53
        .Borders.LineStyle = xlDouble
        .Borders.Color = RGB(63, 63, 63)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    
    'Colocar dados
    Range("A3").Value = "TERRITÓRIO 91 - Condomínio Residencial São Caetano do Sul " & bloco
End Function

Function DoAddress(lineData() As String, initialRange As Range)
    Dim data As Variant
    Dim row As Integer
    
    row = initialRange.row
    
    
    'Irmão
    With Range("A" & row & ":B" & row)
        .Interior.Color = RGB(255, 217, 101)
        .Merge
        .ColumnWidth = 53
        .Borders.LineStyle = xlDouble
        .Borders.Color = RGB(63, 63, 63)
        .Font.Bold = True
        .Font.Italic = True
        .HorizontalAlignment = xlLeft
        .Value = "Irmãos:"
    End With
    
    row = row + 1
    
    'Condominio Nome
    With Range("A" & row & ":B" & row)
        .Interior.Color = RGB(91, 155, 213)
        .Merge
        .ColumnWidth = 53
        .Borders.LineStyle = xlDouble
        .Borders.Color = RGB(63, 63, 63)
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .Value = lineData(0)
    End With
    
    row = row + 1
    
    'Rua
    With Range("A" & row & ":B" & row)
        .Interior.Color = RGB(197, 90, 17)
        .Merge
        .ColumnWidth = 53
        .Borders.LineStyle = xlDouble
        .Borders.Color = RGB(63, 63, 63)
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .Value = lineData(1)
    End With
    
    row = row + 1
    
    'Bloco
    With Range("A" & row & ":B" & row)
        .Interior.Color = RGB(165, 165, 165)
        .Merge
        .ColumnWidth = 53
        .Borders.LineStyle = xlDouble
        .Borders.Color = RGB(63, 63, 63)
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .Value = lineData(2)
    End With
    
    row = row + 1
    
    'Retirado
    With Range("A" & row)
        .Interior.Color = RGB(0, 176, 80)
        .ColumnWidth = 53
        .Borders.LineStyle = xlDouble
        .Borders.Color = RGB(63, 63, 63)
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .Value = "Retirado: "
    End With
    
    'Postado
    With Range("B" & row)
        .Interior.Color = RGB(0, 176, 80)
        .ColumnWidth = 53
        .Borders.LineStyle = xlDouble
        .Borders.Color = RGB(63, 63, 63)
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .Value = "Postado: "
    End With
    
        

End Function
