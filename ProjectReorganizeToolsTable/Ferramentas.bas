Attribute VB_Name = "Ferramentas"
'@Folder("VBAProject")
Option Explicit

'Programa feito pra reorganizar a tabela de ferramentas de forma que
'facilite a visualização. Organizando por ferramenta e não por data


Sub Main()
    Dim numberRows As Integer, iterator As Integer
    Dim lineDate As String
    'Salvar dados
    
    numberRows = Range("B4:" & Cells(Rows.Count, 2).End(xlUp).Address).Count - 1
    
    For iterator = 0 To numberRows
    
        If Not Cells(iterator + 4, 1).Value = "" Then
            lineDate = Cells(iterator + 4, 1).Value
        End If
        
        'Data
        Debug.Print lineDate; Cells(iterator + 4, 2).Value; Cells(iterator + 4, 3).Value; Cells(iterator + 4, 4).Value
        
        'Ferramenta
        'Debug.Print Cells(iterator + 4, 2).Value
        
        'SEQ
        'Debug.Print Cells(iterator + 4, 3).Value
        
        'Peso
        'Debug.Print Cells(iterator + 4, 4).Value
        
        'Tarugos
        'Debug.Print Cells(iterator + 4, 5).Value
        
    Next iterator
End Sub
