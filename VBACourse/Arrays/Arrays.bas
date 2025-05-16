Attribute VB_Name = "Arrays"
'@Folder("VBAProject")
Option Explicit

Sub DynamicArray2D()
    
    Dim arr(11, 1) As Variant
    Dim r As Byte 'Row
    Dim c As Byte 'Column
    
    For r = 0 To 11
        For c = 0 To 1
            arr(r, c) = Cells(r + 5, c + 1).Value
            
        Next c
    Next r
    
End Sub

Sub Activity()
    
    'Colocar dados em um array
    'Criar um novo workbook
    'Colocar os nomes das empresas nos nomes das sheets
    'Colocar o nome do manager na celula A1 de cada sheet
    
    Dim data(7, 1) As String
    Dim r As Byte 'rows
    Dim c As Byte 'columns
    Dim newWorkbook As Workbook
    Dim Sh As Worksheet
    
    ActiveWorkbook.Worksheets("f").Select
    
    For r = 0 To 7
        For c = 0 To 1
            data(r, c) = Cells(r + 7, c + 1).Value
        Next c
    Next r
    
    Set newWorkbook = Workbooks.Add
    With newWorkbook
        .Title = "Activity"
        .Subject = "Sales"
    End With
    
    r = 0
    
    'Vai mudar o nome das tabelas existentes
    For Each Sh In newWorkbook.Worksheets
    
        Sh.Name = data(r, 0)
        Sh.Range("A1") = data(r, 1)
        
        r = r + 1
    Next Sh
    
    'Criar tabelas para os nomes que faltam
    Do While r < 8
    
        newWorkbook.Sheets.Add
        
        ActiveSheet.Range("A1") = data(r, 1)
        ActiveSheet.Name = data(r, 0)
        
        r = r + 1
    Loop
End Sub
