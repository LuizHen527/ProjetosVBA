Attribute VB_Name = "Activity"
'@Folder("VBAProject")
Option Explicit

Sub Activity()
    Dim Sh As Worksheet
    Dim SheetName As String
    Dim firstComment As Comment
    Dim com As Comment
    Dim i As Integer 'Iterador do loop
    Dim rngComment As String
    Dim commentText As String
    Dim rng As Range

    Worksheets.Add
    SheetName = ActiveSheet.Name
    
    Range("A1").Value = "Comentario"
    Range("B1").Value = "Endereço"
    Range("C1").Value = "Autor"

    'Loopar por cada planilha
    'Procurar por comentarios
        'Se não tiver ele pula pra proxima planilha
        'Se tiver ele vai salvando e colocando na planilha e depois procura o proximo(FindNext)
        'Quando ele achar o primeiro denovo ele pula pra proxima planilha
        'Ele termina depois de passar por todas as planilhas
    
    For Each Sh In Worksheets
        Sh.Activate
        
        For Each com In ActiveSheet.Comments
            Worksheets(SheetName).Activate
            
            
            Range("A" & Cells(Rows.Count, 1).End(xlUp).Row).Offset(1, 0).Value = com.Text
            Range("B" & Cells(Rows.Count, 2).End(xlUp).Row).Offset(1, 0).Value = com.Parent.Address
            Range("C" & Cells(Rows.Count, 3).End(xlUp).Row).Offset(1, 0).Value = com.Author
            
        Next com
    Next Sh
End Sub
