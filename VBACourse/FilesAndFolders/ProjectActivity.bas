Attribute VB_Name = "ProjectActivity"
'@Folder("VBAProject")
Option Explicit

Sub ExportToFile()

    Dim ExpRange As Range
    Dim ExpRow As Range
    Dim ExpCell As Range
    Dim myValue As Variant
    Dim fileName As String
    
    Set ExpRange = Range("A6").CurrentRegion.Rows
    fileName = ThisWorkbook.path & "\ProjectActivity.csv"
    
    Open fileName For Output As #1
    For Each ExpRow In ExpRange
        For Each ExpCell In ExpRow.Cells
            myValue = myValue & ExpCell & ";"
        Next ExpCell
        
        myValue = Left(myValue, Len(myValue) - 1)
        
        Print #1, myValue
        
        myValue = ""
        
    Next ExpRow
    Close #1
    
    MsgBox "Tabela importada!"
    
End Sub
