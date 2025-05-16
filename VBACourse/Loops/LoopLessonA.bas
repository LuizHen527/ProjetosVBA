Attribute VB_Name = "LoopLessonA"
'@Folder("VBAProject")
Option Explicit
Const startRow As Byte = 10
Dim lastRow As Long

Sub SimpleFor()

 Dim i As Long
 Dim myValue As Double
 lastRow = Range("A" & startRow).End(xlDown).Row
 
 For i = startRow To lastRow
    myValue = Range("F" & i).Value
    If myValue > 400 Then Range("F" & i).Value = myValue + 10
    If myValue < 0 Then Exit For
 Next i
End Sub

Sub ForNextLoop()

    Dim i As Long
    Dim myValue As String
    Dim NumFound As Long
    Dim TxtFound As String
    Dim r As Long 'For looping through rows
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    
    
    
    For r = startRow To lastRow
    
        myValue = Range("A" & r).Value
        
        For i = 1 To Len(myValue)
            If IsNumeric(Mid(myValue, i, 1)) Then
                NumFound = NumFound & Mid(myValue, i, 1)
            
            ElseIf Not IsNumeric(myValue) Then
                TxtFound = TxtFound & Mid(myValue, i, 1)
            End If
        Next i
        
        Range("H" & r) = TxtFound
        Range("I" & r) = NumFound
        
        TxtFound = ""
        NumFound = 0
        
    Next r
End Sub

Sub Delete_hidden_filtered_rows()

    Dim r As Long
    
    lastRow = Range("A" & startRow).CurrentRegion.Rows.Count + startRow - 2
    
    For r = lastRow To startRow Step -1
        If Rows(r).Hidden Then
            Rows(r).Delete
        End If
    Next r
End Sub

Sub Copy_Filtered_List()

    ActiveSheet.AutoFilter.Range.Copy
    Worksheets.Add
    Range("A1").PasteSpecial
End Sub

Sub Clear()
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    
    Range("H10", "I" & lastRow).Clear
End Sub


