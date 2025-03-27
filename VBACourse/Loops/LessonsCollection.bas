Attribute VB_Name = "LessonsCollection"
Option Explicit

Sub With_Change_Font()

    Dim myRange As Range
    Set myRange = Range("A10", "A" & Cells(Rows.Count, 1).End(xlUp).Row)
    
    With myRange.Font
        .Name = "Arial"
        .Size = 12
        .Bold = True
    End With
End Sub

Sub With_Reset_Font()

    Dim myRange As Range
    Set myRange = Range("A10", "A" & Cells(Rows.Count, 1).End(xlUp).Row)
    
    With myRange.Font
        .Name = "Calibri"
        .Size = 11
        .Bold = False
    End With
End Sub

Sub Unprotect_All_Sheets()

    Dim Sh As Worksheet
    
    For Each Sh In ThisWorkbook.Worksheets
        Sh.Unprotect "test"
        
    Next Sh

End Sub

Sub Protect_All_Sheets()

    Dim Sh As Worksheet
    
    For Each Sh In ThisWorkbook.Worksheets
        
        Sh.Protect Password:="test", AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True
        
    Next Sh

End Sub

Sub Simple_If()

    If Range("B3").Value <> "" Then Range("C3").Value = Range("B3").Value
     
    If Range("b4").Value >= 0 And Range("b4").Value <= 400 Then
        Range("c4").Value = Range("b4").Value
    End If
End Sub

Sub Protect_Specific_Sheets()

    Dim Sh As Worksheet
    
    For Each Sh In ThisWorkbook.Worksheets
        
        If Sh.Name = "Purpose" Then
            Sh.Protect
        ElseIf Sh.CodeName = "Sheet1" Then
            Sh.Unprotect "test"
        Else
        Sh.Protect Password:="test", AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True
        End If
        
    Next Sh

End Sub

Sub Simple_Case()

Select Case Range("b3").Value
    Case 1 To 200
        Range("c3").Value = "Good"
    Case 0
        Range("c3").Value = ""
    Case Is > 200
        Range("c3").Value = "Excellent"
    Case Is < 0
        Range("c3").Value = "Bad"
End Select
End Sub
 
Sub Error_Handling()
    Range("d3").Value = ""
    
    If VBA.IsError(Range("B3").Value) Then GoTo getout
    
    Range("c3").Value = Range("b3").Value
    
    Exit Sub
    
getout:
    Range("c3").Value = ""
    Range("d3").Value = "A celula tem um erro"

End Sub

Sub Activity()
Dim numFormulas As Integer
Dim cell As Range

For Each cell In Sheet3.UsedRange

    If cell.HasFormula Then
    
        numFormulas = numFormulas + 1
        
    End If
    
Next

Range("b6").Value = numFormulas

End Sub
