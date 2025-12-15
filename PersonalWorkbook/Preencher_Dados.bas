Attribute VB_Name = "Preencher_Dados"
Option Base 1

Sub Preencher_Dados()

'Application.ScreenUpdating = False
'Application.DisplayAlerts = False

'Sheets("Base").Select

células = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "K", "L", "M")

For a = 1 To 6
Range(células(a) & "1").Select
    
    For i = 1 To 128
    
        If ActiveCell.Offset(1, 0).Value = "" Then
            ActiveCell.Copy
            ActiveCell.Offset(1, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Else
            ActiveCell.Offset(1, 0).Select
        End If
    Next

Next
End Sub


