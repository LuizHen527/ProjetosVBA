Attribute VB_Name = "Controle_Acesso"
Sub Entrada()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

'Sheets("Molducolor").Select
'    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
'Sheets("Pollux").Select
'    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True

Sheets("ACESSO").Visible = True
Sheets("ACESSO").Select
    Range("A1048576").Select
        Selection.End(xlUp).Select
        ActiveCell.Offset(1, 0).Select
            ActiveCell.Value = Environ("Username")
            ActiveCell.Offset(0, 1).Value = Date & " - " & Time

Sheets("ACESSO").Visible = False
'Sheets("Molducolor").Select


End Sub
 
Sub Saída()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Sheets("ACESSO").Visible = True
Sheets("ACESSO").Select

Range("A1048576").Select
        Selection.End(xlUp).Select
        ActiveCell.Offset(0, 2).Value = Date & " - " & Time

Sheets("ACESSO").Visible = False
'Sheets("Molducolor").Select
        

End Sub
