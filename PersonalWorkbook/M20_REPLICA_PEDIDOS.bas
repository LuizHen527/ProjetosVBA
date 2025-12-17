Attribute VB_Name = "M20_REPLICA_PEDIDOS"
Sub Replicar_Pedidos()

'Replica os pedidos removidos
'Application.ScreenUpdating = False

'Sheets("Macro").Select
    last_line = Cells(Cells.Rows.Count, 1).End(xlUp).Row

    Columns("F:F").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("E:E").Select
        Selection.Copy
    Range("F1").Select
        ActiveSheet.Paste
    Range("F1").Select
        ActiveCell.FormulaR1C1 = "Pedido_2"

    Range("F" & last_line).Select
        Do While ActiveCell.Offset(-1, 0).Value <> "Pedido_2"
            If ActiveCell.Offset(-1, 0).Value = "" Then
                ActiveCell.Copy
                ActiveCell.Offset(-1, 0).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            Else: ActiveCell.Offset(-1, 0).Select
            End If
        Loop

End Sub


