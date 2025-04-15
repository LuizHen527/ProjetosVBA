Attribute VB_Name = "Módulo2"
Option Explicit

Sub Frango()
Attribute Frango.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Frango Macro
'

'
    With Range("A79:C79")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

End Sub
Sub Berinjela()
Attribute Berinjela.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Berinjela Macro
'

'
    Range("A93:C93").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
End Sub
