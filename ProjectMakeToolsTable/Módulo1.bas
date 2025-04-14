Attribute VB_Name = "Módulo1"
Option Explicit

Sub Batata()
Attribute Batata.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Batata Macro
'

'
    Range(Cells(1, Columns.Count).End(xlToLeft)).Select
    Range("A1", Cells(1, Columns.Count).End(xlToLeft)).Select
    Range("BO2").Select
    Selection.AutoFill Destination:=Range("BO2:BO52"), Type:=xlFillDefault
    
End Sub
