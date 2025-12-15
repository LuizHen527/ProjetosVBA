Attribute VB_Name = "Módulo3"
Option Explicit

Sub Papagaio()
Attribute Papagaio.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Papagaio Macro
'

'
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    Range("Tabela3[[#Headers],[SITUAÇÃO]]").Select
    ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=10, Criteria1:= _
        "EM ABERTO"
    ActiveWindow.SmallScroll Down:=60
End Sub
