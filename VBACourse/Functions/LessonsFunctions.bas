Attribute VB_Name = "LessonsFunctions"
Option Explicit

Sub Msg_Box()

    'MsgBox "Oi " & Application.UserName, , "Bem-vindo!"
    
    MsgBox prompt:="Oi " & Application.UserName & "." & VBA.Constants.vbNewLine & "Obrigado por parar por aqui", _
    Title:="Seja bem-vindo"

End Sub

Sub VBA_Excel_Functions()

    With ShVEF
        .Range("B3").Value = VBA.Date
        .Range("B6").Value = VBA.UCase(.Range("A6"))
        .Range("B7").Value = VBA.LCase(.Range("A7"))
        .Range("B8").Value = Excel.WorksheetFunction.Proper(.Range("A9"))
        .Range("B9").Value = VBA.StrConv(.Range("A9"), vbProperCase)
        
    End With
    
    Dim numSet As Range
    
    Set numSet = Range("B17").CurrentRegion
    
    Range("B11").Value = Excel.WorksheetFunction.Max(numSet)
    
End Sub

Sub VBA_Functions()

    With Sheet4
        .Range("B3").Value = VBA.Month(VBA.Date)
        .Range("B4").Value = VBA.MonthName(VBA.Month(VBA.Date))
        .Range("C5").Value = VBA.MonthName(VBA.Month(VBA.Date), True)
        
        .Range("B9").Value = IsEmpty(.Range("A9"))
        .Range("B10").Value = IsEmpty(.Range("A10"))
        .Range("B11").Value = IsEmpty(.Range("A11"))
        
    End With
End Sub

Sub VBA_InputBox()

    Dim CName As String
    Dim LastRow As Long
    
    CName = InputBox("Coloque o nome do cliente", "Clientes")
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row + 1
    Range("A" & LastRow).Value = Excel.WorksheetFunction.Proper(CName)
    
End Sub

Sub Yes_No_MsgBox()

Dim Answer As VbMsgBoxResult

Answer = MsgBox("Tem certeza que quer deletar tudo?", vbYesNo + vbQuestion + vbDefaultButton1, "Limpa celulas")

If Answer = vbYes Then
    Range("A7:B9").Clear
Else
    Exit Sub
End If


End Sub

Sub VBA_Simple_InputBox()

    Dim myImp As String
    myImp = VBA.InputBox("Coloque um subtitulo", "Subtitulo...")
    If myImp = "" Then Exit Sub
    
    Range("A2") = Excel.WorksheetFunction.Proper(myImp)
    
End Sub

Sub Excel_InputBox()

    Dim CName As String
    Dim LastRow As Long
    Dim CAmount As Long
    
    CName = InputBox("Coloque o nome do cliente", "Clientes")
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row + 1
    Range("A" & LastRow).Value = Excel.WorksheetFunction.Proper(CName)
    
    CAmount = Excel.Application.InputBox("Coloque a quantidade", "Qunatidade", Type:=1)
    Cells(LastRow, 2).Value = CAmount
End Sub

Sub Excel_InputBox_Range()

    Dim myRange As Range
    Dim cellBlank As Long, cellNum As Long, cellOther As Long
    On Error GoTo Leave
    
    Set myRange = Application.InputBox("Conte o numero de celulas", "Verifica celulas", , , , , , 8)
    
    cellBlank = Excel.WorksheetFunction.CountBlank(myRange)
    cellNum = Excel.WorksheetFunction.Count(myRange)
    cellOther = Excel.WorksheetFunction.CountA(myRange) - cellNum
    MsgBox cellBlank & "cells are blank" & vbNewLine & _
    cellNum & "cells have numbers." & vbNewLine _
    & cellOther & "cells are non-numeric"
Leave:
End Sub

Sub Activity()

    Dim myRange As Range
    Dim top1 As Double, top2 As Double, top3 As Double
     
    On Error GoTo Leave
    Set myRange = Application.InputBox("Selecione o intervalo de numeros", "Mostrar tres maiores numeros", , , , , , 8)
    
    If Application.WorksheetFunction.Count(myRange) > 2 Then
    
        top1 = Excel.WorksheetFunction.Large(myRange, 1)
        top2 = Excel.WorksheetFunction.Large(myRange, 2)
        top3 = Excel.WorksheetFunction.Large(myRange, 3)
        
        MsgBox "Primeiro maior: " & top1 & vbNewLine & "Segundo maior: " _
        & top2 & vbNewLine & "Terceiro maior: " & top3
    Else
        MsgBox "Selecione pelo menos 3 celulas com numeros", vbInformation, "Erro"
        
    End If
    
Leave:
End Sub
