Attribute VB_Name = "LoopLessonB"
'@Folder("VBAProject")
Option Explicit
Dim StartCell As Integer

Sub Simple_Do_Until_V1()

    StartCell = 8
    
    Do Until Range("A" & StartCell).Value = ""
        Range("B" & StartCell).Value = Range("A" & StartCell).Value + 10
        StartCell = StartCell + 1
    Loop
End Sub

Sub Simple_Do_Until_V2()

    StartCell = 8
    
    Do While Range("A" & StartCell).Value <> ""
        Range("C" & StartCell).Value = Range("A" & StartCell).Value + 10
        StartCell = StartCell + 1
    Loop
End Sub

Sub Simple_Do_Until_V3()

    StartCell = 8
    
    Do Until Range("A" & StartCell).Value = ""
        If Range("A" & StartCell) = 0 Then Exit Do
        Range("D" & StartCell).Value = Range("A" & StartCell).Value + 10
        StartCell = StartCell + 1
    Loop
End Sub

Sub DoLoopExample()

    Dim answer As String
    
    Do While Not IsNumeric(answer)
        answer = InputBox("Digite um numero")
        If IsNumeric(answer) Then MsgBox ("Muito bem!")
    Loop

End Sub

Sub Find()
    
    'Onde Salvar, Salvar, onde colar
    Dim CompId As Range
    Range("D3:D6").ClearContents
    Set CompId = Range("A:A").Find(what:=Range("B3").Value, LookIn:=xlValues, lookat:=xlWhole)
    If CompId Is Nothing Then
        MsgBox "Empresa não existe"
        Exit Sub
    End If
    
    Range("C3").Value = CompId.Offset(, 4).Value
End Sub

Sub Find_ManyMatches()
    
    'Onde Salvar, Salvar, onde colar
    Dim CompId As Range
    Dim i As Byte
    Dim firstAddress As Variant
    Dim start
    
    Range("D3:D6").ClearContents
    i = 3
    
    start = Timer
    
    
    Set CompId = Range("A:A").Find(what:=Range("B3").Value, LookIn:=xlValues, lookat:=xlWhole)
    
    If CompId Is Nothing Then
        MsgBox "Empresa não existe"
    Else
        Range("D" & i).Value = CompId.Offset(, 4).Value
        firstAddress = CompId.Address
        Do
            Set CompId = Range("A:A").FindNext(CompId)
            
            If CompId.Address = firstAddress Then Exit Do
            
            i = i + 1
            
            Range("D" & i).Value = CompId.Offset(, 4).Value
        Loop
    End If
    
    Debug.Print Round(Timer - start, 3)
    'Application.Speech.Speak "Muito bem!" & i - 2 & " correspondências foram encontradas"
    
End Sub
