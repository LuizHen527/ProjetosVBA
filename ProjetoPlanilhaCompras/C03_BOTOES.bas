Attribute VB_Name = "C03_BOTOES"
Public ButtonShapes()
Public BarShapes()

'Sub Menu()
'
'Application.DisplayAlerts = False
'
'On Error GoTo panda
'
'Sheets("BASE").Shapes("M00_MENU").Select
'    Call Delete
'
'If 1 > 2 Then
'panda:
'    Call Botao
'End If
'
'End Sub
'
'Sub Botao()
'
'Application.DisplayAlerts = False
'
'Sheets("BASE").Select
'
''Puxa o bloco dos kits especificos
'Sheets("BOT�ES").Shapes("M00_MENU").Copy
'    Range("A2").Select
'    ActiveSheet.Paste
'Range("A1").Select
'
'
'End Sub

Sub Delete()

Application.DisplayAlerts = False

Sheets("BASE").Select

On Error Resume Next

ActiveSheet.Shapes("M00_MENU").Delete

End Sub

Sub MenuButton()

    'Quando eu apertar
    'Troca o icone do bot�o menu
    
    'Come�a visivel (true)
    BarShapes = Array("IconBar1", "IconBar2", "IconBar3")
    
    'Come�a escondido (false)
    ButtonShapes = Array("UpdateButton", "FilterButton", "CleanButton", "IconClose", "FinanButton", "ClassiButton", "AcompButton")
    
    ActiveSheet.Shapes.Range(BarShapes).Visible = Not ActiveSheet.Shapes.Range(BarShapes).Visible
    
    ActiveSheet.Shapes.Range(ButtonShapes).Visible = Not ActiveSheet.Shapes.Range(ButtonShapes).Visible
    
    'Aparece e desaparce os bot�es debaixo
    
End Sub

Sub SetNotVisible()
    Dim ButtonShapes()
    
    'Come�a hidden
    ButtonShapes = Array("UpdateButton", "FilterButton", "IconClose", "CleanButton")
    
    ActiveSheet.Shapes.Range(ButtonShapes).Visible = Not ActiveSheet.Shapes.Range(ButtonShapes).Visible

End Sub

