Attribute VB_Name = "FazIndiceConteudos"
'@Folder("VBAProject")
Option Explicit

Sub Auto_Table_Contents()

    'Mostrar uma input box pedindo uma celula
    Dim sh As Worksheet
    Dim InitialCell As Range
    Dim EndCell As Range
    Dim ShName As String
    Dim MsgConfirm As VBA.VbMsgBoxResult
    
    
    
    'Quando tem um erro ele continua
    On Error Resume Next
    Set InitialCell = Application.InputBox("Onde quer colocar a tabela?" & vbNewLine & "Selecione uma celula:" _
    , "Tabela automatica", , , , , , 8)
    If Err.Number = 424 Then Exit Sub
    
    On Error GoTo Handle
    Set InitialCell = InitialCell.Cells(1, 1)
    Set EndCell = InitialCell.Offset(Worksheets.Count - 2, 1)
    
    MsgConfirm = MsgBox("O programa vai usar as celulas" & InitialCell.Address & _
    " Até a " & EndCell.Address & ". Quer continuar?", vbOKCancel _
    , "Confirmar")
    If MsgConfirm = vbCancel Then Exit Sub
    
    For Each sh In Worksheets
        ShName = sh.Name
        If ActiveSheet.Name <> ShName And sh.Visible = xlSheetVisible Then
            ActiveSheet.Hyperlinks.Add Anchor:=InitialCell, Address:="", SubAddress:= _
            "'" & ShName & "'!A1", TextToDisplay:=ShName
            
            InitialCell.Offset(0, 1) = sh.Range("A1").Value
            Set InitialCell = InitialCell.Offset(1, 0)
        End If
    
    Next sh
    
    Exit Sub
    
Handle:
    MsgBox "Um erro inesperado ocorreu.", vbCritical, "Erro"
    
End Sub


