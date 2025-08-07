Attribute VB_Name = "AtualizarPerfisPedido"
'@Folder("VBAProject")
Option Explicit

Sub AtualizaPerfisPedido()

    PegarPerfisDoPedido
End Sub

Function PegarPerfisDoPedido() As String()
    Dim cliente As String, pedidoNumero As String, data As String
    Dim rng As Range

    'Pegar nome do perfil
    
    'Pegar Cliente, Numero e Data
    cliente = Range("A3").Value
    pedidoNumero = CStr(Range("B3").Value)
    data = CStr(Range("C3").Value)
    
    
    'Loopar pela coluna de perfis pegando apenas os nomes de perfis
    For Each rng In Range("A6:A" & Cells(Rows.Count, 1).End(xlUp).row)
        Debug.Print rng.Value
    Next rng
    
    
End Function
