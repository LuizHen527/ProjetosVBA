Attribute VB_Name = "AtualizarPerfisPedido"
'@Folder("VBAProject")
Option Explicit

Sub AtualizaPerfisPedido()
    Dim perfisPorPedido() As String
    
    ThisWorkbook.Sheets("perfis_contagem").Activate
    
    perfisPorPedido = PegarPerfisDoPedido
    
    ThisWorkbook.Sheets("perfis_pedido").Activate
      
    InserirDados perfisPorPedido
    
End Sub

Function InserirDados(perfis() As String)
    Dim i As Integer, row As Integer
    Dim initialRange As String
    
    ThisWorkbook.Sheets("perfis_pedido").Activate
    
    initialRange = Range("A" & Cells(Rows.Count, 1).End(xlUp).row).Offset(1, 0).Address
    
    row = 0
    
    For i = 0 To UBound(perfis)
        
        If perfis(i, 0) <> "" Then
            
            'NUMERO
            Range(initialRange).Offset(row, 0) = perfis(i, 0)
            
            'NOME PERFIL
            Range(initialRange).Offset(row, 1) = perfis(i, 1)
            
            'COR
            Range(initialRange).Offset(row, 2) = perfis(i, 2)
            
            'QUANTIDADE
            Range(initialRange).Offset(row, 3) = perfis(i, 3)
            
            'STATUS
            Range(initialRange).Offset(row, 4) = perfis(i, 4)
            
            'DATA
            Range(initialRange).Offset(row, 5) = perfis(i, 5)
            
            row = row + 1
        End If
        
    Next i
    
End Function


Function PegarPerfisDoPedido() As String()
    Dim cliente As String, pedidoNumero As String, data As String
    Dim perfis() As String
    Dim iterator As Integer
    Dim rng As Range, rgnRow As Range

    'Pegar nome do perfil
    
    'Pegar Cliente, Numero e Data
    cliente = Range("A3").Value
    pedidoNumero = CStr(Range("B3").Value)
    data = CStr(Range("C3").Value)
    
    ReDim perfis(200, 5)
    
    iterator = 0
    
    'Loopar pela coluna de perfis pegando apenas os nomes de perfis
    For Each rng In Range("A6:A" & Cells(Rows.Count, 1).End(xlUp).row)
        
        If rng.Value <> "SITUAÇÃO" Then
            
            For Each rgnRow In Range(rng.Offset(0, 1), Cells(rng.row, Columns.Count).End(xlToLeft))
            
                If Cells(5, rgnRow.Column).Value <> "TOTAL" And _
                rgnRow.Value > 0 Then
                
                    'NUMERO PEDIDO
                    perfis(iterator, 0) = pedidoNumero
                    
                    'NOME PERFIL
                    perfis(iterator, 1) = rng.Value
                    
                    'COR
                    perfis(iterator, 2) = Cells(5, rgnRow.Column).Value
                    
                    'QUANTIDADE
                    perfis(iterator, 3) = rgnRow.Value
                    
                    If rgnRow.Offset(1, 0).Value = "N" Then
                    
                        'STATUS
                        perfis(iterator, 4) = "PRODUZIR"
                        
                    Else
                    
                        'STATUS
                        perfis(iterator, 4) = "EM ESTOQUE"
                        
                    End If
                    
                    'DATA
                    perfis(iterator, 5) = Date
                    
                    iterator = iterator + 1
                End If
            Next rgnRow
        End If
    Next rng
    
    'Return
    PegarPerfisDoPedido = perfis
    
    
End Function
