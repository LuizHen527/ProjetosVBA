Attribute VB_Name = "FazRelatorioFerr"
Option Explicit

Sub RelatorioFerramentas()
'Antes precisa corrigir os nomes

'Fazer a coluna de nomes
    'Colocar "PERFIL" no A1
    'Fazer loop que passa pelos nomes
        'Pegar o nome e numero
        'Comparar com a lista de nomes que j� foi colocada. Ver se aquela ferramenta com aquele numero j� n�o foi colocada
            'Se sim, pula pra proxima. Se n�o, cola ela e as informa��es.
        'Pega o nome da empresa, numero da ferramenta, e o que precisa
        'Cola tudo na tabela de ferramentas
'------------------------------------------------------

    Dim fileName As String
    Dim arrDate() As String
    
    fileName = ThisWorkbook.Name
    arrDate = Split(ActiveSheet.Name, "_")
    
    Workbooks(fileName).Activate
    
    Workbooks("HIST�RICO PRODU��O 2022-2024_V5.xlsm").Activate
    Worksheets("01_Base").Select

    
    ActiveSheet.Range("$A$3:$BA$4805").AutoFilter Field:=1, Operator:= _
    xlFilterValues, Criteria2:=Array(1, arrDate(1) & "/10/20" & arrDate(2))
    
    'Criar loop
    'Salvar nome, numero e empresa(Ferramentas)
    'Comparar nome e numero copiados com os que est�o l�
    'Se j� tiver, pula. Se n�o, copia.
    
    
    
    
End Sub
