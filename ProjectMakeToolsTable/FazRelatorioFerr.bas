Attribute VB_Name = "FazRelatorioFerr"
Option Explicit

Sub RelatorioFerramentas()
'Ao usar a macro:
    'deixe a planilha que voc� quer copiar selecionada
    'a planilha precisa estar com esse formato de nome:
    'Mes_Numero do mes_Utimos dois digitos do ano
    'Deve ficar assim: Mar_3_25
    
    'Antes precisa corrigir os nomes
    
'------------------------------------------------------

'Fazer a coluna de nomes
    'Colocar "PERFIL" no A1
    'Fazer loop que passa pelos nomes
        'Pegar o nome e numero
        'Comparar com a lista de nomes que j� foi colocada. Ver se aquela ferramenta com aquele numero j� n�o foi colocada
            'Se sim, pula pra proxima. Se n�o, cola ela e as informa��es.
        'Pega o nome da empresa, numero da ferramenta, e o que precisa
        'Cola tudo na tabela de ferramentas
'------------------------------------------------------

    Dim data() As Variant
    Dim fileName As String, arrDate() As String
    Dim numRows As Integer, colArray As Integer, rowArray As Integer
    
    fileName = ThisWorkbook.Name
    arrDate = Split(ActiveSheet.Name, "_")
    
    Workbooks("HIST�RICO PRODU��O 2022-2024_V5.xlsm").Activate
    Worksheets("01_Base").Select
    
    'Tira filtros aplicados
    ActiveWorkbook.Worksheets("01_Base").AutoFilter.Sort.SortFields.Clear
    ActiveSheet.ShowAllData
 
    'Filtra os dados da base pela data, de acordo com o nome da planilha(Ex:Mar_1_25)
    ActiveSheet.Range("$A$3:$BA$4805").AutoFilter Field:=1, Operator:= _
    xlFilterValues, Criteria2:=Array(1, arrDate(1) & "/10/20" & arrDate(2))
    
    '-------------------------------- SALVANDO DADOS NO ARRAY --------------------------------
    
    'Conta linha visiveis
    numRows = Range("A3", "A" & Cells(Rows.Count, 1).End(xlUp).Row).Rows.SpecialCells(xlCellTypeVisible).Count
    
    ReDim data(numRows, 4) As Variant
    
    
    
    'Loop para colocar o nomes das ferramentas
    'Tem alguma fun��o que n�o coloca duplicadas?
    'Se n�o tiver:
        'Coloca nome da primeira ferramenta.
        'Salva o nome da ferramenta em um array de ferramentas copiadas
        'Pega o nome da proxima, compara com as ferramentas j� copiadas
            'Se j� tiver uma igual, ele pula.
            'Se n�o tiver, ele cola.
    
    'Loop para colocar dados nas colunas
    'Salvar nome, numero e empresa(Ferramentas). Salvar em um array.
    'Ir para planilha de ferramentas
    'Percorrer o array incerindo dados coluna por coluna
    'Se j� tiver, pula. Se n�o, copia.
    
    
    
    
End Sub
