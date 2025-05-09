Attribute VB_Name = "C01_Compilar"
Option Base 1
Sub Atualizar_Pedidos()

'Esconde botões do menu
ActiveSheet.Shapes.Range(BarShapes).Visible = Not ActiveSheet.Shapes.Range(BarShapes).Visible
ActiveSheet.Shapes.Range(ButtonShapes).Visible = Not ActiveSheet.Shapes.Range(ButtonShapes).Visible

'Att em 22/06/22
Application.ScreenUpdating = False
Application.DisplayAlerts = False

'Identificação das variáveis
Dim compras_pasta, compras_arquivo As Object
Dim nome_arquivo, planilha_base, planilha_pedidos As String
'Dim i, j, numero_linhas As Integer
Dim pedidos_cadastrados, pedidos_solicitados, pedidos_novos As Variant

'Define local dos arquivos de compras
Set compras_pasta = CreateObject("Scripting.fileSystemObject").getfolder("\\121.137.1.5\alumitec9\COMPRAS\25_Compras")

'Define nome da planilha Base como variável
planilha_base = ActiveWorkbook.Name

pedidos_cadastrados = Array("")
pedidos_solicitados = Array("")
pedidos_novos = Array("")

'Atualiza os dados da planilha backup com as informações _
anteriores ao processo de adição dos novos pedidos
Sheets("BASE").Select

'REVER ESSA PARTE
'--------------------------------------------------------
ActiveWorkbook.Worksheets("BASE").AutoFilter.Sort.SortFields.Clear
ActiveSheet.ShowAllData
'----------------------------------------------------

'Limpa Valores de AA01
Sheets("BASE").Select
    Range("AAO1").Select
        Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents

'Copia o nome dos pedidos e remove os pedidos duplicados
Range("A2").Select
    If Range("A2").Value <> "" And Range("A3").Value <> "" Then
        Range(Selection, Selection.End(xlDown)).Select
            numero_linhas = Selection.Count
            Selection.Copy
            Range("AAO1").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                        ActiveSheet.Range("$AAO$1:$AAO" & numero_linhas).RemoveDuplicates Columns:=1, Header:=xlNo
    Else
        numero_linhas = 1
        Selection.Copy
            Range("AAO1").Select
                ActiveSheet.Paste
    End If

'Verifica o número de pedidos cadastrados
If Range("AAO1").Value <> "" And Range("AAO2").Value = "" Then
    numero_linhas = 1
ElseIf Range("AAO2").Value <> "" Then
    Range("AAO1").Select
        Range(Selection, Selection.End(xlDown)).Select
            numero_linhas = Selection.Count
End If

'Add os pedidos já cadastrados em um array
For i = 1 To numero_linhas:
   ReDim Preserve pedidos_cadastrados(i)
      pedidos_cadastrados(i) = Range("AAO" & i) & ".xlsm"
Next

'Corrige Variáveis
i = i - 1
j = 1

'Add os nomes dos pedidos solicitados em um array
For Each compras_arquivo In compras_pasta.Files
    ReDim Preserve pedidos_solicitados(j)
       pedidos_solicitados(j) = compras_arquivo.Name
    j = j + 1
Next
    j = j - 1

'Quantidade de pedidos cadastrados
num_pedidos_cadastrados = UBound(pedidos_cadastrados)
'MsgBox num_pedidos_cadastrados

'Quantidade de pedidos solicitados
num_pedidos_solicitados = UBound(pedidos_solicitados)
'MsgBox num_pedidos_solicitados

'numero de pedidos adicionados
num_pedidos_novos = 1

'Add os pedidos novos ("pedidos cadastrados - pedidos solicitados") em um array
For k = 1 To num_pedidos_solicitados
    n = 0
    For l = 1 To num_pedidos_cadastrados
    
        'pedido solicitado = nome do arquivo de pedidos
        'Confere se esse arquivo não é um desses abaixo
        If pedidos_solicitados(k) = "00_BASE COMPRAS V7.xlsm" Or _
            pedidos_solicitados(k) = "00_SOLICITAÇÃO COMPRAS.xlsm" Or _
            pedidos_solicitados(k) = "00_SOLICITAÇÃO DE COMPRAS_25.xlsm" Or _
            pedidos_solicitados(k) = "Thumbs.db" Or _
            pedidos_solicitados(k) = "~$00_SOLICITAÇÃO DE COMPRAS_25.xlsm" Or _
            pedidos_solicitados(k) = "~$00_SOLICITAÇÃO COMPRAS.xlsm" Or _
            pedidos_solicitados(k) = "~$00_BASE COMPRAS V7.xlsm" Then
            
            Exit For
            
        ElseIf pedidos_solicitados(k) = pedidos_cadastrados(l) Then
            
            Exit For
        
        Else
            n = n + 1
        End If

        If n >= num_pedidos_cadastrados Then
        MsgBox ("PEDIDO NOVO " & "//" & pedidos_solicitados(k))
            ReDim Preserve pedidos_novos(num_pedidos_novos)
                pedidos_novos(num_pedidos_novos) = pedidos_solicitados(k)
                num_pedidos_novos = num_pedidos_novos + 1
        End If
    Next
Next

'---------------------------------------------------------------------
  
'Acessa apenas os pedidos listados em pedidos_novos e copia e cola na planilha resumo
num_pedidos_novos = 1

On Error GoTo panda

If pedidos_novos(num_pedidos_novos) <> "" Then
    For i = 1 To UBound(pedidos_novos):
        Workbooks.Open Filename:=compras_pasta & "\" & pedidos_novos(i)
    
        planilha_pedidos = ActiveWorkbook.Name
        Base = planilha_pedidos
        Sheets("Compilado").Select
            
            For y = 3 To 50
                Range("A" & y).Select
                If ActiveCell.Value <> "" Then
                    ActiveCell.Offset(1, 0).Select
                    Else
                        Exit For
                End If
            Next
            'j = numero_linhas do range a ser selecionado na planilha
            w = y - 1
            
        Range("A3:K" & w).Select
            Selection.Copy
            Windows(planilha_base).Activate
                Range("A" & 10000).Select
                    Selection.End(xlUp).Select
                    ActiveCell.Offset(1, 0).Select
                        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                        :=False, Transpose:=False
                                Windows(Base).Activate
                                    If Base = "00_BASE COMPRAS V5.xlsm" Then
                                        Exit For
                                    Else
                                        ActiveWorkbook.Close
                                    End If
    Next
End If

Range("AAO1").Select
    Range(Selection, Selection.End(xlDown)).Select
        Selection.Clear

ActiveSheet.Range("$A$2:$AE$6000").AutoFilter

'Call C03_BOTOES.Delete

Range("A2").Select

If pedidos_novos(num_pedidos_novos) = "" Then
    MsgBox "Nenhum pedido novo"
Else
    MsgBox "Dados Atualizados com Sucesso"
End If


Exit Sub

panda:
    MsgBox ("ERROR")

End Sub



