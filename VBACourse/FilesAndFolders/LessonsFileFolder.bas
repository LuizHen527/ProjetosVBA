Attribute VB_Name = "LessonsFileFolder"
'@Folder("VBAProject")
Option Explicit

'Verifica se um arquivo existe ou não. Se existe ele abre.
Sub FileExists()

    Dim fileName As String
    
    fileName = Dir("C:\Users\Molducolor7\Desktop\CursoVBA\S10_Looping_*Start.xls?")
    
    If fileName = vbNullString Then
    
        MsgBox "O arquivo não existe"
        
    Else
        Workbooks.Open fileName
    End If
    
End Sub

Sub Path_Exists()

    Dim path As String
    Dim folder As String
    Dim Answer As VbMsgBoxResult
    
    path = "C:\Users\Molducolor7\Desktop\CursoVBA\S12"
    
    folder = Dir(path, vbDirectory)
    
    If folder = vbNullString Then
        Answer = MsgBox("Essa pasta não existe. Você quer criar uma nova?", vbYesNo, "Criar nova pasta")
        
        Select Case Answer
            Case vbYes
                MkDir (path) 'Cria nova pasta. Make directory
            Case Else
                Exit Sub
        End Select
        
    Else
        MsgBox "A pasta existe."
    End If
    
End Sub

Sub GetDataFromFile()

    Dim FileToOpen As Variant
    Dim OpenWorkbook As Workbook
    
    Application.ScreenUpdating = False
    
    FileToOpen = Application.GetOpenFilename(Title:="Procure pelo arquivo e selecione o intervalo", FileFilter:="Excel Files(*.xls*),*xls*")

    If FileToOpen <> False Then
        Set OpenWorkbook = Workbooks.Open(FileToOpen)
        OpenWorkbook.Sheet(1).Range("A1:E20").Copy
        ThisWorkbook.Worksheets("SelectFile").Range("A10").PasteSpecial xlPasteValues
        OpenWorkbook.Close False
    End If
    
    Application.ScreenUpdating = True
End Sub

Sub LoopInsideFolder()
    Dim folderDir As String
    Dim fileToList As Variant
    Dim OpenBook As Workbook

    'Achamos o caminho da pasta
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Selecione uma pasta"
        .ButtonName = "Selecionar pasta"
        If .Show = 0 Then
            MsgBox "Nada foi selecionado"
        Else
            folderDir = .SelectedItems(1) & "\"
        End If
    End With
    
    'Achar os arquivos(Loop)
    fileToList = Dir(folderDir & "*xls*")
    
    Do Until fileToList = ""
    
        DoEvents
        
        Set OpenBook = Workbooks.Open(folderDir & fileToList)
        OpenBook.Sheets(1).Copy before:=ThisWorkbook.Worksheets(1)
        
        fileToList = Dir
        
        OpenBook.Close False
        
    Loop
    
End Sub

Sub SaveAsCSV()

    Dim NewBook As Workbook
    Dim fileName As Variant
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    fileName = Application.ThisWorkbook.path & "\TestTextCSV.csv"
    
    Set NewBook = Workbooks.Add
    Sheet8.Copy NewBook.Sheets(1)
    
    'No novo workbook, deleta as primeiras 3 linhas, salva como csv e fecha o workbook.
    With NewBook
        .Sheets(1).Rows("1:3").Delete
        .SaveAs fileName, Excel.xlCSV
        .Close
    End With
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Planilha exportada com sucesso!", vbInformation
End Sub

Sub SimpleTextFile()

    Dim fileName As String
    
    fileName = ThisWorkbook.path & "\TestTextFile.csv"
    
    Open fileName For Output As #1
        Print #1, "salve"
        Write #1, "salve"
        Print #1, 1
        Write #1, 1
        
        Write #1, Range("A1").Value
        Print #1, Range("A1").Value
    Close #1
End Sub

'Abre uma InputBox que permite o usuario inserir o nome do arquivo.
'A função abre e copia certos dados.
Sub Bonus_Get_Data_From_File_InputBox()
    
    Dim FileToOpen As Variant
    Dim OpenBook As Workbook
    Dim ShName As String
    Dim Sh As Worksheet
    On Error GoTo Handle:
    
    FileToOpen = Application.GetOpenFilename(Title:="Browse for your File & Import Range", FileFilter:="Excel Files (*.xls*),*.xls*")
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    If FileToOpen <> False Then
        Set OpenBook = Application.Workbooks.Open(FileToOpen)
        ShName = Application.InputBox("Enter the sheet name to copy", "Enter the sheet name to copy")
        For Each Sh In OpenBook.Worksheets
            If UCase(Sh.Name) Like "*" & UCase(ShName) & "*" Then
                ShName = Sh.Name
            End If
        Next Sh

        'copy data from the specified sheet to this workbook - updae range as you see fit
        OpenBook.Sheets(ShName).Range("A1:CF1100").Copy
        ThisWorkbook.ActiveSheet.Range("A10").PasteSpecial xlPasteValues
        OpenBook.Close False
    End If
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
Handle:
    If Err.Number = 9 Then
        MsgBox "The sheet name does not exist. Please check spelling"
    Else
        MsgBox "An error has occurred."
    End If
    OpenBook.Close False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
