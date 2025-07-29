VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormInsertFileNamesHyperlinks 
   Caption         =   "Создание гипперсилок на документы"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "FormInsertFileNamesHyperlinks.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormInsertFileNamesHyperlinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Dim folderPath As String
    Dim cell As Range
    Dim fileName As String
    Dim i As Integer

    ' Получение пути к папке
    folderPath = TextBox1.Value

    ' Проверка на наличие пути к папке
    If folderPath = "" Then
        MsgBox "Путь к папке не указан. Пожалуйста, введите путь.", vbExclamation
        Exit Sub
    ElseIf Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Указанный путь не существует. Пожалуйста, введите правильный путь.", vbExclamation
        Exit Sub
    End If

    ' Проверка на наличие обратного слеша в конце пути
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    ' Получение начальной ячейки
    On Error Resume Next
    Set cell = Range(RefEdit1.Text)
    On Error GoTo 0

    ' Проверка на правильность выбора ячейки
    If cell Is Nothing Then
        MsgBox "Указан некорректный адрес ячейки. Пожалуйста, выберите правильную ячейку.", vbExclamation
        Exit Sub
    End If

    ' Перебор всех файлов в папке
    fileName = Dir(folderPath & "*.*")
    i = 0
    Do While fileName <> ""
        cell.Offset(i, 0).Value = fileName
        cell.Offset(i, 0).Hyperlinks.Add Anchor:=cell.Offset(i, 0), _
            Address:=folderPath & fileName, _
            TextToDisplay:=fileName
        fileName = Dir
        i = i + 1
    Loop

    ' Сообщение о завершении
    MsgBox "Гиперссылки на файлы добавлены успешно!", vbInformation
    
    ' Закрытие формы
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    ' Закрытие формы без выполнения действий
    Unload Me
End Sub

Private Sub CommandButton3_Click()
    ' Открытие окна выбора папки
    Dim folderPath As FileDialog
    Set folderPath = Application.FileDialog(msoFileDialogFolderPicker)
    folderPath.Title = "Выберите папку с файлами"
    
    ' Проверка, выбрал ли пользователь папку
    If folderPath.Show = -1 Then
        TextBox1.Text = folderPath.SelectedItems(1)
    Else
        MsgBox "Папка не выбрана. Пожалуйста, выберите папку.", vbExclamation
    End If
End Sub

