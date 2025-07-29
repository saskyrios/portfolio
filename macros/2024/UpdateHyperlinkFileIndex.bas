Attribute VB_Name = "UpdateHyperlink"
Sub UpdateHyperlinks()
    Dim ws As Worksheet
    Dim hl As Hyperlink
    Dim newPath As String
    Dim folderDialog As Object
    Dim fso As Object
    Dim folder As Object
    Dim targetName As String
    Dim targetPath As String
    Dim isFolder As Boolean
    Dim indexFilePath As String
    Dim fileIndex As Object
    Dim textStream As Object
    Dim line As String
    Dim parts() As String
    Dim currentKey As Variant
    Dim fileExists As Boolean

    ' Создаем объект FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Выбор новой папки
    Set folderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    folderDialog.Title = "Выберите новую папку для гиперссылок"
    If folderDialog.Show = -1 Then
        newPath = folderDialog.SelectedItems(1) & "\"
    Else
        MsgBox "Папка не была выбрана. Макрос остановлен.", vbExclamation
        Exit Sub
    End If
    
    ' Определяем путь для файла индексации в той же папке, где находится Excel файл
    indexFilePath = ThisWorkbook.Path & "\FileIndex.txt"
    
    ' Если файл индексации не существует, создаем его
    If Not fso.fileExists(indexFilePath) Then
        Set textStream = fso.CreateTextFile(indexFilePath, True, True) ' Создаем файл индекса, Unicode
        MsgBox "Файл индексации был создан.", vbInformation
        textStream.Close ' Закрываем файл после создания
    End If
    
    ' Открываем файл индекса для чтения
    Set textStream = fso.OpenTextFile(indexFilePath, 1, False, -1) ' 1 = ForReading, -1 = Unicode
    
    ' Создаем коллекцию для индексации
    Set fileIndex = CreateObject("Scripting.Dictionary")
    
    ' Чтение индекса из файла
    Do While Not textStream.AtEndOfStream
        line = textStream.ReadLine
        parts = Split(line, ",")
        If UBound(parts) = 1 Then
            currentKey = LCase(parts(0))
            fileIndex.Add currentKey, parts(1)
        End If
    Loop
    textStream.Close ' Закрываем файл после чтения
    
    ' Индексируем все файлы и папки
    If fileIndex.Count = 0 Then
        Set folder = fso.GetFolder(newPath)
        Call IndexFilesAndFolders(folder, fileIndex, fso)
        
        ' Записываем индекс в файл
        Set textStream = fso.OpenTextFile(indexFilePath, 2, False, -1) ' 2 = ForWriting, -1 = Unicode
        For Each currentKey In fileIndex
            textStream.WriteLine currentKey & "," & fileIndex(currentKey)
        Next currentKey
        textStream.Close ' Закрываем файл после записи
    End If
    
    ' Проверка гиперссылок
    For Each ws In ThisWorkbook.Sheets
        Application.StatusBar = "Обработка листа: " & ws.Name
        For Each hl In ws.Hyperlinks
            Debug.Print "Проверка гиперссылки: " & hl.Address
            If Len(hl.Address) > 0 Then ' Проверяем, что гиперссылка содержит адрес
                targetName = LCase(fso.GetFileName(hl.Address))
                Debug.Print "Имя файла: " & targetName ' Проверяем, выводится ли имя файла
                
                isFolder = (Right(hl.Address, 1) = "\" Or fso.FolderExists(hl.Address))
                fileExists = False
                
                ' Проверяем индекс
                If fileIndex.Exists(targetName) Then
                    targetPath = fileIndex(targetName)
                    hl.Address = targetPath ' Обновляем гиперссылку на правильный путь
                    hl.Parent.Font.Color = vbBlack ' Возвращаем чёрный цвет рабочим ссылкам
                    fileExists = True
                    Debug.Print "Файл найден в индексе. Путь: " & targetPath
                Else
                    Debug.Print "Файл не найден в индексе."
                End If
                
                ' Если не найдено, помечаем гиперссылку красным цветом
                If Not fileExists Then
                    hl.Address = "" ' Очищаем адрес, если файл не найден
                    hl.Parent.Font.Color = vbRed
                    Debug.Print "Гиперссылка помечена красным."
                End If
            Else
                Debug.Print "Пустая гиперссылка."
            End If
        Next hl
    Next ws
    
    ' Очистка StatusBar
    Application.StatusBar = False
    
    MsgBox "Обновление гиперссылок завершено.", vbInformation
End Sub

Sub IndexFilesAndFolders(folder As Object, fileIndex As Object, fso As Object)
    Dim file As Object
    Dim subFolder As Object

    ' Индексируем файлы
    For Each file In folder.Files
        fileIndex.Add LCase(fso.GetFileName(file.Path)), file.Path
        ' Отладка: выводим информацию о файлах
        Debug.Print "Файл добавлен в индекс: " & fso.GetFileName(file.Path) & " -> " & file.Path
    Next file
    
    ' Индексируем папки рекурсивно
    For Each subFolder In folder.SubFolders
        Call IndexFilesAndFolders(subFolder, fileIndex, fso)
    Next subFolder
End Sub
