Attribute VB_Name = "IndexAndCreateHyperlink"
Option Explicit

' Определение констант для режима открытия файлов
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

' Основная процедура для обработки файлов
Sub ProcessFiles()
    Dim ws As Worksheet
    Dim indexFolderPath As String ' Папка индексации (куда копируем файлы)
    Dim searchFolderPath As String ' Папка поиска недостающих файлов (откуда копируем файлы)
    Dim useCopyFunctionality As Boolean
    Dim fileNamesDictDestination As Object ' Словарь файлов в папке индексации
    Dim fileNamesDictSource As Object ' Словарь файлов в папке поиска
    Dim fileName As String
    Dim currentRow As Long
    Dim version As Integer
    Dim foundFiles As Collection
    Dim file As Variant
    Dim currentVersion As Integer
    Dim maxVersion As Integer
    Dim indexFilePathDestination As String
    Dim indexFilePathSource As String
    Dim fso As Object
    Dim lastRow As Long
    Dim dlg As FileDialog
    Dim userResponse As VbMsgBoxResult
    Dim excelFilePath As String
    Dim destinationPath As String

    Set ws = ActiveSheet
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Получаем путь к папке, где находится файл Excel
    excelFilePath = ThisWorkbook.Path & "\"

    ' Выбор папки для сканирования (папка индексации)
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    dlg.Title = "Выберите папку для сканирования (папка индексации)"
    If dlg.Show = -1 Then
        indexFolderPath = dlg.SelectedItems(1) & "\"
    Else
        MsgBox "Путь к папке индексации не выбран. Макрос завершен."
        Exit Sub
    End If

    ' Получаем максимальное количество версий через диалог
    maxVersion = InputBox("Введите максимальное количество версий", "Настройка", 10)
    If maxVersion = 0 Then
        MsgBox "Максимальное количество версий должно быть больше нуля.", vbExclamation
        Exit Sub
    End If

    ' Спрашиваем пользователя о необходимости использования функционала копирования недостающих файлов
    userResponse = MsgBox("Хотите ли вы использовать функционал поиска и копирования недостающих файлов?" & vbCrLf & _
                          "Если да, вам будет предложено выбрать папку для поиска." & vbCrLf & _
                          "Обратите внимание, что некоторые файлы будут скопированы в папку индексации.", _
                          vbYesNo + vbQuestion, "Использовать поиск и копирование файлов?")

    If userResponse = vbYes Then
        useCopyFunctionality = True

        ' Выбор папки для поиска недостающих файлов
        Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
        dlg.Title = "Выберите папку для поиска недостающих файлов"
        If dlg.Show = -1 Then
            searchFolderPath = dlg.SelectedItems(1) & "\"
        Else
            MsgBox "Путь к папке поиска не выбран. Макрос завершен."
            Exit Sub
        End If
    Else
        useCopyFunctionality = False
    End If

    ' Определяем пути к файлам индексации (в папке с файлом Excel)
    indexFilePathDestination = excelFilePath & "index_destination.txt"
    If useCopyFunctionality Then
        indexFilePathSource = excelFilePath & "index_source.txt"
    End If

    ' Проверяем, нужно ли обновлять файл индексации для папки индексации
    If Not IsIndexUpToDate(indexFolderPath, indexFilePathDestination, fso) Then
        CreateIndexFile indexFolderPath, indexFilePathDestination, fso
    End If

    ' Загружаем данные из файла индексации в словарь для папки индексации
    Set fileNamesDictDestination = ReadIndexFileToDictionary(indexFilePathDestination, fso)
    If fileNamesDictDestination Is Nothing Then
        MsgBox "Ошибка при загрузке файла индексации для папки индексации.", vbCritical
        Exit Sub
    End If

    ' Если используется функционал копирования, работаем с индексом папки поиска
    If useCopyFunctionality Then
        ' Проверяем, нужно ли обновлять файл индексации для папки поиска
        If Not fso.FileExists(indexFilePathSource) Then
            ' Если файла нет, создаем его
            CreateIndexFile searchFolderPath, indexFilePathSource, fso
        Else
            ' Спросить пользователя о перезаписи
            userResponse = MsgBox("Файл индексации для папки поиска уже существует. Пересканировать папку и обновить индекс?", vbYesNo + vbQuestion, "Обновить индекс?")
            If userResponse = vbYes Then
                CreateIndexFile searchFolderPath, indexFilePathSource, fso
            End If
        End If

        ' Загружаем данные из файла индексации в словарь для папки поиска
        Set fileNamesDictSource = ReadIndexFileToDictionary(indexFilePathSource, fso)
        If fileNamesDictSource Is Nothing Then
            MsgBox "Ошибка при загрузке файла индексации для папки поиска.", vbCritical
            Exit Sub
        End If
    End If

    ' Определяем последнюю строку для обработки
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' Шаг через строки (B2, B4, B6 и т.д.)
    For currentRow = 2 To lastRow Step 2
        fileName = Trim(ws.Cells(currentRow, "B").Value)
        If fileName <> "" Then
            ' Инициализируем foundFiles
            Set foundFiles = New Collection

            ' Ищем файлы в индексе папки индексации
            Set foundFiles = FindFilesByName(fileNamesDictDestination, fileName)
            If foundFiles.Count > 0 Then
                ' Обрабатываем все найденные версии
                For Each file In foundFiles
                    currentVersion = ExtractVersionFromFileName(fso.GetFileName(CStr(file)), maxVersion)
                    If currentVersion <= maxVersion Then
                        InsertHyperlink ws, currentRow, currentVersion, CStr(file), maxVersion
                    End If
                Next file
            ElseIf useCopyFunctionality Then
                ' Файл не найден в индексе папки индексации, пытаемся найти его в индексе папки поиска
                Set foundFiles = FindFilesByName(fileNamesDictSource, fileName)
                If foundFiles.Count > 0 Then
                    ' Предполагаем, что первый найденный файл наиболее подходящий
                    Dim fullFilePath As String
                    fullFilePath = foundFiles(1)
                    ' Копируем файл в папку индексации с использованием определения целевой папки
                    destinationPath = CopyFileToIndexFolder(fullFilePath, indexFolderPath, fso)
                    ' Обновляем индекс папки индексации
                    AppendFileToIndex destinationPath, indexFilePathDestination, indexFolderPath
                    ' Перезагружаем словарь файлов папки индексации
                    Set fileNamesDictDestination = ReadIndexFileToDictionary(indexFilePathDestination, fso)
                    ' Повторяем поиск в индексе папки индексации
                    Set foundFiles = FindFilesByName(fileNamesDictDestination, fileName)
                    If foundFiles.Count > 0 Then
                        ' Обрабатываем найденные файлы
                        For Each file In foundFiles
                            currentVersion = ExtractVersionFromFileName(fso.GetFileName(CStr(file)), maxVersion)
                            If currentVersion <= maxVersion Then
                                InsertHyperlink ws, currentRow, currentVersion, CStr(file), maxVersion
                            End If
                        Next file
                    Else
                        Debug.Print "Файл не найден после копирования: " & fileName
                    End If
                Else
                    Debug.Print "Файл не найден в индексе папки поиска: " & fileName
                End If
            Else
                Debug.Print "Файл не найден в индексе и копирование отключено: " & fileName
            End If
        Else
            ' Ячейка пуста, пропускаем
        End If
    Next currentRow
End Sub

' Функция для проверки, нужно ли обновлять файл индексации
Function IsIndexUpToDate(scanFolderPath As String, indexFilePath As String, fso As Object) As Boolean
    Dim ScanFolder As Object
    Dim indexFile As Object

    Set ScanFolder = fso.GetFolder(scanFolderPath)

    If fso.FileExists(indexFilePath) Then
        Set indexFile = fso.GetFile(indexFilePath)
        ' Проверяем, была ли папка сканирования изменена после создания файла индексации
        If ScanFolder.DateLastModified <= indexFile.DateLastModified Then
            IsIndexUpToDate = True
        Else
            IsIndexUpToDate = False
        End If
    Else
        IsIndexUpToDate = False
    End If
End Function

' Процедура для создания файла индексации
Sub CreateIndexFile(scanFolderPath As String, indexFilePath As String, fso As Object)
    Dim folder As Object
    Dim outputFile As Object

    Set folder = fso.GetFolder(scanFolderPath)
    On Error GoTo CreateIndexError
    Set outputFile = fso.CreateTextFile(indexFilePath, True)

    ' Рекурсивное сканирование файлов
    ScanFolder folder, outputFile, fso

    outputFile.Close
    On Error GoTo 0
    Exit Sub

CreateIndexError:
    MsgBox "Ошибка при создании файла индексации.", vbCritical
End Sub

' Рекурсивная функция для обхода папок и записи файлов в индекс
Sub ScanFolder(folder As Object, outputFile As Object, fso As Object)
    Dim file As Object
    Dim subFolder As Object

    On Error Resume Next
    For Each file In folder.files
        Err.Clear
        Dim fileName As String
        Dim filePath As String
        fileName = file.Name
        filePath = file.Path
        If Err.Number = 0 Then
            outputFile.WriteLine fileName & "|" & filePath
        Else
            Debug.Print "Ошибка при обработке файла: " & file.Path & " - " & Err.Description
        End If
    Next file
    On Error GoTo 0

    On Error Resume Next
    For Each subFolder In folder.SubFolders
        ScanFolder subFolder, outputFile, fso
    Next subFolder
    On Error GoTo 0
End Sub

' Функция для чтения файла индексации и создания словаря
Function ReadIndexFileToDictionary(filePath As String, fso As Object) As Object
    Dim file As Object
    Dim line As String
    Dim fileNamesDict As Object
    Dim parts() As String
    Dim key As String

    On Error GoTo ReadIndexError
    Set file = fso.OpenTextFile(filePath, ForReading)

    Set fileNamesDict = CreateObject("Scripting.Dictionary")

    Do Until file.AtEndOfStream
        line = file.ReadLine
        parts = Split(line, "|")
        If UBound(parts) = 1 Then
            key = parts(0) ' Используем имя файла в качестве ключа
            If Not fileNamesDict.Exists(key) Then
                Set fileNamesDict(key) = New Collection
            End If
            fileNamesDict(key).Add parts(1)
        End If
    Loop

    file.Close
    On Error GoTo 0

    Set ReadIndexFileToDictionary = fileNamesDict
    Exit Function

ReadIndexError:
    MsgBox "Ошибка при чтении файла индексации.", vbCritical
    Set ReadIndexFileToDictionary = Nothing
End Function

' Процедура для добавления нового файла в индекс
Sub AppendFileToIndex(filePath As String, indexFilePath As String, baseFolderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim outputFile As Object
    Set outputFile = fso.OpenTextFile(indexFilePath, ForAppending, True)
    Dim relativePath As String
    relativePath = GetRelativePath(filePath, baseFolderPath)
    Dim fileName As String
    fileName = fso.GetFileName(filePath)
    ' Записываем в индекс путь к файлу в папке индексации
    outputFile.WriteLine fileName & "|" & fso.BuildPath(baseFolderPath, relativePath)
    outputFile.Close
End Sub

' Функция для получения относительного пути файла от базовой папки
Function GetRelativePath(filePath As String, baseFolderPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim absFilePath As String
    Dim absBaseFolderPath As String

    ' Получаем абсолютные пути
    absFilePath = fso.GetAbsolutePathName(filePath)
    absBaseFolderPath = fso.GetAbsolutePathName(baseFolderPath)

    ' Проверяем, что файл находится внутри базовой папки
    If InStr(1, absFilePath, absBaseFolderPath, vbTextCompare) = 1 Then
        GetRelativePath = Mid(absFilePath, Len(absBaseFolderPath) + 1)
    Else
        ' Файл не находится внутри базовой папки
        GetRelativePath = filePath
    End If
End Function

' Функция для копирования файла в папку индексации с определением целевой папки
Function CopyFileToIndexFolder(sourceFilePath As String, indexFolderPath As String, fso As Object) As String
    Dim fileName As String
    fileName = fso.GetFileName(sourceFilePath)
    Dim destinationPath As String
    destinationPath = DetermineDestinationPath(fileName, indexFolderPath, fso)
    Dim destinationFolder As String
    destinationFolder = fso.GetParentFolderName(destinationPath)
    ' Создаем необходимые папки
    CreateFolderHierarchy destinationFolder, fso
    ' Копируем файл
    fso.CopyFile sourceFilePath, destinationPath, True
    ' Возвращаем путь к скопированному файлу
    CopyFileToIndexFolder = destinationPath
End Function

' Рекурсивная процедура для создания иерархии папок
Sub CreateFolderHierarchy(folderPath As String, fso As Object)
    If Not fso.FolderExists(folderPath) Then
        CreateFolderHierarchy fso.GetParentFolderName(folderPath), fso
        On Error Resume Next
        fso.CreateFolder folderPath
        On Error GoTo 0
    End If
End Sub

' Функция для определения пути назначения на основе названия файла
Function DetermineDestinationPath(fileName As String, indexFolderPath As String, fso As Object) As String
    Dim parts() As String
    parts = Split(fileName, "-")
    Dim destinationFolder As String
    destinationFolder = indexFolderPath ' Начинаем с папки индексации

    If UBound(parts) >= 1 Then
        ' Используем вторую часть названия файла для поиска папки
        Dim secondPart As String
        secondPart = parts(1)
        destinationFolder = FindFolderByNamePart(indexFolderPath, secondPart, fso)
        
        If destinationFolder = "" Then
            ' Папка не найдена, помещаем в папку "Unsorted"
            destinationFolder = fso.BuildPath(indexFolderPath, "Unsorted")
            ' Создаём папку "Unsorted", если её нет
            If Not fso.FolderExists(destinationFolder) Then
                fso.CreateFolder destinationFolder
            End If
        End If
    Else
        ' Не удалось разбить название файла, помещаем в "Unsorted"
        destinationFolder = fso.BuildPath(indexFolderPath, "Unsorted")
        If Not fso.FolderExists(destinationFolder) Then
            fso.CreateFolder destinationFolder
        End If
    End If

    ' Возвращаем полный путь к файлу в целевой папке
    DetermineDestinationPath = fso.BuildPath(destinationFolder, fileName)
End Function

' Функция для поиска папки по части названия
Function FindFolderByNamePart(rootFolderPath As String, namePart As String, fso As Object) As String
    Dim folder As Object
    Set folder = fso.GetFolder(rootFolderPath)
    Dim subFolder As Object
    For Each subFolder In folder.SubFolders
        If InStr(1, subFolder.Name, namePart, vbTextCompare) > 0 Then
            ' Папка найдена
            FindFolderByNamePart = subFolder.Path
            Exit Function
        End If
    Next subFolder
    ' Папка не найдена
    FindFolderByNamePart = ""
End Function

' Функция для поиска файлов по имени в словаре
Function FindFilesByName(fileNamesDict As Object, searchTerm As String) As Collection
    Dim foundFiles As New Collection
    Dim key As Variant

    For Each key In fileNamesDict.Keys
        If InStr(1, key, searchTerm, vbTextCompare) > 0 Then
            Dim files As Collection
            Set files = fileNamesDict(key)
            Dim file As Variant
            For Each file In files
                foundFiles.Add file
            Next file
        End If
    Next key

    ' Всегда возвращаем коллекцию, даже если она пустая
    Set FindFilesByName = foundFiles
End Function

' Функция для вставки гиперссылки
Sub InsertHyperlink(ws As Worksheet, rowNumber As Long, version As Integer, filePath As String, maxVersion As Integer)
    Dim colOffset As Integer
    Dim cell As Range

    ' Получаем номер столбца по версии
    colOffset = GetColumnByVersion(version)
    If colOffset = -1 Then
        MsgBox "Некорректная версия: " & version, vbExclamation
        Exit Sub
    End If

    Set cell = ws.Cells(rowNumber, colOffset)

    ' Проверяем, нет ли уже гиперссылки
    If cell.Hyperlinks.Count = 0 Then
        On Error Resume Next
        cell.Hyperlinks.Add Anchor:=cell, Address:=filePath, TextToDisplay:="Ссылка"
        If Err.Number <> 0 Then
            Debug.Print "Ошибка добавления гиперссылки для файла: " & filePath
            Err.Clear
        End If
        On Error GoTo 0
    Else
        Debug.Print "Гиперссылка уже существует для ячейки " & cell.Address
    End If
End Sub

' Функция для получения номера столбца по версии
Function GetColumnByVersion(version As Integer) As Integer
    ' Столбцы F (6), H (8), J (10), ..., AA (27)
    If version >= 0 And version <= 10 Then
        GetColumnByVersion = 6 + (version * 2)
    Else
        GetColumnByVersion = -1
    End If
End Function

' Функция для извлечения версии из имени файла
Function ExtractVersionFromFileName(fileName As String, maxVersion As Integer) As Integer
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim versionNumber As Integer
    Dim versionString As String

    versionNumber = 0

    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True
    regex.Global = True

    ' Шаблоны для поиска версии
    Dim patterns As Variant
    patterns = Array( _
        "[-_](\d{2,3})(?:[_-][A-Z])?", _
        "_(\d{2,3})_", _
        "(\d{2,3})[-_][E]?", _
        "(\d{2,3})$" _
    )

    Dim i As Integer
    For i = LBound(patterns) To UBound(patterns)
        regex.Pattern = patterns(i)
        Set matches = regex.Execute(fileName)
        If matches.Count > 0 Then
            Set match = matches(matches.Count - 1)
            versionString = match.SubMatches(0)
            versionString = RemoveNonNumeric(versionString)
            If versionString <> "" Then
                versionNumber = CInt(versionString)
                Exit For
            End If
        End If
    Next i

    ExtractVersionFromFileName = versionNumber
End Function

' Функция для удаления нечисловых символов из строки
Function RemoveNonNumeric(str As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\D"
    regex.Global = True
    RemoveNonNumeric = regex.Replace(str, "")
End Function

