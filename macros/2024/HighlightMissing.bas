Attribute VB_Name = "HighlightMissing"
Sub FindDocumentsAndHighlightMissing()
    Dim ws As Worksheet
    Dim scanFolders As Variant
    Dim fileDict As Object
    Dim cell As Range
    Dim startRow As Long, lastRow As Long
    Dim startCol As Long, endCol As Long
    Dim docCode As String, cleanCode As String
    Dim found As Boolean
    Dim i As Long, j As Long
    Dim indexFilePath As String

    ' Инициализация
    Set ws = ThisWorkbook.Sheets("Instrument List") ' Укажите ваш лист
    scanFolders = Array("\\files-eco-001\02_дрп\04_СЗ_DRI\Исходные_данные", _
                        "\\Files-eco-001\02_дрп\06_АСУТП_ИТ\DRI\Оборудование\по агрегатам")
    Set fileDict = CreateObject("Scripting.Dictionary")
    indexFilePath = ThisWorkbook.Path & "\file_index.csv" ' Путь для сохранения файла индекса

    ' Загрузка индекса из файла или создание нового
    If Dir(indexFilePath) <> "" Then
        Call LoadFileIndex(indexFilePath, fileDict)
    Else
        Call IndexFiles(scanFolders, fileDict)
        Call SaveFileIndex(indexFilePath, fileDict)
    End If

    ' Указание диапазона поиска
    startRow = 10
    startCol = 10 ' J
    endCol = 23  ' W

    ' Определяем последнюю строку с данными
    lastRow = ws.Cells(ws.Rows.Count, startCol).End(xlUp).Row
    For i = startCol + 1 To endCol
        Dim tempRow As Long
        tempRow = ws.Cells(ws.Rows.Count, i).End(xlUp).Row
        If tempRow > lastRow Then
            lastRow = tempRow
        End If
    Next i

    ' Проход по таблице
    For i = startRow To lastRow
        For j = startCol To endCol
            Set cell = ws.Cells(i, j)
            If Len(Trim(cell.Value)) > 0 Then
                ' Извлекаем код документа, очищаем его
                docCode = cell.Value
                cleanCode = ExtractCode(docCode)

                ' Проверка наличия документа
                found = fileDict.Exists(cleanCode)

                ' Если документ не найден, окрашиваем ячейку в красный цвет
                If Not found Then
                    cell.Interior.Color = RGB(255, 0, 0)
                End If
            End If
        Next j
    Next i

    MsgBox "Проверка завершена", vbInformation
End Sub

' Функция для индексации файлов из указанных папок
Sub IndexFiles(folders As Variant, fileDict As Object)
    Dim folder As Variant
    Dim fso As Object
    Dim folderObj As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    For Each folder In folders
        If fso.FolderExists(folder) Then
            Set folderObj = fso.GetFolder(folder)
            Call RecursiveFileSearch(folderObj, fileDict)
        Else
            MsgBox "Папка не найдена: " & folder, vbExclamation
        End If
    Next folder
End Sub

' Рекурсивная функция для поиска файлов
Sub RecursiveFileSearch(folderObj As Object, fileDict As Object)
    Dim fileObj As Object
    Dim subFolderObj As Object
    Dim cleanName As String

    For Each fileObj In folderObj.Files
        cleanName = ExtractCode(fileObj.Name)
        If Not fileDict.Exists(cleanName) Then
            fileDict.Add cleanName, fileObj.Path
        End If
    Next fileObj

    For Each subFolderObj In folderObj.SubFolders
        Call RecursiveFileSearch(subFolderObj, fileDict)
    Next subFolderObj
End Sub

' Функция для очистки кода документа
Function ExtractCode(docName As String) As String
    Dim parts As Variant
    parts = Split(docName, " -")(0) ' Убираем все после " - "
    parts = Split(parts, "_")(0)    ' Убираем все после "_"
    ExtractCode = Trim(parts)
End Function

' Сохранение индекса в файл CSV
Sub SaveFileIndex(filePath As String, fileDict As Object)
    Dim fileNum As Integer
    Dim key As Variant

    fileNum = FreeFile
    Open filePath For Output As #fileNum

    For Each key In fileDict.Keys
        Print #fileNum, fileDict(key) & ";" & key
    Next key

    Close #fileNum
End Sub

' Загрузка индекса из файла CSV
Sub LoadFileIndex(filePath As String, fileDict As Object)
    Dim fileNum As Integer
    Dim line As String
    Dim parts As Variant

    fileNum = FreeFile
    Open filePath For Input As #fileNum

    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        parts = Split(line, ";")
        If UBound(parts) = 1 Then
            If Not fileDict.Exists(parts(1)) Then
                fileDict.Add parts(1), parts(0)
            End If
        End If
    Loop

    Close #fileNum
End Sub
