Attribute VB_Name = "IndexAndCreateHyperlink"
Option Explicit

' ����������� �������� ��� ������ �������� ������
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

' �������� ��������� ��� ��������� ������
Sub ProcessFiles()
    Dim ws As Worksheet
    Dim indexFolderPath As String ' ����� ���������� (���� �������� �����)
    Dim searchFolderPath As String ' ����� ������ ����������� ������ (������ �������� �����)
    Dim useCopyFunctionality As Boolean
    Dim fileNamesDictDestination As Object ' ������� ������ � ����� ����������
    Dim fileNamesDictSource As Object ' ������� ������ � ����� ������
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

    ' �������� ���� � �����, ��� ��������� ���� Excel
    excelFilePath = ThisWorkbook.Path & "\"

    ' ����� ����� ��� ������������ (����� ����������)
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    dlg.Title = "�������� ����� ��� ������������ (����� ����������)"
    If dlg.Show = -1 Then
        indexFolderPath = dlg.SelectedItems(1) & "\"
    Else
        MsgBox "���� � ����� ���������� �� ������. ������ ��������."
        Exit Sub
    End If

    ' �������� ������������ ���������� ������ ����� ������
    maxVersion = InputBox("������� ������������ ���������� ������", "���������", 10)
    If maxVersion = 0 Then
        MsgBox "������������ ���������� ������ ������ ���� ������ ����.", vbExclamation
        Exit Sub
    End If

    ' ���������� ������������ � ������������� ������������� ����������� ����������� ����������� ������
    userResponse = MsgBox("������ �� �� ������������ ���������� ������ � ����������� ����������� ������?" & vbCrLf & _
                          "���� ��, ��� ����� ���������� ������� ����� ��� ������." & vbCrLf & _
                          "�������� ��������, ��� ��������� ����� ����� ����������� � ����� ����������.", _
                          vbYesNo + vbQuestion, "������������ ����� � ����������� ������?")

    If userResponse = vbYes Then
        useCopyFunctionality = True

        ' ����� ����� ��� ������ ����������� ������
        Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
        dlg.Title = "�������� ����� ��� ������ ����������� ������"
        If dlg.Show = -1 Then
            searchFolderPath = dlg.SelectedItems(1) & "\"
        Else
            MsgBox "���� � ����� ������ �� ������. ������ ��������."
            Exit Sub
        End If
    Else
        useCopyFunctionality = False
    End If

    ' ���������� ���� � ������ ���������� (� ����� � ������ Excel)
    indexFilePathDestination = excelFilePath & "index_destination.txt"
    If useCopyFunctionality Then
        indexFilePathSource = excelFilePath & "index_source.txt"
    End If

    ' ���������, ����� �� ��������� ���� ���������� ��� ����� ����������
    If Not IsIndexUpToDate(indexFolderPath, indexFilePathDestination, fso) Then
        CreateIndexFile indexFolderPath, indexFilePathDestination, fso
    End If

    ' ��������� ������ �� ����� ���������� � ������� ��� ����� ����������
    Set fileNamesDictDestination = ReadIndexFileToDictionary(indexFilePathDestination, fso)
    If fileNamesDictDestination Is Nothing Then
        MsgBox "������ ��� �������� ����� ���������� ��� ����� ����������.", vbCritical
        Exit Sub
    End If

    ' ���� ������������ ���������� �����������, �������� � �������� ����� ������
    If useCopyFunctionality Then
        ' ���������, ����� �� ��������� ���� ���������� ��� ����� ������
        If Not fso.FileExists(indexFilePathSource) Then
            ' ���� ����� ���, ������� ���
            CreateIndexFile searchFolderPath, indexFilePathSource, fso
        Else
            ' �������� ������������ � ����������
            userResponse = MsgBox("���� ���������� ��� ����� ������ ��� ����������. ��������������� ����� � �������� ������?", vbYesNo + vbQuestion, "�������� ������?")
            If userResponse = vbYes Then
                CreateIndexFile searchFolderPath, indexFilePathSource, fso
            End If
        End If

        ' ��������� ������ �� ����� ���������� � ������� ��� ����� ������
        Set fileNamesDictSource = ReadIndexFileToDictionary(indexFilePathSource, fso)
        If fileNamesDictSource Is Nothing Then
            MsgBox "������ ��� �������� ����� ���������� ��� ����� ������.", vbCritical
            Exit Sub
        End If
    End If

    ' ���������� ��������� ������ ��� ���������
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' ��� ����� ������ (B2, B4, B6 � �.�.)
    For currentRow = 2 To lastRow Step 2
        fileName = Trim(ws.Cells(currentRow, "B").Value)
        If fileName <> "" Then
            ' �������������� foundFiles
            Set foundFiles = New Collection

            ' ���� ����� � ������� ����� ����������
            Set foundFiles = FindFilesByName(fileNamesDictDestination, fileName)
            If foundFiles.Count > 0 Then
                ' ������������ ��� ��������� ������
                For Each file In foundFiles
                    currentVersion = ExtractVersionFromFileName(fso.GetFileName(CStr(file)), maxVersion)
                    If currentVersion <= maxVersion Then
                        InsertHyperlink ws, currentRow, currentVersion, CStr(file), maxVersion
                    End If
                Next file
            ElseIf useCopyFunctionality Then
                ' ���� �� ������ � ������� ����� ����������, �������� ����� ��� � ������� ����� ������
                Set foundFiles = FindFilesByName(fileNamesDictSource, fileName)
                If foundFiles.Count > 0 Then
                    ' ������������, ��� ������ ��������� ���� �������� ����������
                    Dim fullFilePath As String
                    fullFilePath = foundFiles(1)
                    ' �������� ���� � ����� ���������� � �������������� ����������� ������� �����
                    destinationPath = CopyFileToIndexFolder(fullFilePath, indexFolderPath, fso)
                    ' ��������� ������ ����� ����������
                    AppendFileToIndex destinationPath, indexFilePathDestination, indexFolderPath
                    ' ������������� ������� ������ ����� ����������
                    Set fileNamesDictDestination = ReadIndexFileToDictionary(indexFilePathDestination, fso)
                    ' ��������� ����� � ������� ����� ����������
                    Set foundFiles = FindFilesByName(fileNamesDictDestination, fileName)
                    If foundFiles.Count > 0 Then
                        ' ������������ ��������� �����
                        For Each file In foundFiles
                            currentVersion = ExtractVersionFromFileName(fso.GetFileName(CStr(file)), maxVersion)
                            If currentVersion <= maxVersion Then
                                InsertHyperlink ws, currentRow, currentVersion, CStr(file), maxVersion
                            End If
                        Next file
                    Else
                        Debug.Print "���� �� ������ ����� �����������: " & fileName
                    End If
                Else
                    Debug.Print "���� �� ������ � ������� ����� ������: " & fileName
                End If
            Else
                Debug.Print "���� �� ������ � ������� � ����������� ���������: " & fileName
            End If
        Else
            ' ������ �����, ����������
        End If
    Next currentRow
End Sub

' ������� ��� ��������, ����� �� ��������� ���� ����������
Function IsIndexUpToDate(scanFolderPath As String, indexFilePath As String, fso As Object) As Boolean
    Dim ScanFolder As Object
    Dim indexFile As Object

    Set ScanFolder = fso.GetFolder(scanFolderPath)

    If fso.FileExists(indexFilePath) Then
        Set indexFile = fso.GetFile(indexFilePath)
        ' ���������, ���� �� ����� ������������ �������� ����� �������� ����� ����������
        If ScanFolder.DateLastModified <= indexFile.DateLastModified Then
            IsIndexUpToDate = True
        Else
            IsIndexUpToDate = False
        End If
    Else
        IsIndexUpToDate = False
    End If
End Function

' ��������� ��� �������� ����� ����������
Sub CreateIndexFile(scanFolderPath As String, indexFilePath As String, fso As Object)
    Dim folder As Object
    Dim outputFile As Object

    Set folder = fso.GetFolder(scanFolderPath)
    On Error GoTo CreateIndexError
    Set outputFile = fso.CreateTextFile(indexFilePath, True)

    ' ����������� ������������ ������
    ScanFolder folder, outputFile, fso

    outputFile.Close
    On Error GoTo 0
    Exit Sub

CreateIndexError:
    MsgBox "������ ��� �������� ����� ����������.", vbCritical
End Sub

' ����������� ������� ��� ������ ����� � ������ ������ � ������
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
            Debug.Print "������ ��� ��������� �����: " & file.Path & " - " & Err.Description
        End If
    Next file
    On Error GoTo 0

    On Error Resume Next
    For Each subFolder In folder.SubFolders
        ScanFolder subFolder, outputFile, fso
    Next subFolder
    On Error GoTo 0
End Sub

' ������� ��� ������ ����� ���������� � �������� �������
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
            key = parts(0) ' ���������� ��� ����� � �������� �����
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
    MsgBox "������ ��� ������ ����� ����������.", vbCritical
    Set ReadIndexFileToDictionary = Nothing
End Function

' ��������� ��� ���������� ������ ����� � ������
Sub AppendFileToIndex(filePath As String, indexFilePath As String, baseFolderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim outputFile As Object
    Set outputFile = fso.OpenTextFile(indexFilePath, ForAppending, True)
    Dim relativePath As String
    relativePath = GetRelativePath(filePath, baseFolderPath)
    Dim fileName As String
    fileName = fso.GetFileName(filePath)
    ' ���������� � ������ ���� � ����� � ����� ����������
    outputFile.WriteLine fileName & "|" & fso.BuildPath(baseFolderPath, relativePath)
    outputFile.Close
End Sub

' ������� ��� ��������� �������������� ���� ����� �� ������� �����
Function GetRelativePath(filePath As String, baseFolderPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim absFilePath As String
    Dim absBaseFolderPath As String

    ' �������� ���������� ����
    absFilePath = fso.GetAbsolutePathName(filePath)
    absBaseFolderPath = fso.GetAbsolutePathName(baseFolderPath)

    ' ���������, ��� ���� ��������� ������ ������� �����
    If InStr(1, absFilePath, absBaseFolderPath, vbTextCompare) = 1 Then
        GetRelativePath = Mid(absFilePath, Len(absBaseFolderPath) + 1)
    Else
        ' ���� �� ��������� ������ ������� �����
        GetRelativePath = filePath
    End If
End Function

' ������� ��� ����������� ����� � ����� ���������� � ������������ ������� �����
Function CopyFileToIndexFolder(sourceFilePath As String, indexFolderPath As String, fso As Object) As String
    Dim fileName As String
    fileName = fso.GetFileName(sourceFilePath)
    Dim destinationPath As String
    destinationPath = DetermineDestinationPath(fileName, indexFolderPath, fso)
    Dim destinationFolder As String
    destinationFolder = fso.GetParentFolderName(destinationPath)
    ' ������� ����������� �����
    CreateFolderHierarchy destinationFolder, fso
    ' �������� ����
    fso.CopyFile sourceFilePath, destinationPath, True
    ' ���������� ���� � �������������� �����
    CopyFileToIndexFolder = destinationPath
End Function

' ����������� ��������� ��� �������� �������� �����
Sub CreateFolderHierarchy(folderPath As String, fso As Object)
    If Not fso.FolderExists(folderPath) Then
        CreateFolderHierarchy fso.GetParentFolderName(folderPath), fso
        On Error Resume Next
        fso.CreateFolder folderPath
        On Error GoTo 0
    End If
End Sub

' ������� ��� ����������� ���� ���������� �� ������ �������� �����
Function DetermineDestinationPath(fileName As String, indexFolderPath As String, fso As Object) As String
    Dim parts() As String
    parts = Split(fileName, "-")
    Dim destinationFolder As String
    destinationFolder = indexFolderPath ' �������� � ����� ����������

    If UBound(parts) >= 1 Then
        ' ���������� ������ ����� �������� ����� ��� ������ �����
        Dim secondPart As String
        secondPart = parts(1)
        destinationFolder = FindFolderByNamePart(indexFolderPath, secondPart, fso)
        
        If destinationFolder = "" Then
            ' ����� �� �������, �������� � ����� "Unsorted"
            destinationFolder = fso.BuildPath(indexFolderPath, "Unsorted")
            ' ������ ����� "Unsorted", ���� � ���
            If Not fso.FolderExists(destinationFolder) Then
                fso.CreateFolder destinationFolder
            End If
        End If
    Else
        ' �� ������� ������� �������� �����, �������� � "Unsorted"
        destinationFolder = fso.BuildPath(indexFolderPath, "Unsorted")
        If Not fso.FolderExists(destinationFolder) Then
            fso.CreateFolder destinationFolder
        End If
    End If

    ' ���������� ������ ���� � ����� � ������� �����
    DetermineDestinationPath = fso.BuildPath(destinationFolder, fileName)
End Function

' ������� ��� ������ ����� �� ����� ��������
Function FindFolderByNamePart(rootFolderPath As String, namePart As String, fso As Object) As String
    Dim folder As Object
    Set folder = fso.GetFolder(rootFolderPath)
    Dim subFolder As Object
    For Each subFolder In folder.SubFolders
        If InStr(1, subFolder.Name, namePart, vbTextCompare) > 0 Then
            ' ����� �������
            FindFolderByNamePart = subFolder.Path
            Exit Function
        End If
    Next subFolder
    ' ����� �� �������
    FindFolderByNamePart = ""
End Function

' ������� ��� ������ ������ �� ����� � �������
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

    ' ������ ���������� ���������, ���� ���� ��� ������
    Set FindFilesByName = foundFiles
End Function

' ������� ��� ������� �����������
Sub InsertHyperlink(ws As Worksheet, rowNumber As Long, version As Integer, filePath As String, maxVersion As Integer)
    Dim colOffset As Integer
    Dim cell As Range

    ' �������� ����� ������� �� ������
    colOffset = GetColumnByVersion(version)
    If colOffset = -1 Then
        MsgBox "������������ ������: " & version, vbExclamation
        Exit Sub
    End If

    Set cell = ws.Cells(rowNumber, colOffset)

    ' ���������, ��� �� ��� �����������
    If cell.Hyperlinks.Count = 0 Then
        On Error Resume Next
        cell.Hyperlinks.Add Anchor:=cell, Address:=filePath, TextToDisplay:="������"
        If Err.Number <> 0 Then
            Debug.Print "������ ���������� ����������� ��� �����: " & filePath
            Err.Clear
        End If
        On Error GoTo 0
    Else
        Debug.Print "����������� ��� ���������� ��� ������ " & cell.Address
    End If
End Sub

' ������� ��� ��������� ������ ������� �� ������
Function GetColumnByVersion(version As Integer) As Integer
    ' ������� F (6), H (8), J (10), ..., AA (27)
    If version >= 0 And version <= 10 Then
        GetColumnByVersion = 6 + (version * 2)
    Else
        GetColumnByVersion = -1
    End If
End Function

' ������� ��� ���������� ������ �� ����� �����
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

    ' ������� ��� ������ ������
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

' ������� ��� �������� ���������� �������� �� ������
Function RemoveNonNumeric(str As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\D"
    regex.Global = True
    RemoveNonNumeric = regex.Replace(str, "")
End Function

