Attribute VB_Name = "CreateDocumentHyperlinks"
Option Explicit

' �������� ��������� ��� ��������� ������
Sub ProcessFiles()
    Dim ws As Worksheet
    Dim folderPath As String
    Dim fileNamesDict As Object
    Dim fileName As String
    Dim currentRow As Long
    Dim version As Integer
    Dim foundFiles As Collection
    Dim file As Variant
    Dim currentVersion As Integer
    Dim maxVersion As Integer
    Dim indexFilePath As String
    Dim fso As Object
    Dim lastRow As Long

    Set ws = ActiveSheet
    folderPath = ThisWorkbook.Path & "\" ' ���� � ����� � �������
    indexFilePath = folderPath & "index.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' �������� ������������ ���������� ������ ����� ������
    maxVersion = InputBox("������� ������������ ���������� ������", "���������", 10)
    If maxVersion = 0 Then
        MsgBox "������������ ���������� ������ ������ ���� ������ ����.", vbExclamation
        Exit Sub
    End If

    ' ���������, ����� �� ��������� ���� ����������
    If Not IsIndexUpToDate(folderPath, indexFilePath, fso) Then
        CreateIndexFile folderPath, indexFilePath, fso
    End If

    ' ��������� ������ �� ����� ���������� � �������
    Set fileNamesDict = ReadIndexFileToDictionary(indexFilePath, fso)
    If fileNamesDict Is Nothing Then
        MsgBox "������ ��� �������� ����� ����������.", vbCritical
        Exit Sub
    End If

    ' ���������� ��������� ������ ��� ���������
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row

    ' ��� ����� ������ (B8, B10, B12 � �.�.)
    For currentRow = 8 To lastRow Step 2
        fileName = Trim(ws.Cells(currentRow, "B").Value)
        If fileName <> "" Then
            Set foundFiles = FindFilesByName(fileNamesDict, fileName)
            If Not foundFiles Is Nothing Then
                If foundFiles.Count > 0 Then
                    ' ������������ ��� ��������� ������
                    For Each file In foundFiles
                        currentVersion = ExtractVersionFromFileName(fso.GetFileName(CStr(file)), maxVersion)
                        If currentVersion <= maxVersion Then
                            InsertHyperlink ws, currentRow, currentVersion, CStr(file), maxVersion
                        End If
                    Next file
                Else
                    Debug.Print "����� �� ������� ���: " & fileName
                End If
            Else
                Debug.Print "����� �� ������� ���: " & fileName
            End If
        Else
            ' ������ �����, ����������
            ' Debug.Print "������ ������ � ������ " & currentRow
        End If
    Next currentRow
End Sub

' �������� ������������� ����� ���������� � ������������� ����������
Function IsIndexUpToDate(folderPath As String, indexFilePath As String, fso As Object) As Boolean
    Dim folder As Object
    Dim indexFile As Object

    Set folder = fso.GetFolder(folderPath)

    If fso.FileExists(indexFilePath) Then
        Set indexFile = fso.GetFile(indexFilePath)
        If folder.DateLastModified <= indexFile.DateLastModified Then
            IsIndexUpToDate = True
        Else
            IsIndexUpToDate = False
        End If
    Else
        IsIndexUpToDate = False
    End If
End Function

' ������� ��� ������ ������ �� ����� �� �������
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

    If foundFiles.Count > 0 Then
        Set FindFilesByName = foundFiles
    Else
        Set FindFilesByName = Nothing
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

' ������� ��� �������� ����� �������
Sub CreateIndexFile(folderPath As String, indexFilePath As String, fso As Object)
    Dim folder As Object
    Dim outputFile As Object

    Set folder = fso.GetFolder(folderPath)
    On Error GoTo CreateIndexError
    Set outputFile = fso.CreateTextFile(indexFilePath, True)

    ' ����������� ������������ ������
    ScanFolder folder, outputFile, fso

    outputFile.Close
    On Error GoTo 0
    Exit Sub

CreateIndexError:
    MsgBox "������ ��� �������� ����� �������.", vbCritical
End Sub

' ����������� ������� ��� ������ �����
Sub ScanFolder(folder As Object, outputFile As Object, fso As Object)
    Dim file As Object
    Dim subfolder As Object
    
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
    For Each subfolder In folder.SubFolders
        ScanFolder subfolder, outputFile, fso
    Next subfolder
    On Error GoTo 0
End Sub

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

' ������� ��� ������ ����� ������� � ���������� ������ � �������
Function ReadIndexFileToDictionary(filePath As String, fso As Object) As Object
    Dim file As Object
    Dim line As String
    Dim fileNamesDict As Object
    Dim parts() As String
    Dim key As String

    On Error GoTo ReadIndexError
    Set file = fso.OpenTextFile(filePath, 1)

    Set fileNamesDict = CreateObject("Scripting.Dictionary")

    Do Until file.AtEndOfStream
        line = file.ReadLine
        parts = Split(line, "|")
        If UBound(parts) = 1 Then
            key = parts(0) ' ���������� ������ ��� ����� � �������� �����
            If Not fileNamesDict.Exists(key) Then
                Set fileNamesDict(key) = New Collection
            End If
            fileNamesDict(key).Add parts(1) ' ��������� ������ ���� � ����� � ���������
        End If
    Loop

    file.Close
    On Error GoTo 0

    Set ReadIndexFileToDictionary = fileNamesDict
    Exit Function

ReadIndexError:
    MsgBox "������ ��� ������ ����� �������.", vbCritical
    Set ReadIndexFileToDictionary = Nothing
End Function
