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

    ' ������� ������ FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' ����� ����� �����
    Set folderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    folderDialog.Title = "�������� ����� ����� ��� �����������"
    If folderDialog.Show = -1 Then
        newPath = folderDialog.SelectedItems(1) & "\"
    Else
        MsgBox "����� �� ���� �������. ������ ����������.", vbExclamation
        Exit Sub
    End If
    
    ' ���������� ���� ��� ����� ���������� � ��� �� �����, ��� ��������� Excel ����
    indexFilePath = ThisWorkbook.Path & "\FileIndex.txt"
    
    ' ���� ���� ���������� �� ����������, ������� ���
    If Not fso.fileExists(indexFilePath) Then
        Set textStream = fso.CreateTextFile(indexFilePath, True, True) ' ������� ���� �������, Unicode
        MsgBox "���� ���������� ��� ������.", vbInformation
        textStream.Close ' ��������� ���� ����� ��������
    End If
    
    ' ��������� ���� ������� ��� ������
    Set textStream = fso.OpenTextFile(indexFilePath, 1, False, -1) ' 1 = ForReading, -1 = Unicode
    
    ' ������� ��������� ��� ����������
    Set fileIndex = CreateObject("Scripting.Dictionary")
    
    ' ������ ������� �� �����
    Do While Not textStream.AtEndOfStream
        line = textStream.ReadLine
        parts = Split(line, ",")
        If UBound(parts) = 1 Then
            currentKey = LCase(parts(0))
            fileIndex.Add currentKey, parts(1)
        End If
    Loop
    textStream.Close ' ��������� ���� ����� ������
    
    ' ����������� ��� ����� � �����
    If fileIndex.Count = 0 Then
        Set folder = fso.GetFolder(newPath)
        Call IndexFilesAndFolders(folder, fileIndex, fso)
        
        ' ���������� ������ � ����
        Set textStream = fso.OpenTextFile(indexFilePath, 2, False, -1) ' 2 = ForWriting, -1 = Unicode
        For Each currentKey In fileIndex
            textStream.WriteLine currentKey & "," & fileIndex(currentKey)
        Next currentKey
        textStream.Close ' ��������� ���� ����� ������
    End If
    
    ' �������� �����������
    For Each ws In ThisWorkbook.Sheets
        Application.StatusBar = "��������� �����: " & ws.Name
        For Each hl In ws.Hyperlinks
            Debug.Print "�������� �����������: " & hl.Address
            If Len(hl.Address) > 0 Then ' ���������, ��� ����������� �������� �����
                targetName = LCase(fso.GetFileName(hl.Address))
                Debug.Print "��� �����: " & targetName ' ���������, ��������� �� ��� �����
                
                isFolder = (Right(hl.Address, 1) = "\" Or fso.FolderExists(hl.Address))
                fileExists = False
                
                ' ��������� ������
                If fileIndex.Exists(targetName) Then
                    targetPath = fileIndex(targetName)
                    hl.Address = targetPath ' ��������� ����������� �� ���������� ����
                    hl.Parent.Font.Color = vbBlack ' ���������� ������ ���� ������� �������
                    fileExists = True
                    Debug.Print "���� ������ � �������. ����: " & targetPath
                Else
                    Debug.Print "���� �� ������ � �������."
                End If
                
                ' ���� �� �������, �������� ����������� ������� ������
                If Not fileExists Then
                    hl.Address = "" ' ������� �����, ���� ���� �� ������
                    hl.Parent.Font.Color = vbRed
                    Debug.Print "����������� �������� �������."
                End If
            Else
                Debug.Print "������ �����������."
            End If
        Next hl
    Next ws
    
    ' ������� StatusBar
    Application.StatusBar = False
    
    MsgBox "���������� ����������� ���������.", vbInformation
End Sub

Sub IndexFilesAndFolders(folder As Object, fileIndex As Object, fso As Object)
    Dim file As Object
    Dim subFolder As Object

    ' ����������� �����
    For Each file In folder.Files
        fileIndex.Add LCase(fso.GetFileName(file.Path)), file.Path
        ' �������: ������� ���������� � ������
        Debug.Print "���� �������� � ������: " & fso.GetFileName(file.Path) & " -> " & file.Path
    Next file
    
    ' ����������� ����� ����������
    For Each subFolder In folder.SubFolders
        Call IndexFilesAndFolders(subFolder, fileIndex, fso)
    Next subFolder
End Sub
