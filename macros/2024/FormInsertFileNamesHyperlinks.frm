VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormInsertFileNamesHyperlinks 
   Caption         =   "�������� ����������� �� ���������"
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

    ' ��������� ���� � �����
    folderPath = TextBox1.Value

    ' �������� �� ������� ���� � �����
    If folderPath = "" Then
        MsgBox "���� � ����� �� ������. ����������, ������� ����.", vbExclamation
        Exit Sub
    ElseIf Dir(folderPath, vbDirectory) = "" Then
        MsgBox "��������� ���� �� ����������. ����������, ������� ���������� ����.", vbExclamation
        Exit Sub
    End If

    ' �������� �� ������� ��������� ����� � ����� ����
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    ' ��������� ��������� ������
    On Error Resume Next
    Set cell = Range(RefEdit1.Text)
    On Error GoTo 0

    ' �������� �� ������������ ������ ������
    If cell Is Nothing Then
        MsgBox "������ ������������ ����� ������. ����������, �������� ���������� ������.", vbExclamation
        Exit Sub
    End If

    ' ������� ���� ������ � �����
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

    ' ��������� � ����������
    MsgBox "����������� �� ����� ��������� �������!", vbInformation
    
    ' �������� �����
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    ' �������� ����� ��� ���������� ��������
    Unload Me
End Sub

Private Sub CommandButton3_Click()
    ' �������� ���� ������ �����
    Dim folderPath As FileDialog
    Set folderPath = Application.FileDialog(msoFileDialogFolderPicker)
    folderPath.Title = "�������� ����� � �������"
    
    ' ��������, ������ �� ������������ �����
    If folderPath.Show = -1 Then
        TextBox1.Text = folderPath.SelectedItems(1)
    Else
        MsgBox "����� �� �������. ����������, �������� �����.", vbExclamation
    End If
End Sub

