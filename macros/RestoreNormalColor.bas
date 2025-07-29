Attribute VB_Name = "RestoreNormalColor"
Sub RestoreNormalColorForSpecificEntries()
    Dim ws As Worksheet
    Dim cell As Range
    Dim startRow As Long, lastRow As Long
    Dim startCol As Long, endCol As Long
    Dim i As Long, j As Long
    Dim valueList As Variant
    Dim cellValue As String
    
    ' ������� ��� ����
    Set ws = ThisWorkbook.Sheets("Instrument List")
    
    ' �������� �������� �� J �� W
    startRow = 10
    startCol = 10 ' J
    endCol = 23   ' W
    
    ' ������ �������� ��� ��������
    valueList = Array("-", "AIH", "AI", "AOA", "REG & SEG", "Safety", "N", "Y", "By Vendor", "REG & SEQ", "AOH", "AI (4-20mA)", "DO", "DI", "Burner Local Panel")
    
    ' ���������� ��������� ������ � �������
    lastRow = ws.Cells(ws.Rows.Count, startCol).End(xlUp).Row
    For i = startCol + 1 To endCol
        Dim tempRow As Long
        tempRow = ws.Cells(ws.Rows.Count, i).End(xlUp).Row
        If tempRow > lastRow Then
            lastRow = tempRow
        End If
    Next i
    
    ' ������ �� �������
    For i = startRow To lastRow
        For j = startCol To endCol
            Set cell = ws.Cells(i, j)
            cellValue = Trim(cell.Value)
            If Len(cellValue) > 0 Then
                If IsInArray(cellValue, valueList) Then
                    ' ��������������� ������� ���� ������
                    cell.Interior.ColorIndex = xlColorIndexNone
                End If
            End If
        Next j
    Next i
    
    MsgBox "���� ����� ��������.", vbInformation
End Sub

' ������� ��� �������� ������� �������� � �������
Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

