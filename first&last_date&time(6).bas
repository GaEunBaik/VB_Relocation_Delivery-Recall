Attribute VB_Name = "Module1"
'���� ���
'I�� (�湮�ڵ�)���� �׷��� ������, �� �湮�ڵ� �׷쿡��:
'J�� = 1�� ���� H�� ���� M���� �ֱ� (ù �ð�)
' �� �湮�ڵ� ������ ������ ����(J���� �ִ밪)�� ���� H�� ���� N���� �ֱ� (������ �ð�)
'������ Sheet1 ~ Sheet5 ���ο� ���� �����մϴ�.


Sub SetFirstAndLastTimePerVisitCode_AllSheets()
    Dim ws As Worksheet
    Dim sheetIndex As Integer
    Dim lastRow As Long
    Dim i As Long
    Dim visitCode As String, timeValue As Variant
    Dim firstRowDict As Object
    Dim lastRowDict As Object
    
    For sheetIndex = 1 To 5
        Set ws = ThisWorkbook.Sheets("Sheet" & sheetIndex)
        Set firstRowDict = CreateObject("Scripting.Dictionary")
        Set lastRowDict = CreateObject("Scripting.Dictionary")
        
        lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
        
        ' 1. �湮�ڵ庰 ù/�� row ��ġ ����
        For i = 2 To lastRow
            visitCode = ws.Cells(i, "I").Value
            
            If visitCode <> "" Then
                Dim seq As Variant
                seq = ws.Cells(i, "J").Value
                
                If seq = 1 Then
                    firstRowDict(visitCode) = i
                End If
                
                ' �湮�ڵ庰 ���� ū ������ ã�� ���� ��� ����
                If Not lastRowDict.exists(visitCode) Then
                    lastRowDict(visitCode) = i
                ElseIf ws.Cells(i, "J").Value > ws.Cells(lastRowDict(visitCode), "J").Value Then
                    lastRowDict(visitCode) = i
                End If
            End If
        Next i
        
        ' 2. �ش� �࿡ H�� ���� M/N���� �Է�
        Dim key As Variant
        For Each key In firstRowDict.Keys
            Dim firstRow As Long, lastRowNum As Long
            firstRow = firstRowDict(key)
            lastRowNum = lastRowDict(key)
            
            ws.Cells(firstRow, "M").Value = ws.Cells(firstRow, "H").Value
            ws.Cells(lastRowNum, "N").Value = ws.Cells(lastRowNum, "H").Value
        Next key
    Next sheetIndex
    
    MsgBox "ù/������ �ð� ��� �Ϸ�!", vbInformation
End Sub
