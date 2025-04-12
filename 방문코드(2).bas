Attribute VB_Name = "Module1"
' Sheet1���� Sheet5���� �ڵ����� �����鼭, �湮�ڵ�(A0000001 �� A0000002 �� ...)�� ���������� �̾���̱�
' Sheet1���� 5���� D�÷��� �������� H�÷��� 30�� �̻��� �� I�÷��� �湮�ڵ� A0000001���� ����

Sub GenerateVisitCodesAcrossSheets()
    Dim ws As Worksheet
    Dim sheetIndex As Integer
    Dim currentCodeIndex As Long
    currentCodeIndex = 1 ' A0000001���� ����
    
    For sheetIndex = 1 To 5
        Set ws = ThisWorkbook.Sheets("Sheet" & sheetIndex)
        
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.count, "D").End(xlUp).Row
        
        Dim visitCode As String
        visitCode = "A" & Format(currentCodeIndex, "0000000")
        
        Dim prevTime As Date
        Dim prevLocation As Variant
        
        Dim i As Long
        For i = 2 To lastRow
            If IsDate(ws.Cells(i, "H").Value) = False Then Exit For
            
            Dim currTime As Date
            currTime = ws.Cells(i, "H").Value
            
            Dim currLocation As Variant
            currLocation = ws.Cells(i, "D").Value
            
            If i = 2 And sheetIndex = 1 Then
                ' Sheet1�� ù ��
                ws.Cells(i, "I").Value = visitCode
                prevTime = currTime
                prevLocation = currLocation
            ElseIf i = 2 Then
                ' �ٸ� ��Ʈ�� ù ��
                ws.Cells(i, "I").Value = visitCode
                prevTime = currTime
                prevLocation = currLocation
            Else
                If currLocation <> prevLocation Or Abs(DateDiff("n", currTime, prevTime)) >= 30 Then
                    currentCodeIndex = currentCodeIndex + 1
                    visitCode = "A" & Format(currentCodeIndex, "0000000")
                End If
                ws.Cells(i, "I").Value = visitCode
                prevTime = currTime
                prevLocation = currLocation
            End If
        Next i
    Next sheetIndex
End Sub
