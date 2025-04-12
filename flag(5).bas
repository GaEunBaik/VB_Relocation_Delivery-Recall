Attribute VB_Name = "Module1"
'���� ���
'I��(�湮�ڵ�) �������� ���� �׷� ������ C��(�����Ź�ȣ)�� �ߺ��� ���
'�� ó�� ��Ÿ�� �ุ ����, �� �Ʒ��� �ݺ��� ���� L���� "��¥"��� ǥ���ϱ�
'(Sheet1~Sheet5 ��ü ��ȸ)

Sub MarkFakeBikesInAllSheets()
    Dim ws As Worksheet
    Dim sheetIndex As Integer
    Dim lastRow As Long
    Dim i As Long
    Dim seenDict As Object
    Dim visitCode As String, bikeID As String, key As String

    For sheetIndex = 1 To 5
        Set ws = ThisWorkbook.Sheets("Sheet" & sheetIndex)
        Set seenDict = CreateObject("Scripting.Dictionary")
        
        lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
        
        For i = 2 To lastRow
            visitCode = ws.Cells(i, "I").Value
            bikeID = ws.Cells(i, "C").Value
            
            If visitCode <> "" And bikeID <> "" Then
                key = visitCode & "|" & bikeID
                
                If seenDict.exists(key) Then
                    ws.Cells(i, "L").Value = "��¥"
                Else
                    seenDict.Add key, True
                    ws.Cells(i, "L").ClearContents ' Ȥ�� ������ ��¥�� ������ ����
                End If
            End If
        Next i
    Next sheetIndex
    
    MsgBox "��� ��Ʈ ó�� �Ϸ�Ǿ����ϴ�!", vbInformation
End Sub
