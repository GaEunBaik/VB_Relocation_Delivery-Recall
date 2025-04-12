Attribute VB_Name = "Module1"
'���� ���
'I��(�湮�ڵ�) �������� ���� �׷� ������ C��(�����Ź�ȣ)�� �ߺ��� ���
'�� ó�� ��Ÿ�� �ุ ����, �� �Ʒ��� �ݺ��� ���� L���� "��¥"��� ǥ���ϱ�
'(Sheet1~Sheet5 ��ü ��ȸ)

Sub CountUniqueBikesByVisitCode_AllSheets()
    Dim ws As Worksheet
    Dim sheetIndex As Integer
    Dim lastRow As Long
    Dim i As Long
    Dim visitBikes As Object
    Dim visitCountDict As Object
    Dim visitCode As String, bikeID As String, key As String
    
    For sheetIndex = 1 To 5
        Set ws = ThisWorkbook.Sheets("Sheet" & sheetIndex)
        Set visitBikes = CreateObject("Scripting.Dictionary")
        Set visitCountDict = CreateObject("Scripting.Dictionary")
        
        lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
        
        ' 1. �湮�ڵ� + �����Ź�ȣ �������� �ߺ� ����
        For i = 2 To lastRow
            visitCode = ws.Cells(i, "I").Value
            bikeID = ws.Cells(i, "C").Value
            
            If visitCode <> "" And bikeID <> "" Then
                key = visitCode & "|" & bikeID
                If Not visitBikes.exists(key) Then
                    visitBikes.Add key, True
                    
                    If visitCountDict.exists(visitCode) Then
                        visitCountDict(visitCode) = visitCountDict(visitCode) + 1
                    Else
                        visitCountDict.Add visitCode, 1
                    End If
                End If
            End If
        Next i
        
        ' 2. �� �࿡ �����Ŵ�� ä���
        For i = 2 To lastRow
            visitCode = ws.Cells(i, "I").Value
            If visitCountDict.exists(visitCode) Then
                ws.Cells(i, "K").Value = visitCountDict(visitCode)
            End If
        Next i
    Next sheetIndex
    
    MsgBox "Sheet1 ~ Sheet5 �湮�ڵ庰 �����Ŵ�� ���� �Ϸ�!", vbInformation
End Sub
