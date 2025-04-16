Attribute VB_Name = "Module1"
'목적 요약
'I열(방문코드) 기준으로 같은 그룹 내에서 C열(자전거번호)이 중복될 경우
'맨 처음 나타난 행만 정상, → 아래에 반복된 행은 L열에 "가짜"라고 표시하기
'(Sheet1~Sheet5 전체 순회)

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
                    ws.Cells(i, "L").Value = "가짜"
                Else
                    seenDict.Add key, True
                    ws.Cells(i, "L").ClearContents ' 혹시 기존에 가짜가 있으면 지움
                End If
            End If
        Next i
    Next sheetIndex
    
    MsgBox "모든 시트 처리 완료되었습니다!", vbInformation
End Sub
