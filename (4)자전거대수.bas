Attribute VB_Name = "Module1"
'목적 요약
'I열(방문코드) 기준으로 같은 그룹 내에서 C열(자전거번호)이 중복될 경우
'맨 처음 나타난 행만 정상, → 아래에 반복된 행은 L열에 "가짜"라고 표시하기
'(Sheet1~Sheet5 전체 순회)

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
        
        ' 1. 방문코드 + 자전거번호 조합으로 중복 제거
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
        
        ' 2. 각 행에 자전거대수 채우기
        For i = 2 To lastRow
            visitCode = ws.Cells(i, "I").Value
            If visitCountDict.exists(visitCode) Then
                ws.Cells(i, "K").Value = visitCountDict(visitCode)
            End If
        Next i
    Next sheetIndex
    
    MsgBox "Sheet1 ~ Sheet5 방문코드별 자전거대수 집계 완료!", vbInformation
End Sub
