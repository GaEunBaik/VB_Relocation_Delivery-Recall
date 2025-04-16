Attribute VB_Name = "Module1"
'목적 요약
'I열 (방문코드)별로 그룹을 나누고, 각 방문코드 그룹에서:
'J열 = 1인 행의 H열 값을 M열에 넣기 (첫 시간)
' 그 방문코드 내에서 마지막 순번(J열의 최대값)인 행의 H열 값을 N열에 넣기 (마지막 시간)
'기준은 Sheet1 ~ Sheet5 전부에 대해 적용합니다.


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
        
        ' 1. 방문코드별 첫/끝 row 위치 저장
        For i = 2 To lastRow
            visitCode = ws.Cells(i, "I").Value
            
            If visitCode <> "" Then
                Dim seq As Variant
                seq = ws.Cells(i, "J").Value
                
                If seq = 1 Then
                    firstRowDict(visitCode) = i
                End If
                
                ' 방문코드별 가장 큰 순번을 찾기 위해 계속 갱신
                If Not lastRowDict.exists(visitCode) Then
                    lastRowDict(visitCode) = i
                ElseIf ws.Cells(i, "J").Value > ws.Cells(lastRowDict(visitCode), "J").Value Then
                    lastRowDict(visitCode) = i
                End If
            End If
        Next i
        
        ' 2. 해당 행에 H열 값을 M/N열에 입력
        Dim key As Variant
        For Each key In firstRowDict.Keys
            Dim firstRow As Long, lastRowNum As Long
            firstRow = firstRowDict(key)
            lastRowNum = lastRowDict(key)
            
            ws.Cells(firstRow, "M").Value = ws.Cells(firstRow, "H").Value
            ws.Cells(lastRowNum, "N").Value = ws.Cells(lastRowNum, "H").Value
        Next key
    Next sheetIndex
    
    MsgBox "첫/마지막 시간 기록 완료!", vbInformation
End Sub
