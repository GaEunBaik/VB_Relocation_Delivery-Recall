Attribute VB_Name = "Module1"
' Sheet1부터 Sheet5까지 자동으로 돌리면서, 방문코드(A0000001 → A0000002 → ...)를 연속적으로 이어붙이기
' Sheet1부터 5까지 D컬럼을 기준으로 H컬럼이 30분 이상일 때 I컬럼에 방문코드 A0000001부터 생성

Sub GenerateVisitCodesAcrossSheets()
    Dim ws As Worksheet
    Dim sheetIndex As Integer
    Dim currentCodeIndex As Long
    currentCodeIndex = 1 ' A0000001부터 시작
    
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
                ' Sheet1의 첫 행
                ws.Cells(i, "I").Value = visitCode
                prevTime = currTime
                prevLocation = currLocation
            ElseIf i = 2 Then
                ' 다른 시트의 첫 행
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
