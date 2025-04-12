Attribute VB_Name = "Module1"
'Sheet1부터 Sheet5까지 순회하며 방문코드별 순번 넣기 (J열)

Sub AssignSequenceByVisitCode()
    Dim ws As Worksheet
    Dim sheetIndex As Integer
    
    For sheetIndex = 1 To 5
        Set ws = ThisWorkbook.Sheets("Sheet" & sheetIndex)
        
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.count, "I").End(xlUp).Row
        
        Dim prevCode As String
        Dim count As Long
        count = 0
        
        Dim i As Long
        For i = 2 To lastRow
            Dim currCode As String
            currCode = ws.Cells(i, "I").Value
            
            If currCode = "" Then Exit For
            
            If currCode <> prevCode Then
                count = 1
            Else
                count = count + 1
            End If
            
            ws.Cells(i, "J").Value = count
            prevCode = currCode
        Next i
    Next sheetIndex
End Sub
