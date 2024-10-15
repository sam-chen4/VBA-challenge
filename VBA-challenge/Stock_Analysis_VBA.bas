Sub Stock_Analysis()
    
    '===========================
    '   DECLARE VARIABLES
    '===========================
    Dim ws As Worksheet
    Dim i As Long
    Dim j As Long
    Dim lastRow As Long
    Dim Ticker_Count As Long
    Dim Percent_Change As Single
    Dim Greatest_Increase As Single
    Dim Greatest_Decrease As Single
    Dim Greatest_Total_Volume As Double
    Dim lastrowJ As Long
    Dim Ticker_Greatest_Increase As String
    Dim Ticker_Greatest_Decrease As String
    Dim Ticker_Greatest_Total_Volume As String
    
    '=========================
    '   SET WORKSHEET
    '=========================
    For Each ws In ThisWorkbook.Worksheets
    
    
    
    '===========================
    '   HEADERS FOR COLUMNS
    '===========================
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Quaterly Change($)"
    ws.Cells(1, 12).Value = "Percentage Change"
    ws.Cells(1, 13).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    '================
    '   SET COUNT
    '================
    Ticker_Count = 2
    j = 2
    
    '=========================
    '   GET LAST ROW
    '=========================
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    '==============================================
    '   LOOP THROUGH ROWS AND SET INTERIOR COLOUR
    '==============================================
    For i = 2 To lastRow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ws.Cells(Ticker_Count, 10).Value = ws.Cells(i, 1).Value
        ws.Cells(Ticker_Count, 11).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
        
        If ws.Cells(Ticker_Count, 11).Value < 0 Then
        ws.Cells(Ticker_Count, 11).Interior.ColorIndex = 3
        
        ElseIf ws.Cells(Ticker_Count, 11).Value > 0 Then
        ws.Cells(Ticker_Count, 11).Interior.ColorIndex = 4
        
        ElseIf ws.Cells(Ticker_Count, 11).Value = 0 Then
        ws.Cells(Ticker_Count, 11).Interior.ColorIndex = xlNone
        
        End If
        
        If ws.Cells(j, 3).Value <> 0 Then
        Percent_Change = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
        
        ws.Cells(Ticker_Count, 12).Value = Format(Percent_Change, "Percent")
        
        Else
        ws.Cells(Ticker_Count, 12).Value = Format(0, "Percent")
        
            End If
        
        ws.Cells(Ticker_Count, 13).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
        
        Ticker_Count = Ticker_Count + 1
        
        j = i + 1
        
            End If
            
    Next i
    
    '==========================
    '  GET LAST ROW
    '==========================
    lastrowJ = ws.Cells(ws.Rows.Count, 10).End(xlUp).Row
    
    '===================
    '   SET VALUES
    '===================
    Greatest_Increase = ws.Cells(2, 12).Value
    Greatest_Decrease = ws.Cells(2, 12).Value
    Greatest_Total_Volume = ws.Cells(2, 13).Value
    Ticker_Greatest_Increase = ws.Cells(2, 10).Value
    Ticker_Greatest_Decrease = ws.Cells(2, 10).Value
    Ticker_Greatest_Total_Volume = ws.Cells(2, 10).Value
    
    '======================
    '   LOOP THROUGH ROWS
    '======================
    For i = 2 To lastrowJ
    
    If ws.Cells(i, 12).Value > Greatest_Increase Then
    Greatest_Increase = ws.Cells(i, 12).Value
    ws.Cells(2, 16).Value = ws.Cells(i, 10).Value
    
    End If

    If ws.Cells(i, 12).Value < Greatest_Decrease Then
    Greatest_Decrease = ws.Cells(i, 12).Value
    ws.Cells(3, 16).Value = ws.Cells(i, 10).Value
    
    End If
    
    If ws.Cells(i, 13).Value > Greatest_Total_Volume Then
    Greatest_Total_Volume = ws.Cells(i, 13).Value
    ws.Cells(4, 16).Value = ws.Cells(i, 10).Value
    
    End If
    
    ws.Cells(2, 17).Value = Format(Greatest_Increase, "Percent")
    ws.Cells(3, 17).Value = Format(Greatest_Decrease, "Percent")
    ws.Cells(4, 17).Value = Format(Greatest_Total_Volume, "Scientific")
    
        Next i
        Next ws
End Sub
