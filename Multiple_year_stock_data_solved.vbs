Attribute VB_Name = "Module2"
Sub TickerLoop()

Dim i, column As Integer
Dim ticker As String
Dim lastvalue, totalstockvol, openvalue As Double
Dim lastRow, summary_pos As Integer

'Running for all worksheets in the workbook
For Each ws In Worksheets

    'get the last row in the spreadsheet with values
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    column = 1
    'set cell to start summary table
    summary_pos = 1
    
    ' sort the data by ticker and date -
    ' if the data is not sorted by ticker and date, the code does not work
    ws.Range("A1", ws.Range("G1").End(xlDown)).Sort Key1:=ws.Range("A1"), Order1:=xlAscending, Key2:=ws.Range("B1"), Order1:=xlAscending, Header:=xlYes
    
    'delete columns we are populating first so it is clean
    ws.Range("I:P").EntireColumn.Delete

    'setting column heading
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage change"
    ws.Cells(1, 12).Value = "Total Stock volume"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    
    'format the titles as bold
    ws.Range("I1:L1").Font.Bold = True
    ws.Range("N1:N4").Font.Bold = True
    
    'starting at row 2 loop through all rows
    For i = 2 To lastRow + 1

        ticker = ws.Cells(i, 1).Value
       
        'check summary table to see if a new ticker
        If ws.Cells(summary_pos, 9).Value <> ticker Then
            'found a new ticker so we start populating table
            summary_pos = summary_pos + 1
            ws.Cells(summary_pos, 9).Value = ticker
        
            'for a new ticker, so we need to get the open value and first stock volume
            openvalue = ws.Cells(i, 3).Value
            totalstockvol = ws.Cells(i, 7).Value
    
        Else
            'when it is not a new ticker then just do a running sum of the stock volume
            totalstockvol = totalstockvol + ws.Cells(i, 7).Value
                
        End If
    
        'check next ticker as well
        'if the next ticker after the current one is not the same
        'populate the summary with the current ticker values
        'before moving to the next ticker calculations
        If ws.Cells(i + 1, 1) <> ticker Then
    
            'populate the total stock volume
            ws.Cells(summary_pos, 12).Value = totalstockvol
                    
            'calculate yearly change
            ws.Cells(summary_pos, 10).Value = ws.Cells(i, 6).Value - openvalue
            
            'colour cells depending on the sign
            If ws.Cells(summary_pos, 10).Value >= 0 Then
                'Use RGB colour to shade the cell Green
                ws.Cells(summary_pos, 10).Interior.Color = RGB(0, 150, 68)
            Else
                'Use RGB colour to shade the cell Red
                ws.Cells(summary_pos, 10).Interior.Color = RGB(255, 0, 0)
            End If
                    
            'calculate the % changed - check for div 0
            If openvalue <> 0 Then
                ws.Cells(summary_pos, 11).Value = (ws.Cells(i, 6).Value - openvalue) / openvalue
            Else
                ws.Cells(summary_pos, 11).Value = 0
            End If
            ws.Cells(summary_pos, 11).Value = Format(ws.Cells(summary_pos, 11), "##.##%")


        End If
      
    Next i

    'Calculate and set the Greatest % Increase, Greatest % Decrease and Greatest Total Volume
    ws.Cells(2, 16) = Application.WorksheetFunction.Max(ws.Range("k:k"))
    ws.Cells(2, 15).Value = ws.Cells(Application.WorksheetFunction.Match(ws.Cells(2, 16).Value, ws.Range("K1:K" & summary_pos), 0), 9).Value
    ws.Cells(2, 16).Value = Format(ws.Cells(2, 16), "##.##%")

    ws.Cells(3, 16) = Application.WorksheetFunction.Min(ws.Range("k:k"))
    ws.Cells(3, 15).Value = ws.Cells(Application.WorksheetFunction.Match(ws.Cells(3, 16).Value, ws.Range("K1:K" & summary_pos), 0), 9).Value
    ws.Cells(3, 16).Value = Format(ws.Cells(3, 16), "##.##%")

    ws.Cells(4, 16) = Application.WorksheetFunction.Max(ws.Range("L:L"))
    ws.Cells(4, 15).Value = ws.Cells(Application.WorksheetFunction.Match(ws.Cells(4, 16).Value, ws.Range("L1:L" & summary_pos), 0), 9).Value

    'autofit the columns to make it easy to read
    ws.UsedRange.EntireColumn.AutoFit
    
Next ws

MsgBox "yes!!"

End Sub

