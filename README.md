
Sub MultipleStockData():

    For Each ws In Worksheets
        Dim WorksheetName As String
        Dim i As Long
        Dim j As Long
        Dim TickCount As Long
        Dim LastRowA As Long
        Dim LastRowI As Long
        Dim PerChange As Double
        Dim GreatIncr As Double
        Dim GreatDecr As Double
        Dim GreatVol As Double
        
        'Getting the Names of worksheet
        
        WorksheetName = ws.Name
        
        'Creating headers for the column
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(1, 15).Value = "Greatest % Increase"
        ws.Cells(1, 15).Value = "Greatest % Decrease"
        ws.Cells(1, 15).Value = "Greatest Total Volumne"
        
        'set Ticker Counter to first row
        TickCount = 2
        
        'Set strat row to 2
        j = 2
        
        'Find the last non-blank cell in Colum A
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox("Last row in column A is" & LasrRowA)
        
        'Loop through all rows
        For i = 2 To LastRowA
        
            'Check if ticker name changed
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'write ticker in column I
            ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value
            
            'Calculate and write Yearly Change in column J
             ws.Cells(TickCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
             
                'conditional formatting
                If ws.Cells(TickCount, 10).Value < 0 Then
                
                'set cell backgroud color to red
                 ws.Cells(TickCount, 10).Interior.ColorIndex = 3
                 
                 Else
                 
                   'set cell backgroud color to green
                 ws.Cells(TickCount, 10).Interior.ColorIndex = 4
                 
                 End If
                 
                 'Calculate and write percent change in column K
                 If ws.Cells(j, 3).Value <> 0 Then
                 PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                 
                 'Percent formatting
                 ws.Cells(TickCount, 11).Value = Format(PerChange, "Percent")
                 
                 Else
                 
                 ws.Cells(TickCount, 11).Value = Format(0, "Percent")
                 
                 End If
                 
                'Calculate and write total volumn is column L
                ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                'Increase TickCount by 1
                
                TickCount = TickCount + 1
                
                'Set new strat row of the ticker block
                j = i + 1
                
                End If
                
            Next i
            
            'Find last non-blank cell in colum I
            LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
            'MsgBox ( "Last row in column I is" & LastRowI)
            
            'Prepare for Summary
            GreatVol = ws.Cells(2, 12).Value
            GreatIncr = ws.Cells(2, 11).Value
            GreatDecr = ws.Cells(2, 11).Value
            
                'Loop for summary
                For i = 2 To LastRowI
                
                    'For greatest total volume--check if its greater or not
                    If ws.Cells(i, 12).Value > GreatVol Then
                    GreatVol = ws.Cells(i, 12).Value
                    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                    
                    Else
                    
                    GreatVol = GreatVol
                    
                    End If
                    
                    'For greatest increase--check if its greater or not
                    If ws.Cells(i, 11).Value > GreatIncr Then
                    GreatIncr = ws.Cells(i, 11).Value
                    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                    
                    Else
                    
                    GreatIncr = GreatIncr
                    
                    End If
                    
                    'For greatest decrease--check if its greater or not
                    If ws.Cells(i, 11).Value < GreatDecr Then
                    GreatDecr = ws.Cells(i, 11).Value
                    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                    
                    Else
                    
                    End If
                'write summary results
                ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
                ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
                ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
                    
                    Next i
                    
                'adjust column width authomatically
                Worksheets(WorksheetName).Columns("a:Z").AutoFit
            
            Next ws

End Sub
