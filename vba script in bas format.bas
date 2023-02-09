Sub Stockyear():

    For Each ws In Worksheets
    
        Dim WorksheetName As String
        'Present Row
        Dim a As Long
        'Ticker block starting row
        Dim b As Long
        'Counter of index to fill ticker row
        Dim TC As Long
        'A column end row
        Dim LRA As Long
        'I column end row
        Dim LRI As Long
        'Percentage change variable
        Dim PC As Double
        'Instance to calculate the greatest increase
        Dim GI As Double
        'Instance to calculate the greatest decrease
        Dim GD As Double
        'Instances to calculate the greatest volume
        Dim GV As Double
        
        'Save the name of the workbook in a variable
        WN = ws.Name
        
        'Giving names to the header of the column
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Change in year"
        ws.Cells(1, 11).Value = "Change in percentage"
        ws.Cells(1, 12).Value = "Total volume of stock"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Ticker counter has to be set at 2nd row
        TC = 2
        
        'Starting row is set to be at 2
        b = 2
        
        'At last find a blank cell
        LRA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            'Iterative through the rows
            For a = 2 To LRA
            
                'Evaluate the name change of ticker
                If ws.Cells(a + 1, 1).Value <> ws.Cells(a, 1).Value Then
                
                'Ticker is written in column I
                ws.Cells(TC, 9).Value = ws.Cells(a, 1).Value
                
                'Write the yearly change in the Column J
                ws.Cells(TC, 10).Value = ws.Cells(a, 6).Value - ws.Cells(b, 3).Value
                
                    'Conditional formating
                    If ws.Cells(TC, 10).Value < 0 Then
                
                    'Change the colour of background as red
                    ws.Cells(TC, 10).Interior.ColorIndex = 3
                
                    Else
                
                    'Change the background colour as green
                    ws.Cells(TC, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    'Write the change in term of percent in the column k
                    If ws.Cells(b, 3).Value <> 0 Then
                    PC = ((ws.Cells(a, 6).Value - ws.Cells(b, 3).Value) / ws.Cells(b, 3).Value)
                    
                    'Percent formating
                    ws.Cells(TC, 11).Value = Format(PC, "Percent")
                    
                    Else
                    
                    ws.Cells(TC, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                'Write the total volume of the stock in the L column
                ws.Cells(TC, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(b, 7), ws.Cells(a, 7)))
                
                'Increase the value by 1
                TC = TC + 1
                
                'Create the new row for ticker
                b = a + 1
                
                End If
            
            Next a
            
        'At end find the non-blank cell
        LRI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Get the values
        GV = ws.Cells(2, 12).Value
        GI = ws.Cells(2, 11).Value
        GD = ws.Cells(2, 11).Value
        
            'Iterate for the summary
            For a = 2 To LRI
            
                If ws.Cells(a, 12).Value > GV Then
                GV = ws.Cells(a, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(a, 9).Value
                
                Else
                
                GV = GV
                End If
                
                If ws.Cells(a, 11).Value > GI Then
                GI = ws.Cells(a, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(a, 9).Value
                
                Else
                
                GI = GI
                
                End If
                

                If ws.Cells(a, 11).Value < GD Then
                GD = ws.Cells(a, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(a, 9).Value
                
                Else
                
                GD = GD
                
                End If
                
            'Set the values in the column in the ws cells
            ws.Cells(2, 17).Value = Format(GI, "Percent")
            ws.Cells(3, 17).Value = Format(GD, "Percent")
            ws.Cells(4, 17).Value = Format(GV, "Scientific")
            
            Next a
            
        'Adjust the width
        Worksheets(WN).Columns("A:Z").AutoFit
            
    Next ws
        
End Sub
