Attribute VB_Name = "Module1"
Sub StockDataChallenge()
    
    For Each ws In Worksheets
    
        Dim WorksheetName As String
        Dim i As Long
        Dim j As Long
        Dim TickCount As Long
        Dim LastRowA As Long
        Dim LAstROwI As Long
        Dim PerChange As Double
        Dim GreatIncr As Double
        Dim GreatDecr As Double
        Dim GreatVol As Double
        Dim TotalVol As Double
  

        
        'Obtain Worksheetname
        WorksheetName = ws.Name
        
        'Label column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        
        
        'Initialize Variables
        TotalVol = 0
        
        
        'Set Ticker Counter to first row
        TickCount = 2
        
        'Set start row to 2
        j = 2
        
        'FInd the last non blank cell in Column A
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            'Loop Through all rows
            For i = 2 To LastRowA
            
            
            
                TotalVol = TotalVol + ws.Cells(i, 7).Value
            
                'Check if ticker name has changed
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                
                
                
                    'Write ticker in Column I
                    ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value
                    
                    'Write yearly change in column J
                    ws.Cells(TickCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                    
                    'Conditional Formatting for color(Red and Green)
                        If ws.Cells(TickCount, 10).Value < 0 Then
                        ws.Cells(TickCount, 10).Interior.ColorIndex = 3
                        
                        Else
                        ws.Cells(TickCount, 10).Interior.ColorIndex = 4
                        
                        End If
                    
                    'Calculate percent change and write in column K
                    If ws.Cells(j, 3).Value <> 0 Then
                    PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                    ws.Cells(TickCount, 11).Value = Format(PerChange, "Percent")
                    
                    Else
                    
                    ws.Cells(TickCount, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                    
                    'Write total volume in column L
                    
                    ws.Cells(TickCount, 12).Value = TotalVol
                    
                    
                    'Update Variables
                    TickCount = TickCount + 1
                    TotalVol = 0
                    j = i + 1
                End If
                
            
            Next i
        
        'Find the last non blank cell in column I
        LAstROwI = ws.Cells(Rows.Count, 9).End(xlUp).Row
            
            'Loop for greatest value
            For i = 2 To LAstROwI
            
                'For greatest total volume-check if next value is larger and populate ws.Cells
                If ws.Cells(i, 12).Value > GreatVol Then
                GreatVol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
             
                Else
                
                GreatVol = GreatVol
                
                End If
                
                'For greatest increase-check if next value is larger and populate ws.Cells
                If ws.Cells(i, 11).Value > GreatIncr Then
                GreatIncr = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatIncr = GreatIncr
                
                End If
                
                'For greatest decrease-check if next value is smaller and populate ws.Cells
                If ws.Cells(i, 11).Value < GreatDecr Then
                GreatDecr = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatDecr = GreatDecr
                
                End If
                
            'Print Greatest values
            ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
            ws.Cells(4, 17).Value = GreatVol
            
            Next i
            
        Next ws
End Sub






