Attribute VB_Name = "Module1"
Sub vba_macro_test():
    
    'Setting up for Each worksheets
    For Each ws In Worksheets
        'Setting Labels
        ws.Range("J1,Q1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Total Stock Volume"
        ws.Range("R1").Value = "Value"
        ws.Range("P2").Value = "Greatest % Increase"
        ws.Range("P3").Value = "Greatest % Decrease"
        ws.Range("P4").Value = "Greatest Total Volume"
        
        'Setting variables
        Dim WorksheetName As String
        Dim Rownumber As Long
        Dim Columnnumber As Long
        Dim SummaryTable As Long
        Dim LastRowA As Long
        Dim LastRowI As Long
        Dim PercentChange As Double
        Dim GreatestI As Double
        Dim GreatestD As Double
        Dim GreatestVol As Double
        
        'Getting worksheet names
        WorksheetName = ws.Name
        'Setting SummaryTable to read the second Row and start for yearly change
        SummaryTable = 2
        'starting new row for Yearly change and percent change calculation
        Start = 2

        'Finding last row
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Calculations for Yearly_Change, Percent_Change, & Total_Stock _Volume + Loop
            For Rownumber = 2 To LastRowA
            '"+1" is looking forward in a string "<>" is not equal too
                If Cells(Rownumber + 1, 1).Value <> Cells(Rownumber, 1).Value Then
                
                    'Setting Ticker
                ws.Cells(SummaryTable, 10).Value = ws.Cells(Rownumber, 1).Value
                
                    'Setting Yearly Change
                ws.Cells(SummaryTable, 11).Value = ws.Cells(Rownumber, 6).Value - ws.Cells(Start, 3).Value
                     
                'Setting Conditional formating red and green
                    If ws.Cells(SummaryTable, 11).Value < 0 Then
                
                    ws.Cells(SummaryTable, 11).Interior.ColorIndex = 3

                    Else
                    ws.Cells(SummaryTable, 11).Interior.ColorIndex = 4
                
                    End If
                    'Setting PercentChange
                    If ws.Cells(Start, 3).Value <> 0 Then
                    PercentChange = ((ws.Cells(Rownumber, 6).Value - ws.Cells(Start, 3).Value) / ws.Cells(Start, 3).Value)
                    
                    'Setting percent change formattting
                    ws.Cells(SummaryTable, 12).Value = Format(PercentChange, "Percent")
                    
                    Else
                    
                    ws.Cells(SummaryTable, 12).Value = Format(0, "Percent")
                    
                    End If
                
                'Setting Total Volume
                ws.Cells(SummaryTable, 13).Value = WorksheetFunction.Sum(Range(ws.Cells(Start, 7), ws.Cells(Rownumber, 7)))
                
                'Increasing SummaryTable by 1
                SummaryTable = SummaryTable + 1
            
                'Setting new start row
                Start = Rownumber + 1
                
                End If
            
            Next Rownumber
        
        'Find last cell in Column J
        LastRowI = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
        'Setting up the Summary of values
        GreatestVol = ws.Cells(2, 13).Value
        GreatestI = ws.Cells(2, 12).Value
        GreatestD = ws.Cells(2, 12).Value
        
            'Summary portion loop
            For Rownumber = 2 To LastRowI
            
                'Setting greatest total Volumn
                If ws.Cells(Rownumber, 13).Value > GreatestVol Then
                GreatestVol = ws.Cells(Rownumber, 13).Value
                ws.Cells(4, 17).Value = ws.Cells(Rownumber, 10).Value
                
                Else
                
                GreatestVol = GreatestVol
                
                End If
                
                'Setting Greatest Increase
                If ws.Cells(Rownumber, 12).Value > GreatestI Then
                GreatestI = ws.Cells(Rownumber, 12).Value
                ws.Cells(2, 17).Value = ws.Cells(Rownumber, 10).Value
                
                Else
                
                GreatestI = GreatestI
                
                End If
                
                'Setting Greatest Decrease
                If ws.Cells(Rownumber, 12).Value < GreatestD Then
                GreatestD = ws.Cells(Rownumber, 12).Value
                ws.Cells(3, 17).Value = ws.Cells(Rownumber, 10).Value
                
                Else
                
                GreatestD = GreatestD
                
                End If
                
            'Write summary results in ws.Cells with formatting
            ws.Cells(2, 18).Value = Format(GreatestI, "Percent")
            ws.Cells(3, 18).Value = Format(GreatestD, "Percent")
            ws.Cells(4, 18).Value = Format(GreatestVol, "Scientific")
            
            Next Rownumber
         'Adjusting Column formatting
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
    Next ws
        
End Sub

