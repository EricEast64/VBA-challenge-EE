Attribute VB_Name = "Module1"

   
Sub stock_analysis()
    
    'Create variables
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim AnnualChange As Double
    Dim TotalVolume As LongLong
    
    'Create variable for populating summary table rows
    Dim FillRow As Integer

    OpenPrice = 0
    ClosePrice = 0
    TotalVolume = 0
    
    'Loop through worksheets in excel file
    For Each ws In Worksheets
    
        'Find the last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Print the header for the summary table
        ws.Range("I1").Value = "Ticker Symbol"
        ws.Range("J1").Value = "Annual Growth"
        ws.Range("K1").Value = "Annual Growth (%)"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Set FillRow to 2 as the file contains headers
        FillRow = 2
        
        'Begin loop through the worksheet
        For i = 2 To LastRow
            
            'Check if the ticker value has changed from preceding row
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                'Collect the opening price here as we know it is the first price for that stock
                OpenPrice = ws.Cells(i, 3).Value
                       
            'Check if the ticker value will change in following row
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Set the Ticker as value if last row of that ticker
                Ticker = ws.Cells(i, 1).Value
                
                'Fill in ticker symbols
                ws.Range("I" & FillRow).Value = Ticker
                
                'Keep a running total of stock volume
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                    
                'Collect closing price as we know it is the last price for that stock
                ClosePrice = ws.Cells(i, 6).Value
            
                'Add the annual stock growth (absolute)
                AnnualChange = ClosePrice - OpenPrice
                ws.Range("J" & FillRow).Value = AnnualChange
                
                'Add the annual stock growth (% change)
                If OpenPrice = 0 Then
                    ws.Range("K" & FillRow).Value = 0
                Else
                    ws.Range("K" & FillRow).Value = AnnualChange / OpenPrice
                End If
                            
                'Conditional formatting for the colors for stock growth or loss in value
                If AnnualChange > 0 Then
                    ws.Range("J" & FillRow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & FillRow).Interior.ColorIndex = 3
                End If
                                
                'Fill in the total stock volume
                ws.Range("L" & FillRow).Value = TotalVolume
                          
                'Bump down to the next summary table row after filling in all columns
                FillRow = FillRow + 1
             
                'Reset the values before moving to next ticker
                OpenPrice = 0
                ClosePrice = 0
                TotalVolume = 0
                               
                If ws.Cells(i, 1) = LastRow Then
                FillRow = 0
                End If
                
            'If the cell immediately following a row is the same TS
            Else
            
                'Keep a running total of stock volume
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                
            End If
            
        Next i
        
    Next ws
    
End Sub


