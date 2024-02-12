Sub StockStats()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim summaryRow As Integer 
    Dim i As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    
    'Loop through each worksheet
    'since all worksheets have identical structure, the same code can be executed on every worksheet w/o any changes
    'ws variable references the current worksheet
    For Each ws In ThisWorkbook.Sheets
    
        'Find the last row of data
        'Assumption that data is continuos on the page and all column have the same length
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        'Set headers of two summary tables
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % increase"
        ws.Range("O3").Value = "Greatest % decrease"
        ws.Range("O4").Value = "Greatest total volume"
        
        'Set the start row in a summary table
        summaryRow = 2
        
        'Initialize variables for greatest increase, decrease, and volume
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        
        'Outer loop through all rows on a sheet.
        'All the calculations are performed inside this loop, therefore the code is iterated on every row in the sheet only once
        'This is critical to provide good performance
        i = 2 'Assuming data starts from row 2
        Do While i <= lastRow
            'The assumption here that a data is sorted by ticker and date ascending (otherwise, the code should be more complicated)
            'That is, the first row for a ticker is corresponded to the first bank day of a year, while the last row for the ticker is corresponded to the last bank Day
            'Therefore, to calculate an yearly change, we should take the opening price from the first row and closing price from the last row
            
            'Initialize variables for a new ticker (at the beginning is it the very first ticker)
            ticker = ws.Cells(i, 1).Value 		'column A
            openPrice = ws.Cells(i, 3).Value	'column C
            totalVolume = 0
            
            'Inner loop through all records with the same ticker value (or until the last row on the sheet)
            'Get the closing price and accumulate a total volume
            Do While ws.Cells(i, 1).Value = ticker And i <= lastRow
                closePrice = ws.Cells(i, 6).Value                   'column F
                totalVolume = totalVolume + ws.Cells(i, 7).Value    'column G
                i = i + 1                                           'increase the variable to move to the next row
            Loop
            
            'Calculate yearly change in absolute values and in per cents. The percentage is calculated based on the opening price of a year
            yearlyChange = closePrice - openPrice
            If openPrice <> 0 Then
                percentageChange = (yearlyChange / openPrice) * 100
            Else
                percentageChange = 0
            End If
            
            'Fill out the first summary table
            ws.Cells(summaryRow, 9).Value = ticker 					'column I
            ws.Cells(summaryRow, 10).Value = yearlyChange 			'column J
            ws.Cells(summaryRow, 11).Value = percentageChange & "%"	'column K
            ws.Cells(summaryRow, 12).Value = totalVolume 			'column L
            
            'Format columns Yearly Change and Percentage Change based on their values. 0 is included into a "green zone"
            If yearlyChange >= 0 Then
                ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0) 'Green color for column J
                ws.Cells(summaryRow, 11).Interior.Color = RGB(0, 255, 0) 'Green color for column K
            Else
                ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0) ' Red color for column J
                ws.Cells(summaryRow, 11).Interior.Color = RGB(255, 0, 0) ' Red color for column K
            End If
            
            'Adjust varible for the next first summary table's row
            summaryRow = summaryRow + 1
                                       
            'Check for greatest increase, decrease, and volume. Assign the current values to variables, if appropriate
            If percentageChange > greatestIncrease Then
                greatestIncrease = percentageChange
                greatestIncreaseTicker = ticker
            End If
            
            If percentageChange < greatestDecrease Then
                greatestDecrease = percentageChange
                greatestDecreaseTicker = ticker
            End If
            
            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                greatestVolumeTicker = ticker
            End If
        Loop
        
        'Fill out the second summary table
        ws.Range("P2") = greatestIncreaseTicker
        ws.Range("Q2").Value = greatestIncrease & "%"
        ws.Range("P3").Value = greatestDecreaseTicker
        ws.Range("Q3").Value = greatestDecrease & "%"
        ws.Range("P4").Value = greatestVolumeTicker
        ws.Range("Q4").Value = greatestVolume
        
        'Autofit columns to adjust their width for long numbers
        ws.Columns.AutoFit
     
    Next ws
End Sub
