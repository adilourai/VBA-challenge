Attribute VB_Name = "FINALStocksCode"
Sub stocks()
    
    'Declare worksheet variable
    Dim ws As Worksheet
    
    'Begin for loop to loop through all worksheets in workbook
    For Each ws In Worksheets
    
    'Declare variables for first batch of calculated data
        Dim stockcount As Integer
        Dim openprice As Single
        Dim closeprice As Single
        Dim volume As Double
        
    'Add column headers for the first batch of calculated data
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
    'Determine length of dataset in each worksheet
        lastrow1 = ws.Range("A1").End(xlDown).Row
        
    'Set stockcount variable to 1 in order to set destination cells for newly calculated data
        stockcount = 1
        
    'Begin for loop to gather data on each stock's yearly change in price and total volume for the year
        For i = 2 To lastrow1
    'Find starting row for stock based on first date of the year, grab stock ticker symbol, opening price, and begin totaling annual volume
            If ws.Cells(i, 2).Value Like "????0102" Then
                ws.Cells(stockcount + 1, 9).Value = ws.Cells(i, 1).Value
                openprice = ws.Cells(i, 3).Value
                volume = ws.Cells(i, 7).Value
    'Continue summing volume from all the dates between start and end of the year
            ElseIf Not (ws.Cells(i, 2).Value Like "????01012") And Not (ws.Cells(i, 2).Value Like "????1231") Then
                volume = volume + ws.Cells(i, 7).Value
    'Find ending row for stock based on last date of the year, grab closing price, finish summing volume, calculate yearly change, and place calculated values in new columns
            ElseIf ws.Cells(i, 2).Value Like "????1231" Then
                closeprice = ws.Cells(i, 6).Value
                volume = volume + ws.Cells(i, 7).Value
                ws.Cells(stockcount + 1, 12).Value = volume
                ws.Cells(stockcount + 1, 10).Value = closeprice - openprice
    'Apply number formatting to new data
                ws.Cells(stockcount + 1, 10).NumberFormat = "0.00"
                ws.Cells(stockcount + 1, 11).Value = FormatPercent((ws.Cells(stockcount + 1, 10).Value) / openprice)
    'Add 1 to stockcount so next stock's data gets added to the right row
                stockcount = stockcount + 1
            End If
        Next i
        
    'Determine length of dataset from the newly calculated summary data
        Dim lastrow2 As Long
        lastrow2 = ws.Range("I1").End(xlDown).Row
        
    'Begin for loop to add conditional formatting - green for positive change and red for negative change in price
        For l = 2 To (lastrow2)
            If ws.Cells(l, 10).Value < 0 Then
                ws.Cells(l, 10).Interior.Color = RGB(255, 0, 0)
            ElseIf ws.Cells(l, 10).Value > 0 Then
                ws.Cells(l, 10).Interior.Color = RGB(0, 255, 0)
            End If
        Next l
        
    'Add column headers for second batch of calculated data
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
    
    'Declare variables for second batch of calculated data
        Dim maxincrease As Single
        Dim maxdecrease As Single
        Dim maxvolume As Double
        Dim maxincreaseticker As String
        Dim maxdecreaseticker As String
        Dim maxvolumeticker As String
    
    'Declare arrays to store percent change and volume, using ReDim to account for the length of the newly calculated dataset
        Dim percentchange As Variant
        ReDim percentchange(2 To lastrow2)
        Dim totalvolume As Variant
        ReDim totalvolume(2 To lastrow2)
    
    'Fill arrays with the percent change and volume values from the first calculated dataset
        For j = 2 To lastrow2
            percentchange(j) = ws.Cells(j, 11).Value
            totalvolume(j) = ws.Cells(j, 12).Value
        Next
        
    'Calculate the max increase and max decrease in percent change, and the max volume
        maxincrease = WorksheetFunction.Max(percentchange)
        maxdecrease = WorksheetFunction.Min(percentchange)
        maxvolume = WorksheetFunction.Max(totalvolume)
    
    'Begin for loop to grab the ticker symbols for the stocks with max increase, max decrease, and max volume
        For k = 2 To lastrow2
            If ws.Cells(k, 11).Value = maxincrease Then
                ws.Range("P2").Value = ws.Cells(k, 9).Value
            ElseIf ws.Cells(k, 11).Value = maxdecrease Then
                ws.Range("P3").Value = ws.Cells(k, 9).Value
            ElseIf ws.Cells(k, 12).Value = maxvolume Then
                ws.Range("P4").Value = ws.Cells(k, 9).Value
            End If
        Next k
    
    'Place values for second calculated set of data in proper cells, formatted correctly
        ws.Range("Q2").Value = FormatPercent(maxincrease)
        ws.Range("Q3").Value = FormatPercent(maxdecrease)
        ws.Range("Q4").Value = maxvolume
        
    'Adjust the column widths to fit all the new data precisely
        ws.Columns("I:Q").AutoFit
    
    'Loop through next worksheet
    Next ws
    
End Sub

