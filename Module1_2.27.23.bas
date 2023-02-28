Attribute VB_Name = "Module1"
Sub stocks()
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
        Dim i As Integer
        Dim stockcount As Integer
        Dim openprice As Single
        Dim closeprice As Single
        Dim volume As Double
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        lastrow1 = ws.Range("A1").End(xlDown).Row
        
        stockcount = 1
        
        For i = 2 To lastrow1
            If ws.Cells(i, 2).Value Like "????0102" Then
                ws.Cells(stockcount + 1, 9).Value = ws.Cells(i, 1).Value
                openprice = ws.Cells(i, 3).Value
                volume = ws.Cells(i, 7).Value
            ElseIf Not (ws.Cells(i, 2).Value Like "????01012") And Not (ws.Cells(i, 2).Value Like "????1231") Then
                volume = volume + ws.Cells(i, 7).Value
            ElseIf ws.Cells(i, 2).Value Like "????1231" Then
                closeprice = ws.Cells(i, 6).Value
                volume = volume + ws.Cells(i, 7).Value
                ws.Cells(stockcount + 1, 12).Value = volume
                ws.Cells(stockcount + 1, 10).Value = Format((closeprice - openprice), "#.00")
                ws.Cells(stockcount + 1, 11).Value = FormatPercent((ws.Cells(stockcount + 1, 10).Value) / openprice)
                stockcount = stockcount + 1
            End If
        Next i
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        Dim lastrow2 As Long
        lastrow2 = ws.Range("I1").End(xlDown).Row
        
        For l = 2 To (lastrow2 - 1)
            If ws.Cells(l, 11).Value < 0 Then
                ws.Cells(l, 11).Interior.Color = RGB(255, 0, 0)
            ElseIf ws.Cells(l, 11).Value > 0 Then
                ws.Cells(l, 11).Interior.Color = RGB(0, 255, 0)
            End If
        Next l
        
        Dim maxincrease As Single
        Dim maxdecrease As Single
        Dim maxvolume As Double
        Dim maxincreaseticker As String
        Dim maxdecreaseticker As String
        Dim maxvolumeticker As String
        
        Dim percentchange As Variant
        ReDim percentchange(2 To lastrow2)
        Dim totalvolume As Variant
        ReDim totalvolume(2 To lastrow2)
        
        For j = 2 To lastrow2
            percentchange(j) = ws.Cells(j, 11).Value
            totalvolume(j) = ws.Cells(j, 12).Value
        Next
        
        maxincrease = WorksheetFunction.Max(percentchange)
        maxdecrease = WorksheetFunction.Min(percentchange)
        maxvolume = WorksheetFunction.Max(totalvolume)
        
        For k = 2 To lastrow2
            If ws.Cells(k, 11).Value = maxincrease Then
                ws.Range("P2").Value = ws.Cells(k, 9).Value
            ElseIf ws.Cells(k, 11).Value = maxdecrease Then
                ws.Range("P3").Value = ws.Cells(k, 9).Value
            ElseIf ws.Cells(k, 12).Value = maxvolume Then
                ws.Range("P4").Value = ws.Cells(k, 9).Value
            ElseIf ws.Cells(k, 11).Value < 0 Then
                ws.Cells(k, 11).Interior.Color = RGB(255, 0, 0)
            ElseIf ws.Cells(k, 11).Value >= 0 Then
                ws.Cells(k, 11).Interior.Color = RGB(0, 255, 0)
            End If
        Next k
        
        ws.Range("Q2").Value = FormatPercent(maxincrease)
        ws.Range("Q3").Value = FormatPercent(maxdecrease)
        ws.Range("Q4").Value = maxvolume
        
        ws.Columns("I:Q").AutoFit
    
    Next ws
    
End Sub
