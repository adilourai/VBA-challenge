Attribute VB_Name = "Module1"
Sub stocks()

    Dim i As Integer
    Dim stockcount As Integer
    Dim openprice As Single
    Dim closeprice As Single
    Dim volume As Double
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    lastrow1 = Range("A1").End(xlDown).Row
    
    stockcount = 1
    
    For i = 2 To lastrow1
        If Cells(i, 2).Value Like "????0102" Then
            Cells(stockcount + 1, 9).Value = Cells(i, 1).Value
            openprice = Cells(i, 3).Value
            volume = Cells(i, 7).Value
        ElseIf Not (Cells(i, 2).Value Like "????01012") And Not (Cells(i, 2).Value Like "????1231") Then
            volume = volume + Cells(i, 7).Value
        ElseIf Cells(i, 2).Value Like "????1231" Then
            closeprice = Cells(i, 6).Value
            volume = volume + Cells(i, 7).Value
            Cells(stockcount + 1, 12).Value = volume
            Cells(stockcount + 1, 10).Value = Format((closeprice - openprice), "#.00")
            Cells(stockcount + 1, 11).Value = FormatPercent((Cells(stockcount + 1, 10).Value) / openprice)
            stockcount = stockcount + 1
        End If
    Next i
    
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    Dim lastrow2 As Long
    lastrow2 = Range("I1").End(xlDown).Row
    
    For l = 2 To (lastrow2 - 1)
        If Cells(l, 11).Value < 0 Then
            Cells(l, 11).Interior.Color = RGB(255, 0, 0)
        ElseIf Cells(l, 11).Value > 0 Then
            Cells(l, 11).Interior.Color = RGB(0, 255, 0)
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
        percentchange(j) = Cells(j, 11).Value
        totalvolume(j) = Cells(j, 12).Value
    Next
    
    maxincrease = WorksheetFunction.Max(percentchange)
    maxdecrease = WorksheetFunction.Min(percentchange)
    maxvolume = WorksheetFunction.Max(totalvolume)
    
    For k = 2 To lastrow2
        If Cells(k, 11).Value = maxincrease Then
            Range("P2").Value = Cells(k, 9).Value
        ElseIf Cells(k, 11).Value = maxdecrease Then
            Range("P3").Value = Cells(k, 9).Value
        ElseIf Cells(k, 12).Value = maxvolume Then
            Range("P4").Value = Cells(k, 9).Value
        ElseIf Cells(k, 11).Value < 0 Then
            Cells(k, 11).Interior.Color = RGB(255, 0, 0)
        ElseIf Cells(k, 11).Value >= 0 Then
            Cells(k, 11).Interior.Color = RGB(0, 255, 0)
        End If
    Next k
    
    Range("Q2").Value = FormatPercent(maxincrease)
    Range("Q3").Value = FormatPercent(maxdecrease)
    Range("Q4").Value = maxvolume
    
    Columns("I:Q").AutoFit
    
End Sub
