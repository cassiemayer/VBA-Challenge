Sub StockDataAnalysis()
    
    'Declare variables
    Dim ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    
    'Initialize variables
    ticker = ""
    OpenPrice = 0
    ClosePrice = 0
    YearlyChange = 0
    PercentChange = 0
    TotalStockVolume = 0
    
    'Find last row of data
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Create summary table headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    'Loop through each row of data
    For i = 2 To lastRow
        'Check if we are still in the same ticker
        If Cells(i, 1).Value <> ticker Then
            'If we are not in the same ticker, record the summary data for the previous ticker
            If ticker <> "" Then
                'Calculate the yearly change, percent change, and record the summary data in the summary table
                YearlyChange = ClosePrice - OpenPrice
                If OpenPrice <> 0 Then
                    PercentChange = YearlyChange / OpenPrice
                Else
                    PercentChange = 0
                End If
                Range("I" & j + 1).Value = ticker
                Range("J" & j + 1).Value = YearlyChange
                If YearlyChange <= 0 Then
                    Range("J" & j + 1).Interior.Color = vbRed
                Else
                    Range("J" & j + 1).Interior.Color = vbGreen
                End If
                Range("K" & j + 1).Value = PercentChange
                Range("L" & j + 1).Value = TotalStockVolume
            End If
            'Record the new ticker and reset the summary data variables
            ticker = Cells(i, 1).Value
            OpenPrice = Cells(i, 3).Value
            ClosePrice = Cells(i, 6).Value
            TotalStockVolume = Cells(i, 7).Value
            j = j + 1
        Else
            'If we are in the same ticker, add the volume to the total and update the closing price
            TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
            ClosePrice = Cells(i, 6).Value
        End If
    Next i
    
    'Format the summary table
    Range("J2:J" & j).NumberFormat = "0.00"
    Range("K2:K" & j).NumberFormat = "0.00%"
    
    Cells(1, "P").Value = "Ticker"
    Cells(1, "Q").Value = "Value"
    Cells(2, "O").Value = "Greatest % Increase"
    Cells(3, "O").Value = "Greatest % Decrease"
    Cells(4, "O").Value = "Greatest Total Volume"
    
    ' Determine the last row of data
    Dim last_sum_Row As Long
    last_sum_Row = Cells(Rows.Count, "I").End(xlUp).Row
    
    ' Initialize variables to store maximum and minimum values and ticker symbols
    Dim maxPercentChange As Double
    Dim minPercentChange As Double
    Dim maxVolume As Double
    Dim minVolume As Double
    Dim maxYearlyChange As Double
    Dim minYearlyChange As Double
    Dim maxPercentChangeTicker As String
    Dim minPercentChangeTicker As String
    Dim maxVolumeTicker As String
    Dim minVolumeTicker As String
    Dim maxYearlyChangeTicker As String
    Dim minYearlyChangeTicker As String
    
    ' Loop through the data and find the maximum and minimum values and ticker symbols
    For i = 2 To last_sum_Row
        ' Percent change
        If Cells(i, "K").Value > maxPercentChange Then
            maxPercentChange = Cells(i, "K").Value
            maxPercentChangeTicker = Cells(i, "I").Value
        ElseIf Cells(i, "K").Value < minPercentChange Then
            minPercentChange = Cells(i, "K").Value
            minPercentChangeTicker = Cells(i, "I").Value
        End If
        
        ' Volume
        If Cells(i, "L").Value > maxVolume Then
            maxVolume = Cells(i, "L").Value
            maxVolumeTicker = Cells(i, "I").Value
        ElseIf Cells(i, "L").Value < minVolume Then
            minVolume = Cells(i, "L").Value
            minVolumeTicker = Cells(i, "I").Value
        End If
        
        ' Yearly change
        If Cells(i, "J").Value > maxYearlyChange Then
            maxYearlyChange = Cells(i, "J").Value
            maxYearlyChangeTicker = Cells(i, "I").Value
        ElseIf Cells(i, "J").Value < minYearlyChange Then
            minYearlyChange = Cells(i, "J").Value
            minYearlyChangeTicker = Cells(i, "I").Value
        End If
    Next i
    
    Cells(2, "P").Value = maxPercentChangeTicker
    Cells(2, "Q").Value = Format(maxPercentChange, "0.00%")
    Cells(3, "P").Value = minPercentChangeTicker
    Cells(3, "Q").Value = Format(minPercentChange, "0.00%")
    Cells(4, "P").Value = maxVolumeTicker
    Cells(4, "Q").Value = Format(maxVolume, "#,##0")
    
End Sub


