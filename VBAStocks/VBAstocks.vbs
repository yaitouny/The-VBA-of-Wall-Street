Sub Stocks()

    For Each ws In Worksheets
    
        Dim ticker As String
        Dim yearlyChange As Double
        Dim percentChange As Double
        Dim totalVolume As Double
        Dim lastRow As Long
        Dim lastColumn As Integer
        Dim summaryTableRow As Integer
        Dim worksheetName As String
        Dim yearOpen As Double
        Dim yearClose As Double
        Dim i As Long
        Dim greatestIncrease As Double
        Dim greatestDecrease As Double
        Dim greatestVolume As Double
        Dim maxValue As Double
        Dim minValue As Double
        Dim maxVolume As Double
        Dim maxTicker As String
        Dim minTicker As String
        Dim maxVolumeTicker As String
        
        worksheetName = ws.Name
        summaryTableRow = 2
        totalVolume = 0
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        lastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        maxValue = 0
        minValue = 0
        maxVolume = 0
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest % Total Volume"
        
        yearOpen = ws.Cells(2, 3).Value
        
        For i = 2 To lastRow
            yearClose = ws.Cells(i, 6).Value
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                yearlyChange = yearClose - yearOpen
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
                If yearOpen = 0 Then
                    percentChange = yearClose - yearOpen
                Else:
                    percentChange = yearlyChange / yearOpen
                End If
                
                ws.Range("I" & summaryTableRow).Value = ticker
                ws.Range("J" & summaryTableRow).Value = yearlyChange
                ws.Range("K" & summaryTableRow).Value = percentChange
                ws.Range("L" & summaryTableRow).Value = totalVolume
                
                If yearlyChange > 0 Then
                    ws.Range("J" & summaryTableRow).Interior.ColorIndex = 4
                Else:
                    ws.Range("J" & summaryTableRow).Interior.ColorIndex = 3
                End If
                
                summaryTableRow = summaryTableRow + 1
                yearOpen = ws.Cells(i + 1, 3).Value
                
                If percentChange > maxValue Then
                    maxValue = percentChange
                    maxTicker = ticker
                    'MsgBox ("found the max value")
            
                ElseIf percentChange < minValue Then
                    minValue = percentChange
                    minTicker = ticker
                    'MsgBox ("found the min value")
                End If
                
                If totalVolume > maxVolume Then
                    maxVolume = totalVolume
                    maxVolumeTicker = ticker
                    'MsgBox ("found the max volume")
                End If
            
                percentChange = 0
                totalVolume = 0
                
            Else:
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
            
        Next i
        
        ws.Columns("K").Style = "Percent"
        ws.Range("P2").Value = maxTicker
        ws.Range("Q2").Value = maxValue
        ws.Range("Q2").Style = "Percent"
        ws.Range("P3").Value = minTicker
        ws.Range("Q3").Value = minValue
        ws.Range("Q3").Style = "Percent"
        ws.Range("P4").Value = maxVolumeTicker
        ws.Range("Q4").Value = maxVolume
        
    Next ws
    
End Sub


