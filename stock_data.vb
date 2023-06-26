Sub stockData()

For Each ws In Worksheets

    'define the variables
    Dim ticker As String
    Dim maxTicker As String
    Dim minTicker As String
    Dim greatVol As String
    Dim tickerRow As Integer
    Dim openS As Double
    Dim yearlyChange As Double
    Dim perChange As Double
    Dim volume As Double
    Dim maxPer As Double
    Dim minPer As Double
    Dim maxVol As Double
    
    'initializes the variables
    tickerRow = 2
    openS = 0
    yearlyChange = 0
    perChange = 0
    volume = 0
    maxPer = 0
    minPer = 0
    maxVol = 0
    
    'add headers
    ws.Range("I1").Value = "ticker"
    ws.Range("J1").Value = "yearly change"
    ws.Range("K1").Value = "percent change"
    ws.Range("L1").Value = "total stock volume"
    ws.Range("P1").Value = "ticker"
    ws.Range("Q1").Value = "value"
    ws.Range("O2").Value = "greatest % increase"
    ws.Range("O3").Value = "greatest % decrease"
    ws.Range("O4").Value = "greatest total volume"
    
    'loop through the ticker to combine them all
    For i = 2 To 753001
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'moves through the ticker cell
            ticker = ws.Cells(i, 1).Value
            
            'calculate the yearly change
            yearlyChange = ws.Cells(i, 6).Value - openS
            
            'calculate the percent change
            perChange = yearlyChange / openS
            
            'calculate the total volume
            volume = volume + ws.Cells(i, 7).Value
            
            'max percent
            If (perChange > maxPer) Then
                maxPer = perChange
                maxTicker = ticker
            End If
            
            'min percent
            If (perChange < minPer) Then
                minPer = perChange
                minTicker = ticker
            End If
            
            'max volume
            If (volume > maxVol) Then
                maxVol = volume
                greatVol = ticker
            End If
            
            'adds each variable to the columns
            ws.Range("I" & tickerRow).Value = ticker
            ws.Range("J" & tickerRow).Value = yearlyChange
            ws.Range("K" & tickerRow).Value = FormatPercent(perChange)
            ws.Range("L" & tickerRow).Value = volume
            
            'add color for positive or negative values
            If (yearlyChange > 0) Then
                ws.Range("J" & tickerRow).Interior.ColorIndex = 4
            ElseIf (yearlyChange < 0) Then
                ws.Range("J" & tickerRow).Interior.ColorIndex = 3
            End If
            
            If (perChange > 0) Then
                ws.Range("K" & tickerRow).Interior.ColorIndex = 4
            ElseIf (perChange < 0) Then
                ws.Range("K" & tickerRow).Interior.ColorIndex = 3
            End If

            'goes to the next row
            tickerRow = tickerRow + 1
            
            'resets the variables
            yearlyChange = 0
            perChange = 0
            volume = 0
            openS = ws.Cells(i + 1, 3).Value
            
        ElseIf i = 2 Then
            'stores the first opening value for each ticker
            openS = ws.Cells(i + 1, 3).Value
            
        Else
            'adds the total volume for each ticker
            volume = volume + ws.Cells(i, 7).Value
            
        End If
        
    Next i
    
    'assign the min and max to the cells
    ws.Range("P2").Value = maxTicker
    ws.Range("P3").Value = minTicker
    ws.Range("P4").Value = greatVol
    ws.Range("Q2").Value = FormatPercent(maxPer)
    ws.Range("Q3").Value = FormatPercent(minPer)
    ws.Range("Q4").Value = maxVol
    
Next ws

End Sub