Attribute VB_Name = "Module1"
Sub market()
    
    For Each ws In Worksheets
    
    'Variables
    Dim row As Integer
    Dim vol, vol_total, volume As Variant
    Dim ticker As String
    Dim openvalue, endvalue, increase, decrease As Double
    
    'SUMMARY TABLE
    
    'Initial values
    row = 2
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
    ws.Range("J1").Value = "Ticker"
    ws.Range("K1").Value = "Yearly change"
    ws.Range("L1").Value = "Percent change"
    ws.Range("M1").Value = "Total stock volume"
    openvalue = Range("C" & 2).Value
    vol_total = 0
    vol = 0

    'Loop looking for tickers
    For i = 2 To lastrow
        
        ticker = ws.Range("A" & i).Value
        vol = ws.Range("G" & i).Value
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Range("J" & i + 1).Value 'Take new ticker
            endvalue = ws.Range("F" & i).Value 'Take new end value
            YearlyChange = endvalue - openvalue 'Yearly change calculation
            ws.Range("K" & row).Value = YearlyChange 'Insert yearly change
            PercentChange = YearlyChange / openvalue 'Yearly change percent calculation
            ws.Range("L" & row).Value = FormatPercent(PercentChange) 'Insert change percent
            vol_total = vol_total + vol
            ws.Range("M" & row).Value = vol_total
            vol_total = 0
                
                'Change format yearly change
                If YearlyChange < 0 Then
                ws.Range("K" & row).Interior.ColorIndex = 3
                ws.Range("L" & row).Interior.ColorIndex = 3
                Else
                ws.Range("K" & row).Interior.ColorIndex = 4
                ws.Range("L" & row).Interior.ColorIndex = 4
                End If
                
            openvalue = ws.Range("C" & i + 1) 'Take new open value
            row = row + 1
            ws.Range("J" & row).Value = ticker 'Insert new ticker on final table
        Else
            ws.Range("J" & row).Value = ticker 'Insert new ticker on final table
            vol_total = vol_total + vol
        End If

    Next i
    
    'BONUS
    
    'Initial values
    increase = 0
    decrease = 0
    volume = 0
    lastrow_bonus = ws.Cells(Rows.Count, 12).End(xlUp).row
    ws.Range("O1").Value = "Greatest % increase"
    ws.Range("O2").Value = "Greatest % decrease"
    ws.Range("O3").Value = "Greatest total volume"
    
    'Loop checking values
    For j = 2 To lastrow_bonus
        'Change positive
        If ws.Range("L" & j).Value > 0 Then
            If ws.Range("L" & j).Value > increase Then
                increase = ws.Range("L" & j)
            Else
                increase = increase
            End If
        'Change negative
        Else
            If ws.Range("L" & j).Value < decrease Then
                decrease = ws.Range("L" & j)
                'MsgBox ("HERE " & decrease)
            Else
                decrease = decrease
            End If
        End If
        'Greatest volume
        If ws.Range("M" & j).Value > volume Then
            volume = ws.Range("M" & j)
        Else
            volume = volume
        End If

    Next j
    
        ws.Range("P1").Value = FormatPercent(increase)
        ws.Range("P2").Value = FormatPercent(decrease)
        ws.Range("P3").Value = FormatNumber(volume)
            
    Next ws
    
End Sub


