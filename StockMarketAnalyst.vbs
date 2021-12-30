Attribute VB_Name = "StockMarketAnalyst"
Sub StockAnalyst()
    
    'Create Column Headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    'Determine number of rows
    rownum = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Define variable to track which ticker rows have already been filled
    Dim tickrow As Integer
    tickrow = 2
    
    'Define variable to track volume per ticker
    Dim totvol As Long
    totvol = 0
    
    'Define variables for opening price at beginning of year and closing price at EOY
    Dim openprice As Double
    Dim closeprice As Double
    openprice = Cells(2, 3).Value
    
    For i = 2 To rownum
        'List Ticker Symbols
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            Cells(tickrow, 9).Value = Cells(i, 1).Value 'Input Ticker Value
            closeprice = Cells(i, 6).Value 'Sets to closing price at EOY
            Cells(tickrow, 10).Value = closeprice - openprice 'Input Yearly Change Value
                'Conditional Formatting Based on Pos or Neg Change
                If Cells(tickrow, 10).Value > 0 Then
                    Cells(tickrow, 10).Interior.ColorIndex = 4
                Else
                    Cells(tickrow, 10).Interior.ColorIndex = 3
                End If
            If openprice = 0 Then
                Cells(tickrow, 11).Value = 0 'Dividing by 0 will throw error
            Else
                Cells(tickrow, 11).Value = (closeprice - openprice) / openprice 'Input % Change Value
            End If
            Cells(tickrow, 11).NumberFormat = "0.00%" 'Format % Change as %
            Cells(tickrow, 12).Value = totalvol + Cells(i, 7).Value 'Input Total Volume for Ticker
            tickrow = tickrow + 1 'Moves to next row in I to prevent overwrite
            totalvol = 0 'Reset totalvol since ticker changed
            openprice = Cells(i + 1, 3).Value 'Sets to opening price for following ticker
        Else
            totalvol = totalvol + Cells(i, 7).Value 'If ticker hasn't changed, add volume to total volume
        End If
    Next i

End Sub
