Attribute VB_Name = "StockMarketAnalyst"
Sub StockAnalyst()
    
    'Loop through each tab
    For Each ws In Worksheets
    
        'Create Column Headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
        'Determine number of rows
        rownum = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Define variable to track which ticker rows have already been filled
        Dim tickrow As Integer
        tickrow = 2
    
        'Define variable to track volume per ticker
        Dim totvol As Long
        totvol = 0
    
        'Define variables for opening price at beginning of year and closing price at EOY
        Dim openprice As Double
        Dim closeprice As Double
        openprice = ws.Cells(2, 3).Value
    
        'Produce Moderate Solution Outputs
        For i = 2 To rownum
            'List Ticker Symbols
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ws.Cells(tickrow, 9).Value = ws.Cells(i, 1).Value 'Input Ticker Value
                closeprice = ws.Cells(i, 6).Value 'Sets to closing price at EOY
                ws.Cells(tickrow, 10).Value = closeprice - openprice 'Input Yearly Change Value
                    'Conditional Formatting Based on Pos or Neg Change
                    If ws.Cells(tickrow, 10).Value > 0 Then
                        ws.Cells(tickrow, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(tickrow, 10).Interior.ColorIndex = 3
                    End If
                If openprice = 0 Then
                    ws.Cells(tickrow, 11).Value = 0 'Dividing by 0 will throw error
                Else
                    ws.Cells(tickrow, 11).Value = (closeprice - openprice) / openprice 'Input % Change Value
                End If
                ws.Cells(tickrow, 11).NumberFormat = "0.00%" 'Format % Change as %
                ws.Cells(tickrow, 12).Value = totalvol + ws.Cells(i, 7).Value 'Input Total Volume for Ticker
                tickrow = tickrow + 1 'Moves to next row in I to prevent overwrite
                totalvol = 0 'Reset totalvol since ticker changed
                openprice = ws.Cells(i + 1, 3).Value 'Sets to opening price for following ticker
            Else
                totalvol = totalvol + ws.Cells(i, 7).Value 'If ticker hasn't changed, add volume to total volume
            End If
        Next i
        
        'Produce Hard Solution Outputs
        'Create Column/Row Headers
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Determine number of tickers
        ticknum = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Create variables to store outputs
        Dim greatinc As Double
        Dim greatdec As Double
        Dim greatvol As Double
        Dim tick1, tick2, tick3 As String
        
        'Initialize Variables
        greatinc = ws.Cells(2, 11).Value
        greatdec = ws.Cells(2, 11).Value
        greatvol = ws.Cells(2, 12).Value
        
        'Compare % Values and Volumes
        For i = 2 To ticknum
            If ws.Cells(i, 11).Value >= greatinc Then
                greatinc = ws.Cells(i, 11).Value
                tick1 = ws.Cells(i, 9).Value
            ElseIf ws.Cells(i, 11).Value <= greatdec Then
                greatdec = ws.Cells(i, 11).Value
                tick2 = ws.Cells(i, 9).Value
            End If
            If ws.Cells(i, 12).Value >= greatvol Then
                greatvol = ws.Cells(i, 12).Value
                tick3 = ws.Cells(i, 9).Value
            End If
        Next i
    
        'Output Results
        ws.Range("Q2").Value = greatinc
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("P2").Value = tick1
        ws.Range("Q3").Value = greatdec
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("P3").Value = tick2
        ws.Range("Q4").Value = greatvol
        ws.Range("P4").Value = tick3
    
    Next ws

End Sub
