Attribute VB_Name = "Module2"
Sub ALLticker_summary()
    'A script that will loop through all the stocks for each year
    'for each run and take the following information:
    '1. The ticker symbol.
    '2. Yearly change from opening price at the beginning of a given year
    '     to the closing price at the end of that year.
    '3. The percent change from opening price at the beginning of a given year
    '     to the closing price at the end of that year.
    '4. The total stock volume of the stock.
    'Challenge 1 - Return the stock with the:
    '"Greatest % increase", "Greatest % Decrease" and "Greatest total volume"
    
    
    'Variables
    Dim current_ticker As String  'The current ticker being considered.
    Dim i As Long  'Row indexer
    Dim result_i As Integer  'Indexer for displayed results
    Dim open_price As Double  'Opening price of the stock
    Dim close_price As Double  'Closing price of the stock
    Dim deltaY As Double  'Yearly change in price (close_price - open_price)
    Dim deltaPercent As Double  'Percent change (deltaY / open_price * 100)
    Dim total As Double  'Total stock volume
    'Challenge 1 Variables
    Dim gpi As Double  'Greatest % Increase
    Dim gpi_ticker As String
    Dim gpd As Double  'Greatest % Decrease
    Dim gpd_ticker As String
    Dim gtv As Double  'Greatest Total Volume
    Dim gtv_ticker As String
    
    Dim ws As Worksheet  'For summarizing all of the years/sheets
    
    For Each ws In Sheets
    
        'Create Headers
        ws.Range("I1:R1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume", , , , , "Ticker", "Value")
        ws.Range("O2:O4").Value = Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume")
    
        current_ticker = ws.Range("A2").Value  'Initialize current_ticker to be the first ticker
        i = 2
        result_i = 2
    
        'Initialize results display
        ws.Cells(result_i, 9).Value = current_ticker
    
    
        Do While ws.Cells(i, 1).Value <> ""
            open_price = ws.Cells(i, 3).Value
            total = 0
            Do While ws.Cells(i, 1).Value = current_ticker
                total = total + ws.Cells(i, 7).Value
                i = i + 1
            Loop
        
            'Yearly Change
            close_price = ws.Cells(i - 1, 6).Value
            deltaY = close_price - open_price
            ws.Cells(result_i, 10).Value = deltaY
        
            'Percent Change
            If open_price = 0 Then
                deltaPercent = 0
            Else
                deltaPercent = deltaY / open_price
            End If
            ws.Cells(result_i, 11).Value = deltaPercent
        
            'Total Stock Volume
            ws.Cells(result_i, 12).Value = total
        
            'Set new current_ticker now that we've finished the previous one
            current_ticker = ws.Cells(i, 1).Value
            result_i = result_i + 1
            ws.Cells(result_i, 9).Value = current_ticker
        Loop
    
        'Challenge 1 Calculations
    
        result_i = 2  'Reset result_i for new processes
    
        Do While ws.Cells(result_i, 9).Value <> ""
            If ws.Cells(result_i, 11).Value > gpi Then
                gpi = ws.Cells(result_i, 11).Value
                gpi_ticker = ws.Cells(result_i, 9).Value
            End If
            If ws.Cells(result_i, 11).Value < gpd Then
                gpd = ws.Cells(result_i, 11).Value
                gpd_ticker = ws.Cells(result_i, 9).Value
            End If
            If ws.Cells(result_i, 12).Value > gtv Then
                gtv = ws.Cells(result_i, 12).Value
                gtv_ticker = ws.Cells(result_i, 9).Value
            End If
            result_i = result_i + 1
        Loop
    
        'Display Challenge 1 Results
        ws.Range("Q2:R2").Value = Array(gpi_ticker, gpi)
        ws.Range("Q3:R3").Value = Array(gpd_ticker, gpd)
        ws.Range("Q4:R4").Value = Array(gtv_ticker, gtv)
        
        'Format columns I:R
        ws.Range("I1:R1").Columns.AutoFit
    
    Next ws
    
End Sub

