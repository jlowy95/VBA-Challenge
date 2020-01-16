Attribute VB_Name = "Module1"
Sub ticker_summary()
    'A script that will loop through all the stocks for one year
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
    
    
    current_ticker = Range("A2").Value  'Initialize current_ticker to be the first ticker
    i = 2
    result_i = 2
    
    'Initialize results display
    Cells(result_i, 9).Value = current_ticker
    
    
    Do While Cells(i, 1).Value <> ""
        open_price = Cells(i, 3).Value
        total = 0
        Do While Cells(i, 1).Value = current_ticker
            total = total + Cells(i, 7).Value
            i = i + 1
        Loop
        
        'Yearly Change
        close_price = Cells(i - 1, 6).Value
        deltaY = close_price - open_price
        Cells(result_i, 10).Value = deltaY
        
        'Percent Change
        deltaPercent = deltaY / open_price
        Cells(result_i, 11).Value = deltaPercent
        
        'Total Stock Volume
        Cells(result_i, 12).Value = total
        
        'Set new current_ticker now that we've finished the previous one
        current_ticker = Cells(i, 1).Value
        result_i = result_i + 1
        Cells(result_i, 9).Value = current_ticker
    Loop
    
    'Challenge 1 Calculations
    
    result_i = 2  'Reset result_i for new processes
    
    Do While Cells(result_i, 9).Value <> ""
        If Cells(result_i, 11).Value > gpi Then
            gpi = Cells(result_i, 11).Value
            gpi_ticker = Cells(result_i, 9).Value
        End If
        If Cells(result_i, 11).Value < gpd Then
            gpd = Cells(result_i, 11).Value
            gpd_ticker = Cells(result_i, 9).Value
        End If
        If Cells(result_i, 12).Value > gtv Then
            gtv = Cells(result_i, 12).Value
            gtv_ticker = Cells(result_i, 9).Value
        End If
        result_i = result_i + 1
    Loop
    
    'Display Challenge 1 Results
    Range("Q2:R2").Value = Array(gpi_ticker, gpi)
    Range("Q3:R3").Value = Array(gpd_ticker, gpd)
    Range("Q4:R4").Value = Array(gtv_ticker, gtv)
    
    
End Sub
