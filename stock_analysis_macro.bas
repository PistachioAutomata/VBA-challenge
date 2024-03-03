Attribute VB_Name = "Module1"
Sub stocksummary()
    For Each ws In Worksheets
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' initial variable for ticker
    Dim ticker As String
    
    ' initial variable for opening price
    Dim openprice As Double
    openprice = ws.Cells(2, 3).Value
    
    ' initial variable for closing price
    Dim closeprice As Double
    closeprice = 0
    
    ' variable for change in price over year
    Dim pricechange As Double
    pricechange = 0
    
    ' variable for % change in price
    Dim percchange As Double
    percchange = 0
    
    ' variable for volume total
    Dim volume As Double
    volume = 0
    
    ' variable for locating each ticker in the summary table
    Dim summary_table_row As Integer
    summary_table_row = 2
    
    ' setting variable for stock ticker column
    Dim column As Integer
    column = 1
    
    ' loop through rows in column, search for when values change
    For I = 2 To lastrow
        If ws.Cells(I + 1, column).Value <> ws.Cells(I, column) Then
        
            ' set ticker name
            ticker = ws.Cells(I, column).Value
        
            ' set close price
            closeprice = ws.Cells(I, 6).Value
        
            ' find change in price
            pricechange = closeprice - openprice
        
            ' find % change in price
            percchange = pricechange / openprice
        
            ' add volume total
            volume = volume + ws.Cells(I, 7).Value
        
            ' Print values to the summary table
            ws.Range("I" & summary_table_row).Value = ticker
            ws.Range("J" & summary_table_row).Value = pricechange
            ws.Range("K" & summary_table_row).Value = percchange
            ws.Range("L" & summary_table_row).Value = volume
        
            ' add 1 to summary_table_row
            summary_table_row = summary_table_row + 1
        
            ' set new opening price
            openprice = ws.Cells(I + 1, 3).Value
            
            'reset volume
            volume = 0
        
        Else
            volume = volume + ws.Cells(I, 7).Value
            
        End If
    Next I
    
    ' Format Summary Table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Price Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Volume"
    ws.Range("K2:K10000").NumberFormat = "0.00%"
    ws.Range("J2:J10000").NumberFormat = "$0.00"
    
    ' set variable for finding Greatest % increase
    Dim greatest_increase As Double
    greatest_increase = 0
    Dim greatest_increase_ticker As String
    
    'set variable for finding greatest % decrease
    Dim greatest_decrease As Double
    greatest_decrease = 0
    Dim greatest_decrease_ticker As String
    
    'set variable for finding largest volume
    Dim highvolume As Double
    highvolume = 0
    Dim highvolume_ticker As String
    
    lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    For I = 2 To lastrow2
        If greatest_increase < ws.Cells(I, 11).Value Then
        greatest_increase = ws.Cells(I, 11).Value
        greatest_increase_ticker = ws.Cells(I, 9).Value
        End If
    Next I
    
    For I = 2 To lastrow2
        If greatest_decrease > ws.Cells(I, 11).Value Then
        greatest_decrease = ws.Cells(I, 11).Value
        greatest_decrease_ticker = ws.Cells(I, 9).Value
        End If
    Next I
    
     For I = 2 To lastrow2
        If highvolume < ws.Cells(I, 12).Value Then
        highvolume = ws.Cells(I, 12).Value
        highvolume_ticker = ws.Cells(I, 9).Value
        End If
    Next I
    
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Largest Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    ws.Range("P2").Value = greatest_increase_ticker
    ws.Range("P3").Value = greatest_decrease_ticker
    ws.Range("P4").Value = highvolume_ticker
    
    ws.Range("Q2").Value = greatest_increase
    ws.Range("Q3").Value = greatest_decrease
    ws.Range("Q4").Value = highvolume
    
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    Next ws
End Sub
