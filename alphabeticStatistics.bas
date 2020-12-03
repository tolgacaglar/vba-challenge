Attribute VB_Name = "Module1"
Sub writeSummaryStatistics()

    'Declare variables
    Dim unique_ticker_array(10000) As String    'unique tickers found from the first column
    Dim first_day_row_array(10000) As Long 'row of the beginning year for each ticker
    Dim last_day_row_array(10000) As Long 'row of the ending year of each ticker
    Dim worksheet_array(10000) As String  'the worksheet of the corresponding row indices and tickers
    Dim first_day_price_array(10000) As Double  'Opening stock price of the first day of the unique ticker
    Dim last_day_price_array(10000) As Double  'Closing stock price of the last day of the unique ticker
    Dim total_volume_array(10000) As Double       'total volume of each ticker
    
    Dim unique_ticker_counter As Long: unique_ticker_counter = 0  'Counting the number of unique tickers
    
    Dim last_row As Long    'last row of the current worksheet
    'Run through each worksheet
    ''' DATA COLLECTION
    For Each ws In Worksheets
        'row of the first day of the year of the first ticker in the worksheet
        first_day_row_array(unique_ticker_counter) = 2
        first_day_price_array(unique_ticker_counter) = ws.Cells(2, 3).Value
        
        
        'The first unique ticker in the worksheet is the first entry
        unique_ticker_array(unique_ticker_counter) = ws.Cells(2, 1).Value
        
        'Unique ticker is inside the worksheet ws
        worksheet_array(unique_ticker_counter) = ws.Name
        
        'run through each line to obtain the beginning and the end of the year and calculate total volume
        last_row = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row   'find the last row of the worksheet
        Dim rix As Long
        For rix = 2 To last_row + 1   'start the first row of the first entry, then move downwards
            'Check if the two consecutive tickers are different
            If ws.Cells(rix, 1).Value <> ws.Cells(rix + 1, 1).Value Then
                'This ticker
                last_day_row_array(unique_ticker_counter) = rix        'Row of the last day of the year
                last_day_price_array(unique_ticker_counter) = ws.Cells(rix, 6).Value  'Closing price of the current ticker
                total_volume_array(unique_ticker_counter) = total_volume + ws.Cells(rix, 7)  'Calculated total volume + current ticker's volume
                
                'Next ticker
                unique_ticker_counter = unique_ticker_counter + 1         'Increment unique row, since a new ticker is found
                unique_ticker_array(unique_ticker_counter) = ws.Cells(rix + 1, 1).Value 'Save the next ticker
                total_volume = 0    'Reset the total volume for a new calculation of the next ticker
                first_day_row_array(unique_ticker_counter) = rix + 1   'First row of the next ticker
                worksheet_array(unique_ticker_counter) = ws.Name            'The next ticker is still inside this worksheet
                first_day_price_array(unique_ticker_counter) = ws.Cells(rix + 1, 3).Value 'opening price of the next ticker
            Else
                total_volume = total_volume + ws.Cells(rix, 7).Value      'Add the volume of the current ticker to the rest
            End If
        Next rix
        
        'final row of the last ticker
        'The current ticker is the last ticker, at row last_row
        last_day_row_array(unique_ticker_counter) = last_row
        last_day_price_array(unique_ticker_counter) = ws.Cells(last_row, 6).Value
        total_volume_array(unique_ticker_counter) = total_volume
        'There is no next ticker!
    Next
    
    
    ''' ANALYSIS & PRINT INTO A SHEET
    'Use the first worksheet to put in the analysis
    'A better approach is to create a new sheet, and add the analysis into this new sheet, but I will use the first worksheet for now.
    Dim ws_analysis As Worksheet: Set ws_analysis = Worksheets(1) 'declare worksheet for analysis
    Dim yearly_change As Double     'Yearly change of the stock price: last-first
    Dim percent_change As Double    'Percent change of the stock price: yearly_change/first
    
    'Create the headers for the analysis
    ws_analysis.Range("I1").Value = "Ticker"            'id of the stock
    ws_analysis.Range("J1").Value = "Yearly Change"     'overall change in one year
    ws_analysis.Range("K1").Value = "Percent Change"    'overall percent change in one year
    ws_analysis.Range("L1").Value = "Total Stock Volume"    'total stock volume over one year
    
    'Headers/Titles for the ''bonus section''
    ws_analysis.Range("O2").Value = "Greatest % Increase"
    ws_analysis.Range("O3").Value = "Greatest % Decrease"
    ws_analysis.Range("O4").Value = "Greatest Total Volume"
    ws_analysis.Range("P1").Value = "Ticker"
    ws_analysis.Range("Q1").Value = "Value"
    
    'Run through each ticker and write to the analysis worksheet
    'Find the maximum of the percent increase,decrease and total volume
    Dim maximum_increase As Double: maximum_increase = -1       'maximum increase is by default -1, since increase is positive
    Dim maximum_increase_ticker As String                       'corresponding ticker of the maximum increase
    Dim maximum_decrease As Double: maximum_decrease = 1        'maximum decrease is by default +1, since decrease is negative
    Dim maximum_decrease_ticker As String                       'corresponding ticker of the maximum decrease
    Dim maximum_volume As Double: maximum_volume = -1           'maximum volume is by default -1, since volume is always positive
    Dim maximum_volume_ticker As String                       'corresponding ticker of the maximum volume
    For ticker_idx = 0 To unique_ticker_counter - 1
        'Calculate the yearly change and the percent change for each ticker
        yearly_change = last_day_price_array(ticker_idx) - first_day_price_array(ticker_idx)
        If yearly_change = 0 Then
            percent_change = 0
        Else
            percent_change = yearly_change / first_day_price_array(ticker_idx)
        End If
        total_volume = total_volume_array(ticker_idx)
        
        ws_analysis.Cells(ticker_idx + 2, 9).Value = unique_ticker_array(ticker_idx)    'Unique ticker ID at column I
        ws_analysis.Cells(ticker_idx + 2, 10).Value = yearly_change                     'Yearly change at column J
        If yearly_change < 0 Then       'If stock price decreased
            ws_analysis.Cells(ticker_idx + 2, 10).Interior.Color = RGB(255, 0, 0) 'Red
        Else    'if stock price increased
            ws_analysis.Cells(ticker_idx + 2, 10).Interior.Color = RGB(0, 255, 0) 'Green
        End If
            
        ws_analysis.Cells(ticker_idx + 2, 11).Value = percent_change                    'Percent change at column K
        'Change formatting of the current cell to percent
        ws_analysis.Cells(ticker_idx + 2, 11).NumberFormat = "0.00%"
        ws_analysis.Cells(ticker_idx + 2, 12).Value = total_volume                      'Total stock volume at column L
        
        'Check if the current ticker has a higher increase/decrease then the saved maximums
        If percent_change > maximum_increase Then
            maximum_increase = percent_change
            maximum_increase_ticker = unique_ticker_array(ticker_idx)
        End If
        If percent_change < maximum_decrease Then
            maximum_decrease = percent_change
            maximum_decrease_ticker = unique_ticker_array(ticker_idx)
        End If
        'Check if the current ticker has the highest total volume
        If total_volume > maximum_volume Then
            maximum_volume = total_volume
            maximum_volume_ticker = unique_ticker_array(ticker_idx)
        End If
    Next
    
    'Print the maximum increasing/decreasing/volume ticker with its value
    ws_analysis.Range("P2").Value = maximum_increase_ticker
    ws_analysis.Range("Q2").Value = maximum_increase
    ws_analysis.Range("P3").Value = maximum_decrease_ticker
    ws_analysis.Range("Q3").Value = maximum_decrease
    ws_analysis.Range("P4").Value = maximum_volume_ticker
    ws_analysis.Range("Q4").Value = maximum_volume
End Sub

