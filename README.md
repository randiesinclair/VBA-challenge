# VBA challenge
This project is about using Excel VBA to create a script that loops through all stocks for one year and outputs the ticker symbol, yearly change from the opening price, percentage change from the opening price, and total stock volume of the stock. Additionally, the script should return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The final solution should also enable the script to run on every worksheet at once and include conditional formatting to highlight positive and negative changes in green and red respectively.

# Technical Skills Required
- VBA scripting
- Understanding of Excel and working with worksheets
- Knowledge of loops and conditional formatting

# Project Parameters
- The script should be able to loop through all stocks for one year
- Output the ticker symbol, yearly change, percentage change, and total stock volume
- Return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
- The script should run on every worksheet at once and include conditional formatting to highlight positive and negative changes in green and red respectively
- The script should be tested with the sheet alphabetical_testing.xlsx and run under 3 to 5 minutes
- The final solution should be submitted to GitHub/GitLab with Screenshots of the results, separate VBA script files, and a README file.


# Final Analysis
Sub Stock_Analyst()

    'Make macro run on all sheets
    Dim ws As Worksheet
    For Each ws In Worksheets

        'Set variables
        Dim stock_name As String
        Dim opening_price As Double
        Dim percent_change As Double
        Dim stock_volume As Double
        Dim closing_price As Double
        Dim yearly_change As Double
          
        'initialize variables
        Dim summary_table_row As Integer
        summary_table_row = 2
        tickervolume = 0
        opening_price = Cells(2, 3).Value

        'Create Row Titles
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        'Determine the number of rows
        RowCount = Cells(Rows.Count, 1).End(xlUp).Row

        'Create For Loop to collect and print stock names, trade volume, and yearly change

        For i = 2 To RowCount

            'Looking for stock names, if different
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'Set the stock name
            stock_name = Cells(i, 1).Value

            'Print the stock name to the summary table
            ws.Range("I" & summary_table_row).Value = stock_name

            'Add up the stock volume
            stock_volume = stock_volume + ws.Cells(i, 7).Value

            'Print the stock volume to the summary table
            ws.Range("L" & summary_table_row).Value = stock_volume

            'Get the closing price data
            closing_price = ws.Cells(i, 6).Value

            'Calculate the yearly change
            yearly_change = (closing_price - opening_price)
              
            'Print the yearly change for each stock to the summary table
            ws.Range("J" & summary_table_row).Value = yearly_change

            'If statement for percent change
            If (opening_price = 0) Then
            percent_change = 0

                Else
                    
                percent_change = yearly_change / opening_price
                
                End If

              'Print the yearly change for each ticker in the summary table and move down
              ws.Range("K" & summary_table_row).Value = percent_change
              ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
              summary_table_row = summary_table_row + 1

              'Reset
              stock_volume = 0
              opening_price = Cells(i + 1, 3)
            
            Else
              
               'Add up the stock volume
              stock_volume = stock_volume + ws.Cells(i, 7).Value

            
            End If
        
        Next i

    'Determine Last row
    lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
    
    'Color format the yearly change data
    
    For i = 2 To lastrow_summary_table
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 10
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
    Next i

Next ws


End Sub
