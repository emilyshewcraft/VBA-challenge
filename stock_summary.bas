Attribute VB_Name = "Module1"
' Instructions:
'
' Create a script that loops through all the stocks for one year
' and outputs the following information:
'       * The ticker symbol
'       * Yearly change from the opening price at the beginning of
'       a given year to the closing price at the end of that year.
'       * The percentage change from the opening price at the
'       beginning of a given year to the closing price at the end
'       of that year.
'       * The total volume of the stock.
'
' Add functionality to your script to return the stock with the
' "Greatest % increase", "Greatest % decrease", and "Greatest
' total volume".
'
' Make the appropriate adjustments to your VBA script to enable it
' to run on every worksheet (that is, every year) at once.
' --------------------------------------------------------------

Sub StockSummary()

    ' Loop through all sheets
    For Each ws In Worksheets

        ' ------------------------------------------------------
        ' SET UP SUMMARY TABLE
        ' ------------------------------------------------------
        ' Print column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ' Apply number formatting to cells
        ws.Range("K:K").NumberFormat = "0.00%"
        ws.Range("L:L").NumberFormat = "#,###"
    
        ' Auto-fit column width
        ws.Range("J1:K1").EntireColumn.AutoFit

        ' ------------------------------------------------------
        ' DEFINE VARIABLES
        ' ------------------------------------------------------

        ' Set variable for holding stock ticker
        Dim Ticker As String

        ' Keep track of the loaction for each stock ticker in
        ' the summary table
        Dim Summary_Table_Row As Integer
        ' Set first row of summary table
        Summary_Table_Row = 2

        ' Set variables for opening and closing price
        Dim Opening_Price As Double
        Dim Closing_Price As Double

        ' Set variables for yearly change and percent change
        Dim Yearly_Change As Double
        Dim Percent_Change As Double

        ' Set counter for holding stock volume
        Dim Total_Volume As LongLong
        Total_Volume = 0

        ' Find last row of data
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' ------------------------------------------------------
        ' COMPLETE SUMMARY TABLE
        ' ------------------------------------------------------

        ' Loop through stock data
        For i = 2 To LastRow

            ' Find each unique ticker symbol
            ' If the next row is a different stock, then...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Store the ticker symbol
                Ticker = ws.Cells(i, 1).Value
                
                ' Store the closing price:
                Closing_Price = ws.Cells(i, 6).Value

                ' Add stock volume to counter
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value

                ' Calculate yearly change
                Yearly_Change = Closing_Price - Opening_Price

                ' Calculate percent change
                Percent_Change = Yearly_Change / Opening_Price

                ' Print the following to the summary table...
                ' Ticker:
                ws.Cells(Summary_Table_Row, 9).Value = Ticker

                ' Yearly Change:
                ws.Cells(Summary_Table_Row, 10).Value = Yearly_Change

                ' Percent Change:
                ws.Cells(Summary_Table_Row, 11).Value = Percent_Change

                ' Total Volume:
                ws.Cells(Summary_Table_Row, 12).Value = Total_Volume

                ' Before continuing to next stock...

                ' Add a new row to the summary table
                Summary_Table_Row = Summary_Table_Row + 1

                ' Reset opening price, closing price, and
                ' total stock volume for next stock ticker
                Opening_Price = 0
                Closing_Price = 0
                Total_Volume = 0
            
            ' If the next row is still part of the same stock...
            Else

                ' Check if this is the first row for this ticker
                ' If the previous row is a different stock, then...
                If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then

                    ' Store the opening price:
                    Opening_Price = ws.Cells(i, 3).Value
                
                End If

                ' For all rows:
                ' Add stock volume to counter
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value

            End If

        ' Close loop through data
        Next i

        ' Auto-fit column width
        ws.Range("L:L").EntireColumn.AutoFit

        ' -----------------------------------------------------
        ' ADD CONDITIONAL FORMATTING TO SUMMARY TABLE
        ' -----------------------------------------------------
    
        ' Loop through summary table
        For i = 2 To (Summary_Table_Row - 1)

            ' If yearly change is 0 or positive...
            If ws.Cells(i, 10).Value >= 0 Then

                ' Fill the cell green
                ws.Cells(i, 10).Interior.Color = RGB(144, 174, 83)

            ' If yearly change is negative...
            Else

                ' Fill the cell red
                ws.Cells(i, 10).Interior.Color = RGB(236, 74, 90)

            End If

        ' Closes summary table loop
        Next i

        ' -----------------------------------------------------
        ' SET UP SECOND SUMMARY TABLE
        ' -----------------------------------------------------
    
        ' Print headers and row labels
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"

        ' Define variables for highest and lowest %, volume
        Dim Highest_Percent_Change As Double
        Dim Highest_Percent_Ticker As String
        Dim Lowest_Percent_Change As Double
        Dim Lowest_Percent_Ticker As String
        Dim Highest_Total_Volume As LongLong
        Dim Highest_Volume_Ticker As String

        Highest_Percent_Change = 0
        Lowest_Percent_Change = 0
        Highest_Total_Volume = 0

        ' Apply number formatting to cells
        ws.Range("P2:P3").NumberFormat = "0.00%"
        ws.Range("P4").NumberFormat = "#,###"

        ' Auto-fit column width
        ws.Range("N2:N4").EntireColumn.AutoFit

        ' ----------------------------------------------------
        ' FIND AND PRINT VALUES
        ' ----------------------------------------------------

        ' Loop through summary table
        For i = 2 To (Summary_Table_Row - 1)

            ' Search for highest % change
            ' If % change is higher than current highest, then...
            If ws.Cells(i, 11) > Highest_Percent_Change Then

                'Then store new % change and ticker #
                Highest_Percent_Change = ws.Cells(i, 11)
                Highest_Percent_Ticker = ws.Cells(i, 9)

            End If

            ' Search for lowest % change
            ' If % change is lower than current lowest, then...
            If ws.Cells(i, 11) < Lowest_Percent_Change Then

                'Then store new % change and ticker #
                Lowest_Percent_Change = ws.Cells(i, 11)
                Lowest_Percent_Ticker = ws.Cells(i, 9)

            End If

            ' Search for highest volume
            If ws.Cells(i, 12) > Highest_Total_Volume Then

                'Then store new highest volume and ticker #
                Highest_Total_Volume = ws.Cells(i, 12)
                Highest_Volume_Ticker = ws.Cells(i, 9)

            End If

        ' Closes summary table loop
        Next i

        ' Print the following values to the table...
        ' Greatest Percent Increase:
        ws.Cells(2, 15).Value = Highest_Percent_Ticker
        ws.Cells(2, 16).Value = Highest_Percent_Change

        ' Greatest Percent Decrease:
        ws.Cells(3, 15).Value = Lowest_Percent_Ticker
        ws.Cells(3, 16).Value = Lowest_Percent_Change

        ' Greatest Total Volume:
        ws.Cells(4, 15).Value = Highest_Volume_Ticker
        ws.Cells(4, 16).Value = Highest_Total_Volume

        ' Auto-fit column width
        ws.Range("P2:P4").EntireColumn.AutoFit
        

    ' Closes worksheet loop
    Next ws

    ' Display message when script is complete
    MsgBox ("Summary Complete")

End Sub
