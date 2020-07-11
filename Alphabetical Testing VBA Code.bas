Attribute VB_Name = "Module1"
Sub TickerSummary()
Dim closingprice As Single
Dim openingprice As Single
Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet

'Iterate through each worksheet in workbook
For Each ws In ThisWorkbook.Worksheets
    ws.Activate

        'Find last row with data in table set as var lastrow
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row

        'Set Total Stock Volume sum, Opening Price, and Closing price to initial value of 0
        totalstockvolume = 0
        openingprice = 0
        closingprice = 0

        'The data will be summarized in a summary table. The first row of that table is set to 2. This allows the code to iterate and add a new line to the summary table with each loop.
        summarytablerow = 2

        'Add Headers to Summary Table
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"

        'Loop through data
        For Row = 2 To lastrow

            'Create variables for current row, prior row, and next row. This allows code to compare values in each row. Differences in the rows' values will drive the code.
            currentrowticker = Cells(Row, 1).Value
            nextrowticker = Cells(Row + 1, 1).Value
            priorrowticker = Cells(Row - 1, 1).Value

            'Create variables for stock volume of each ticker (Column G) and ticker name (Column A)
            volume = Cells(Row, 7).Value
            tickername = Cells(Row, 1).Value

           'Establish calculator to track total stock volume. The var totalstockvolume will store the sum of additional units of daily volume to the total sum.
           totalstockvolume = totalstockvolume + volume


                'Logic Statement: As the code reads each row, if current row's ticker is different than the prior row's ticker, store the opening price on the current row. This allows the earliest date's opening price to be stored.
                If priorrowticker <> currentrowticker Then
                
                    'Set variable for opening price, stored in Column C
                    openingprice = Cells(Row, 3).Value

                End If

                 'Logic Statement: As the code reads each row, if the current row's ticker is different than the next row's ticker, store the closing price of the current row. This allows the latest date's closing price to be stored.
                If nextrowticker <> currentrowticker Then
                    closingprice = Cells(Row, 6).Value
                    
                    'Print ticker name, Total Stock Volume, Difference in Closing and Opening Price, and % change of Closing Price to Summary table
                    Cells(summarytablerow, 9).Value = tickername
                    Cells(summarytablerow, 12).Value = totalstockvolume
                    Cells(summarytablerow, 10).Value = (closingprice - openingprice)
                    Cells(summarytablerow, 11).Value = (closingprice - openingprice) / openingprice

                    'Format Results
                    Cells(summarytablerow, 10).NumberFormat = "0.00"
                    Cells(summarytablerow, 11).NumberFormat = "0.00%"
                    Cells(summarytablerow, 12).NumberFormat = "#,##0"


                        'Apply red shading to negative Yearly Change results
                        If Cells(summarytablerow, 10).Value > 0 Then
                            Cells(summarytablerow, 10).Interior.ColorIndex = 4

                        'Apply green shading to positive yearly change results
                        ElseIf Cells(summarytablerow, 10).Value < 0 Then
                        Cells(summarytablerow, 10).Interior.ColorIndex = 3

                        End If

                'Next unique ticker's data will print on next row of Summary Table
                summarytablerow = summarytablerow + 1

                'reset Total Stock Volume calculatorto recalculate for next unique ticker
                totalstockvolume = 0
                openingprice = 0
                closingprice = 0


                End If

            Next Row


'CHALLENGE CODE:

        'Add Headers to a new summary table that summarizes the Max % Gain, Max % Loss, and Greatest Total Volume.
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"

        'Find last row with data in table set as var lastrow
        lastrowgreatest = Cells(Rows.Count, 9).End(xlUp).Row
        
        'Set Max, Min, and Greatest Volume variables to initial value of 0.
        Max = 0
        Min = 0
        GreatestVolume = 0
        
        'The data will be summarized in a new summary table. The first row of that table is set to 2. This allows the code to iterate and add a new line to the summary table with each loop.
        greatestsummarytablerow = 2
        
        'Loop through data. The last row is set to 100 for
        For SummaryRowMax = 2 To lastrowgreatest
            
            tickername = Cells(SummaryRowMax, 9).Value
            currentrowpercentchange = Cells(SummaryRowMax, 11).Value
                    
            'Find Max % change by looping through summary table (columns K), print value and tickername in Greatest Summary Table
            If currentrowpercentchange > Max Then
                Max = currentrowpercentchange
                Else: Max = Max
                
                Cells(greatestsummarytablerow, 16) = tickername
                Cells(greatestsummarytablerow, 17) = Max
                Cells(greatestsummarytablerow, 17).NumberFormat = "0.00%"
                
                
            End If
        Next SummaryRowMax

                 greatestsummarytablerow = greatestsummarytablerow + 1
        
        For SummaryRowMin = 2 To 100

            tickername = Cells(SummaryRowMin, 9).Value
            currentrowpercentchange = Cells(SummaryRowMin, 11).Value
                        
                If currentrowpercentchange < Min Then
                    Min = currentrowpercentchange
                                
                    Cells(greatestsummarytablerow, 16) = tickername
                    Cells(greatestsummarytablerow, 17) = Min
                    Cells(greatestsummarytablerow, 17).NumberFormat = "0.00%"
        
                End If
            Next SummaryRowMin
                    
                    greatestsummarytablerow = greatestsummarytablerow + 1
        
        For SummaryRowGreatestVolume = 2 To 100

            tickername = Cells(SummaryRowGreatestVolume, 9).Value
            currentrowGreatestVolume = Cells(SummaryRowGreatestVolume, 12).Value
        
                If currentrowGreatestVolume > GreatestVolume Then
                    GreatestVolume = currentrowGreatestVolume
                    Cells(greatestsummarytablerow, 16) = tickername
                    Cells(greatestsummarytablerow, 17) = GreatestVolume
                    Cells(greatestsummarytablerow, 17).NumberFormat = "#,##0"
        
                End If
            Next SummaryRowGreatestVolume
                    

Next ws

'starting_ws.Activate


End Sub
