Attribute VB_Name = "Module2"
'Sub allstocksanalysis()

'    Worksheets("All Stocks Analysis test").Activate
    
 '   Range("A1").Value = "All Stocks (2018)"
    'Create header for the worksheet
  '  Cells(3, 1).Value = "Ticker"
  '  Cells(3, 2).Value = "Total Daily Volume"
  '  Cells(3, 3).Value = "Return"
    
'End Sub

Sub All_Stocks_Analysis()

    Dim startTime As Single
    Dim endTime  As Single

    Worksheets("All Stocks Analysis test").Activate
    
    yearvalue = InputBox("What year would you like to run the analysis on?")
    
    startTime = Timer
    
    Range("A1").Value = "All Stocks (" + yearvalue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    'Initialize the array of Tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    Worksheets(yearvalue).Activate
    'Initialize the variables
    rowstart = 2
    'Delete : rowend = 3013
    'rowend code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
    'to find the number of rows to loop over
    rowend = Cells(Rows.Count, "A").End(xlUp).Row

    Dim startingprice As Single
    Dim endingprice As Single
    
    'Loop through the tickers
    For i = 0 To 11
        ticker = tickers(i)
        totalvolume = 0
    
        'loop over all the rows
        Worksheets(yearvalue).Activate
        For j = rowstart To rowend
    
            'increase totalvolume by the value in the current row
            If Cells(j, 1).Value = ticker Then
                totalvolume = totalvolume + Cells(j, 8).Value
            End If
            'calculating starting and ending price for current ticker
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1) <> ticker Then
                startingprice = Cells(j, 6).Value
            End If
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1) <> ticker Then
                endingprice = Cells(j, 6).Value
            End If
        
        Next j
        
        Worksheets("All Stocks Analysis test").Activate
    
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalvolume
        Cells(4 + i, 3).Value = endingprice / startingprice - 1
        
    Next i
        
        endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearvalue)
    
    
End Sub

Sub formatAllStocksAnalysisTable()

'Formatting
    Worksheets("All Stocks Analysis test").Activate
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.00%"
    Columns("B").AutoFit
    
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd
        If Cells(i, 3).Value > 0 Then
        Cells(i, 3).Interior.Color = vbGreen
        
        ElseIf Cells(i, 3).Value < 0 Then
        Cells(i, 3).Interior.Color = vbRed
        
        Else
        Cells(i, 3).Interior.Color = xlNone

        End If

    Next i
    
    
End Sub

Sub clearworksheet()

    Worksheets("All Stocks Analysis test").Activate
    
    Cells.Clear
    
End Sub

'Sub yearvalueAnalysis()

'yearvalue = InputBox("What year would you like to run the analysis on?")

'Call All_Stocks_Analysis(yearvalue)

'End Sub





