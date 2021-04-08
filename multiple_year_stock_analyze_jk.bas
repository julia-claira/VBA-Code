Attribute VB_Name = "Module1"
Sub stockAnalyze()


For Each ws In Worksheets
    'set variables
    
    Dim currentTicker As String 'this variable stores the current ticker
    Dim beginYearClose As Double 'the closing value for a Ticker at the beginning of the year
    Dim endYearClose As Double 'the closing value for a Ticker at end of year
    Dim yearlyChange As Double 'the amount of change for a ticker over a year
    Dim yearlyChangePer As Double 'the yearly percentage change of a Ticker
    Dim totalStockVolume As Double 'keeps track of the total stock for current ticker
    Dim tickerRowCounter 'keeps track of the next row to put a new ticker's information
    
    '---sets initial values-----
    tickerRowCounter = 2 'initializes the row counter
    currentTicker = ws.Range("A2").Value 'sets the first ticker
    beginYearClose = ws.Range("F2").Value 'sets the closing value for the first ticker at beginning
    totalStockVolume = ws.Range("G2") 'sets the initial stock volume for the first ticker
    ws.Cells(tickerRowCounter, 9).Value = currentTicker 'populates the first ticker symbol in the new ticker field
    
    '-----Bonust Variables-------
    Dim greatestIncrease As Double
    Dim greatestIncTicker As String
    Dim greatestDecrease As Double
    Dim greatestDecTicker As String
    Dim greatestVolume As Double
    Dim greatestVolTicker As String
    
    
    '------------------------------loops through the rows and analyzes the information----------
   For i = 3 To ws.Cells(Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then 'if not equal then the next cell contains a new ticker symbol
        
            'outputs the final yearly change and percentage change for the current ticker
            endYearClose = ws.Cells(i, 6).Value
            yearlyChange = endYearClose - beginYearClose
            ws.Cells(tickerRowCounter, 10).Value = yearlyChange
            
            If beginYearClose <> 0 Then 'will not calculate a stock that has 0 in all fields - like plnt in 2014 - which avoids an error
                yearlyChangePer = yearlyChange / beginYearClose
            Else
                yearlyChangePer = 0
            End If
            
            ws.Cells(tickerRowCounter, 11).Value = yearlyChangePer
            totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
            ws.Cells(tickerRowCounter, 12).Value = totalStockVolume
            
            'changes cell color of yearly change to green if change is positive and red if negative
            'for one sheet I used Conditional Formatting in case that was how we were supposed to do it
            If yearlyChange >= 0 Then
                 ws.Cells(tickerRowCounter, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(tickerRowCounter, 10).Interior.ColorIndex = 3
            End If
            
            'checks to see if current ticker has any greatest values and if it does it stores the information--------
            If yearlyChangePer > greatestIncrease Or tickerRowCounter = 2 Then
                greatestIncrease = yearlyChangePer
                greatestIncTicker = currentTicker
            ElseIf yearlyChangePer < greatestDecrease Or tickerRowCounter = 2 Then
                greatestDecrease = yearlyChangePer
                greatestDecTicker = currentTicker
            End If
            
            If totalStockVolume > greatestVolume Then
                greatestVolume = totalStockVolume
                greatestVolTicker = currentTicker
            End If
            
            'finished with the current ticker, this resets and initialize the new ticker's values--------------
            currentTicker = ws.Cells(i + 1, 1).Value
            tickerRowCounter = tickerRowCounter + 1
            ws.Cells(tickerRowCounter, 9).Value = currentTicker
            beginYearClose = ws.Cells(i + 1, 6).Value
            totalStockVolume = 0 'resets the total stock volume for the next ticker
               
        Else 'run this if the current ticker has more columns
            
            If beginYearClose = 0 Then 'if first cell for a ticker's closing value is 0, ensures that it keeps looking for the first cell with a closing value
                beginYearClose = ws.Cells(i, 6).Value
            End If
                
            totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value 'adds the volume of cell to the current stock total
            
        End If
        
    Next i
    
    '----bonus---output greatest values-----------
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(2, 16).Value = greatestIncTicker
    ws.Cells(2, 17).Value = greatestIncrease
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(3, 16).Value = greatestDecTicker
    ws.Cells(3, 17).Value = greatestDecrease
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(4, 16).Value = greatestVolTicker
    ws.Cells(4, 17).Value = greatestVolume
    
    'set headers for new rows------------
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    '----format rows--------
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    ws.Range("K2").EntireColumn.NumberFormat = "0.00%"
    ws.Range("O:Q").EntireColumn.AutoFit
    ws.Range("I:L").EntireColumn.AutoFit
    
    
 Next ws
    
    '--------------------------------------------------------------------------------


End Sub
