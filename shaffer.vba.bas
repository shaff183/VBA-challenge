Attribute VB_Name = "Module1"
'From the information given, loop through and output different quantities
Sub stockCreater()

    For Each ws In Worksheets
    
                'formating each new worksheet with the proper headings for each category
                ws.Cells(1, 9).value = "Ticker"
                ws.Cells(1, 10).value = "Yearly Change"
                ws.Cells(1, 11).value = "Percent Change"
                ws.Cells(1, 12).value = "Total Stock Volume"
                
                'getting the last row of the data set and storing it in a variable
                Dim lastRow As Long
                lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                
                'creating a variable that holds the position of the summary table
                Dim sumTable As Integer
                sumTable = 2
                
                'Variables we will need to hold the information of the stocks
                Dim ticker As String
                Dim yearlyChange As Double
                Dim openingPrice As Double
                Dim percentChange As Double
                Dim stockVolume As LongPtr
                Dim closingPrice As Double
                
                'setting the initial opening price for the first stock
                openingPrice = ws.Cells(2, 3).value
                
                    'creating a for loop that will go through the whole worksheet
                    For i = 2 To lastRow
                
                                'testing if the ticker symbol matches the one below it
                                 If (ws.Cells(i, 1).value <> ws.Cells(i + 1, 1).value) Then 'ticker symbol doesnt match
            
                                            'get the closing price
                                            closingPrice = ws.Cells(i, 6).value
                                            
                                            'calculating the yearly change and percent change from opening and closing prices
                                            yearlyChange = closingPrice - openingPrice
                                            
                                            If (openingPrice = 0) Then
                                                    percentChange = 0
                                            Else
                                                    percentChange = (yearlyChange / openingPrice)
                                            End If
                                                                       
                                            'output the data into the summary table
                                            stockVolume = stockVolume + ws.Cells(i, 7).value
                                            ws.Cells(sumTable, 12).value = stockVolume
                                            
                                            ticker = ws.Cells(i, 1).value
                                            ws.Cells(sumTable, 9).value = ticker
                                            ws.Cells(sumTable, 10).value = yearlyChange
                                            ws.Cells(sumTable, 11).value = FormatPercent(percentChange, 2)
                                                          
                                            'adding 1 to the sumtable variable to move it down a row
                                            sumTable = sumTable + 1
                                        
                                            'reseting variables
                                            stockVolume = 0
                                            yearlyChange = 0
                                            openingPrice = 0
                                            closingPrice = 0
                                            
                                        
                                            'get the opening price for the next stock and hold it
                                            openingPrice = ws.Cells(i + 1, 3).value
                                            
                                Else 'when the ticker symbol matches the one directly below it
                                            stockVolume = stockVolume + ws.Cells(i, 7).value
                                End If
                    Next i
                    
                'conditional formatting for the yearly change, green if positive (or 0), and red if negative
                Dim rangeLength As Long
                rangeLength = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
                For n = 2 To rangeLength
                        'adding first rule
                        ws.Cells(n, 10).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, _
                                Formula1:="=0"
                        ws.Cells(n, 10).FormatConditions(1).Interior.Color = vbGreen
                        'adding second rule
                        ws.Cells(n, 10).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                                Formula1:="0"
                        ws.Cells(n, 10).FormatConditions(2).Interior.Color = vbRed
                Next n
                    
                'BONUS SECTION: return the Greatest % increase, greatest % decrease, and greatest total volume
                'formatting for the new section
                ws.Range("O2").value = "Greatest % Increase"
                ws.Range("O3").value = "Greatest % Decrease"
                ws.Range("O4").value = "Greatest Total Volume"
                ws.Range("P1").value = "Ticker"
                ws.Range("Q1").value = "Value"
               
                'variables for the bonus section
                Dim bonusTicker1 As String
                Dim bonusTicker2 As String
                Dim bonusTicker3 As String
                Dim greatestIncrease As Double
                Dim greatestDecrease As Double
                Dim greatestVolume As LongPtr
                
                bonusTicker1 = ws.Cells(2, 9).value
                greatestIncrease = ws.Cells(2, 11).value
                
                bonusTicker2 = ws.Cells(2, 9).value
                greatestDecrease = ws.Cells(2, 11).value
                
                bonusTicker3 = ws.Cells(2, 9).value
                greatestVolume = ws.Cells(2, 12).value
                
                For j = 2 To rangeLength
        
                        If (ws.Cells(j, 11).value > greatestIncrease) Then
                                bonusTicker1 = ws.Cells(j, 9).value
                                greatestIncrease = ws.Cells(j, 11).value
                        End If
                        
                        If (ws.Cells(j, 11).value < greatestDecrease) Then
                                bonusTicker2 = ws.Cells(j, 9).value
                                greatestDecrease = ws.Cells(j, 11).value
                        End If
                        
                        If (ws.Cells(j, 12).value > greatestVolume) Then
                                bonusTicker3 = ws.Cells(j, 9).value
                                greatestVolume = ws.Cells(j, 12).value
                        End If
        
                Next j
                
                'setting each value to the appropriate cell
                ws.Range("P2").value = bonusTicker1
                ws.Range("Q2").value = FormatPercent(greatestIncrease, 2)
                
                ws.Range("P3").value = bonusTicker2
                ws.Range("Q3").value = FormatPercent(greatestDecrease, 2)
                
                ws.Range("P4").value = bonusTicker3
                ws.Range("Q4").value = greatestVolume
        
    Next ws
                
End Sub
