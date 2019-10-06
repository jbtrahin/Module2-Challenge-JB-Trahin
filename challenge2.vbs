'Challenge 2'

Sub AllStocksAnalysis()
    yearValue = InputBox("What year would you like to run the analysis on?")
    
	'Format the output sheet on the “All Stocks Analysis” worksheet.'
    Worksheets("All Stocks Analysis").Activate

    Range("A1").Value = "All Stocks (" + yearValue + ")"

    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    

    'Set a tickerIndex equal to zero before the loop. If the next row’s ticker doesn’t match, increase the tickerIndex. Here, we’re taking advantage of the fact that the tickers are all in alphabetical order.'
    Dim tickers() As String
	tickerIndex = 0
    ReDim Preserve tickers(tickerIndex)
	
	'Create arrays for all the volumes, starting prices, and ending prices.'
    Dim volumes() As Long
    Dim startingPrice() As Double
    Dim endingPrice() As Double
    ReDim Preserve tickers(tickerIndex)
    ReDim Preserve volumes(tickerIndex)
    ReDim Preserve startingPrice(tickerIndex)
    ReDim Preserve endingPrice(tickerIndex)

    'Activate the data worksheet.'
    Worksheets(yearValue).Activate
	
	'Find number of rows (before both loops)'
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'use value to increment output writing' 
    i = 0
        
        'Loop through the tickers.'
    	'Use tickerIndex to access the correct index across the four different arrays you’ll be using (the tickers array and the three output arrays). tickers(tickerIndex) will give you the current ticker you’re working on.'
	    'Once you’re done collecting the data, loop through your arrays to output all of the information you’ve collected. The code responsible for formatting should still work without needing to be changed.'
	    'Once you’re done collecting the data, loop through your arrays to output all of the information you’ve collected. The code responsible for formatting should still work without needing to be changed.'
	    For j = 2 To RowCount

	    	If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then

	    		Worksheets(yearValue).Activate
		        
		        volumes(tickerIndex) = volumes(tickerIndex) + Cells(j, 8).Value
	    		endingPrice(tickerIndex) = Cells(j, 6).Value

		        'Activate outcome worksheet'
		       	Worksheets("All Stocks Analysis").Activate
		        
		        Cells(4 + i, 1).Value = tickers(tickerIndex)
		        Cells(4 + i, 2).Value = volumes(tickerIndex)
		        Cells(4 + i, 3).Value = (endingPrice(tickerIndex) / startingPrice(tickerIndex)) - 1
		        Cells(4 + i, 4).Value = startingPrice(tickerIndex)	
		        Cells(4 + i, 5).Value = endingPrice(tickerIndex)
		        
		        i = i + 1
		        
		        Worksheets(yearValue).Activate
		        
		        tickerIndex = tickerIndex + 1
		        
		        ReDim Preserve tickers(tickerIndex)
		        ReDim Preserve startingPrice(tickerIndex)
		        ReDim Preserve endingPrice(tickerIndex)
		        ReDim Preserve volumes(tickerIndex) 
      
      		ElseIf Cells(j - 1, 1).Value <> Cells(j, 1).Value Then
				
				Worksheets(yearValue).Activate
		        
              	tickers(tickerIndex) = Cells(j, 1).Value
              	startingPrice(tickerIndex) = Cells(j, 3).Value
              	volumes(tickerIndex) = volumes(tickerIndex) + Cells(j, 8).Value

            Else

            	Worksheets(yearValue).Activate
		        
		        volumes(tickerIndex) = volumes(tickerIndex) + Cells(j, 8).Value

      		End If

		Next j 

    'Formatting'
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    
    dataRowStart = 4
    dataRowEnd = 15
    
    For i = dataRowStart To dataRowEnd
      
      	If Cells(i, 3) > 0 Then
       		'Color the cell green
        	Cells(i, 3).Interior.Color = vbGreen
      
      	ElseIf Cells(i, 3) < 0 Then  
    		'Color the cell red
          	Cells(i, 3).Interior.Color = vbRed
      
      	Else
          	'Clear the cell color
          	Cells(i, 3).Interior.Color = xlNone
      
      	End If
  
  	Next i

End Sub