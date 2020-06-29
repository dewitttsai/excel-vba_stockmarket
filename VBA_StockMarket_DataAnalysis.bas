"Attribute VB_Name = ""Module1"""
Sub stockdata_analysis_sheetloop()

//////Loop - Starting
Dim ws As Worksheet
For Each ws In Worksheets

"'First, let's create the column titles for Ticker, Yearly Change, Percent Change, and Total Stock Volume using a loop..."
'with a series of if-then-elseifs to populate the titles depending on the looped value
For I = 9 To 12

  If I = 9 Then
"    ws.Cells(1, I).Value = ""Ticker"""
  
  ElseIf I = 10 Then
"    ws.Cells(1, I).Value = ""Yearly Change"""
    
  ElseIf I = 11 Then
"    ws.Cells(1, I).Value = ""Percent Change"""
    
  Else
"    ws.Cells(1, I).Value = ""Total Stock Volume"""
    
  End If

Next I

"'Second, let's set up a script that helps me find the bottom of the sheet and stores the end-row's value."
"'This utilizes the find function, using ""*"" to find the last possible cell with any value that isn't a blank and spits out its row."

Dim totalrowcount As Long

"  totalrowcount = ws.Cells.Find(What:=""*"", SearchDirection:=xlPrevious).Row"
  
'Now... here is the start of the actual macro.
  
Dim year_startopening As Double
Dim year_endclosing As Double
Dim year_change As Double
Dim year_changepercentage As Double
Dim stock_volume As Double
Dim summary_rowcounter As Double
summary_rowcounter = 2

'Going further: Defining Variables and setting their starting values.

'values to store the tickers for the respective data
Dim inc_topticker As String
Dim dec_topticker As String
Dim total_topticker As String

'values to store the requested data
Dim inc_topchange As Double
Dim dec_topchange As Double
Dim total_topvalue As Double

'setting the values at a neutral value
inc_topchange = 0
dec_topchange = 0
total_topvalue = 0
                
'Kick-off code to record the first data set's opening price at the start of the year.
"year_startopening = ws.Cells(2, 3).Value"

'Main beef of the assignment
For I = 2 To totalrowcount

"  'At the moment the ticker symbol is about to change, it will FIRST CHECK if the stock_volume = 0."
"  'If so, it will skip and publish the data set as a field of 0's..."
  '...while setting up the next summary row counter
  '...and setting up the variables to be used for the next ticker.
"  If ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1).Value And stock_volume = 0 Then"
        
        'ticker symbol
"        ws.Cells(summary_rowcounter, 9).Value = ws.Cells(I, 1).Value"
        
        'yearly change
"        ws.Cells(summary_rowcounter, 10).Value = 0"
        
        'percent change
"        ws.Cells(summary_rowcounter, 11) = Format(0, ""0.00%"")"
        
        'total stock volume
"        ws.Cells(summary_rowcounter, 12).Value = stock_volume"
        
        'green condition
"        ws.Cells(summary_rowcounter, 10).Interior.ColorIndex = 4"
        
        'Rest of the script to publish data is towards the end of the macro in this if-condition
"        'Once the summary of the current ticker's stock info has been published, the macros below will scrub and clean some values to be ready for use for the next ticker symbol data set."
        
        'holds the year's starting date's opening stock price to use for the next ticker symbol.
"        year_startopening = ws.Cells(I + 1, 3).Value"
        
        'resets the total volume count to zero.... just in case xD
        stock_volume = 0
        
        'To prepare the next line of the summary to write the next ticker's info...
    
            'we need to increase the counter by 1.
            summary_rowcounter = summary_rowcounter + 1

"  ElseIf ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1).Value Then"
    
    'When the ticker symbol does not match the symbol found in the below...
    
        'The macro below will first save the value of the end-of-year's closing-stock-price to be used in the final calculations.
"        year_endclosing = ws.Cells(I, 6).Value"
    
    'Listed below are the macros that will start calculations prior to publishing the calculated data.
    
        'calculates the change and the percentage of change when compared to the opening stock price at the start of the year.
        year_change = year_endclosing - year_startopening
        year_changepercentage = year_change / year_startopening
        
        'tops off the total stock volume bucket with the last volume value for the current ticker.
"        stock_volume = stock_volume + ws.Cells(I, 7).Value"
    
    'populates the data into the summary section.
    
        'ticker symbol
"        ws.Cells(summary_rowcounter, 9).Value = ws.Cells(I, 1).Value"
        
        'yearly change
"        ws.Cells(summary_rowcounter, 10).Value = year_change"
        
        'percent change
"        ws.Cells(summary_rowcounter, 11) = Format(0, ""0.00%"")"
"        ws.Cells(summary_rowcounter, 11).Value = year_changepercentage"
        
        'total stock volume
"        ws.Cells(summary_rowcounter, 12).Value = stock_volume"
        
"    'now, let's add some conditional formatting for the yearly change value."
    
        'let's use a if-else boolean script that will be based off the yearly change value.
        
"        If ws.Cells(summary_rowcounter, 10).Value < 0 Then"
            
                'red condition
"                ws.Cells(summary_rowcounter, 10).Interior.ColorIndex = 3"
            
            Else
                
                'green condition
"                ws.Cells(summary_rowcounter, 10).Interior.ColorIndex = 4"
        
        End If
    
            'Additional script start

            
            'Greatest % increase macro
            If year_changepercentage > inc_topchange Then
                inc_topchange = year_changepercentage
"                inc_topticker = ws.Cells(I, 1).Value"
            End If
            
            'Greatest % decrease macro
            If year_changepercentage < dec_topchange Then
                dec_topchange = year_changepercentage
"                dec_topticker = ws.Cells(I, 1).Value"
            End If
            
            'Greatest total volume
            If stock_volume > total_topvalue Then
                total_topvalue = stock_volume
"                total_topticker = ws.Cells(I, 1).Value"
            End If
            
            'Additional script end. Rest of the script to publish data is towards the end of the macro.
    
"    'Once the summary of the current ticker's stock info has been published, the macros below will scrub and clean some values to be ready for use for the next ticker symbol data set."
    
        'holds the year's starting date's opening stock price to use for the next ticker symbol.
"        year_startopening = ws.Cells(I + 1, 3).Value"
        
        'resets the total volume count to zero.
        stock_volume = 0
        
    'To prepare the next line of the summary to write the next ticker's info...
    
        'we need to increase the counter by 1.
        summary_rowcounter = summary_rowcounter + 1
        
  Else
  
    'The macros listed below will execute whenever the current cell's ticker is the same as the ticker symbol for the next cell below it.
    
"    stock_volume = stock_volume + ws.Cells(I, 7).Value"
    
    'Addition: Script will check if the current stored value of year_startopening is 0.
        'this will then make it equal to the stock-opening-value located in the row located under the current active cell.
        'Goal: This will help make sure we don't ever divide by 0 for the total change equation and cause a bug.
        
        If year_startopening = 0 Then
            
"            year_startopening = ws.Cells(I + 1, 3).Value"
        
        End If
        
  End If

Next I


'Additional Continuation Start

    'Generating the bonus table's column labels.
        For d = 16 To 17
        
          If d = 16 Then
"            ws.Cells(1, d).Value = ""Ticker"""
          
          ElseIf d = 17 Then
"            ws.Cells(1, d).Value = ""Value"""
            
          End If
        
        Next d
        
    'Generating the bonus table's row labels.
        For e = 2 To 4
        
          If e = 2 Then
"            ws.Cells(e, 15).Value = ""Greatest % Increase"""
"            ws.Cells(e, 16).Value = inc_topticker"
"            ws.Cells(e, 17) = Format(0, ""0.00%"")"
"            ws.Cells(e, 17).Value = inc_topchange"
            
          ElseIf e = 3 Then
"            ws.Cells(e, 15).Value = ""Greatest % Decrease"""
"            ws.Cells(e, 16).Value = dec_topticker"
"            ws.Cells(e, 17) = Format(0, ""0.00%"")"
"            ws.Cells(e, 17).Value = dec_topchange"
          
          ElseIf e = 4 Then
"            ws.Cells(e, 15).Value = ""Greatest Total Volume"""
"            ws.Cells(e, 16).Value = total_topticker"
"            ws.Cells(e, 17).Value = total_topvalue"
            
          End If
            
        Next e

'Additional Continuation End
                
Next
                            
End Sub
