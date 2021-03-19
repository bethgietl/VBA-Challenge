'Create a script that will loop through all the stocks for one year and output the following information.
    'The ticker symbol.
    'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The total stock volume of the stock.
'You should also have conditional formatting that will highlight positive change in green and negative change in red.

Sub Stock_Data()

'----------1. Define variables: use double for decimals, integer as data type, LongLong for really long numbers, string as sequence of characters
    ' set inital variable for holding the ticker
    Dim Ticker As String
    
    ' set inital variable for holding the total volume
    Dim Total_Volume As LongLong
    ' declare variable = 0 (duplicate)
    Total_Volume = 0
    
    ' set initial variables for holding the stock open, stock close, yearly change and percent change
    Dim Year_Open, Year_Close, Yearly_Change, Percent_Change As Double
    ' declare yearly change = 0, the rest of the variables are below the For Each
    Yearly_Change = 0

'------------2. start with going through all the worksheets
For Each ws In Worksheets
    
'------------3. insert column headers to put the extracted data, format columns, set values inside loop, keep track of location of table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'format columns to adjust the width of the columns I-L depending on text
    ws.Range("I:L").Columns.AutoFit
    ws.Columns("K").NumberFormat = "0.00%"

    'declare variables inside for loop to move to the beginning of next ticker
        'year open needs to be declared the first cell in order to move through each ticker
    Year_Open = ws.Cells(2, 3).Value
    
    Year_Close = 0
    Percent_Change = 0

    'keep track of the location for each ticker in the summary table (used the Ticker activity as a reference)
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

'------------4. find last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
'------------5. loop through all Tickers, when the ticker value changes down every line in the column
                'Create a For i loop to summarize Tickers, yearly change, percent change and total volume
                'used credit_charges VBA activity
    
    For i = 2 To LastRow
        'The indent makes this now a nested loop inside the For Each ws loop
      
'------------6. conditional statement to check if we are still within the same Ticker, if it is not...
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
          ' Set the Ticker name
          Ticker = ws.Cells(i, 1).Value
    
          ' Add to the Volume Total
          Total_Volume = Total_Volume + ws.Cells(i, 7).Value
          
          ' Print the Ticker in the Summary Table
          ws.Range("I" & Summary_Table_Row).Value = Ticker
    
          ' Print the Volume Total Amount to the Summary Table
          ws.Range("L" & Summary_Table_Row).Value = Total_Volume
                    
          ' Set the Yearly Change components
          Year_Close = ws.Cells(i, 6).Value
            'Year Open was defined in variables
          
          ' Find yearly change
          Yearly_Change = (Year_Close - Year_Open)
    
          ' Print the yearly change Amount to the Summary Table
          ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
          
          'Conditional formatting yearly change - green=positive, red=negative
            If Yearly_Change > 0 Then
                
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            
            Else
                
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            
            End If
          
            ' Find percent change
                'kept getting error 11: division by zero error so need to create a conditional statement
            If Year_Open = 0 Then
                
                Percent_Change = 0
            
            Else
                
                Percent_Change = (Yearly_Change / Year_Open)
            
            End If
            
          ' Print the Percent change to the Summary Table
          ws.Range("K" & Summary_Table_Row).Value = Percent_Change
          
          ' Add one to the summary table row
          Summary_Table_Row = Summary_Table_Row + 1
          
          ' Reset the Volume Total
          Total_Volume = 0
          
          ' Reset the Year open
          Year_Open = ws.Cells(i + 1, 3).Value
    
        ' If the cell immediately following a row is the same ticker...
        Else
    
          ' Add to the Volume Total
          Total_Volume = Total_Volume + ws.Cells(i, 7).Value
       
        End If
 
      Next i
    
    Next ws

End Sub
