Attribute VB_Name = "Module1"
Sub stock_data_analysis()

'INSTRUCTION:Create a script that loops through all the stocks for one year and outputs the following information
    
    'define variable
ws = Worksheet
    'Enable script to run on multiple worksheets
For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    
   
        'define the output variable and location (row 1 is a header)
        output_row = 2
     
    'initialize loop, variables defined
        yearly_change = 0
        counter = 0
        opening_price = 0
        closing_price = 0
        percent_change = 0
        total_volume = 0
     
    'for loop set range
    For row_ = 2 To 759002
    
        'Increment the counter
        counter = counter + 1
        
        'INSTRUCTION-TOTAL VOLUME
        total_volume = total_volume + Cells(row_, 7).Value
        
        'where to put total stock volume
        Cells(output_row, 12).Value = total_volume
        
        If counter = 1 Then opening_price = Cells(row_, 3).Value
    
        
        'INSTRUCTION-THE TICKER SYMBOL i.e. when does it change from one to another?
            
            'If the values don't match
             If Cells(row_ + 1, 1).Value <> Cells(row_, 1).Value Then
              
            'put the original value in the designated output location
            Cells(output_row, 9).Value = Cells(row_, 1).Value
             
               
        'INSTRUCTION-YEARLY CHANGE
            'assigning a value
            closing_price = Cells(row_, 6).Value
              
            'calculating yearly change
            yearly_change = closing_price - opening_price
            'where to put yearly_change- assigning location
            Cells(output_row, 10).Value = yearly_change
             
             
        'INSTRUCTION-PERCENT CHANGE
            'calculate percent change
            percent_change = closing_price / opening_price - 1
             
            'where to put percent change
            Cells(output_row, 11).Value = percent_change
                
               
                 'move to another row in the output location for each different value
                 output_row = output_row + 1

             
             'resetting counter before new ticker
             counter = 0
             total_volume = 0
                
        End If
    Next row_
    
'INSTRUCTION return the stock with the greatesst % increase, greatest % decerase, greatest total volume
    
      'identify ranges
      change_range = Range("K2:K759001")
      volume_range = Range("L2:L759001")
      ticker_range = Range("I2:I759001")
      

        'find min and max change values in range
        min_change = Application.WorksheetFunction.Min(change_range)
        max_change = Application.WorksheetFunction.Max(change_range)

            'location for result of min max
            Cells(2, 16).Value = max_change
            Cells(3, 16).Value = min_change

        'find max total volume
        max_volume = Application.WorksheetFunction.Max(volume_range)
            
            'location for result of max_volume
            Cells(4, 16).Value = max_volume


    'find rows of greatest increase, decrease, greatest total volume and extract ticker value to new column

    Set increase_cell = Range("K2:K759001").Find(max_change)
    increase_row = Application.WorksheetFunction.Match(max_change, change_range, 0)

    'ticker of greatest percent increase
    ticker_of_increase = Cells(increase_row + 1, 9).Value
    Cells(2, 15).Value = ticker_of_increase

    'greatest percent decrease
    Set decrease_cell = Range("K2:K759001").Find(min_change)
    decrease_row = Application.WorksheetFunction.Match(min_change, change_range, 0)
   
   'ticker of greastest percent decrease
    ticker_of_decrease = Cells(decrease_row + 1, 9).Value
    Cells(3, 15).Value = ticker_of_decrease

    'total greatest volume
    
    Set volume_cell = Range("L2:L759001").Find(max_volume)
    volume_row = Application.WorksheetFunction.Match(max_volume, volume_range, 0)
    
    'ticker of greastest total volume
    ticker_of_volume = Cells(volume_row + 1, 9).Value
    Cells(4, 15).Value = ticker_of_volume

    
    Next ws

End Sub
