Sub Stock_Data()
    
    'loop through all worksheets
    For Each ws In Worksheets
        'Create column headers for data ouput
        [I1:L1] = Split("Ticker Yearly_Change Percent_Change Total_Stock_Volume")
        
        'set variable for ticker
        Dim Ticker As String
        'set variable for yearly change
        Dim Yearly As Double
        ' set variable for total stock volume
        Dim Vol As Variant
        'set variables to = 0
        Yearly = 0
    
        Vol = 0
    
    
    
        endRow = Cells(Rows.count, "A").End(x1up).Row
        
        'Loop Through all stocks
        For i = 2 To endRow
        
            'check if we are still within same ticker symbol group
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'Set Ticker Symbol
            Ticker = Cells(i, 1).Value