Attribute VB_Name = "Module1"
Sub Stock_Data()

    ' SET CURRENT WORKSHEET AS A WORKSHEET OBJECT VARIABLE
    Dim WS As Worksheet


' LOOP THROUGH WORKSHEETS
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
    
        'ADD SUMMARY TABLE COLUMN HEADERS
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percentage Change"
        Cells(1, 12).Value = "Total Stock Volume"
        'CREATE NEW SUMAMRY TABLE LAYOUT FOR BONUS MATERIAL
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"

        'SET FIRST PRIORTY VARIABLES
        Dim Ticker As String
        Dim Yearly_Change, Percent_Change, Open_Price, Close_Price, Vol As Double
        'SET INITIAL VALUE OF PRIORTY NUMBER VARIABLES TO 0
        Yearly_Change = 0
        Open_Price = 0
        Close_Price = 0
        Vol = 0
        Percent_Change = 0
        
        'SET BONUS VARIABLES AND SET INITIAL VALUES TO ZERO IF NUMBERS
        Dim Greatest_Increase_Ticker, Greatest_Decrease_Ticker, Greatest_Total_Vol As String
        Dim Increase_Value, Decrease_Value, Most_Vol As Double
        Greatest_Increase_Ticker = " "
        Greatest_Decrease_Ticker = " "
        Greatest_Total_Vol = " "
        Increase_Value = 0
        Decrease_Value = 0
        Most_Vol = 0
        
        'KEEP TRACK OF THE LOCATION OF EACH TICKER SYMBOL IN THE SUMMARY TABLE
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        'DETERMINE FINAL ROW
        Dim lastrow As Long
        Dim i As Long

        lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        'SET INITIAL VALUE OF OPEN PRICE FOR FIRST TICKER
        Open_Price = Cells(2, 3).Value

        'LOOP THROUGH ALL STOCKS
        For i = 2 To lastrow
            
            'CHECK IF SAME TICKER SYMBOL
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                '========================FIND, AND CALCULATE PRIORTY AND BONUS VARIABLE VALUES===================
                'PRIORTY VARIABLE VALUES

                'SET TICKER SYMBOL
                Ticker = Cells(i, 1).Value
                
                'SET CLOSE PRICE LOOP VALUE
                Close_Price = WS.Cells(i, 6).Value
                'FIND YEARLY CHANGE
                Yearly_Change = Close_Price - Open_Price
                'CHECK DIVISION BY 0 CONDITION
                'IF NOT ZERO
                If Open_Price <> 0 Then
                    Percent_Change = (Yearly_Change / Open_Price) * 100
                Else
                    'IF ZERO
                    Percent_Change = 0
                End If
                'ADD TOTAL STOCK VOLUME
                Vol = Vol + Cells(i, 7).Value

                'BONUS VARIABLE VALUES
                'GREATEST PERCENT INCREASE
                If (Percent_Change > Increase_Value) Then
                    Increase_Value = Percent_Change
                    Greatest_Increase_Ticker = Ticker
                'GREATEST PERCENT DECREASE
                ElseIf (Percent_Change < Decrease_Value) Then
                    Decrease_Value = Percent_Change
                    Greatest_Decrease_Ticker = Ticker
                End If
                'GREATEST TOTAL VOLUME
                If (Vol > Most_Vol) Then
                    Most_Vol = Vol
                    Greatest_Total_Vol = Ticker
                End If
                '================================================================================================================================================

                '=====================PRINT PRIORTY VARIABLE VALUES TO SUMMARY TABLE
                '    VALUE TO SUMMARY TABLE
                'PRINT TICKER SYMBOL
                Range("I" & Summary_Table_Row).Value = Ticker
                'PRINT YEARLY CHANGE OF CURRENT TICKER
                Range("J" & Summary_Table_Row).Value = Yearly_Change
                'FILL CELL FOR YEARLY CHANGE WITH COLORS GREEN FOR POSITIVE OUTCOME AND RED FOR NEGATIVE OUTCOME
                If (Yearly_Change > 0) Then
                    'POSTIVE OUTCOME
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf (Yearly_Change <= 0) Then
                    'NEGATIVE OUTCOME
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                'PRINT PERECENT CHANGE
                Range("K" & Summary_Table_Row).Value = (CStr(Percent_Change) & "%")
                'PRINT TOTAL STOCK VOLUME
                Range("L" & Summary_Table_Row).Value = Vol
                
                'ADD 1 TO THE SUMMARY TABLE ROW
                Summary_Table_Row = Summary_Table_Row + 1
                'RESET VARIABLE VALUES FOR NEW TICKER
                Yearly_Change = 0
                Close_Price = 0
                Open_Price = Cells(i + 1, 3).Value
                Vol = 0

            'IF CELL IMMEDIATELY FOLLOWING A ROW IS THE SAME TICKER...
            Else
                'ADD TO TOTAL STOCK VOLUME
                Vol = Vol + Cells(i, 7).Value
            End If
                
                

        Next i

            
        '   VALUES TO NEW BONUS TABLE LAYOUT
                'PRINT TICKER SYMBOL FOR GREATEST PERCENT INCREASE
                Range("P2").Value = Greatest_Increase_Ticker
                'PRINT TICKER SYMBOL FOR GREATEST PERCENT DECREASE
                Range("P3").Value = Greatest_Decrease_Ticker
                'PRINT TICKER SYMBOL FOR GREATEST TOTAL VOLUME
                Range("P4").Value = Greatest_Total_Vol
                'PRINT GREATEST PERCENT INCREASE
                Range("Q2").Value = (CStr(Increase_Value) & "%")
                'PRINT GREATES PERCENT DECREASE
                Range("Q3").Value = (CStr(Decrease_Value) & "%")
                'PRINT GREATES TOTAL VOLUME
                Range("Q4").Value = Most_Vol

    Next WS

End Sub
