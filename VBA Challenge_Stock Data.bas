Attribute VB_Name = "Module1"
Sub Stock_Totals()
    
    'Loop through all years in each worksheet
    For Each ws In Worksheets
        
        ' Create a Totals_Table
        
        'Create a variable for Ticker symbols
        Dim Ticker  As String
        
        'Create a variable for total stock Volume
        Dim Volume  As Double
        Volume = 0
        
        'Create a variable to correctly place new Ticker data in the appropriate, following row on the Totals_Table
        Dim Totals_Table_Row As Integer
        Totals_Table_Row = 2
        
        'Loop through all stocks to pull data into Totals_Table
        For r = 2 To Cells(i, 1).End(xlDown)
            
            'Evaluate each new Ticker symbol
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                'Set the Ticker symbol
                Ticker = Cells(i, 1).Value
                
                'Add to the total stock Volume
                Volume = Volume + Cells(i, 7).Value
                
                'Print the Ticker symbol into the Totals_Table
                Range("I" & Totals_Table_Row).Value = Ticker
                
                'Print the total stock Volume into the Totals_Table
                Range("L" & Totals_Table_Row).Value = Volume
                
                'Add one to the Totals_Table_Row to move into the next row when we loop through a new Ticker symbol
                Totals_Table_Row = Totals_Table_Row + 1
                
                'Reset the total stock Volume
                Volume = 0
                
            Else
                'If Cells(i + 1, 1).Value = Cells(i, 1).Value Then...
                Dim Close_Start As Double
                Dim Close_End As Double
                
                Close_Start = Cells(i, 5).Value
                Close_End = Cells(i, 5).End(xlDown)
                
                'Calculate yearly change
                Range("J" & Totals_Table_Row).Value = Close_Start - Close_End
                
                'Calculate percent change
                Range("K" & Totals_Table_Row).Value = Close_Start - Close_End / Close_Start
                
                'Add to the total stock Volume
                Volume = Volume + Cells(i, 7).Value
                
            End If
            
        Next r
        
    Next ws
    
End Sub


