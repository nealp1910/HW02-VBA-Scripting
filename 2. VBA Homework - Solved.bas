Attribute VB_Name = "Module1"
Sub year_stock()

    Dim ws As Worksheet
        For Each ws In ActiveWorkbook.Worksheets
        ws.Activate

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'Code For ticker and Total Stock Volume
    '---------------------------------------------------------
    
        Dim Ticker As String
    
        'Set an initial varialble for holding total stock volumer per ticker
        Dim Total_Stock As Double
        Total_Stock = 0
    
        'Keep track of location for each ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    
        'counting the number of rows
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
        'Loop through all tickers
        For i = 2 To lastrow
    
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
                'Set ticker name
                Ticker = Cells(i, 1).Value
            
                'Add total stock volume
                Total_Stock = Total_Stock + Cells(i, 7).Value
                
                'Print the Ticker Symbol in the Summary Table
                Range("I" & Summary_Table_Row).Value = Ticker
                          
                'Print total stock value in the summary table
                Range("L" & Summary_Table_Row).Value = Total_Stock
                
                'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Reset stock volume
                Total_Stock = 0
            
         'If the cell immediately following a row as the same ticker...
         Else
         
             'Add total stock volume
             Total_Stock = Total_Stock + Cells(i, 7).Value
                        
        End If
        
        Next i
    
    
    'Code For Yearly Change and Percentage Change
    '---------------------------------------------------------
    'Set an initial variable for yearly change per ticker
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    
    'Resetting Summary Table
    Summary_Table_Row = 2
    
    'Grab Open Price Per Ticker
    OpenPrice = Cells(2, 3).Value
        
    'Loop through all tickers (lastrow already called above)
    For i = 2 To lastrow
    
        'Check to see if same Ticker name used
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'Add Closing Price to Variable
            ClosePrice = Cells(i, 6).Value
            
            'Yearly Change Formula = Last Day - First Day
            YearlyChange = ClosePrice - OpenPrice
            
            'Equation for Percent
            PercentChange = YearlyChange / OpenPrice
            
            'Change Open Price
            OpenPrice = Cells(i + 1, 3).Value
            
            'Print Year Change in the Summary Table
            Range("J" & Summary_Table_Row).Value = YearlyChange
            
            'Print Percent Change in the Summary Table
            Range("K" & Summary_Table_Row).Value = PercentChange
            
            '% Format for Percent Change
            Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        
            'Add 1 to the Summary Table
            Summary_Table_Row = Summary_Table_Row + 1
                        
        End If
               
    Next i
    
    'Color Code for the Yearly Change
    '----------------------------------------------------------------
    For j = 2 To lastrow
        
            If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
                Cells(j, 10).Interior.ColorIndex = 4
                
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
                
            End If
            
    Next j
    
    Next ws
    
End Sub

