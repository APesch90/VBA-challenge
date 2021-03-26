Attribute VB_Name = "Module2"
Sub Stock_Market_Analysis():

    'Define the ticker variable and set up column header
    Dim Ticker As String
    Cells(1, "I").Value = "Ticker"
                      
    'Define the yearly change variable and set up column header
    Dim Yearly_Change As Double
    Cells(1, "J").Value = "Yearly Change"
        
    'Define the percent change variable and set up column header
    Dim Percent_Change As Double
    Cells(1, "K").Value = "Percent Change"
        
    'Define the total stock volume variable as a counter to hold the total stock volume per ticker and set up column header
    Dim Total_Counter As Double
    Cells(1, "L").Value = "Total Stock Volume"
    
    'Start the counter at 0
    Total_Counter = 0
    
    'Define the location of the beginning point for populating the data starting with row 2
    Dim Summary_table_row As Long 'Long because there are >32,767 rows
    Summary_table_row = 2
    
    'Define first value of the year
    Dim First_value As Double
    
    'Define closing value of the year
    Dim End_value As Double
    
    'Define lastrow
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Define the first value to hold the very first stock price before the loop starts (where there's a change recognized)
    First_value = Cells(2, "C").Value
    
    'Loop through all rows under the column headers
    For i = 2 To lastrow
    
        'Use a conditional to verify whether the ticker is still the same, let's start with if it is NOT...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
                    'Keep track of the ticker symbol
                    Ticker = Cells(i, 1).Value
                
                    'Print the ticker symbols in the Summary Table
                    Range("I" & Summary_table_row).Value = Ticker
                                      
                    'Establish the closing value of the ticker
                    End_value = Cells(i, 6).Value
                    
                    'Establish the yearly change calculation
                    Yearly_Change = End_value - First_value
                    
                    'Print the yearly change in the Summary Table
                    Range("J" & Summary_table_row).Value = Yearly_Change
                    
                    'To avoid a bug for Percent_Change for zero as the first value in the formula, set up a condition to check for this first, then do the full calculation
                    
                        If First_value = 0 Then
                        
                            Percent_Change = 0
                            
                        Else
                        
                            'Establish the percent change calculation
                            Percent_Change = (End_value - First_value) / First_value
                            
                        End If
            
                    'Print the percent change in the Summary Table
                    Range("K" & Summary_table_row).Value = Format(Percent_Change, "Percent")
                                 
                    'Add to the Total Counter
                    Total_Counter = Total_Counter + Cells(i, 7).Value
                    
                    'Print the total stock volume in the Summary Table
                    Range("L" & Summary_table_row).Value = Total_Counter
                    
                    'Reset the total counter
                    Total_Counter = 0
                    
                    'Establish the next ticker's first value
                    First_value = Cells(i + 1, 3).Value
                    
                    'Create a conditional for the red and green cell fill formatting (negative is red, positive is green)
                
                    If Yearly_Change < 0 Then
                   
                        Worksheets("2015").Range("J" & Summary_table_row).Interior.ColorIndex = 3
                        
                    Else
                
                        Worksheets("2015").Range("J" & Summary_table_row).Interior.ColorIndex = 4
                       
                    End If
                    
                        'Add one to the summary table row
                    Summary_table_row = Summary_table_row + 1
                    
                Else
                         
                    'Add to the Total Counter
                    Total_Counter = Total_Counter + Cells(i, 7).Value
                                  
        End If

    Next i

End Sub

