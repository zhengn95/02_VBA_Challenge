Attribute VB_Name = "Module1"
Sub MultYear_Stock_Data():

'Loop across worksheets
Dim ws As Worksheet
For Each ws In Worksheets

'Part 1: Create a script that loops through all the stocks for one year and outputs: the ticker, yearly change, percent change, and total stock volume

'Create headers for the outputs listed above
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'Track the summary table row location for each variable
Dim summary_row As Integer
summary_row = 2

'Set an initial variable for holding the ticker initials, yearly change, percent change, total stock volume
Dim ticker As String
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock As Double
    'set total_stock to 0
    total_stock = 0

'Set opening price variable
    Dim opening_price As Double
    opening_price = Cells(2, 3).Value

'Set last_row variable
Dim last_row As Long
Set ws = ActiveSheet

last_row = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

'loop through all ticker initials
For i = 2 To last_row
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        'Set the ticker to the first column
        ticker = Cells(i, 1).Value
    
        'Set closing_price to the 6th columns. Subtract closing price from opening price
        closing_price = Cells(i, 6).Value
        yearly_change = (closing_price - opening_price)
            
        'Add up all total stock volumes
        total_stock = total_stock + Cells(i, 7).Value
        
            'Calculate percent change. Avoid Division by zero by including a nested conditional statement
            If opening_price <> 0 Then
                percent_change = (yearly_change / opening_price)
            
            Else
                percent_change = 0
            
            End If


        'Print the ticker symbols, yearly change, and percent change into the output listed with the corresponding header
        Range("I" & summary_row).Value = ticker
        Range("J" & summary_row).Value = yearly_change
        Range("K" & summary_row).Value = percent_change
            'Format percent change to percent
            Range("K2:K" & last_row).NumberFormat = "0.00%"
            Range("K2:K" & last_row).Value = Range("K2:K" & last_row).Value
        
        Range("L" & summary_row).Value = total_stock
    
       'Update variables. These variables are updated outside the initial calculations above. They are updated to perform similar calculations in the rows within the loop
        summary_row = summary_row + 1
        opening_price = Cells(i + 1, 3).Value
        
        'Reset total stock volume
        total_stock = 0
        
    Else
        total_stock = total_stock + Cells(i, 7).Value
    
    End If
    
Next i
    
'-----------------
'Part 2: Greatest Values Table

'Set titles for greatest % increase, greatest % decrease, and greatest total volume
Range("N2").Value = "Greatest % Increase"
Range("N3").Value = "Greatest % Decrease"
Range("N4").Value = "Greatest Total Volume"

'Set corresponding headers to input values to the above variables
Range("O1") = "Ticker"
Range("P1") = "Value"

'Set last_row variable
last_row = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row

'Format greatest increase and greatest decrease as percent
Range("P2").NumberFormat = "0.00%"
Range("P2").Value = Range("P2").Value

Range("P3").NumberFormat = "0.00%"
Range("P3").Value = Range("P3").Value
    
'Find the greatest % increase by using the max function
Cells(2, 16).Value = WorksheetFunction.Max(Range("K2", "K" & last_row).Value)
      
'Calculate the greatest % decrease
Cells(3, 16).Value = WorksheetFunction.Min(Range("K2", "K" & last_row).Value)

'Calculate the greatest total volume
Cells(4, 16).Value = WorksheetFunction.Max(Range("L2", "L" & last_row).Value)

'Index and match the ticker to values
Cells(2, 15).Value = WorksheetFunction.Index(Range("I2:I" & last_row), WorksheetFunction.Match(Cells(2, 16), Range("K2:K" & last_row), 0))
Cells(3, 15).Value = WorksheetFunction.Index(Range("I2:I" & last_row), WorksheetFunction.Match(Cells(3, 16), Range("K2:K" & last_row), 0))
Cells(4, 15).Value = WorksheetFunction.Index(Range("I2:I" & last_row), WorksheetFunction.Match(Cells(4, 16), Range("L2:L" & last_row), 0))

'------------
'Part 3: Conditional Formatting

'Set last_row variable
last_row = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row

'yearly change
For j = 2 To last_row
    
    If Cells(j, 10).Value > 0 Then
        Cells(j, 10).Interior.ColorIndex = 4
        
    ElseIf Cells(j, 10).Value < 0 Then
        Cells(j, 10).Interior.ColorIndex = 3
      
    Else
        Cells(j, 10).Interior.ColorIndex = 6
        
    End If

Next j

'percent change
For k = 2 To last_row
    
    If Cells(k, 11).Value > 0 Then
        Cells(k, 11).Interior.ColorIndex = 4
        
    ElseIf Cells(k, 10).Value < 0 Then
        Cells(k, 11).Interior.ColorIndex = 3
      
    Else
        Cells(k, 11).Interior.ColorIndex = 6
        
    End If

Next k

'--------------------
Next ws
End Sub
