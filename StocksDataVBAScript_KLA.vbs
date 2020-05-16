Attribute VB_Name = "SummaryCalcs"
Option Explicit

Sub STOCKSDATAWSLOOPING()

'+++++++++++++++++++++++++++++++++++++++++++++
'CREATING LOOP THRU ALL WORKSHEETS
'+++++++++++++++++++++++++++++++++++++++++++++
 
Dim ws As Worksheet
For Each ws In Worksheets

'Declaring Variables for entire project
    
    Dim Ticker As String
    Dim Yearly_Chg As Double
    Dim Percent_Chg As Double
    Dim Total_Stock As LongLong
    Dim i As LongLong
    Dim SummaryRow As Integer
    Dim Lastrow As Long
    Dim Open_stock As Double
    Dim close_stock As Double
    Dim highest_chg As Double
    Dim lowest_chg As Double
    Dim greatest_total As LongLong
    
    
'Setting up Summary Table and initial values of calculations
    
    SummaryRow = 2
    Total_Stock = 0
    Yearly_Chg = 0
    Percent_Chg = 0
    
'Establishing beginning values for calculating cells
    lowest_chg = 0
    highest_chg = 0
    greatest_total = 0
    
    'Establishing Last Row
     Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Setting Opening stock rate outside of loop.
     Open_stock = ws.Cells(2, 3).Value                          'set certain values outside of loop , same ticker logic
    
    'Setting headers in Table
     ws.Cells(1, 10) = "Ticker"
     ws.Cells(1, 11) = "Yearly Change"
     ws.Cells(1, 12) = "Percent Change"
     ws.Cells(1, 13) = "Total Stock Volume"
     ws.Cells(2, 15) = "Greatest % Increase"
     ws.Cells(3, 15) = "Greatest % Decrease"
     ws.Cells(4, 15) = "Greatest Total Volume"
     ws.Cells(1, 15) = "Bonus Data"
     ws.Cells(1, 16) = "Ticker"
     ws.Cells(1, 17) = "Value"
     
    
               
    '+++++++++++++++++++++++++++++++++++++++++++++
    'CREATING LOOP THRU ALL ROWS OF DATA
    '+++++++++++++++++++++++++++++++++++++++++++++
   
        For i = 2 To Lastrow                                                                    'Loop through all ticker signs to plug into Summary Table
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then                                  '(if same ticker, then)
                Ticker = ws.Cells(i, 1).Value                                                      ' Set the Ticker name
                ws.Range("J" & SummaryRow).Value = Ticker                                          ' Print the ticker name to the Summary
                
                Total_Stock = Total_Stock + ws.Cells(i, 7).Value                                   'calculate total stock sum
                ws.Range("m" & SummaryRow).Value = Total_Stock                                     ' Print the yearly change to the Summary
                        
                close_stock = ws.Cells(i, 6).Value
                ws.Range("k" & SummaryRow).Value = close_stock - Open_stock
        
        'Conditional formatting on Yearly_chg that will color positive change in green
                    If ws.Range("k" & SummaryRow).Value > 0 Then
                        ws.Range("k" & SummaryRow).Interior.Color = RGB(0, 255, 0)
                    Else
                    End If
        'Conditional formatting on Yearly_chg that will color negative change in red
                    If ws.Range("k" & SummaryRow).Value < 0 Then
                        ws.Range("k" & SummaryRow).Interior.Color = RGB(255, 0, 0)
                    Else
                    End If
        
        'To fix division by 0 error
                If Open_stock > 0 Then                                                              'if, then, else to avoid dividing by 0
                    Percent_Chg = (close_stock - Open_stock) / Open_stock
                    ws.Range("l" & SummaryRow).Value = Percent_Chg
                Else
                    ws.Range("l" & SummaryRow).Value = 0
                End If
         
        'Calculating highest and lowest % change and Highest Total Stock Values
                     If Percent_Chg > highest_chg Then                                              'calculating highest % change
                        highest_chg = Percent_Chg
                        ws.Range("q2") = highest_chg
                        ws.Cells(2, 16) = Ticker
                     Else
                     End If
                    
                     If Percent_Chg < lowest_chg Then                                               'calculating lowest % change
                        lowest_chg = Percent_Chg
                        ws.Range("q3") = lowest_chg
                        ws.Cells(3, 16) = Ticker
                     Else
                     End If
                     
                     If Total_Stock > greatest_total Then                                           'calculating Highest Total Stock
                        greatest_total = Total_Stock
                        ws.Range("q4") = greatest_total
                        ws.Cells(4, 16) = Ticker
                     Else
                     End If
               
               
                SummaryRow = SummaryRow + 1                                                         'Add one row to the summary
                Total_Stock = 0                                                                     'Reset the total
                Open_stock = ws.Cells(i + 1, 3).Value
                
        Else
            Total_Stock = Total_Stock + ws.Cells(i, 7).Value                                        'If the cell immediately following a row is the same ticker...
        End If
          
    Next i
    

'FORMATING SECTION
ws.Range("K2:k" & Lastrow).NumberFormat = "$#,##0.00"
ws.Range("L2:L" & Lastrow).NumberFormat = "0.00%"
ws.Range("A1:Q1").Font.FontStyle = "bold"
ws.Range("q2:q3").NumberFormat = "0.00%"
ws.Range("a1:Z" & Lastrow).EntireColumn.AutoFit
                    

'Closing the Worksheet Loop
Next

MsgBox ("Now that is some spectacular coding that deserves an A+!")
MsgBox ("Click OK if you agree.")
MsgBox ("Just Kidding!")
MsgBox ("Sort of...=))")

End Sub


