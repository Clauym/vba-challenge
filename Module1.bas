Attribute VB_Name = "Module1"
Sub Calculating()
' Subrutine that displays opening and closing price for the year. Percentage change from the opening price at the beginning of a year and the closing price at the end of that year. _
        The total stock volume of the stock . During the year?

Dim All_wks As Worksheet
Dim Stk_vol As Double, EOY_price As Double, BOY_price As Double, Yr_chg As Double, Pct_chg As Double
Dim w As Integer, s As Integer, r As Double
Dim Ticker As String

'Call routine that clears data
Call Initialize

'Check every worsheet, calculates changes and copy values to summary table
For Each All_wks In Worksheets
    'Count rows of database
    Num_rows = All_wks.Range("A:A").SpecialCells(xlCellTypeLastCell).Row
  
    'Set initial values for the variables
    Ticker = ""
    r = 1
    s = 1
    
    'Select ticker and price at the beginning of the year. Begins adding up the volume
    Do While r <= Num_rows
    
        If All_wks.Cells(r + 1, 1).Value <> Ticker Then
            Ticker = All_wks.Cells(r + 1, 1).Value
            BOY_price = All_wks.Cells(r + 1, 3).Value
            Stk_vol = All_wks.Cells(r + 1, 7).Value
            r = r + 1
            
        Else    'Obtains last price of the year and adds up the volume
            Do While All_wks.Cells(r + 1, 1).Value = Ticker
                EOY_price = All_wks.Cells(r + 1, 6).Value 'allwks
                Stk_vol = All_wks.Cells(r + 1, 7).Value + Stk_vol   'allwks
                r = r + 1
            Loop
        
            Yr_chg = EOY_price - BOY_price
            Pct_chg = Yr_chg / BOY_price
       
            All_wks.Cells(s + 1, 9).Value = Ticker
            All_wks.Cells(s + 1, 10).Value = Yr_chg
    
            All_wks.Cells(s + 1, 11).Value = Pct_chg
            All_wks.Cells(s + 1, 12).Value = Stk_vol
            
            'Condittional formating for cells. Red negative values, green positive values.
            If Yr_chg < 0 Then
                All_wks.Cells(s + 1, 10).Interior.ColorIndex = 3  'vbRed
            Else
                All_wks.Cells(s + 1, 10).Interior.ColorIndex = 4  'vbGreen
            End If
            
            'Counter for the table with results
            s = s + 1
        
        End If
    Loop
Next All_wks
 
 'Call routine that calculate the stock with greatest changes
 Call Greatest

End Sub

Sub Initialize()
' Routine that begins the process by clearing all the data in the output columns.

' Initial message
Ini = MsgBox("Calculate the changes in price and volume for the stocks in each of the years from 2018 to 2020", vbOKOnly, "Created by Claudia Yurrita")

Dim wks As Worksheet

'Clear the range of calculations in all the worksheets and write down the name of the columns being calculated

For Each wks In Worksheets

    wks.Range("I:U").Clear

    wks.Range("I1").Value = "Ticker"
    wks.Range("J1").Value = "Yearly Change"
    wks.Range("K1").Value = "Percent Change"
    wks.Range("L1").Value = "Total Stock Volume"
    
    wks.Range("P1").Value = "Ticker"
    wks.Range("Q1").Value = "Value"
    
    wks.Range("O2").Value = "Greatest % Increase"
    wks.Range("O3").Value = "Greatest % Decrease"
    wks.Range("O4").Value = "Greatest Total Volume"
    
    wks.Range("J:J").NumberFormat = "#0.00"
    wks.Range("K:K,Q2:Q3").NumberFormat = "0.00%"
    wks.Range("L:L,Q4").NumberFormat = "#,000"
    
    wks.Range("I:I").ColumnWidth = 12
    wks.Range("J:L, Q:Q").ColumnWidth = 18
    wks.Range("O:O").ColumnWidth = 22
    
    
Next wks

End Sub

Sub Greatest()
'Displays the stocks with the greatest changes during each year

Dim wks As Worksheet, Max_Row As Double, i As Double, t As Double
Dim Great_Incr As Double, Great_Decr As Double, Great_Vol As Double, GIncr_Stock As String, GDecr_Stock As String, GVol_Stock As String

'Finds the stocks with biggest changes for each Worksheet in the Workbook
For Each wks In Worksheets
    'Count the rows
    Max_Row = wks.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Obtain the biggest values
    wks.Activate
    Great_Incr = Application.Max(Range(Cells(2, 11), Cells(Max_Row, 11)))
    Great_Decr = Application.Min(Range(Cells(2, 11), Cells(Max_Row, 11)))
    Great_Vol = Application.Max(Range(Cells(2, 12), Cells(Max_Row, 12)))
    
    'Check every cell to obtain the name of the stocks with the biggest changes and oeprtaed volumen
    For i = 1 To Max_Row
    
        If wks.Cells(i, 11).Value = Great_Incr Then                        '
            GIncr_Stock = wks.Cells(i, 9).Value
        End If
    
        If wks.Cells(i, 11).Value = Great_Decr Then
            GDecr_Stock = wks.Cells(i, 9).Value
        End If
        
        If wks.Cells(i, 12).Value = Great_Vol Then                        '
            GVol_Stock = wks.Cells(i, 9).Value
        End If
    Next i
      
    'Copy the name  of the selected stocks and their corresponding values
    wks.Cells(2, 16).Value = GIncr_Stock
    wks.Cells(2, 17).Value = Great_Incr
    wks.Cells(3, 16).Value = GDecr_Stock
    wks.Cells(3, 17).Value = Great_Decr
    wks.Cells(4, 16).Value = GVol_Stock
    wks.Cells(4, 17).Value = Great_Vol

Next wks
    
End Sub


