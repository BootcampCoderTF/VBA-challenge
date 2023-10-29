Sub Summarise_Data()


Dim ws As Worksheet

'repeat for every worksheet in the workbook
For Each ws In ThisWorkbook.Worksheets

'Add the table headers to the worksheet
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percentage Change"
ws.Range("L1").Value = "Total Stock Volume"

'Add row headers and table headers for greatest values
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

'Defines the types of the variable value used in the table
Dim LR As Long
Dim ticker_name As String
Dim yearly_change As Double
Dim percentage_change As Double
Dim stock_total As Double

Dim close_sumRange As Range
Dim close_criteriaRange As Range
Dim close_criteria As String
Dim open_sumRange As Range
Dim open_criteriaRange As Range
Dim open_criteria As String
Dim close_price_total As Double
Dim open_price_total As Double

Dim greatest_percent_inc_ticker As String
Dim greatest_percent_inc_value As Double
Dim greatest_percent_dec_ticker As String
Dim greatest_percent_dec_value As Double
Dim greatest_total_ticker As String
Dim greatest_total_value As Double

'------------------------------------------------------------
'first we filter the tickers
    'set the ticker total to nothing
    stock_total = 0
    
    'Keep track of the location for each ticker in the summary table
    Dim summary_table_row As Integer
    summary_table_row = 2

    'variable to count the number of rows
    LR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    
    'loop to filter the tickers
    For j = 2 To LR
        'Check if we are still within the same ticker, if it is not...
        If ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value Then
        
        'Set the ticker name
        ticker_name = ws.Cells(j, 1).Value
        
        'We now can work out the yearly change between close price and open price
        
        Set close_sumRange = ws.Range("F:F")
        Set close_criteriaRange = ws.Range("A:A")
        close_criteria = ticker_name
        
        close_price_total = WorksheetFunction.SumIfs(close_sumRange, close_criteriaRange, close_criteria)
        
        Set open_sumRange = ws.Range("C:C")
        Set open_criteriaRange = ws.Range("A:A")
        open_criteria = ticker_name
        
        open_price_total = WorksheetFunction.SumIfs(open_sumRange, open_criteriaRange, open_criteria)
        
        yearly_change = close_price_total - open_price_total
        
        'We then work out the percentage change between close price and open price
        percentage_change = ((Range("F" & summary_table_row).Value - Range("C" & summary_table_row).Value) / Range("C" & summary_table_row).Value) * 100
        
        'Add to the stock total
        stock_total = stock_total + ws.Cells(j, 7).Value
        
        '------------------------------------------------------------
        'Print the ticker name in the Summary Table
        ws.Range("I" & summary_table_row).Value = ticker_name
        
        'Print the yearly change in the Summary Table
        ws.Range("J" & summary_table_row).Value = yearly_change
        If ws.Range("J" & summary_table_row).Value >= 0 Then
            ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
            Else: ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
        End If
        
        'Print the percentage change in the Summary Table
        ws.Range("K" & summary_table_row).Value = FormatPercent(percentage_change)
        
        'Print the stock total to the Summary Table
        ws.Range("L" & summary_table_row).Value = stock_total
        
        'Add one to the summary table row
        summary_table_row = summary_table_row + 1
        '------------------------------------------------------------

        'Reset the stock total
        stock_total = 0
    
        'If the cell immediately following a row is the same ticker...
        Else
            'Add to the stock total
            stock_total = stock_total + ws.Cells(j, 7).Value
        End If
    Next j
    
    greatest_percent_inc_value = WorksheetFunction.Max(ws.Range("J:J").Value)
    ws.Cells(2, 17).Value = greatest_percent_inc_value
    ws.Cells(2, 16).Value = greatest_percent_inc_ticker
    
    greatest_percent_dec_value = WorksheetFunction.Min(ws.Range("J:J").Value)
    ws.Cells(3, 17).Value = greatest_percent_dec_value
    ws.Cells(3, 16).Value = greatest_percent_inc_ticker
    
    greatest_total_value = WorksheetFunction.Max(ws.Range("L:L").Value)
    ws.Cells(4, 17).Value = greatest_total_value
    ws.Cells(4, 16).Value = greatest_total_ticker
    
    
    
Next ws
    
End Sub

