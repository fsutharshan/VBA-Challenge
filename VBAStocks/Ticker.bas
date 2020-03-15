Attribute VB_Name = "Module1"
Sub ticker_symbol()
On Error Resume Next

Dim ws As Worksheet
Dim ticker As String
Dim vol As Double
Dim year_open As Double
Dim year_close As Double
Dim percent_change As Double
Dim yearly_change As Double
Dim Summary_Table_Row As Integer
Dim next_ticker_begin_row As Integer
Dim LastRow As Long



'set headers
Set ws = Application.ActiveSheet
ws.Cells(1, 9).Value = " Ticker"
ws.Cells(1, 10).Value = " Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
 'set up integers for loop
Summary_Table_Row = 2
next_ticker_begin_row = 2
vol = 0

LastRow = Cells(Rows.Count, 1).End(xlUp).Row
  'loop
For i = 2 To LastRow
     vol = vol + ws.Cells(i, 7).Value
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            
            year_open = ws.Cells(next_ticker_begin_row, 3).Value
            year_close = ws.Cells(i, 6).Value
            yearly_change = year_close - year_open
            percent_change = (year_close - year_open) / year_close

            'Insert values into summary

            ws.Cells(Summary_Table_Row, 9).Value = ticker
            ws.Cells(Summary_Table_Row, 10).Value = yearly_change
            
            
            ' After calcualting yearly change and updating cell we immediately change its background color based on the calculated value
            If yearly_change < 0 Then
                ws.Cells(Summary_Table_Row, 10).Interior.Color = RGB(255, 0, 0)
            Else
               ws.Cells(Summary_Table_Row, 10).Interior.Color = RGB(0, 255, 0)
            End If
            
            ws.Cells(Summary_Table_Row, 11).Value = percent_change
            ws.Cells(Summary_Table_Row, 12).Value = vol
            Summary_Table_Row = Summary_Table_Row + 1
            next_ticker_begin_row = i + 1
            vol = 0
     End If
Next i
ws.Columns("K").NumberFormat = "0.00%"

 
'For Challenge question

Dim greatest_percent_increase_ticker As String
Dim greatest_percent_increase_value As Double
Dim greatest_percent_decrease_ticker As String
Dim greatest_percent_decrease_value As Double
Dim greatest_total_volume_ticker As String
Dim greatest_total_volume_value As Double



LastRow = Cells(Rows.Count, 11).End(xlUp).Row

greatest_percent_increase_ticker = ws.Cells(2, 9).Value
greatest_percent_increase_value = ws.Cells(2, 11).Value
greatest_percent_decrease_ticker = ws.Cells(2, 9).Value
greatest_percent_decrease_value = ws.Cells(2, 11).Value
greatest_total_volume_ticker = ws.Cells(2, 9).Value
greatest_total_volume_value = ws.Cells(2, 12)


For i = 2 To LastRow
    If greatest_percent_increase_value < ws.Cells(i, 11) Then
         greatest_percent_increase_value = ws.Cells(i, 11).Value
         greatest_percent_increase_ticker = ws.Cells(i, 9).Value
    End If
    If greatest_percent_decrease_value > ws.Cells(i, 11) Then
         greatest_percent_decrease_value = ws.Cells(i, 11).Value
         greatest_percent_decrese_ticker = ws.Cells(i, 9).Value
    End If
     
    If greatest_total_volume_value < ws.Cells(i, 12).Value Then
       greatest_total_volume_value = ws.Cells(i, 12).Value
       greatest_total_volume_ticker = ws.Cells(i, 9).Value
    End If
Next i
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % increase"
ws.Cells(3, 15).Value = "Greatest % decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(2, 16).Value = greatest_percent_increase_ticker
ws.Cells(2, 17).Value = greatest_percent_increase_value
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).NumberFormat = "0.00%"
ws.Cells(3, 16).Value = greatest_percent_decrease_ticker
ws.Cells(3, 17).Value = greatest_percent_decrease_value
ws.Cells(4, 16).Value = greatest_total_volume_ticker
ws.Cells(4, 17).Value = greatest_total_volume_value
End Sub





