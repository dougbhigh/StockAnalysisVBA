Attribute VB_Name = "Module2"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  RUT-SOM-DATA-PT-06-2020-U-C  VBA Project                         Douglas High '
'                                                                  June 20, 2020 '
'  macro 2 of 4 : VBAHWv2                                                        '
'   - reads in spreadsheet of stock records sorted by date within ticker symbol. '
'   - produces summary of each stock for the year in columns I:L.                '
' v2- produces second table under summary showing stocks with greates increase   '
'     and decrease(%) and largest volume.                                        '
'    it assumes the first record for each stock has an opening price and the     '
'    last record has a closing price.                                            '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub VBAHWv2()
'''''''''''''''''''''''''''''
'      Data Definitions     '
'''''''''''''''''''''''''''''
Dim r As Long                               'row count for input
Dim ticker_symbol As String
Dim year_open As Single
Dim year_close As Single
Dim year_change As Single
Dim year_change_per As Single
Dim total_volume As LongLong
Dim greatest_volume As LongLong
Dim greatest_volume_ticker As String
Dim greatest_increase As Single
Dim greatest_increase_ticker As String
Dim greatest_decrease As Single
Dim greatest_decrease_ticker As String
Dim lastrow As Long
Dim output_row As Integer                  ' row count for output
Dim first_ticker As Boolean
'''''''''''''''''''''''''''''''''''''
'          Initializations          '
'''''''''''''''''''''''''''''''''''''
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
output_row = 2                            'set ouput under headers
first_ticker = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   set up headers for output table and format columns for numbers   '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Volume"
Range("I1:L1").Font.FontStyle = "Bold"
Columns("J").NumberFormat = "0.00"
Columns("k").NumberFormat = "0.00%"
Columns("L").NumberFormat = "#,##0"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   loop through symbols and accumulate total volume        '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
For r = 2 To lastrow
    If Cells(r, 1) = Cells(r + 1, 1) Then        ' not last occurence
        If Cells(r, 1) <> Cells(r - 1, 1) Then   ' first occurence
            year_open = Cells(r, 3)
            total_volume = Cells(r, 7)
        Else
            total_volume = total_volume + Cells(r, 7) 'not first or last occurence
        End If
    Else                                    ' last occurence
        total_volume = total_volume + Cells(r, 7)
        year_close = Cells(r, 6)
        Cells(output_row, 9) = Cells(r, 1)     'move ticker symbol to output table
        year_change = year_close - year_open
        year_change_per = year_close / year_open - 1
        Cells(output_row, 10) = year_change
        If Cells(output_row, 10) > 0 Then
            Cells(output_row, 10).Interior.Color = vbGreen
        ElseIf Cells(output_row, 10) < 0 Then
            Cells(output_row, 10).Interior.Color = vbRed
        End If
        Cells(output_row, 11) = year_change_per
        Cells(output_row, 12) = total_volume
        output_row = output_row + 1
                  ' set greatest variables at first ticker summary
        If first_ticker = True Then
            first_ticker = False
            greatest_volume = total_volume
            greatest_volume_ticker = Cells(r, 1)
            greatest_increase = year_change_per
            greatest_increase_ticker = Cells(r, 1)
            greatest_decrease = year_change_per
            greatest_decrease_ticker = Cells(r, 1)
        End If
                 ' check for greatest values
        If total_volume > greatest_volume Then
            greatest_volume = total_volume
            greatest_volume_ticker = Cells(r, 1)
        End If
        If year_change_per > greatest_increase Then
            greatest_increase = year_change_per
            greatest_increase_ticker = Cells(r, 1)
        ElseIf year_change_per < greatest_decrease Then
            greatest_decrease = year_change_per
            greatest_decrease_ticker = Cells(r, 1)
        End If
    End If
Next r

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'     setup greatest change table under summary table                   '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
output_row = output_row + 1           'adds blank line
Cells(output_row, 11).Value = "Ticker"
Cells(output_row, 12).Value = "Value"
Cells(output_row, 11).Font.FontStyle = "Bold"
Cells(output_row, 12).Font.FontStyle = "Bold"
output_row = output_row + 1
Cells(output_row, 9).Value = "Greatest % Increase"
Cells(output_row, 9).Font.FontStyle = "Bold"
Cells(output_row, 11).Value = greatest_increase_ticker
Cells(output_row, 12).Value = greatest_increase
Cells(output_row, 12).NumberFormat = "0.00%"
output_row = output_row + 1
Cells(output_row, 9).Font.FontStyle = "Bold"
Cells(output_row, 9).Value = "Greatest % Decrease"
Cells(output_row, 11).Value = greatest_decrease_ticker
Cells(output_row, 12).Value = greatest_decrease
Cells(output_row, 12).NumberFormat = "0.00%"
output_row = output_row + 1
Cells(output_row, 9).Font.FontStyle = "Bold"
Cells(output_row, 9).Value = "Greatest Volume"
Cells(output_row, 11).Value = greatest_volume_ticker
Cells(output_row, 12).Value = greatest_volume
Columns("I:L").AutoFit
End Sub
