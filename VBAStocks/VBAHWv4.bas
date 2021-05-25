Attribute VB_Name = "Module3"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  RUT-SOM-DATA-PT-06-2020-U-C  VBA Project                         Douglas High '
'                                                                  June 20, 2020 '
'  macro 4 of 4 : VBAHWv4                                                        '
'   - reads in spreadsheet of stock records sorted by date within ticker symbol. '
'   - produces summary of each stock for the year in columns I:L.                '
' v2- produces second table under summary showing stocks with greates increase   '
'     and decrease(%) and largest volume.                                        '
' v3- added additional loop to cycle through all worksheets within workbook.     '
'   - added code to check for an opening price of zero.                          '
' v4- changed opening price of zero logic to loop through all records of stock   '
'     in search of a record with an opening balance.                             '
'   - added code to check for closing price of zero when opening price is        '
'     nonzero, also added error processing for above two conditions.             '
'  *- 5/2021 repo name change, added teting files, no code changes               '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub VBAHWv4()
'''''''''''''''''''''''''''''
'      Data Definitions     '
'''''''''''''''''''''''''''''
Dim r As Long                               ' row count for input
Dim s As Byte                               ' sheetcount
Dim r2 As Integer
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
Dim sheetcount As Byte
Dim lastrow As Long
Dim output_row As Integer                  ' row count for output
Dim first_ticker As Boolean
Dim error_msg As String
Dim error_row_ct As Integer
Dim error_sw As Boolean

'''''''''''''''''''''''''''''''''''''
'    Initializations for workbook   '
'''''''''''''''''''''''''''''''''''''
sheetcount = Worksheets.Count

'''''''''''''''''''''''''''''''''''''''''
'  loop through worksheets in workbook  '
'''''''''''''''''''''''''''''''''''''''''
    For s = 1 To sheetcount        ' yes, i'm still indenting the main loop instead of everything else :)
    
'''''''''''''''''''''''''''''''''''''
'    Initializations for worksheet  '
'''''''''''''''''''''''''''''''''''''
Worksheets(s).Activate
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
output_row = 2                            'set ouput under headers
first_ticker = True
error_row_ct = 2
error_sw = False
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
    If Cells(r, 1) = Cells(r + 1, 1) Then             ' not last occurence
        If Cells(r, 1) <> Cells(r - 1, 1) Then        ' first occurence
            year_open = Cells(r, 3)
            total_volume = Cells(r, 7)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '    if the first record has an opening price of zero, loop through    '
            '  all occurences (dates) of that stock to find if there is a date     '
            '  with a non-zero open price.                                         '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If Cells(r, 3) = 0 Then                      '  opening price = 0 on first record
                For r2 = 1 To 366                        ' in case its a leap year and stock market is open 7 days a week that year
                    If r2 <> 1 Then                      '  we already moved cell r,7 here if it was the first record
                        total_volume = total_volume + Cells(r, 7)  ' probably alot of volume whith a price of $0.00, it's free, take some.
                    End If
                    If Cells(r, 1) <> Cells(r + 1, 1) Then   '   at last record, still looking for an opening price
                        If Cells(r, 3) = 0 Then              ' all records have zero open price
                            If Cells(r, 6) = 0 Then          '    and last record close price = 0
                                year_change_per = 0
                            Else                             ' no non-zero opening price and close price is non-zero, percent change is undefined
                                error_msg = "stock " & Cells(r, 1) & " has no opening price on any record and has a closing price at year end of $" & Cells(r, 6)
                                Cells(error_row_ct, 14) = error_msg
                                error_row_ct = error_row_ct + 1
                                error_sw = True                  ' show error in the yearly percent change cell when we process last record below
                            End If
                        Else                                 ' last record, finally found a non-zero open price
                            year_open = Cells(r, 3)
                        End If
                        r = r - 1                            ' set row back 1 to process last record below
                        total_volume = total_volume - Cells(r, 7) ' will add back on at normal last record section
                        r2 = 366                             ' at last record, get out of internal loop
                    Else                                     ' not at last record
                        If Cells(r, 3) = 0 Then              ' open price still zero, bump up row count to check next record
                            r = r + 1
                        Else
                            year_open = Cells(r, 3)          ' found non-zero open price, set year_open
                            r2 = 366                         '  and get out of internal loop
                        End If
                    End If
                Next r2
            End If
        Else
            total_volume = total_volume + Cells(r, 7) 'not first or last occurence
        End If
    Else                                              ' last occurence - when a non-zero open price was found
        total_volume = total_volume + Cells(r, 7)     '   or if got to last record searching for non-zero
        year_close = Cells(r, 6)
        Cells(output_row, 9) = Cells(r, 1)                    ''''''''''''''''''''''''''''''''''''
        year_change = year_close - year_open                  '  fill output table section       '
        Cells(output_row, 10) = year_change
        If Cells(output_row, 10) > 0 Then                     '''''''''''''''''''''''''''''''''''''''
            Cells(output_row, 10).Interior.Color = vbGreen    '  format yearly change cell color    '
        ElseIf Cells(output_row, 10) < 0 Then                 '    - green if positive change       '
            Cells(output_row, 10).Interior.Color = vbRed      '    - red if negative                '
        End If                                                '''''''''''''''''''''''''''''''''''''''
        If year_open <> 0 Then
            year_change_per = year_close / year_open - 1
        End If
        If year_open <> 0 And year_close = 0 Then
            error_msg = "stock " & Cells(r, 1) & " has a non zero opening price and closing price of 0"
            Cells(error_row_ct, 14) = error_msg        '
            error_row_ct = error_row_ct + 1            ' errmsg for close price of 0 after we found open price ne 0
            Cells(output_row, 11) = "0-ERR"            '
            Cells(output_row, 11).Font.Color = vbRed
        ElseIf error_sw = True Then
            Cells(output_row, 11) = "ERROR"
            Cells(output_row, 11).Font.Color = vbRed
            error_sw = False
        Else
            Cells(output_row, 11) = year_change_per
        End If
        Cells(output_row, 12) = total_volume
        output_row = output_row + 1
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     set greatest variables at first ticker summary     '
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If first_ticker = True Then
            first_ticker = False
            greatest_volume = total_volume
            greatest_volume_ticker = Cells(r, 1)
            greatest_increase = year_change_per
            greatest_increase_ticker = Cells(r, 1)
            greatest_decrease = year_change
            greatest_decrease_ticker = Cells(r, 1)
        End If
        ''''''''''''''''''''''''''''''''''''''''''''
        '  check current values against greatest   '
        ''''''''''''''''''''''''''''''''''''''''''''
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
If error_row_ct > 2 Then                   ''''''''''''''''''''''''''''
    Columns("M").ColumnWidth = 3           '  if errors were written  '
    Columns("N").Font.Color = vbRed        '    format columnm        '
    Columns("N").AutoFit                   '    and write header      '
    Cells(1, 14) = "ERRORS"                ''''''''''''''''''''''''''''
End If
    Next s
End Sub