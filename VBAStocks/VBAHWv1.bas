Attribute VB_Name = "Module1"
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  RUT-SOM-DATA-PT-06-2020-U-C  VBA Project                        Douglas High '
'                                                                 June 20, 2020 '
'  macro 1 of 4 : VBAHWv1                                                       '                                                                               '
'   - reads in spreadsheet of stock records sorted by date within ticker symbol '
'   - produces summary of each stock for the year in columns I:L                '
'  *- 5/2021 repo name change, added teting files, no code changes              '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub VBAHWv1()

Dim r As Long
Dim ticker_symbol As String
Dim year_open As Single
Dim year_close As Single
Dim total_volume As LongLong
Dim lastrow As Long
Dim output_row As Integer

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
output_row = 2                     'set ouput row to first occurence

' set up headers for output table and format columns for numbers
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Volume"
Columns("J").NumberFormat = "0.00"
Columns("k").NumberFormat = "0.00%"
Columns("L").NumberFormat = "#,##0"

' loop through symbols and accumulate total volume
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
        Cells(output_row, 9) = Cells(r, 1)
        Cells(output_row, 10) = year_close - year_open
        If Cells(output_row, 10) > 0 Then
            Cells(output_row, 10).Interior.Color = vbGreen
        ElseIf Cells(output_row, 10) < 0 Then
            Cells(output_row, 10).Interior.Color = vbRed
        End If
        Cells(output_row, 11) = year_close / year_open - 1
        Cells(output_row, 12) = total_volume
        output_row = output_row + 1
    End If
Next r
End Sub