Attribute VB_Name = "Module1"
Sub tickers_print()
Dim v As Double
Dim o_p As Double
Dim c_p As Double
Dim t_n As String

'Set ws = ActiveSheet ' debug to disable loop , use it only on one tab

For Each ws In Worksheets   'loop through all worksheets
ws.Activate
'rows count
Dim rows_c As Long
rows_c = ws.Range("A1", ws.Range("A1").End(xlDown)).Rows.Count
'coulmn count
Dim columns_c As Long
columns_c = ws.Range(ws.Range("A1").End(xlToRight), "A1").Columns.Count
'_______________________________________formating and sorting
Set my_range = Range(Cells(2, 1), Cells(rows_c, columns_c))
my_range.Sort Key1:=Range("B2"), Order1:=xlAscending, Header:=xlNo
my_range.Sort Key1:=Range("A2"), Order1:=xlAscending, Header:=xlNo

'formatting output summary table for all tickers
    'conditional formating  green - positive , red - negative
    Set mc = Range(Cells(2, 10), Cells(rows_c, 10))
    mc.ColumnWidth = 15
     mc.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    mc.FormatConditions(mc.FormatConditions.Count).SetFirstPriority
    With mc.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    mc.FormatConditions(1).StopIfTrue = False
    
    mc.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    mc.FormatConditions(mc.FormatConditions.Count).SetFirstPriority
    With mc.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
    End With
' total volume format
  mc.FormatConditions(1).StopIfTrue = False
  mc.NumberFormat = "_($* #,##0.00_)"
  Columns("K:K").Style = "Percent"
  Columns("L:L").ColumnWidth = 15
  Columns("L:L").NumberFormat = "_($* #,##0_)"

'headers
  Rows("1:1").Font.Bold = True
  Cells(1, 9) = "Ticker"
  Cells(1, 10) = "Yearly Change"
  Cells(1, 11) = "% Change"
  Cells(1, 12) = "Total Stock Volume"
    
'formatting second output summary table for best /worst ticekrs
Cells(2, 15) = "Greastest % Increase"
Cells(3, 15) = "Greatest % Decrease"
Cells(4, 15) = "Greatest Total Volume"
Range("o2:o4").ColumnWidth = 20
Range("o2:o4").Font.Bold = True
Cells(1, 16) = "Ticker"
Cells(1, 17) = "Value"
Range("P1:Q1").ColumnWidth = 15
'range("P1:Q1").Font.Bold = True
Range("Q2:Q3").Style = "Percent"
Range("q4").NumberFormat = "_($* #,##0_)"
'_______________________________________end of formating and sorting



'__________________________________________finding yearly change , % change and total volume for each unique ticker____________
' set original variables and counter
t_n = Cells(2, 1).Value                                     'ticker name
t_c = 2                                                     'tickers counter.... how many total tickers
x = 0                                                      'ticker-trading days counter ... how many records for one ticker
' find unique tickers & record into output table
For i = 2 To rows_c
   
    v = v + Cells(i, 7).Value                               'calculating total volume
    If t_n <> Cells(i + 1, 1).Value Then                    ' on condition if tickters name is changed

        Cells(t_c, columns_c + 2).Value = t_n               ' write data into summary table
        h = i - x                                           'finding earliest record  for
        o_p = Cells(h, 3).Value                             ' open price
        c_p = Cells(i, 6).Value                             ' closing price
            If o_p <> 0 Then
            per_ch = (c_p - o_p) / o_p     ' percent change calc with div by zero error check
            Else
            per_ch = "ERROR"
            End If
        Cells(t_c, columns_c + 3).Value = c_p - o_p
        Cells(t_c, columns_c + 4).Value = per_ch
        
        Cells(t_c, columns_c + 5).Value = v
        v = 0                                                'zeroing counters - getting ready for next ticker
        o_p = 0
        c_p = 0
        x = -1
        t_c = t_c + 1
        t_n = Cells(i + 1, 1).Value
                                   
                                        
    End If
 x = x + 1
                                    
Next i
'_______________________________________finding greatest increase,decrease and volume ticker_______
 max_ch = Cells(2, 11)                                         ' maximum change value , set first to frist row of data
 max_ch_n = Cells(2, 9)                                         'maximum change name
 min_ch = Cells(2, 11)
 min_ch_n = Cells(2, 9)
 max_v = Cells(2, 12)
 max_v_n = Cells(2, 9)
 
 For i = 3 To t_c + 1
    If Cells(i, 11) > max_ch And Cells(i, 11) <> "ERROR" Then
    max_ch = Cells(i, 11)
    max_ch_n = Cells(i, 9)
    End If
    If Cells(i, 11) < min_ch And Cells(i, 11) <> "ERROR" Then
    min_ch = Cells(i, 11)
    min_ch_n = Cells(i, 9)
    End If
    If Cells(i, 12) > max_v And Cells(i, 11) <> "ERROR" Then
    max_v = Cells(i, 12)
    max_v_n = Cells(i, 9)
    End If
 Next i

'______________________________________plotting results in "greatest for year table"_______
Cells(2, 16) = max_ch_n
Cells(2, 17) = max_ch
Cells(3, 16) = min_ch_n
Cells(3, 17) = min_ch
Cells(4, 16) = max_v_n
Cells(4, 17) = max_v
Next ws
End Sub



