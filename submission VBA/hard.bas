Attribute VB_Name = "hard"
'Declare Global Variables here:
Public volume As Range
Public condition As Range
Public ticker_date As Range
Public dict As Object
Public min_dict As Object
Public max_dict As Object
Public open_dict As Object
Public close_dict As Object




'initialize the global variables
Sub init()
    Set condition = Range("A:A") 'Range("A1:A10") 'Range("A:A")
    Set volume = Range("G:G") 'Range("G1:G10") 'Range("G:G")
    Set ticker_date = Range("B:B")
    wl_get_ticker_dict
    wl_yearly_dict
End Sub

'WL - get unique set of tickers
Sub wl_get_ticker_dict():
    Set dict = CreateObject("Scripting.Dictionary")
    Dim rng As Range
    'rng = condition 'Set rng = Range("A1:A3")
    'For Each cell In rng
    For Each cell In condition
        'MsgBox (cell)
        cell = CStr(cell)
        If dict.exists(cell) Then
        'do nothing
            'MsgBox ("I did nothing")
        ElseIf cell Like "*<*" Then 'code to not add headings
            'do noting
        Else 'add to dict
            dict.Add cell, cell
        End If
        
        
    Next cell
    
    'Dim key As Variant
    'For Each key In dict.keys
        'MsgBox ("get_ticker:" + key) 'debug
    'Next key
    'Set wl_get_ticker_tick = dict
    
    'Bug - I added the empty key.  So I must remove it.
    If dict.exists("") Then
        
        dict.Remove ""
        'MsgBox ("empty key")
    End If
   
End Sub

'print the dict
Sub wl_print()
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    Dim key As Variant
    Dim i As Integer
    Dim t_open As Double
    Dim t_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
        
    i = 1
    
    'loop through the dictionary and print the key/value pair = {ticker:volume} in the worksheet
    For Each key In dict.keys
        If key = "" Then
         Exit For
        End If
        
        i = i + 1
        t_open = open_dict(key)
        t_close = close_dict(key)
        yearly_change = t_close - t_open
        'MsgBox ((102 - 5) / 4)
        
        'error handling division by 0
        If t_open <> 0 Then
            percent_change = (t_close - t_open) / t_open
        ElseIf t_open = 0 Then
            percent_change = 0
        End If
        
        'percent_change = (open_dict(key) - close_dict(key)) / open_dict(key)
        Cells(i, 9).Value = CStr(key) 'ticker
        Cells(i, 10).Value = yearly_change 'yearly change
        Cells(i, 11).Value = percent_change 'percent change
        Cells(i, 12).Value = CStr(dict(key)) 'total volume
        
        'coloring + formating
        Cells(i, 11).NumberFormat = "0.00%"
        If yearly_change >= 0 Then
            Cells(i, 10).Interior.ColorIndex = 4 'green
        Else
            Cells(i, 10).Interior.ColorIndex = 3 'red
        End If
        
    Next key
    
End Sub

'clear the work columns
Sub wl_clear()
    Range("I:I").Value = ""
    Range("J:J").Value = ""
    Range("K:K").Value = ""
    Range("L:L").Value = ""
    Range("O:O").Value = ""
    Range("P:P").Value = ""
    Range("Q:Q").Value = ""
End Sub

'Start Main
Sub wl_main()
    'clear the workspace
    wl_clear
    'initilize global variables
    init
    'populate the dictionary
    populate_dict
    'print the results to the current worksheet
    wl_print
    'populate the sheet with part 3 - hard data
    hard
End Sub

'Apply main to all worksheets
Sub wl_main_all()
    For i = 1 To ThisWorkbook.Sheets.Count
        ThisWorkbook.Sheets(i).Select
        wl_main
    Next i
End Sub

'clear all worksheets
Sub wl_clear_all()
        For i = 1 To ThisWorkbook.Sheets.Count
        ThisWorkbook.Sheets(i).Select
        wl_clear
    Next i
End Sub

'Yearly dict
Sub wl_yearly_dict()
Dim lastrow As Long
Dim key As Variant
Dim min_date As Long
Dim ticker As String
Dim max_date As Long
Dim t_open As Double
Dim t_close As Double


Set min_dict = CreateObject("scripting.dictionary")
Set open_dict = CreateObject("scripting.dictionary")
Set max_dict = CreateObject("scripting.dictionary")
Set close_dict = CreateObject("scripting.dictionary")
lastrow = Cells(Rows.Count, 1).End(xlUp).row
For i = 2 To lastrow
    ticker = Cells(i, 1).Value
    min_date = Cells(i, 2).Value
    max_date = Cells(i, 2).Value
    t_open = Cells(i, 3).Value
    t_close = Cells(i, 6).Value
    If Not min_dict.exists(ticker) Then
        min_dict(ticker) = min_date
        open_dict(ticker) = t_open
    Else
        If min_dict(ticker) > min_date Then 'set new min_date
            min_dict(ticker) = min_date
            open_dict(ticker) = t_open
        End If
    End If
    'max close
    If Not max_dict.exists(ticker) Then
        max_dict(ticker) = max_date
        close_dict(ticker) = t_close
    Else
        If max_dict(ticker) < max_date Then 'set new min_date
            max_dict(ticker) = max_date
            close_dict(ticker) = t_close
        End If
    End If
    
Next i
'MsgBox (min_dict("A"))
'MsgBox (close_dict("A"))
End Sub

'Part 3 - hard: return min/max
Sub hard()
    'max volume
    Dim max_volume As Double
    Dim volume_row As Integer
    Dim volume_ticker As String
    max_volume = Application.Max(Range("L:L"))
    volume_row = Application.Match(max_volume, Range("L:L"), 0)
    volume_ticker = Cells(volume_row, 9)
    'MsgBox (volume_ticker)
    'MsgBox (max_volume)
    Range("P4").Value = volume_ticker
    Range("Q4").Value = max_volume
    
    'max % increase
    Dim increase As Double
    Dim increase_row As Integer
    Dim increase_ticker As String
    increase = Application.Max(Range("K:K"))
    increase_row = Application.Match(increase, Range("K:K"), 0)
    increase_ticker = Cells(increase_row, 9)
    'MsgBox (increase_ticker)
    'MsgBox (increase)
    Range("p2").Value = increase_ticker
    Range("Q2").Value = increase
    
    'max % decrease
    Dim decrease As Double
    Dim decrease_row As Integer
    Dim decrease_ticker As String
    decrease = Application.Min(Range("K:K"))
    decrease_row = Application.Match(decrease, Range("K:K"), 0)
    decrease_ticker = Cells(decrease_row, 9)
    'MsgBox (decrease_ticker)
    'MsgBox (decrease)
    Range("p3").Value = decrease_ticker
    Range("Q3").Value = decrease
  
    'formating
    Range("Q2").NumberFormat = "0.00%"
    Range("Q3").NumberFormat = "0.00%"
End Sub

'Example - add dictionary details
Function details(yearly_change As Integer)
    Set details = CreateObject("scripting.dictionary")
    details.Add "yearly_change", yearly_change
End Function

'Example populate dict
Sub populate_dict()
    'initilize local variables
    Dim total As Double
    Dim key As Variant
    Dim yearly_change As Double
    Dim yearly_percentage As Double
    
    'loop through the dictionary and assign the key/value pair = {ticker:volume}
    For Each key In dict.keys
        total = Application.SumIf(condition, key, volume)
        dict(key) = total
        
        'MsgBox (dict.Item(key).Item("yearly_change"))
        'MsgBox (CStr(key) + " " + CStr(dict(key))) 'debug
    Next key
End Sub




'Example - get min date
Sub min_date()
    Dim min_date As Long
    'MsgBox (Application.Min(Range("B1:B3")))
    'MsgBox (Application.Max(Range("B1:B3")))
    min_date = Application.Min(Range("B1:B3"))
    'msgbox(type(Application.Min(Range("B1:B3"))))
    'MsgBox (min_date)
    'MsgBox (TypeName("wow"))
    'MsgBox (TypeName(Application.Min(Range("B1:B3"))))
    Dim row As Double
    'MsgBox (Application.Match(min_date, Range("B1:B3"), 0))
    'MsgBox (Application.Match(Application.Min(Range("B1:B3")), "B1:B3", 0))
    row = Application.Match(min_date, Range("B1:B3"), 0)
    Dim output As Double
    output = Cells(row, 3) 'get the opening using column 3
    MsgBox (output)
End Sub

'Example - get max date
Sub max_date()
    Dim max_date As Long
    Dim row As Double
    Dim output As Double
    max_date = Application.Max(Range("B1:B3"))
    row = Application.Match(max_date, Range("B1:B3"), 0)
    output = Cells(row, 6) 'get the closing using column 6
    MsgBox (output)
End Sub

'Example - dictionary + array key unnamed
Sub dictionary()
    Dim key As Variant
    Set d = CreateObject("scripting.dictionary")
    d.Add "a", Array("e1", "e2", "e3")
    'loop through the dictionary and assign the key/value pair = {ticker:volume}
    For Each key In d.keys
        Dim a() As Variant
        'MsgBox (Join(d(key))) 'debug
        a = d(key)
        
        Dim item As Variant
        For Each item In a
            MsgBox (item)
        Next item
    Next key
End Sub

'example - min_dict
Sub min_max_dict()
    Dim key As Variant
    Set m_d = CreateObject("scripting.dictionary")
    For Each key In dict
        
    Next key
    
End Sub

'example range condition
Sub range_condition()
    Dim r As Range
    Dim cell As Range
    Set r = Range("A" & Range("A1:A10"))
    'r = Range("A:A")
    'u = Range("A" & r)
    'Range("K:K").Value = Range("A" & Range("A:A"))
    'For Each cell In Range("A" & Range("A1:A15"))
    '    MsgBox (cell)
    'Next cell
    For Each cell In r
        MsgBox (CStr(cell.Value))
    Next cell
    
End Sub





