Attribute VB_Name = "Simple"
'Declare Global Variables here:
Public volume As Range
Public condition As Range
Public dict As Object

Sub sumif()
    Dim vol As Range
    Dim cond As Range
    Dim ticker As String
    Set cond = Range("A:A") 'Range("A1:A10")
    Set vol = Range("G:G") 'Range("G1:G10")
    ticker = "A"
    Dim output As Double
    output = Application.sumif(cond, ticker, vol)
    Dim f_out As Double
    'f_out = Format(output, "Currency")
    MsgBox (Format(output, "Currency"))
    
End Sub

'Get all the distinct ticker in a column
Sub get_ticker()
    
End Sub

'Iterate through a range  EXAMPLE
Sub iterate()
    Dim rng As Range
    Set rng = Range("A1:A3")
    For Each cell In rng
        MsgBox (CStr(cell))
    Next cell
    
End Sub

'Example dictionary
Sub d()
    Set dict = CreateObject("Scripting.Dictionary")
    Dim rng As Range
    Set rng = Range("A1:A3")
    For Each cell In rng
        'MsgBox (cell)
        cell = CStr(cell)
        If dict.Exists(cell) Then
        'do nothing
            'MsgBox ("I did nothing")
        Else 'add to dict
            dict.Add cell, cell
        End If
        
        
    Next cell
    
    Dim key As Variant
    For Each key In dict.keys
        MsgBox (key)
    Next key
End Sub

'Example LIKE
Sub contains()
    Dim rng As Range
    Set rng = Range("A1:A3")
    For Each cell In rng
        'MsgBox (CStr(cell))
        If cell Like "*<*" Then
            MsgBox ("I did nothing")
        Else
            MsgBox (CStr(cell))
        End If
        
    Next cell
End Sub


'initialize the global variables
Sub init()
    Set condition = Range("A:A") 'Range("A1:A10") 'Range("A:A")
    Set volume = Range("G:G") 'Range("G1:G10") 'Range("G:G")
    wl_get_ticker_dict
End Sub


Sub wl_get_ticker_dict():
    Set dict = CreateObject("Scripting.Dictionary")
    Dim rng As Range
    'rng = condition 'Set rng = Range("A1:A3")
    'For Each cell In rng
    For Each cell In condition
        'MsgBox (cell)
        cell = CStr(cell)
        If dict.Exists(cell) Then
        'do nothing
            'MsgBox ("I did nothing")
        ElseIf cell Like "*<*" Then
            'do noting
        Else 'add to dict
            dict.Add cell, cell
        End If
        
        
    Next cell
    
    Dim key As Variant
    For Each key In dict.keys
        'MsgBox ("get_ticker:" + key) 'debug
    Next key
    'Set wl_get_ticker_tick = dict
End Sub

'print the dict
Sub wl_print()
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Total Stock Volume"
    
    Dim key As Variant
    Dim i As Integer
    i = 1
    
    'loop through the dictionary and print the key/value pair = {ticker:volume} in the worksheet
    For Each key In dict.keys
        i = i + 1
        Cells(i, 9).Value = CStr(key) 'ticker
        Cells(i, 10).Value = CStr(dict(key)) 'total volume
    Next key
    
End Sub

'clear the work columns
Sub wl_clear()
    Range("I:I").Value = ""
    Range("J:J").Value = ""
End Sub

'Start Main
Sub wl_main()
    'clear the workspace
    wl_clear
    'initilize global variables
    init
    'initilize local variables
    Dim total As Double
    Dim key As Variant
    
    'loop through the dictionary and assign the key/value pair = {ticker:volume}
    For Each key In dict.keys
        total = Application.sumif(condition, key, volume)
        dict(key) = total
        'MsgBox (CStr(key) + " " + CStr(dict(key))) 'debug
    Next key
        
    'print the results to the current worksheet
    wl_print
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

'Example - get worksheet count
Sub ws_count()
   MsgBox (ThisWorkbook.Sheets.Count)
   
End Sub

'Example - get worksheet name
Sub ws_name()
    For i = 1 To ThisWorkbook.Sheets.Count
        MsgBox (ThisWorkbook.Sheets(i).Name)
    Next i
End Sub

'Example - iterate worksheet
Sub ws_iterate()
    For i = 1 To ThisWorkbook.Sheets.Count
        'MsgBox (ThisWorkbook.Sheets(i).Name)
        'Dim this As Sheets
        'this = ThisWorkbook.Sheets(i)
        'ThisWorkbook.Sheets(i).Range("J1").Value = "Hi, I got updated" 'DEBUG
        ThisWorkbook.Sheets(i).Select
        wl_main
    Next i
End Sub





