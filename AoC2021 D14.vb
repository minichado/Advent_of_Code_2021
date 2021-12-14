Sub Polymer2()
Dim i, j, k, length As Single
Dim poly, test, test2, pair, rule As String
Dim Alpha As String
Dim t As Single
t = Timer
Application.ScreenUpdating = False


'Dim min, max As Single
'clear output
Range("F3:G102").ClearContents
cycles = Cells(7, 9).Value
r1 = 2
c1 = 1
rulesrow = 102

i = 1
j = 1
test2 = Cells(1, 1).Value

'initial cycle

For i = 1 To rulesrow
    pair = Cells(r1 + i, 2)
    
    'for first cycle only, use raw input, catches FFF
    For j = 1 To Len(test2) - 1
        test = Mid(test2, j, 2)
        If test = pair Then
            countpair = countpair + 1
        End If
    Next j
    
    
    'countpair = (Len(test2) - Len(Replace(test2, pair, ""))) / 2
    'Cells(r1 + i, 6).Value = Cells(r1 + i, 6) + countpair
    If countpair > 0 Then
        'for every pair found, adds to 2 new pairs
        'ex rule cb->h adds 1 to ch and hb
        rule = Cells(r1 + i, 4) 'UPDATE FOR FULL INPUT
        Row = Application.WorksheetFunction.Match(rule, Range(Cells(3, 2), Cells(rulesrow, 2)), 0)
        Cells(r1 + Row, 6).Value = Cells(r1 + Row, 6).Value + countpair
        rule = Cells(r1 + i, 5) 'UPDATE FOR FULL INPUT
        Row = Application.WorksheetFunction.Match(rule, Range(Cells(3, 2), Cells(rulesrow, 2)), 0)
        Cells(r1 + Row, 6).Value = Cells(r1 + Row, 6).Value + countpair
    End If
    countpair = 0
Next i
'subsequent cycles use input in column 6/F
'Range("F3:F18").Copy Range("G3")

For k = 2 To cycles
    For i = 1 To 100
        pair = Cells(r1 + i, 2)
        countpair = Cells(r1 + i, 6)
        'Cells(r1 + i, 6).Value = Cells(r1 + i, 6) + countpair
        If countpair > 0 Then
            'for every pair found, adds to 2 new pairs
            'ex rule cb->h adds 1 to ch and hb
            rule = Application.WorksheetFunction.VLookup(Cells(r1 + i, 4), Range(Cells(3, 2), Cells(rulesrow, 3)), 1, True) 'UPDATE FOR FULL INPUT
            Row = Application.WorksheetFunction.Match(rule, Range(Cells(3, 2), Cells(rulesrow, 2)), 0)
            Cells(r1 + Row, 7).Value = Cells(r1 + Row, 7).Value + countpair
            rule = Application.WorksheetFunction.VLookup(Cells(r1 + i, 5), Range(Cells(3, 2), Cells(rulesrow, 3)), 1, True) 'UPDATE FOR FULL INPUT
            Row = Application.WorksheetFunction.Match(rule, Range(Cells(3, 2), Cells(rulesrow, 2)), 0)
            Cells(r1 + Row, 7).Value = Cells(r1 + Row, 7).Value + countpair
        End If
    Next i
    Range(Cells(3, 7), Cells(rulesrow, 7)).Copy Range(Cells(3, 6), Cells(rulesrow, 6))
    'MsgBox (k & " " & Application.WorksheetFunction.Sum(Range(Cells(3, 7), Cells(rulesrow, 7))))
    Range(Cells(3, 7), Cells(rulesrow, 7)).ClearContents
Next k
'now total letters
Alpha = "BCFHKNOPSV"
max = 0
min = Application.WorksheetFunction.Sum(Range(Cells(3, 6), Cells(rulesrow, 6))) + 1
For i = 1 To 26
    Count = 0
    test = Mid(Alpha, i, 1)
    For j = 1 To rulesrow
        test2 = Left(Cells(r1 + j, 2), 1)
        If test = test2 Then
            Count = Count + Cells(r1 + j, 6)
        End If
    Next j

    If test = Right(Cells(1, 1), 1) Then
        Count = Count + 1
    End If
        'MsgBox (test & " " & Count)
    If Count > max Then
        max = Count
        Cells(4, 11).Value = test
        Cells(5, 11).Value = max
    End If
    If Count < min And Count <> 0 Then
        min = Count
        Cells(4, 12).Value = test
        Cells(5, 12).Value = min
    End If
Next i

answer = max - min


'MsgBox answer
Cells(4, 10).Value = answer
Cells(5, 10).Value = Timer - t
Application.ScreenUpdating = True

End Sub
Sub Polymer()
Dim i, j, k, length As Single
Dim poly, test, test2, pair, rule As String
Dim Alpha As String
Dim min, max As Double
Dim t As Single
t = Timer

rulesrow = 102
cycles = 10
i = 1
j = 1
test2 = Cells(1, 1).Value
For k = 1 To cycles
    poly = test2
    test2 = poly

    While i < Len(poly)
        test = test2
        pair = Mid(poly, i, 2)
        rule = Application.WorksheetFunction.VLookup(pair, Range(Cells(3, 2), Cells(rulesrow, 3)), 2, True)
        If IsError(rule) = False Then
            'insert rule here
            test = Left(test, j) & rule & Right(poly, Len(test) - j)
            test2 = test
            j = j + 1
        Else: test2 = test
        End If
        
        j = j + 1
        i = i + 1
        'MsgBox test2
    Wend
    'Cells(1, 2).Value = test2
i = 1
j = 1
Next k

'for answer, count occurance all letters
Alpha = "BCFHKNOPSV"
max = 0
min = Len(test2)
For i = 1 To 26
    Count = Len(test2) - Len(Replace(test2, Mid(Alpha, i, 1), ""))
    If Count > max Then
        max = Count
        Cells(2, 11).Value = Mid(Alpha, i, 1)
        Cells(3, 11).Value = max
    End If
    If Count < min And Count <> 0 Then
        min = Count
        Cells(2, 12).Value = Mid(Alpha, i, 1)
        Cells(3, 12).Value = min
    End If
Next i

answer = max - min
'MsgBox answer
Cells(2, 10).Value = answer
Cells(3, 10).Value = Timer - t
End Sub

