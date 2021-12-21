Sub dirac()
Dim i, j, k As Double
Dim p1, p2 As Double

flag = True

'scores p1 p2
p1 = 0
p2 = 0
'initial positions
'pos1 = 4    'test input
'pos2 = 8
pos1 = 7   'actual input
pos2 = 10

roun = 0 'round number but round is a function
roll = 1 'counter for dice, mod 100
posi1 = pos1
posi2 = pos2
Do While flag
    '~~~~~~~~ p1
    For i = roll To roll + 2
        move1 = move1 + (i Mod 100)
    Next i
    
    If ((posi1 + move1) Mod 10) <> 0 Then
        p1 = p1 + ((posi1 + move1) Mod 10)
    Else: p1 = p1 + 10
    End If
    
    If ((posi1 + move1) Mod 10) <> 0 Then
        posi1 = ((posi1 + move1) Mod 10)
    Else: posi1 = 10
    End If
    roll = roll + 3
    If p1 >= 1000 Then
        Exit Do
    End If
    '~~~~~~~~ p2
    For i = roll To roll + 2
        move2 = move2 + (i Mod 100)
    Next i
    
    If ((posi2 + move2) Mod 10) <> 0 Then
        p2 = p2 + ((posi2 + move2) Mod 10)
    Else: p2 = p2 + 10
    End If
    
    If ((posi2 + move2) Mod 10) <> 0 Then
        posi2 = ((posi2 + move2) Mod 10)
    Else: posi2 = 10
    End If
    
    If p2 >= 1000 Then
        Exit Do
    End If
    roll = roll + 3
    
    roun = roun + 1

'    Cells(roun, 1) = roun
'    Cells(roun, 2) = roll - 1
'    Cells(roun, 3) = p1
'    Cells(roun, 4) = p2
    move1 = 0
    move2 = 0
Loop

MsgBox ("player 1:" & p1 & vbNewLine & "player 2:" & p2 & vbNewLine & "Rolls:" & roll - 1)
Cells(1, 9).Value = Application.WorksheetFunction.Min(p1, p2) * (roll - 1)


End Sub
Sub dirac2()
Set gamestates = CreateObject("System.Collections.Arraylist")

t = Timer()
'Application.ScreenUpdating = False

p1 = 7
p2 = 10

x = diracmode(p1 - 1, p2 - 1, 0, 0)
Cells(2, 15).Value = Timer() - t
'Application.ScreenUpdating = True

MsgBox ("part 2 answer=" & x(0) & " " & x(1))

End Sub
Function diracmode(p1, p2, s1, s2) As Variant
    Dim output(2) As Variant
    Dim gstates(2) As Variant
    
    If s1 >= 21 Then
        output(0) = 1
        output(1) = 0
        diracmode = output
        Exit Function
    End If
    If s2 >= 21 Then
        output(0) = 0
        output(1) = 1
        diracmode = output
        Exit Function
    End If
    rowwrite = Cells(1, 14).Value
    test = CStr(p1) + "," + CStr(p2) + "," + CStr(s1) + "," + CStr(s2)
    'sorter unecessary
    intuple = Application.WorksheetFunction.IsError(Application.Match(test, Range("K:K"), 0))
    'MsgBox intuple
    If intuple = False Then
        output(0) = Cells(Application.Match(test, Range("K:K"), 0), 12)
        output(1) = Cells(Application.Match(test, Range("K:K"), 0), 13)
        diracmode = output
        Exit Function
    End If
    output(0) = 0
    output(1) = 0
    ans = output
    For i = 1 To 3
        For j = 1 To 3
            For k = 1 To 3
                np1 = (p1 + i + j + k) Mod 10
                ns1 = s1 + np1 + 1
                newans = diracmode(p2, np1, s2, ns1)
                ans(0) = ans(0) + newans(1)
                ans(1) = ans(1) + newans(0)
                
            Next k
        Next j
    Next i
    
    
    Cells(rowwrite, 11).Value = test
    Cells(rowwrite, 12).Value = ans(0)
    Cells(rowwrite, 13).Value = ans(1)
    diracmode = ans
    

End Function





