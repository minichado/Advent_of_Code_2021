Sub decode()
Dim i, j, k As Integer
Dim test As String
Dim one, two, three, four, five, six, seven, eight, nine, zero As String
Dim Sone, Stwo, Sthree, Sfour, Sfive, Ssix, Sseven, Seight, Snine, Szero As String
Dim Code(10) As String
Dim test1, test2 As String


Dim R1, C1 As Integer
R1 = 3
C1 = 23
c2 = 34
For i = R1 To R1 + 200
'first find known values
    For j = C1 To C1 + 10
        
        If Len(Cells(i, j).Value) = 2 Then
            one = Cells(i, j).Value
            'Sone = fncSortStr(Cells(i, j).Value)
            Code(1) = one
        ElseIf Len(Cells(i, j).Value) = 3 Then
            seven = Cells(i, j).Value
            'Sseven = fncSortStr(Cells(i, j).Value)
            Code(7) = seven
        ElseIf Len(Cells(i, j).Value) = 4 Then
            four = Cells(i, j).Value
            'Sfour = fncSortStr(Cells(i, j).Value)
            Code(4) = four
        ElseIf Len(Cells(i, j).Value) = 7 Then
            eight = Cells(i, j).Value
            'Seight = fncSortStr(Cells(i, j).Value)
            Code(8) = eight
        End If
        'now hve 1,4,7,8 and sorted versions
        'made new contain function, don't need to sort
    Next j
    'MsgBox (Sone & " " & Sfour & " " & Sseven & " " & Seight)
    'now evaluate 6 digit
    'MsgBox "made it here"
    
    'loop to find 6/9/0
    For j = C1 To C1 + 10
        If Len(Cells(i, j).Value) = 6 Then
            
            test = Cells(i, j).Value
            'MsgBox test
            'MsgBox Len(Contains(test, seven))
            If Len(Contains(test, four)) = 2 Then 'finds 9
                nine = test
                'Snine = fncSortStr(Cells(i, j).Value)
                Code(9) = nine
            ElseIf Len(Contains(test, seven)) = 4 Then 'finds 6
                six = test
                'Ssix = fncSortStr(Cells(i, j).Value)
                Code(6) = six
            Else: zero = test
            End If
            'MsgBox ("nine" & " " & nine & " " & "six" & " " & six & " " & "zero" & " " & zero)
        End If
    Next j
    Code(10) = zero
    'loop to find 2/3/5
    For j = C1 To C1 + 10
        If Len(Cells(i, j).Value) = 5 Then
            test = Cells(i, j).Value
            If Len(Contains(test, one)) = 3 Then 'finds 3, which contains 1
                three = test
                Code(3) = three
            ElseIf Len(Contains(test, four)) = 2 Then 'finds 5
                five = test
                Code(5) = five
            Else: two = test
            
            End If
            
       
        End If
        
    Next j
    Code(2) = two
    'MsgBox (one & " " & two & " " & three & " " & four & " " & five & " " & six & " " & seven & " " & eight & " " & nine & " " & zero)
    'now decode
    For j = 1 To UBound(Code)
        Cells(i, 37 + j).Value = fncSortStr(Code(j))
    Next j

    For k = 1 To 4
        Cells(i, 47 + k).Value = fncSortStr(Cells(i, 33 + k))
    Next k

Next i


MsgBox "We are done!"



End Sub
Function fncSortStr(strPassed As String) As String
Dim Temp As String
Dim i As Integer

While Len(strPassed)
    Temp = Left(strPassed, 1)
    
        For i = 2 To Len(strPassed)
            If Mid(strPassed, i, 1) < Temp Then
                Temp = Mid(strPassed, i, 1)
            End If
        Next i
    
    fncSortStr = fncSortStr & Temp
    strPassed = Left(strPassed, InStr(1, strPassed, Temp) - 1) & _
                        Mid(strPassed, InStr(1, strPassed, Temp) + 1)
Wend

End Function

Function Contains(Str, Contain) As String
Dim test As String
Dim i As Integer
test = Str
'"dgceab", "fcge"
    For i = 1 To Len(Contain)
        test = Replace(test, Mid(Contain, i, 1), "")
    Next i
    Contains = test



End Function