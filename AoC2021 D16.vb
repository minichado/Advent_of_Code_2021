Sub BitExpansion()
Dim test As String
Dim length As Integer
length = Len(Cells(1, 1).Value)
t = timer()
'convert intput to binary
'L1 = literallength("1000001000000000101111000110000010001101000000")

test = ""
For i = 1 To length
    test = test & Cells(3 + i, 2).Value
Next i

Cells(2, 1).Value = test

x = VersionNumbers(test)

Cells(3, 5).Value = x
elapsed = timer() - t
Cells(3, 6).Value = elapsed


'Cells(2, 1).Value = test
y = OperatorID(test)

'd16p2, new timer, reset value of test before running!

Cells(3, 9).Value = y
Cells(3, 10).Value = timer() - t - elapsed
'MsgBox ("The sum of all version numbers is " & X)



End Sub
Function OperatorID(test As String) As Double
   'currently works on all sample inputs, fails on my input
    VersionNum = Bin2Dec(Left(test, 3))
    typeid = Bin2Dec(Mid(test, 4, 3))
    bit = ""
    SumVersion = 0
    If typeid = 4 Then
        OperatorID = decodeliteral(CStr(test))
    'other typeIDs handled here
        
    End If
    If typeid <> 4 Then
        test2 = Right(test, Len(test) - 6)
        '0 id, next 15 is length of bits in sub packet
        If Left(test2, 1) = 0 Then
            test2 = Right(test2, Len(test2) - 1)
            bitlength = Bin2Dec(Left(test2, 15)) 'outputs 27 correctly
            bitlength1 = bitlength
            test2 = Right(test2, Len(test2) - 15)
            'need to use bitlength to tell when to stop counting literals with version numbers
            While bitlength > 0
                'add function here to do part 2 operation on literals
                If typeid = 0 Then
                    OperatorID = OperatorID + OperatorID(CStr(test2))
                ElseIf typeid = 1 Then
                   'need catch case where only one bitsin
                    logi1 = OperatorID(CStr(test2))
                    
                    If bitlength1 = literallength(CStr(test2)) Then
                        OperatorID = logi1
                    ElseIf bitlength1 = bitlength Then
                        OperatorID = logi1
                    Else: OperatorID = OperatorID * logi1
                    End If
                ElseIf typeid = 2 Then
                    min1 = OperatorID(CStr(test2))
                    If bitlength1 <> bitlength Then
                        OperatorID = min(minlast, min1)
                        minlast = OperatorID
                    End If
                    If bitlength1 = bitlength Then
                        minlast = min1
                        OperatorID = minlast
                    End If
                ElseIf typeid = 3 Then
                    max1 = OperatorID(CStr(test2))
                    If bitlength1 <> bitlength Then
                        OperatorID = max(maxlast, max1)
                        maxlast = OperatorID
                    End If
                    If bitlength1 = bitlength Then
                        maxlast = max1
                        OperatorID = maxlast
                    End If
                ElseIf typeid = 5 Then
                    If bitlength1 = bitlength Then
                        logi1 = OperatorID(CStr(test2))
                    Else
                        logi2 = OperatorID(CStr(test2))
                        If logi1 > logi2 Then
                            OperatorID = 1
                        Else: OperatorID = 0
                        End If
                    End If
                ElseIf typeid = 6 Then
                    If bitlength1 = bitlength Then
                        logi1 = OperatorID(CStr(test2))
                    Else
                        logi2 = OperatorID(CStr(test2))
                        If logi1 < logi2 Then
                            OperatorID = 1
                        Else: OperatorID = 0
                        End If
                    End If
                ElseIf typeid = 7 Then
                    If bitlength1 = bitlength Then
                        logi1 = OperatorID(CStr(test2))
                    Else
                        logi2 = OperatorID(CStr(test2))
                        If logi1 = logi2 Then
                            OperatorID = 1
                        Else: OperatorID = 0
                        End If
                    End If
                End If
                'end function here for part 2
                litlen = literallength(CStr(test2))
                test2 = Right(test2, Len(test2) - litlen)
                bitlength = bitlength - litlen
            Wend
         
'            While Z <= bitlength
'
'            Wend
        '1, next 11 is number of subpackets
        ElseIf Left(test2, 1) = 1 Then
            test2 = Right(test2, Len(test2) - 1)
            bitsin = Bin2Dec(CStr(Left(test2, 11)))
            bitsin1 = bitsin
            test2 = Right(test2, Len(test2) - 11)
            While bitsin > 0
                SumVersion = SumVersion + VersionNumbers(CStr(test2))
                'add function here to do part 2 operation on literals
                If typeid = 0 Then
                    OperatorID = OperatorID + OperatorID(CStr(test2))
                ElseIf typeid = 1 Then
                    'need catch case where only one bitsin
                    logi1 = OperatorID(CStr(test2))
                    
                    If bitsin1 = 1 Then
                        OperatorID = logi1
                    ElseIf bitsin1 = bitsin Then
                        OperatorID = logi1
                    Else: OperatorID = OperatorID * logi1
                    End If
                ElseIf typeid = 2 Then
                    min1 = OperatorID(CStr(test2))
                    If bitsin1 <> bitsin Then
                        OperatorID = min(minlast, min1)
                        minlast = OperatorID
                    End If
                    If bitsin1 = bitsin Then
                        minlast = min1
                        OperatorID = minlast
                    End If
                ElseIf typeid = 3 Then
                    max1 = OperatorID(CStr(test2))
                    If bitsin1 <> bitsin Then
                        OperatorID = max(maxlast, max1)
                        maxlast = OperatorID
                    End If
                    If bitsin1 = bitsin Then
                        maxlast = max1
                        OperatorID = maxlast
                    End If
                ElseIf typeid = 5 Then
                    If bitsin = 2 Then
                        logi1 = OperatorID(CStr(test2))
                    ElseIf bitsin = 1 Then
                        logi2 = OperatorID(CStr(test2))
                        If logi1 > logi2 Then
                            OperatorID = 1
                        Else: OperatorID = 0
                        End If
                    End If
                ElseIf typeid = 6 Then
                    If bitsin = 2 Then
                        logi1 = OperatorID(CStr(test2))
                    ElseIf bitsin = 1 Then
                        logi2 = OperatorID(CStr(test2))
                        If logi1 < logi2 Then
                            OperatorID = 1
                        Else: OperatorID = 0
                        End If
                    End If
                ElseIf typeid = 7 Then
                    If bitsin = 2 Then
                        logi1 = OperatorID(CStr(test2))
                    ElseIf bitsin = 1 Then
                        logi2 = OperatorID(CStr(test2))
                        If logi1 = logi2 Then
                            OperatorID = 1
                        Else: OperatorID = 0
                        End If
                    End If
                End If
                'end function here for part 2
                Debug.Print SumVersion
                litlen = literallength(CStr(test2))
                test2 = Right(test2, Len(test2) - litlen)
                bitsin = bitsin - 1
            Wend

        End If
        
    End If
    'return the output of the operantor step here

End Function

Function VersionNumbers(test As String) As Double
   
    VersionNum = Bin2Dec(Left(test, 3))
    typeid = Bin2Dec(Mid(test, 4, 3))
    bit = ""
    SumVersion = 0
    
    If typeid = 4 Then
        SumVersion = SumVersion + VersionNum
        Debug.Print SumVersion
        'MsgBox ("version number: " & VersionNum & "  Sumversion: " & SumVersion)
    'other typeIDs handled here
        
    End If
    If typeid <> 4 Then
        SumVersion = SumVersion + VersionNum
        Debug.Print SumVersion
        test2 = Right(test, Len(test) - 6)
        '0 id, next 15 is length of bits in sub packet
        If Left(test2, 1) = 0 Then
            test2 = Right(test2, Len(test2) - 1)
            bitlength = Bin2Dec(Left(test2, 15)) 'outputs 27 correctly
            test2 = Right(test2, Len(test2) - 15)
            'need to use bitlength to tell when to stop counting literals with version numbers
            While bitlength > 0
                SumVersion = SumVersion + VersionNumbers(CStr(test2))
                Debug.Print SumVersion
                litlen = literallength(CStr(test2))
                test2 = Right(test2, Len(test2) - litlen)
                bitlength = bitlength - litlen
            Wend
         
'            While Z <= bitlength
'
'            Wend
        '1, next 11 is number of subpackets
        ElseIf Left(test2, 1) = 1 Then
            test2 = Right(test2, Len(test2) - 1)
            bitsin = Bin2Dec(CStr(Left(test2, 11)))
            test2 = Right(test2, Len(test2) - 11)
            While bitsin > 0
                SumVersion = SumVersion + VersionNumbers(CStr(test2))
                Debug.Print SumVersion
                litlen = literallength(CStr(test2))
                test2 = Right(test2, Len(test2) - litlen)
                bitsin = bitsin - 1
            Wend

        End If
        
    End If
    VersionNumbers = SumVersion

End Function
Function literallength(test As String) As Double
    VersionNum = Left(test, 3)
    typeid = Bin2Dec(Mid(test, 4, 3))
    go = 1
    bit = ""
    length = 6
    'literal numbers
    If typeid = 4 Then
        test2 = Right(test, Len(test) - 6)
        While go = 1
            'read 5 bits at a time, first bit is go, next 4 are append
            bits = Left(test2, 5)
            go = Left(bits, 1)
            bit = bit + Mid(bits, 2, 4)
            test2 = Right(test2, Len(test2) - 5)
            length = length + 5
        Wend
        literallength = length
    End If
    
    'if it's package, need to return length of entire package
    'this should fix one of the example problems, test on all preceeding
    If typeid <> 4 Then
        test2 = Right(test, Len(test) - 6)
        If Left(test2, 1) = 0 Then
            test2 = Right(test2, Len(test2) - 1)
            bitlength = Bin2Dec(Left(test2, 15))
            literallength = 6 + 15 + bitlength + 1
            'if type 0, remove 5, remove 15 read, and remove all contained bits of length
            'bitlength
            
        End If
        
        If Left(test2, 1) = 1 Then
            test2 = Right(test2, Len(test2) - 1)
            bitsin = Bin2Dec(CStr(Left(test2, 11)))
            test2 = Right(test2, Len(test2) - 11)
            While bitsin > 0
                litlen = literallength(CStr(test2))
                literallength = literallength + litlen
                test2 = Right(test2, Len(test2) - litlen)
                bitsin = bitsin - 1
                
            Wend
            literallength = literallength + 6 + 11 + 1
        End If
        
        
    End If

    
    

End Function
Function decodeliteral(test As String) As Double
    VersionNum = Bin2Dec(Left(test, 3))
    typeid = Bin2Dec(Mid(test, 4, 3))
    go = 1
    bit = ""
    'literal numbers
    If typeid = 4 Then
        test2 = Right(test, Len(test) - 6)
        While go = 1
            'read 5 bits at a time, first bit is go, next 4 are append
            bits = Left(test2, 5)
            go = Left(bits, 1)
            bit = bit + Mid(bits, 2, 4)
            test2 = Right(test2, Len(test2) - 5)
        Wend
        decodeliteral = Bin2Dec(CStr(bit))
    End If
    

End Function

'Decimal to Binary
' =================
Function Dec2Bin(ByVal DecimalIn As Variant, _
              Optional NumberOfBits As Variant) As String
    Dec2Bin = ""
    DecimalIn = Int(CDec(DecimalIn))
    Do While DecimalIn <> 0
        Dec2Bin = Format$(DecimalIn - 2 * Int(DecimalIn / 2)) & Dec2Bin
        DecimalIn = Int(DecimalIn / 2)
    Loop
    If Not IsMissing(NumberOfBits) Then
       If Len(Dec2Bin) > NumberOfBits Then
          Dec2Bin = "Error - Number exceeds specified bit size"
       Else
          Dec2Bin = Right$(String$(NumberOfBits, _
                    "0") & Dec2Bin, NumberOfBits)
       End If
    End If
End Function

'Binary To Decimal
' =================
Function Bin2Dec(BinaryString As String) As Variant
    Dim x As Integer
    For x = 0 To Len(BinaryString) - 1
        Bin2Dec = CDec(Bin2Dec) + Val(Mid(BinaryString, _
                  Len(BinaryString) - x, 1)) * 2 ^ x
    Next
End Function


Public Function max(x, y As Variant) As Variant
  max = IIf(x > y, x, y)
End Function

Public Function min(x, y As Variant) As Variant
   min = IIf(x < y, x, y)
End Function



