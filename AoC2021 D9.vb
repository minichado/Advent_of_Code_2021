Sub basins()
Dim i, j, k As Integer
Dim test, testhome As Single
R1 = 2
C1 = 2
R2 = 2
C2 = 203

'create 1002x1002 matrix with 9s, to pad the output first, and make the edge stuff below unecessary
'Range("nines").Value = 9
'initialize test matrix
'creates 100x100 with 1's at all valleys
For i = R1 + 1 To R1 + 100
    For j = C1 + 1 To C1 + 100
        If Cells(i, j + 101).Value > 0 Then
            Cells(i, j + 202).Value = 1
        Else: Cells(i, j + 202).Value = ""
        End If
    Next j
Next i

'run tests now on test matrix
For k = 1 To 20
    For i = R1 + 1 To R1 + 100
        For j = C2 + 2 To C2 + 101
            test = Cells(i, j).Value
            testhome = Cells(i, j - C2 + 1).Value
            'center
            
            If test > 0 Then
                If Cells(i, j + 1 - C2 + 1) > testhome And Cells(i, j + 1 - C2 + 1) <> 9 Then
                    Cells(i, j + 1).Value = Cells(i, j + 1 - C2 + 1).Value
                End If
                If Cells(i, j - 1 - C2 + 1) > testhome And Cells(i, j - 1 - C2 + 1) <> 9 Then
                    Cells(i, j - 1).Value = Cells(i, j - 1 - C2 + 1).Value
                End If
                If Cells(i + 1, j - C2 + 1) > testhome And Cells(i + 1, j - C2 + 1) <> 9 Then
                    Cells(i + 1, j).Value = Cells(i + 1, j - C2 + 1).Value
                End If
                If Cells(i - 1, j - C2 + 1) > testhome And Cells(i - 1, j - C2 + 1) <> 9 Then
                    Cells(i - 1, j).Value = Cells(i - 1, j - C2 + 1).Value
                End If
            End If
        Next j
    Next i
'MsgBox k
Next k
End Sub


