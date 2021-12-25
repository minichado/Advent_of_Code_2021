Sub cucumber()
Dim i, j, k, x, y, z As Double
Dim stepcount As Double
t = Timer() - t
stepcount = 0
Change = True
xrange = 137
yrange = 139
Offset = 150

'need to initialize output with periods
Cells(1, 1 + 150).Value = "test"
Do While Change
    Direction = 2
    Change = False
    Range("input").Copy Range("output")
    While Direction > 0
        For i = 1 To xrange
            For j = 1 To yrange
            'case for >
            If Cells(i, j).Value = ">" And Direction = 2 Then
                compare = IIf(j = yrange, Cells(i, 1).Value, Cells(i, j + 1).Value)
                If compare = "." And j <> yrange Then
                    'handle offset for edges, not done
                    Cells(i, j + 1 + Offset).Value = ">"
                    Cells(i, j + Offset).Value = "."
                    Change = True
                ElseIf compare = "." And j = yrange Then
                
                    'need to handle wrap case properly
                    Cells(i, 1 + Offset).Value = ">"
                    Cells(i, j + Offset).Value = "."
                    Change = True
                End If
            End If
            If Cells(i, j).Value = "v" And Direction = 1 Then
                compare = IIf(i = xrange, Cells(1, j).Value, Cells(i + 1, j).Value)
                If compare = "." And i <> xrange Then
                    'handle offset case for edges , not done
                    Cells(i + 1, j + Offset).Value = "v"
                    Cells(i, j + Offset).Value = "."
                    Change = True
                ElseIf compare = "." And i = xrange Then
                    Cells(1, j + Offset).Value = "v"
                    Cells(i, j + Offset).Value = "."
                    Change = True
                End If
            End If
       
            Next j
        Next i
        Range("output").Copy Range("A1")
        Direction = Direction - 1
    Wend
    stepcount = stepcount + 1
    If stepcount = 1000 Then
        Change = False
    End If
Loop
Cells(1, 148).Value = stepcount
Cells(1, 149).Value = Timer() - t
MsgBox (stepcount)
End Sub
