Sub trajectories()
Dim i, j, k, t As Integer
Dim inx, iny As Boolean
s = Timer()
'input
xmin = 88
xmax = 125
ymin = -157
ymax = -103
'example
'xmin = 20
'xmax = 30
'ymin = -10
'ymax = -5

countpairs = 0
For i = ymin To Abs(ymin)
    For j = 1 To xmax
        For t = 0 To 315
            inx = False
            iny = False
            x = posX(j, t)
            If x >= xmin And x <= xmax Then
            
                inx = True

            End If
            y = posY(i, t)
            If y >= ymin And y <= ymax Then
            
                iny = True

            End If
            'if pair ever crosses the target, exit time calculations,
            'move to next pair
            If inx And iny Then
                countpairs = countpairs + 1

                Exit For
            End If
        Next t
    Next j
Next i
Cells(1, 1).Value = countpairs
Cells(2, 1).Value = Timer() - s
End Sub

Function posX(x, t) As Double
    v = x
    posX = 0
    For i = 0 To t
        If v = 0 Then
            Exit For
        End If
        posX = posX + v
        v = v - 1
    Next i
    
End Function

Function posY(y, t) As Double
    v = y
    posY = 0
    For i = 0 To t
        posY = posY + v
        v = v - 1
    Next i
    
End Function
