Attribute VB_Name = "Module3"
    'this module is used for calculating score
Public ScorePoR As Long
Public ScorePoC As Long
Public score As Long
Public SnakeEat As Boolean



Sub setScoreSpace()
    score = 0
    ScorePoR = posiRI       'same row as the first row of the selected range
    ScorePoC = posiCF + 2   'two more row than the last column of the selected range
    
    Cells(ScorePoR, ScorePoC).Value = "Score"
    ScorePoR = ScorePoR + 1
End Sub


Sub ScoreCount()
    RowC = posiRF - posiRI
    ColC = posiCF - posiCI
    score = score + Application.Max(1, Int(500 / Sqr(RowC * ColC)))
    
    If SnakeEat Then
        score = score + 1000
        SnakeEat = False
    End If
End Sub


Sub didSnakeEat()
    Dim item As Variant
    For Each item In foodColle
        If item = positionR & ", " & positionC Then
            SnakeEat = True
            threeFood = threeFood + 1
        End If
    Next item

End Sub


Sub ScoreShow()
    ScoreCount
    Cells(ScorePoR, ScorePoC).Value = score
End Sub
