Attribute VB_Name = "Module1"

Public positionR As Long        'snake's position
Public positionC As Long


Public gameEnd As Boolean
Public gameRunning As Boolean
Public speedInterval As Double
Public lastMoveTime As Double
'the following are four variables for setting up the selected range
Public posiRI As Long
Public posiRF As Long
Public posiCI As Long
Public posiCF As Long

Sub gameStart()
    gameEnd = False
    gameRunning = True
    lastMoveTime = Timer
End Sub

Sub setRangeNumIn()
    posiRI = Selection.row
    posiRF = Selection.row + Selection.Rows.Count - 1
    posiCI = Selection.Column
    posiCF = Selection.Column + Selection.Columns.Count - 1
End Sub

Public Function checkArea() As Boolean
    Dim RowC As Long
    Dim ColC As Long
    RowC = Selection.Rows.Count
    ColC = Selection.Columns.Count
    
    ' To ensure that the backgroundArea is more than 10*10
    If RowC <= 10 Then
        MsgBox "Not enought Rows"
        checkArea = False
    ElseIf ColC <= 10 Then
        MsgBox "Not enought Columns"
        checkArea = False
    Else
        checkArea = True
    End If
End Function

Sub clearBackground()
    MsgBox "clear"
    Selection.Interior.Color = xlNone
End Sub

Sub CreateBackground()
    Dim checkingArea As Boolean
    ' using checkArea() to check if the area is enough to generate a background.
    checkingArea = checkArea()
    If checkingArea Then
        Selection.ColumnWidth = 2.3
        Selection.RowHeight = 15
        ' Selection is the area selected <3
        Selection.Borders.LineStyle = xlNone
        With Selection.BorderAround(Weight:=xlMedium, ColorIndex:=xlAutomatic)
        End With
    Else
        MsgBox "Select Area That is More Than 10*10"
    End If
End Sub

Sub zoomIn()
    Selection.Select
    ActiveWindow.Zoom = True
    Application.Goto Selection, True
End Sub

Sub createSnake()
    Set snakeBodyR = New Collection
    Set snakeBodyC = New Collection
    positionR = Selection.row
    positionC = Selection.Column
    direction = "RIGHT"
    Cells(positionR, positionC).Interior.Color = vbRed
    snakeBodyR.Add positionR                            'Add this to prolong snake length 6/5
    snakeBodyC.Add positionC
    threeFood = 0
End Sub

Sub setInitialSpeed()
    speedInterval = 0.51
End Sub


Sub didGameEnd()
    If positionR > posiRF Then
        gameEnd = True
    ElseIf positionR < posiRI Then
        gameEnd = True
    ElseIf positionC > posiCF Then
        gameEnd = True
    ElseIf positionC < posiCI Then
        gameEnd = True
    End If
    
    Dim i As Long
    For i = 1 To snakeBodyR.Count - 1
        If snakeBodyR(i) = positionR And snakeBodyC(i) = positionC Then
            gameEnd = True
        End If
    Next i
End Sub

Sub stopGame()
    gameRunning = False
    With Range(Cells(posiRI, posiCI), Cells(posiRF, posiCF))

        .Interior.Color = xlNone
        .Borders.LineStyle = xlNone

    End With
    
End Sub

Sub checkEverything()
    
    
    Dim checkingArea As Boolean
    checkingArea = checkArea()      'check selected range
    
    gameStart
    setRangeNumIn               'set the boundary position
    If Not checkingArea Then        'if the range selected didn't exceed 10*10, then an error message will popup
        MsgBox "Select Area That is More Than 10*10"
    End If
    
    clearBackground             'Background setup
    CreateBackground
    zoomIn
    createSnake
    setInitialSpeed
    setControl
    setScoreSpace

    
    gameLoop                'Game start process
    
    
End Sub
