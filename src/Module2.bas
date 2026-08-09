Attribute VB_Name = "Module2"
'this module is to control the snake moving

Public direction As String
Public snakeBodyR As New Collection
Public snakeBodyC As New Collection


Sub setControl()
    Application.OnKey "{UP}", "movingUp"
    Application.OnKey "{DOWN}", "movingDown"
    Application.OnKey "{LEFT}", "movingLeft"
    Application.OnKey "{RIGHT}", "movingRight"

End Sub

Sub movingUp()
    If direction <> "DOWN" Then
        direction = "UP"
    End If
End Sub

Sub movingDown()
    If direction <> "UP" Then
        direction = "DOWN"
    End If
    Debug.Print "DOwn"
End Sub

Sub movingRight()
    If direction <> "LEFT" Then
        direction = "RIGHT"
    End If
End Sub

Sub movingLeft()
    If direction <> "RIGHT" Then
        direction = "LEFT"
    End If
End Sub

Sub changeDirection()
    If direction = "UP" Then
        positionR = positionR - 1
    ElseIf direction = "DOWN" Then
        positionR = positionR + 1
    ElseIf direction = "LEFT" Then
        positionC = positionC - 1
    ElseIf direction = "RIGHT" Then
        positionC = positionC + 1
    End If
End Sub

Sub createNewSnake()                    'to prolong snake length 6/5
    
    changeDirection
    snakeBodyR.Add positionR            'new snake head
    snakeBodyC.Add positionC
    
    didGameEnd
    If gameEnd Then
        Exit Sub
    End If
    
    didSnakeEat
    
    
    If threeFood < 3 Then
        Cells(snakeBodyR(1), snakeBodyC(1)).Interior.Color = xlNone
        snakeBodyR.Remove 1
        snakeBodyC.Remove 1
    Else
        threeFood = threeFood - 3
    End If
    
    drawSnake
    
End Sub

Sub drawSnake()                     'to prolong snake length 6/5
    Dim i As Long
    For i = 1 To snakeBodyR.Count
        Cells(snakeBodyR(i), snakeBodyC(i)).Interior.Color = vbRed
    Next i

End Sub

Sub update()                        'to prolong snake length 6/5
    createNewSnake
    generateFood
    didSnakeEat
    
End Sub
    

Sub gameLoop()
    If gameEnd Then
        MsgBox "Game End, You Lose"
        stopGame
        Exit Sub
    End If
    
    update
    ScoreShow
    
    If gameEnd Then
        MsgBox "Game End, You Lose"
        stopGame
        Exit Sub
    End If
    
    nextTick = Now + (speedInterval / 86400)
    Application.OnTime EarliestTime:=nextTick, Procedure:="gameLoop"
End Sub
