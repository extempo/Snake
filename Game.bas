'Snake v1.0, 2014-01-02
'By Matt Carleton

'Module "Game" - the primary code of the game


Option Explicit

Dim gameBoard As Range 'range featuring all cells in the board, including walls

'Game status, defined as one of the following constants
Dim gameStatus As Integer
Const STAT_PLAY As Integer = 1
Const STAT_CLEAR As Integer = 2
Const STAT_GAME_OVER As Integer = 3
Const STAT_PAUSE As Integer = 4

'Number of cycles since game started
Dim tick As Long

'Relevant sheets
Dim screenSheet As Worksheet
Dim resourceSheet As Worksheet
Dim scoreSheet As Worksheet

'Game board attributes
Const MIN_X = 1
Const MIN_Y = 1
Const MAX_X = 27
Const MAX_Y = 22

'Directions
Const DIR_LEFT As Integer = 1
Const DIR_RIGHT As Integer = 2
Const DIR_UP As Integer = 3
Const DIR_DOWN As Integer = 4

Dim wallColor As Integer, snakeColor As Integer, headColor As Integer, foodColor As Integer, emptyColor As Integer

Dim currentDir As Integer

Dim currentX As Integer
Dim currentY As Integer

Dim foodX As Integer 'where is the food
Dim foodY As Integer 'where is the food??

Dim segmentCounter As Integer 'lifetime segment counter for snake
Dim snakeLength As Integer 'current length of the snake
Dim segmentCollection As Collection 'active snake segments
Dim foodFound As Boolean 'true when food is found. used to increase length

Dim speed As Integer 'speed. 1 = higher minMillis
Dim minMillis As Integer    'minimum number of milliseconds per tick.

Dim newHighScore As Boolean

Dim sw As StopWatch


Sub snake()
    bootGame
    
    Do While gameStatus <> STAT_GAME_OVER
        mainLoop
    Loop
    
    shutDownGame
    
End Sub

Private Sub bootGame()
    Debug.Print "Booting game"
    
    'Turn off screen updating/automatic calculations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    'set initial variables
    setColors
    setSpeed 1
    
    Set screenSheet = Worksheets("Screen")
    Set resourceSheet = Worksheets("Resources")
    Set scoreSheet = Worksheets("Scores")

    Set gameBoard = screenSheet.Range("gameBoard")
    
    Set sw = New StopWatch
    
    'unlock the screen sheet and leave it unprotected for duration of game
    unlockSheet screenSheet, "snake"
    
    'In case it previously said 'new high score!'
    screenSheet.Range("scoreMessage").Value = ""
    newHighScore = False
    
    'Remove the 'newest' tag from existing scores, to avoid bold-faced scores during gameplay
    unlockSheet scoreSheet, "snake"
    scoreSheet.Range("D1:D" & scoreSheet.Range("numScores").Value).ClearContents
    lockSheet scoreSheet, "snake"
    
    'copy board from template
    resourceSheet.Range("gameBoardTemplate").Copy screenSheet.Range("A1")
    Application.CutCopyMode = False
    
    
    screenSheet.Cells(1, 1).Select

    
    'create snake
    Set segmentCollection = New Collection
    createSnake
    
    'place the initial bit of food
    placeFood
    
    
    'Show everything on screen
    updateScreen
    
    'set game status to play
    gameStatus = STAT_PLAY
    tick = 0
    
    Debug.Print "Booting complete."
    Debug.Print "---"
    Debug.Print

    'cellAge as
'    cellAge(1, 1) = 2
'    Debug.Print cellAge(1, 1)
    
    
End Sub

Private Sub mainLoop()
    tick = tick + 1
    
    sw.StartTimer  'starts timer
    Debug.Print "Entering main loop. Tick: " & tick
            
    foodFound = False
                
'    'check for pause button
'    If GetAsyncKeyState(vbKeySpace) <> 0 Then
'        If gameStatus = STAT_PAUSE Then
'            gameStatus = STAT_PLAY
'        Else: gameStatus = STAT_PAUSE
'        End If
'    End If
    
    If gameStatus = STAT_PAUSE Then
        '
    
    Else    'if not paused, continue
        'check for other user input
        If GetAsyncKeyState(vbKeyLeft) <> 0 And currentDir <> DIR_RIGHT And currentDir <> DIR_LEFT Then
            currentDir = DIR_LEFT
        ElseIf GetAsyncKeyState(vbKeyRight) <> 0 And currentDir <> DIR_RIGHT And currentDir <> DIR_LEFT Then
            currentDir = DIR_RIGHT
        ElseIf GetAsyncKeyState(vbKeyUp) <> 0 And currentDir <> DIR_UP And currentDir <> DIR_DOWN Then
            currentDir = DIR_UP
        ElseIf GetAsyncKeyState(vbKeyDown) <> 0 And currentDir <> DIR_UP And currentDir <> DIR_DOWN Then
            currentDir = DIR_DOWN
        End If
        
        'Move
        If currentDir = DIR_LEFT Then
            currentX = currentX - 1
        ElseIf currentDir = DIR_RIGHT Then
            currentX = currentX + 1
        ElseIf currentDir = DIR_UP Then
            currentY = currentY - 1
        Else
            currentY = currentY + 1
        End If
        
        'check for wall collisions
        If currentY <= MIN_Y Or currentY >= MAX_Y Or currentX <= MIN_X Or currentX >= MAX_X Then
            'gameOver
            gameStatus = STAT_GAME_OVER
            Debug.Print "final x/y: " & currentX, currentY
        ElseIf gameBoard.Cells(currentY, currentX).Interior.ColorIndex = snakeColor Then
            'check for self collisions
            gameStatus = STAT_GAME_OVER
            Debug.Print "final x/y: " & currentX, currentY
        End If
        
        'check for food collision
        If currentX = foodX And currentY = foodY Then
            foodFound = True
            placeFood
        End If
        
        'if not game over, add new segment
        
        If gameStatus <> STAT_GAME_OVER Then
            addSegment currentX, currentY, foodFound
            Dim seg As SnakeSegment
            For Each seg In segmentCollection
              '  Debug.Print "segmentCounter: " & segmentCounter, seg.index, snakeLength
                If segmentCounter - seg.index > snakeLength Then
                    gameBoard.Cells(seg.y, seg.x).Interior.ColorIndex = emptyColor
               '     Debug.Print "Deleting segment!"
                    segmentCollection.Remove "" & seg.index
                    Set seg = Nothing
                End If
            Next seg
        End If
    End If 'end 'if not paused'
    
    updateScreen
    
    While sw.EndTimer < minMillis
        'Debug.Print "sleeping for " & minMillis & " minus " & sw.EndTimer
        Sleep (minMillis - sw.EndTimer)
    Wend
    'Debug.Print "Time taken: " & sw.EndTimer

End Sub


Private Sub setColors()
    wallColor = 0
    snakeColor = 4
    headColor = 3
    foodColor = 23
    emptyColor = 1

End Sub

Private Sub setSpeed(newSpeed As Integer)
    speed = newSpeed
    
    Select Case speed
        Case 1 To 12
            minMillis = 110 - 5 * speed
            'minMillis = 110 - 15 * speed
        'Case
        Case Else:
            minMillis = 45
    End Select
End Sub


Private Sub createSnake()
    segmentCounter = 0
    snakeLength = 0
    
    currentX = 5
    currentY = 5
    
    addSegment currentX - 2, currentY, True
    addSegment currentX - 1, currentY, True
    addSegment currentX, currentY, True
    
    currentDir = DIR_RIGHT

End Sub

Private Sub placeFood()
    setSpeed (speed + 1)
    Dim spaceFound As Boolean, randX As Integer, randY As Integer
    spaceFound = False
    
    While spaceFound <> True
        randX = Int((MAX_X - MIN_X + 1) * Rnd() + MIN_X)
        randY = Int((MAX_Y - MIN_Y + 1) * Rnd() + MIN_Y)
        ' randY = 5
        If gameBoard.Cells(randY, randX).Interior.ColorIndex = emptyColor Then
            spaceFound = True
        End If
    Wend
    
    foodX = randX
    foodY = randY
    gameBoard.Cells(foodY, foodX).Interior.ColorIndex = foodColor
'TODO fix
End Sub
Private Sub addSegment(x As Integer, y As Integer, increaseLength As Boolean)
    
    segmentCounter = segmentCounter + 1
    If increaseLength Then
        snakeLength = snakeLength + 1
        screenSheet.Range("score").Value = snakeLength + 1
        
        If snakeLength + 1 > screenSheet.Range("highScore").Value Then
            newHighScore = True
        End If
    End If
    
    Dim seg As SnakeSegment
    Set seg = New SnakeSegment
    seg.x = x
    seg.y = y
    seg.index = segmentCounter
    
    gameBoard.Cells(seg.y, seg.x).Interior.ColorIndex = snakeColor
    
    segmentCollection.Add seg, "" & seg.index
    
End Sub


Private Sub shutDownGame()
    Debug.Print "Game over, man."
    Debug.Print
    
    If newHighScore Then
        screenSheet.Range("scoreMessage").Value = "New High Score!"
    End If
    
    addScore screenSheet.Range("playerName"), screenSheet.Range("score"), Format(Now(), "yyyy-mm-dd hh:nn:ss")
    
    'Reprotect the screen sheet
    lockSheet Sheets("Screen"), "snake"
    
    'drawItem "gameOver", 65, 2
    MsgBox "Game over, man"
    
    'Turn off screen updating/automatic calculations
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub

Private Sub updateScreen()
    'gameBoard.Cells(currentY, currentX).Select
    Application.ScreenUpdating = True
    DoEvents 'WOOO
    gameBoard.Cells(currentY, currentX).Select
    Application.ScreenUpdating = False
End Sub

Private Sub addScore(playerName As String, score As Integer, timeStamp As String)
    unlockSheet scoreSheet, "snake"
    
    'how many rows exist so far?
    Dim newRow As Integer: newRow = scoreSheet.Range("numScores").Value + 1
    
    'clear old 'newest score' value
    scoreSheet.Range("D1:D" & newRow).ClearContents
    
    'add score to score sheet
    scoreSheet.Range("A" & newRow).Value = playerName
    scoreSheet.Range("B" & newRow).Value = score
    scoreSheet.Range("C" & newRow).Value = timeStamp
    scoreSheet.Range("D" & newRow).Value = "newest"
    
    
    'Sort by high scores (then by newest)
    With scoreSheet.sort
        .SetRange Range("A1:D" & newRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    lockSheet scoreSheet, "snake"
    
End Sub
Public Sub clearScores()
    'TODO
    Dim yesNo As String, numScores As Integer
    yesNo = MsgBox("Are you sure you want to clear the high score list?", vbYesNo, "Clear scores?")
    If yesNo = vbYes Then
        With Sheets("Scores")
            unlockSheet Sheets("Scores"), "snake"
            numScores = .Range("numScores").Value
            If numScores > 1 Then
                .Range("A2:D" & numScores).ClearContents
            End If
            lockSheet Sheets("Scores"), "snake"
        End With
        
        Sheets("Screen").Range("score").Value = "0"
        
    End If
End Sub

