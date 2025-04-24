Attribute VB_Name = "Module1"
Public birdRow As Integer
Public birdCol As Integer
Dim isGameRunning As Boolean
Dim prevRow As Integer
Dim prevCol As Integer
Sub DrawHeadingAndInstructions()
    With ThisWorkbook.Sheets(1)
        ' Heading
        .Range("A1:Z1").Merge
        .Range("A1").Value = "BALL GAME"
         ' or Arial, etc.
        .Range("A1").Font.Size = 22
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = RGB(255, 255, 255) ' White
        .Range("A1").Interior.Color = RGB(255, 182, 193) ' Light red-pink
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").VerticalAlignment = xlCenter
        ' Instruction above Start button
        
        .Range("E9:E9").Value = "STOP"
        .Range("O9:O9").Value = "START"
        .Range("B3").Font.Size = 12
        .Range("B3").Font.Color = RGB(255, 69, 0) ' OrangeRed
    End With
End Sub

Sub SetBackground()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    
    ' F
    ws.Range("A1:Z25").Interior.Color = RGB(255, 200, 200) ' Red color
End Sub

Sub CancelGravity()
    On Error Resume Next
    Application.OnTime EarliestTime:=Now + TimeValue("00:00:01"), _
                       Procedure:="ApplyGravity", _
                       Schedule:=False
    On Error GoTo 0
End Sub
Sub startgame()
    CancelGravity
    ClearPreviousGame
    DrawHeadingAndInstructions
    SetBackground
    birdRow = 17
    birdCol = 10
    isGameRunning = True
    Application.OnKey " ", "Flap"
    DrawBird
    ScheduleNextTick
End Sub
Sub ClearPreviousGame()
    ' Clear all previous game content before restarting
    With ThisWorkbook.Sheets(1)
        .Cells.ClearContents
        .Cells.Interior.ColorIndex = -4142 ' Clear any background color
    End With
    prevRow = 0
    prevCol = 0
End Sub
Sub DrawBird()
    On Error Resume Next
    With ThisWorkbook.Sheets(1)
        If prevRow > 0 And prevCol > 0 Then
            .Cells(prevRow, prevCol).ClearContents
        End If
        ' Clear previous bird
        

        ' Update current bird position
        If birdRow < 1 Then birdRow = 1
        If birdCol < 1 Then birdCol = 1

        With .Cells(birdRow, birdCol)
            .Value = "l"
            .Font.Name = "Wingdings"
            .Font.Size = 24
            .Font.ColorIndex = RGB(255, 102, 102) ' yellow
        End With

        ' Remember current for next erase
        prevRow = birdRow
        prevCol = birdCol
    End With
    On Error GoTo 0
End Sub

Sub Flap()
    If birdRow > 2 Then
       birdRow = birdRow - 2
       DrawBird
     End If
End Sub
Sub ApplyGravity()
    birdRow = birdRow + 1 ' gravity pulls down
      If birdRow > 20 Then
        MsgBox "Game Over!"
        isGameRunning = False
        Exit Sub
        
      End If
      DrawBird
      If isGameRunning Then
         ScheduleNextTick
      End If
End Sub
Sub StopGame()
    On Error Resume Next
    Application.OnTime EarliestTime:=Now + TimeValue("00:00:01"), _
                       Procedure:="ApplyGravity", _
                       Schedule:=False
    On Error GoTo 0
    isGameRunning = False
    MsgBox "Game Stopped!"
End Sub
Sub ScheduleNextTick()
    If isGameRunning Then
        Application.OnTime Now + TimeValue("00:00:01"), "ApplyGravity"
     End If
End Sub

        

