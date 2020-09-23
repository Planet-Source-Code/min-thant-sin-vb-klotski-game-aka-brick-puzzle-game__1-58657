Attribute VB_Name = "basDataFunctions"
Option Explicit

Public Function PuzzleSolved() As Boolean
    Dim i As Integer, xpos As Integer, ypos As Integer
    Dim tmpBrick As CBrick
    
    PuzzleSolved = True
    For i = 1 To DestList.Count
        Set tmpBrick = DestList.Item(i)
        xpos = tmpBrick.XCoord
        ypos = tmpBrick.YCoord
                
        If Board(xpos, ypos).BrickType <> MASTER_BRICK Then
            PuzzleSolved = False
            Exit For
        End If
    Next i
End Function

Public Function GetBrick(ByVal XCoord As Integer, ByVal YCoord As Integer) As CBrick
    Set GetBrick = Nothing
    
    If XCoord >= 0 And XCoord <= (BoardDimX - 1) Then
        If YCoord >= 0 And YCoord <= (BoardDimY - 1) Then
            Set GetBrick = Board(XCoord, YCoord).Brick
        End If
    End If
End Function

Public Function GetBrickInfo(ByVal XCoord As Integer, ByVal YCoord As Integer, ByRef nType As Integer, ByRef nGroupID As Integer) As Boolean
    nType = -1
    nGroupID = -1
                                                     
    GetBrickInfo = False
    
    If XCoord >= 0 And XCoord <= (BoardDimX - 1) Then
        If YCoord >= 0 And YCoord <= (BoardDimY - 1) Then
            nType = Board(XCoord, YCoord).BrickType
            nGroupID = Board(XCoord, YCoord).GroupID
        End If
    End If
End Function

Public Sub GetCoordFromCursorPos(ByVal cursorX As Single, ByVal cursorY As Single, ByRef col As Integer, ByRef row As Integer)
    If cursorX > 0 And cursorX < (BoardWidth - 1) Then
        If cursorY > 0 And cursorY < (BoardHeight - 1) Then
            col = Int(cursorX / GridWidth)
            row = Int(cursorY / GridHeight)
        End If
    End If
End Sub

Public Sub CleanItUp()
    If Not boolLoadSuccessful Then Exit Sub
    
    On Error GoTo ErrorHandler
        
    Dim i As Integer
    Dim row As Integer, col As Integer
    
    Set CurrentList = Nothing
    Set DestList = Nothing
    
    For i = 1 To NumGroups
        Set BrickLists(i) = Nothing
    Next i
    
    For row = 0 To (BoardDimY - 1)
        For col = 0 To (BoardDimX - 1)
            Board(col, row).BrickType = EMPTY_SQUARE
            Board(col, row).GroupID = -1
            Board(col, row).DestGID = -1
            Set Board(col, row).Brick = Nothing
            Set Board(col, row).DestBrick = Nothing
        Next col
    Next row
    
    Erase BrickLists()
    Erase Board()
    
    Exit Sub
ErrorHandler:
    MsgBox "Sub CleanItUp() error", vbInformation, Err.Description
    
End Sub
