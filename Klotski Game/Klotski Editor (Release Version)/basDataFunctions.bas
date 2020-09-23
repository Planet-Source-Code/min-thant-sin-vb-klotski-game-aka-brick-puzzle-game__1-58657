Attribute VB_Name = "basDataFunctions"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/// The following functions and subs are self-explanatory, and they are very easy to understand.
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Public Function GetBrick(ByVal XCoord As Integer, ByVal YCoord As Integer) As CBrick
    Set GetBrick = Nothing
    
    If XCoord >= 0 And XCoord <= (BoardDimX - 1) Then
        If YCoord >= 0 And YCoord <= (BoardDimY - 1) Then
            Set GetBrick = Board(XCoord, YCoord).Brick
        End If
    End If
End Function

Public Function GetBrickInfo(ByVal XCoord As Integer, ByVal YCoord As Integer, _
                                      ByRef nType As Integer, ByRef nGroupID As Integer) As Boolean
    nType = -1:    nGroupID = -1
    
    GetBrickInfo = False
    If XCoord >= 0 And XCoord <= (BoardDimX - 1) Then
        If YCoord >= 0 And YCoord <= (BoardDimY - 1) Then
            nType = Board(XCoord, YCoord).BrickType
            nGroupID = Board(XCoord, YCoord).GID
            
            GetBrickInfo = True
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

Public Sub InitializeBoard()
    On Error GoTo ErrorHandler
    
    Dim col As Integer, row As Integer
                    
    CurrentGID = 0
    
    boolFirstBrickPlaced = False
    boolCanPlaceBrick = False
    
    ReDim Board(0 To BoardDimX - 1, 0 To BoardDimY - 1)
    
    For row = 0 To (BoardDimY - 1)
        For col = 0 To (BoardDimX - 1)
            Board(col, row).BrickType = EMPTY_GRID
            Board(col, row).GID = -1
            Board(col, row).DestGID = -1
            Set Board(col, row).Brick = Nothing
        Next col
    Next row
    
    Exit Sub
ErrorHandler:
    MsgBox "Unexpected and unwanted error has occurred!", vbInformation, Err.Description

End Sub
