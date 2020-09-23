Attribute VB_Name = "basLoadingGame"
Option Explicit

Sub LoadLevel(ByVal FileName As String)
    On Local Error GoTo ErrorHandler
    
    Dim GameData As String, DataChunk As String
    Dim strNumBricks As String, strNumGroups As String
    Dim strBoardDimX As String, strBoardDimY As String
    Dim strDataLength As String
    
    Dim DataLength As Integer, FileNum As Integer
    Dim col As Integer, row As Integer
    Dim i As Integer, j As Integer
    
    Dim tmpBrick As New CBrick
        
    boolGameStarted = False
    frmKlotski.picBackBuffer.Cls
        
    'First clean it up
    Call CleanItUp
    
    boolLoadSuccessful = False
    
    FileNum = FreeFile()
    Open FileName For Input As #FileNum
        Line Input #FileNum, strBoardDimX
        Line Input #FileNum, strBoardDimY
        Line Input #FileNum, strNumBricks
        Line Input #FileNum, strNumGroups
        Line Input #FileNum, strDataLength
        
        If Val(strNumGroups) <= 0 Then
            Exit Sub
        End If
        
        BoardDimX = Val(strBoardDimX)
        BoardDimY = Val(strBoardDimY)
        NumBricks = Val(strNumBricks)
        NumGroups = Val(strNumGroups)
        DataLength = Val(strDataLength)
        
        '////////////////////////////////////////////////////////////////////////////////////////////////
        
        ReDim BrickLists(1 To NumGroups)
        ReDim Board(0 To BoardDimX - 1, 0 To BoardDimY - 1)
        Call ReDimensionBoard
        
        For row = 0 To (BoardDimY - 1)
            For col = 0 To (BoardDimX - 1)
                Board(col, row).BrickType = EMPTY_SQUARE
                Board(col, row).GroupID = -1
                Board(col, row).DestGID = -1
                Set Board(col, row).Brick = Nothing
                Set Board(col, row).DestBrick = Nothing
            Next col
        Next row
    
        '////////////////////////////////////////////////////////////////////////////////////////////////
        
        
        For i = 1 To NumGroups
            Line Input #FileNum, GameData
            
            For j = 1 To Len(GameData) Step DataLength
                Set tmpBrick = New CBrick
                
                DataChunk = Mid$(GameData, j, DataLength)
                
                With tmpBrick
                    .XCoord = Val(Mid(DataChunk, 1, 2))            '[2 chars]
                    .YCoord = Val(Mid(DataChunk, 3, 2))            '[2 chars]
                    .BrickType = Val(Mid(DataChunk, 5, 2))         '[2 chars]
                    .GroupID = Val(Mid(DataChunk, 7, 3))           '[3 chars] ***
                    
                    .Locked = CBool(.BrickType = BARRIER_BRICK)
                                                          
                    If tmpBrick.BrickType = DEST_SQUARE Then
                        If Board(tmpBrick.XCoord, tmpBrick.YCoord).BrickType = EMPTY_SQUARE Then
                            Board(tmpBrick.XCoord, tmpBrick.YCoord).BrickType = DEST_SQUARE
                            Board(tmpBrick.XCoord, tmpBrick.YCoord).GroupID = tmpBrick.GroupID
                            Board(tmpBrick.XCoord, tmpBrick.YCoord).DestGID = tmpBrick.GroupID
                            Set Board(tmpBrick.XCoord, tmpBrick.YCoord).Brick = tmpBrick
                        Else
                            Board(tmpBrick.XCoord, tmpBrick.YCoord).DestGID = tmpBrick.GroupID
                        End If
                    
                        DestList.Add tmpBrick
                    
                    Else
                    
                        Board(tmpBrick.XCoord, tmpBrick.YCoord).GroupID = tmpBrick.GroupID
                        Board(tmpBrick.XCoord, tmpBrick.YCoord).BrickType = tmpBrick.BrickType
                        Set Board(tmpBrick.XCoord, tmpBrick.YCoord).Brick = tmpBrick
                    End If
                
                End With
                    
                If tmpBrick.BrickType <> DEST_SQUARE Then
                    BrickLists(tmpBrick.GroupID).Add tmpBrick
                End If
            Next j
        Next i
    Close #FileNum
    
    
    Call UpdateGameBoard
    boolGameStarted = True
    boolLoadSuccessful = True
        
    Exit Sub
ErrorHandler:
    boolLoadSuccessful = False
    MsgBox "Sub LoadLevel() error", vbInformation, Err.Description
End Sub
