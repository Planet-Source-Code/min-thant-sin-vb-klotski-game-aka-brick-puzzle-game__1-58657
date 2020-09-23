Attribute VB_Name = "basSavingAndLoading"
Option Explicit

Public Sub SaveLevel(ByVal FileName As String)
    On Error GoTo ErrorHandler
    
    Dim i As Integer, j As Integer, FileNum As Integer
    
    'Note that XCoord, YCoord, ... are declared as String type here
    Dim XCoord As String, YCoord As String
    Dim GID As String, BrickType As String
    Dim GameData As String
    
    Dim DataLength As Integer
    Dim tmpBrick As CBrick
    
    DataLength = 9
    
    FileNum = FreeFile()
    Open FileName For Output As #FileNum
        Print #FileNum, BoardDimX
        Print #FileNum, BoardDimY
        Print #FileNum, NumBricks
        Print #FileNum, NumGroups
        Print #FileNum, DataLength
        
        For i = 1 To NumGroups
            Lists(i).MoveFirst
            
            GameData = ""   'This is important!!
            
            For j = 1 To Lists(i).NumBricks
                Set tmpBrick = Lists(i).CurrentBrick
                
                XCoord = Format$(tmpBrick.XCoord, "00")               'Mid(Data, 1, 2)   [2 chars]
                YCoord = Format$(tmpBrick.YCoord, "00")               'Mid(Data, 3, 2)    [2 chars]
                BrickType = Format$(tmpBrick.BrickType, "00")        'Mid(Data, 5, 2)    [2 chars]
                GID = Format$(tmpBrick.GID, "000")                      'Mid(Data, 7, 3)   * [3 chars] *
                
                GameData = GameData & XCoord & YCoord & BrickType & GID
                
                Lists(i).MoveNext
            Next j
            
            Print #FileNum, GameData
        Next i
        
    Close #FileNum
    
    Exit Sub
ErrorHandler:
    MsgBox "Error loading file", vbInformation, Err.Description

End Sub

Public Sub LoadLevel(ByVal FileName As String)
    On Error GoTo ErrorHandler
    
    Dim i As Integer, j As Integer, FileNum As Integer
    Dim GameData As String, DataChunk As String
    Dim strNumBricks As String, strNumGroups As String
    Dim strBoardDimX As String, strBoardDimY As String
    Dim strDataLength As String
    
    Dim col As Integer, row   As Integer
    Dim DataLength As Integer
    
    Dim tmpBrick As New CBrick
    
    'First clean it up
    Call CleanItUp
        
    FileNum = FreeFile()
    Open FileName For Input As #FileNum
        Line Input #FileNum, strBoardDimX
        Line Input #FileNum, strBoardDimY
        Line Input #FileNum, strNumBricks
        Line Input #FileNum, strNumGroups
        Line Input #FileNum, strDataLength
        
        BoardDimX = Val(strBoardDimX)
        BoardDimY = Val(strBoardDimY)
        NumBricks = Val(strNumBricks)
        NumGroups = Val(strNumGroups)
        DataLength = Val(strDataLength)
            
        Call InitializeBoard
        Call ReDimensionBoard
        ReDim Lists(1 To NumGroups)
    
        For i = 1 To NumGroups
            Line Input #FileNum, GameData
            
            For j = 1 To Len(GameData) Step DataLength
                Set tmpBrick = New CBrick
                
                DataChunk = Mid$(GameData, j, DataLength)
                
                With tmpBrick
                    .XCoord = Val(Mid(DataChunk, 1, 2))            '[2 chars]
                    .YCoord = Val(Mid(DataChunk, 3, 2))            '[2 chars]
                    .BrickType = Val(Mid(DataChunk, 5, 2))         '[2 chars]
                    .GID = Val(Mid(DataChunk, 7, 3))                 '[3 chars] ***
                End With
                                                        
                If tmpBrick.BrickType = DEST_SQUARE Then
                    If Board(tmpBrick.XCoord, tmpBrick.YCoord).BrickType = EMPTY_GRID Then
                        Board(tmpBrick.XCoord, tmpBrick.YCoord).BrickType = DEST_SQUARE
                        Board(tmpBrick.XCoord, tmpBrick.YCoord).GID = tmpBrick.GID
                        Board(tmpBrick.XCoord, tmpBrick.YCoord).DestGID = tmpBrick.GID
                        Set Board(tmpBrick.XCoord, tmpBrick.YCoord).Brick = tmpBrick
                    Else
                        Board(tmpBrick.XCoord, tmpBrick.YCoord).DestGID = tmpBrick.GID
                    End If
                    
                    DestList.Add tmpBrick
                    
                Else
                    
                    Board(tmpBrick.XCoord, tmpBrick.YCoord).GID = tmpBrick.GID
                    Board(tmpBrick.XCoord, tmpBrick.YCoord).BrickType = tmpBrick.BrickType
                    Set Board(tmpBrick.XCoord, tmpBrick.YCoord).Brick = tmpBrick
                    
                End If
                
                Lists(tmpBrick.GID).AddBrick tmpBrick
                
            Next j
        Next i
    Close #FileNum
    
    Call DisplayBoard
    boolCanPlaceBrick = True
    
    Exit Sub
ErrorHandler:
    MsgBox "Error loading file", vbInformation, Err.Description
End Sub
