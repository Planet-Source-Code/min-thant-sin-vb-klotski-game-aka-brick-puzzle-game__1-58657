Attribute VB_Name = "basMiscellaneous"
Option Explicit

Public Function AddASlash(ByVal strIn As String) As String
    AddASlash = strIn
    If Right$(strIn, 1) <> "\" Then AddASlash = strIn & "\"
End Function

Public Sub CleanItUp()
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    
    For i = 1 To NumGroups
        Set Lists(i) = Nothing
    Next i
    
    Erase Lists()
    
    Do Until DestList.Count = 0
        DestList.Remove 1
    Loop
    
    Set DestList = Nothing
    
    Exit Sub
ErrorHandler:
    MsgBox "Unexpected and unwanted error has occurred!", vbInformation, Err.Description
        
End Sub

Public Sub DisplayInfo()
    With frmEditor
        .lblCurrentGID = "Group ID : " & CurrentGID
        .lblCurrentBrickType = "Brick Type : " & CurrentBrickType
        
        .lblBoardDimX = "BoardDimX : " & BoardDimX
        .lblBoardDimY = "BoardDimY : " & BoardDimY
        .lblNumBricks = "Num Bricks : " & NumBricks
        .lblNumGroups = "Num Groups : " & NumGroups
    End With
End Sub

'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/// Resizes picture boxes based on BoardDimX, BoardDimY and the client area of frmEditor
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub ReDimensionBoard()
    On Error GoTo ErrorHandler
    
    Dim ClientWidth As Integer, ClientHeight As Integer
    Dim MenuHeight As Integer, TitleBarHeight As Integer
    Dim ClientRect As RECT
    
    TitleBarHeight = GetSystemMetrics(SM_CYCAPTION)
    MenuHeight = GetSystemMetrics(SM_CYMENU)
    
    Call GetClientRect(frmEditor.hwnd, ClientRect)
    
    ClientWidth = (ClientRect.Right - ClientRect.Left) - (FrameWidth * 5 / 4)
    ClientHeight = (ClientRect.Bottom - ClientRect.Top) - (TitleBarHeight + MenuHeight)
    
    If ClientWidth = 0 Or ClientHeight = 0 Then Exit Sub
    
    GridWidth = ClientWidth / BoardDimX
    GridHeight = ClientHeight / BoardDimY
           
    If GridWidth < MIN_GRID_SIZE Or GridHeight < MIN_GRID_SIZE Then
        GridWidth = MIN_GRID_SIZE
        GridHeight = MIN_GRID_SIZE
    Else
        If GridHeight <= GridWidth Then
            GridWidth = GridHeight
        Else
            GridHeight = GridWidth
        End If
    End If
 
    BrickWidth = Int(GridWidth * BRICK_WIDTH_PERCENT)
    BrickHeight = Int(GridHeight * BRICK_HEIGHT_PERCENT)
    
    BrickThickness = (GridWidth - BrickWidth) - 1
    
    BoardWidth = GridWidth * BoardDimX
    BoardHeight = GridHeight * BoardDimY
    
    'Resize the picBoard
    With frmEditor.picBoard
        .Width = .ScaleX(BoardWidth, vbPixels, vbTwips)
        .Height = .ScaleY(BoardHeight, vbPixels, vbTwips)
    End With
          
    'Resize the picBackBuffer
    With frmEditor.picBackBuffer
        .Width = .ScaleX(BoardWidth, vbPixels, vbTwips)
        .Height = .ScaleY(BoardHeight, vbPixels, vbTwips)
    End With
          
    'Resize the picBlank
    With frmEditor.picBlank
        .Width = .ScaleX(GridWidth, vbPixels, vbTwips)
        .Height = .ScaleY(GridHeight, vbPixels, vbTwips)
    End With
            
    Exit Sub
ErrorHandler:
    MsgBox "Unexpected and unwanted error has occurred!", vbInformation, Err.Description
            
End Sub
