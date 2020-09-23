Attribute VB_Name = "basMiscellaneous"
Option Explicit

'Save the index of the current game level
Public Sub SaveLevelStatus()
      SaveSetting MY_APP, MY_SECTION, MY_KEY, CStr(CurrentLevel)
End Sub

Public Function AddASlash(ByVal strIn As String) As String
      If Right$(strIn, 1) = "\" Then
            AddASlash = strIn
      Else
            AddASlash = strIn & "\"
      End If
End Function

Public Sub UpdateGameBoard()
    On Local Error GoTo ErrorHandler
    
    frmKlotski.picBoard.Visible = False
    
    Dim i As Integer, j As Integer
    Dim rgnBrick As Long
    Dim hBrush As Long
    
    Dim tmpBrick As CBrick
    
    frmKlotski.picBackBuffer.Cls
    
    'Draw destination squares
    For i = 1 To DestList.Count
        Set tmpBrick = DestList.Item(i)
        If Board(tmpBrick.XCoord, tmpBrick.YCoord).BrickType = DEST_SQUARE Then
            DrawDestination frmKlotski.picBackBuffer, tmpBrick.XCoord, tmpBrick.YCoord
        End If
    Next i
            
    'Draw the rest
    For i = 1 To NumGroups
        For j = 1 To BrickLists(i).Count
            Set tmpBrick = BrickLists(i).Item(j)
            If tmpBrick.BrickType <> DEST_SQUARE Then
                DrawBrick frmKlotski.picBackBuffer, tmpBrick
            End If
        Next j
    Next i
            
    For i = 1 To NumGroups
        Call OutlineGroup(i)
    Next i
                
    frmKlotski.picBoard_Paint
    frmKlotski.picBoard.Visible = True
    
    Exit Sub
ErrorHandler:
    MsgBox "Sub UpdateGameBoard() error", vbInformation, Err.Description
    
End Sub

Public Sub ReDimensionBoard()
    Dim ClientWidth As Integer, ClientHeight As Integer
    Dim MenuHeight As Integer, TitleBarHeight As Integer
    Dim ClientRect As RECT
    
    frmKlotski.picBoard.Visible = False
    
    TitleBarHeight = GetSystemMetrics(SM_CYCAPTION)
    MenuHeight = GetSystemMetrics(SM_CYMENU)
    
    Call GetClientRect(frmKlotski.hwnd, ClientRect)
    
    ClientWidth = (ClientRect.Right - ClientRect.Left)
    ClientHeight = (ClientRect.Bottom - ClientRect.Top) - (TitleBarHeight + MenuHeight)
    
    If ClientWidth = 0 Or ClientHeight = 0 Then Exit Sub
    
    GridWidth = ClientWidth / BoardDimX
    GridHeight = ClientHeight / BoardDimY
        
    If GridHeight < MIN_GRID_SIZE Or GridWidth < MIN_GRID_SIZE Then
        GridHeight = MIN_GRID_SIZE
        GridWidth = MIN_GRID_SIZE
    End If
        
    If GridHeight <= GridWidth Then
        GridWidth = GridHeight
    Else
        GridHeight = GridWidth
    End If
    
    BrickWidth = Int(GridWidth * BRICK_WIDTH_PERCENT)
    BrickHeight = Int(GridHeight * BRICK_HEIGHT_PERCENT)
    
    BrickThickness = (GridWidth - BrickWidth)
    
    BoardWidth = GridWidth * BoardDimX
    BoardHeight = GridHeight * BoardDimY
    
    'Resize the picBoard
    With frmKlotski.picBoard
        .Width = .ScaleX(BoardWidth, vbPixels, vbTwips)
        .Height = .ScaleY(BoardHeight, vbPixels, vbTwips)
        .Left = (frmKlotski.Width - .Width) \ 2
    End With
          
    'Resize the picBackBuffer
    With frmKlotski.picBackBuffer
        .Width = .ScaleX(BoardWidth, vbPixels, vbTwips)
        .Height = .ScaleY(BoardHeight, vbPixels, vbTwips)
    End With
          
    'Resize the picBlank
    With frmKlotski.picBlank
        .Width = .ScaleX(GridWidth, vbPixels, vbTwips)
        .Height = .ScaleY(GridHeight, vbPixels, vbTwips)
    End With
            
    frmKlotski.picBoard.Visible = True
End Sub
