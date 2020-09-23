Attribute VB_Name = "basGraphicsFunctions"
Option Explicit

Public Sub DisplayBoard()
    On Error GoTo ErrorHandler
    
    Dim i As Integer, j As Integer
    Dim tmpBrick As CBrick
    
    frmEditor.picBackBuffer.Cls
    Call DrawGrid(frmEditor.picBackBuffer, RGB(192, 192, 192))
    
    For i = 1 To DestList.Count
        Set tmpBrick = DestList.Item(i)
        DrawDestination frmEditor.picBackBuffer, tmpBrick.XCoord, tmpBrick.YCoord
    Next i
    
    'Display bricks
    For i = 1 To NumGroups
        Lists(i).MoveFirst
        For j = 1 To Lists(i).NumBricks
            Set tmpBrick = Lists(i).CurrentBrick
            
            'Make sure the dest square is NOT drawn on other bricks
            If tmpBrick.BrickType <> DEST_SQUARE Then
                DrawBrick frmEditor.picBackBuffer, tmpBrick
            End If
            
            Lists(i).MoveNext
        Next j
    Next i
                                
    For i = 1 To NumGroups
        OutlineGroup i
    Next i
    frmEditor.picBoard_Paint
    
    Exit Sub
ErrorHandler:
    MsgBox "Unexpected and unwanted error has occurred!", vbInformation, Err.Description

End Sub

Public Sub DrawGrid(ByRef obj As PictureBox, ByVal color As Long)
    On Error GoTo ErrorHandler
    
    Dim row As Integer, col As Integer
    
    obj.FillStyle = vbTransparent
    obj.ForeColor = color
    
    'Vertical lines
    For col = 1 To (BoardDimX - 1)
        MoveToEx obj.hdc, (col * GridWidth), 0, ByVal 0&
        LineTo obj.hdc, (col * GridWidth), BoardHeight
    Next col
    
    'Horizontal lines
    For row = 1 To (BoardDimY - 1)
        MoveToEx obj.hdc, 0, (row * GridHeight), ByVal 0&
        LineTo obj.hdc, BoardWidth, (row * GridHeight)
    Next row
    
    Exit Sub
ErrorHandler:
    MsgBox "Unexpected and unwanted error has occurred!", vbInformation, Err.Description
    
End Sub

Public Sub DrawBrick(ByRef obj As PictureBox, ByRef Brick As CBrick)
    If (Brick Is Nothing) Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    Dim col As Integer, row As Integer
    Dim bw As Integer, bh As Integer    'Brick's width and height
    Dim bx As Integer, by As Integer     'Brick's x and y positions
    Dim hd As Integer, wd As Integer
    Dim lbType As Integer, lbGID As Integer
    Dim tbType As Integer, tbGID As Integer
    Dim ltbType As Integer, ltbGID As Integer
    Dim BrickGID As Integer
    
    Dim lb As CBrick    'Left brick
    
    hd = GridHeight - BrickHeight
    wd = GridWidth - BrickWidth
    
    BrickGID = Brick.GID
    col = Brick.XCoord
    row = Brick.YCoord
    
    Call GetBrickInfo(col - 1, row, lbType, lbGID)               'Left block
    Call GetBrickInfo(col, row - 1, tbType, tbGID)             'Top block
    Call GetBrickInfo(col - 1, row - 1, ltbType, ltbGID)       'Left-top block
    
    bx = (col * GridWidth)      'brick x
    by = (row * GridHeight)    'brick y
    bw = BrickWidth              'brick width
    bh = BrickHeight              'brick height
    
                
    If lbGID = BrickGID And tbGID = BrickGID And ltbGID = BrickGID Then
        bx = (col * GridWidth) - wd
        by = (row * GridHeight) - hd
        bw = GridWidth
        bh = GridHeight
    Else
        If lbGID = BrickGID And tbGID = BrickGID Then
            bx = (col * GridWidth) - wd
            by = (row * GridHeight)
            bw = GridWidth
            bh = BrickHeight
            DrawIt obj, bx + 1, by + 1, bw, bh, Brick
            
            Set lb = Board(col - 1, row).Brick
            
            If Not lb Is Nothing Then
                lb.Width = GridWidth + wd
            End If
            Set lb = Nothing
            
            bx = (col * GridWidth)
            by = (row * GridHeight) - hd
            bw = BrickWidth
            bh = GridHeight
            DrawIt obj, bx + 1, by + 1, bw, bh, Brick
        Else
    
            If lbGID = BrickGID Then
                bx = (col * GridWidth) - wd
                by = (row * GridHeight)
                bw = GridWidth
                bh = BrickHeight
            End If
            
            If tbGID = BrickGID Then
                bx = (col * GridWidth)
                by = (row * GridHeight) - hd
                bw = BrickWidth
                bh = GridHeight
            End If
        End If
    End If
    
    Brick.Left = bx
    Brick.Top = by
    Brick.Width = bw
    Brick.Height = bh
                    
    DrawIt obj, bx + 1, by + 1, bw, bh, Brick
                        
    Exit Sub
ErrorHandler:
    MsgBox "Unexpected and unwanted error has occurred!", vbInformation, Err.Description
                    
End Sub

Sub DrawIt(ByRef obj As PictureBox, _
                ByVal bx As Integer, ByVal by As Integer, _
                ByVal bw As Integer, ByVal bh As Integer, ByVal Brick As CBrick)
            
    On Error GoTo ErrorHandler
    
    Dim pts(0 To 3) As POINTAPI
    
    obj.FillStyle = vbSolid
    obj.FillColor = ColorTable(Brick.BrickType)
    obj.ForeColor = obj.FillColor
    
    Rectangle obj.hdc, bx, by, (bx + bw), (by + bh)
        
    'Horizontal shade
    obj.ForeColor = RGB(0, 0, 0)
    obj.FillColor = obj.ForeColor
    
    pts(0).x = bx
    pts(0).y = (by + bh)
    pts(1).x = pts(0).x + bw
    pts(1).y = pts(0).y
    pts(2).x = pts(1).x + BrickThickness - 1
    pts(2).y = pts(1).y + BrickThickness - 1
    pts(3).x = pts(0).x + BrickThickness - 1
    pts(3).y = pts(0).y + BrickThickness - 1
    
    Polygon obj.hdc, pts(0), 4
    
    'Vertical shade
    obj.ForeColor = RGB(190, 190, 190)
    obj.FillColor = obj.ForeColor
        
    pts(0).x = (bx + bw)
    pts(0).y = (by + bh)
    pts(1).x = pts(0).x + BrickThickness - 1
    pts(1).y = pts(0).y + BrickThickness - 1
    pts(2).x = (bx + bw) + BrickThickness - 1
    pts(2).y = by + BrickThickness - 1
    pts(3).x = (bx + bw)
    pts(3).y = by
    
    Polygon obj.hdc, pts(0), 4
    
    Exit Sub
ErrorHandler:
    MsgBox "Unexpected and unwanted error has occurred!", vbInformation, Err.Description
            
End Sub

Public Sub DrawDestination(ByRef obj As PictureBox, ByVal XCoord As Integer, ByVal YCoord As Integer)
    On Error GoTo ErrorHandler
    
    Dim bx As Integer, by As Integer, bw As Integer, bh As Integer
        
    bx = (XCoord * GridWidth) + BrickThickness
    by = (YCoord * GridHeight) + BrickThickness
    bw = (BrickWidth - BrickThickness) + 1
    bh = (BrickHeight - BrickThickness) + 1
    
    obj.FillColor = RGB(200, 0, 0)
    obj.FillStyle = vbSolid
    obj.ForeColor = obj.FillColor
    
    Rectangle obj.hdc, bx, by, (bx + bw), (by + bh)
    
    Exit Sub
ErrorHandler:
    MsgBox "Unexpected and unwanted error has occurred!", vbInformation, Err.Description
    
End Sub

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/// I could've done it using complex calculations and reduced processing time.
'/// I just don't have the know-how to do it yet.
'/// For example: I could store the vertices of the shape in an array and use it whenever I need,
'                       but then that would require a much more complicated algorithm to implement it.
'/// So the following algorithm is totally inefficient.
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub OutlineGroup(ByVal GID As Integer)
    
    Dim j As Integer
    Dim xc As Integer, yc As Integer
    Dim xpos As Integer, ypos As Integer
    Dim lbType As Integer, lbGID As Integer
    Dim rbType As Integer, rbGID As Integer
    Dim bbType As Integer, bbGID As Integer
    Dim tbType As Integer, tbGID As Integer
    
    Dim tmpBrick As New CBrick
    
    Lists(GID).MoveFirst
    For j = 1 To Lists(GID).NumBricks
        Set tmpBrick = Lists(GID).CurrentBrick
        
        'No outlining for you BARRIER_BRICK!! Get out!
        If tmpBrick.BrickType = BARRIER_BRICK Then Exit Sub
        
        'No outlining for you too DEST_SQUARE! Get lost!
        If tmpBrick.BrickType = DEST_SQUARE Then Exit Sub
        
        'Not efficent to check brick type like this, but what the heck...
        frmEditor.picBackBuffer.ForeColor = vbYellow
        If tmpBrick.BrickType = FRAME_BRICK Then
            frmEditor.picBackBuffer.ForeColor = vbCyan
        End If
        
        xc = tmpBrick.XCoord
        yc = tmpBrick.YCoord
        xpos = xc * GridWidth + 1
        ypos = yc * GridHeight + 1
        GID = tmpBrick.GID
    
        'I have a penchant to using "Call", don't you think?
        
        Call GetBrickInfo(xc - 1, yc, lbType, lbGID)
        Call GetBrickInfo(xc + 1, yc, rbType, rbGID)
        Call GetBrickInfo(xc, yc + 1, bbType, bbGID)
        Call GetBrickInfo(xc, yc - 1, tbType, tbGID)
    
        '/////////////////////////////////////////////////////////////////////////////////////////////////////////
        '/// Vertical lines
        '/////////////////////////////////////////////////////////////////////////////////////////////////////////
        
        'Left vertical lines
        If lbGID <> GID Then
            If tbGID = GID Then
                Call MoveToEx(frmEditor.picBackBuffer.hdc, xpos, (ypos - BrickThickness), ByVal 0&)
            Else
                Call MoveToEx(frmEditor.picBackBuffer.hdc, xpos, ypos, ByVal 0&)
            End If
        
            If bbGID = GID Then
                Call LineTo(frmEditor.picBackBuffer.hdc, xpos, (ypos + GridHeight))
            Else
                Call LineTo(frmEditor.picBackBuffer.hdc, xpos, (ypos + BrickHeight))
            End If
        End If
        
        
        'Right vertical lines
        If rbGID <> GID Then
            If tbGID = GID Then
                Call MoveToEx(frmEditor.picBackBuffer.hdc, (xpos + BrickWidth), (ypos - BrickThickness - 1), ByVal 0&)
            Else
                Call MoveToEx(frmEditor.picBackBuffer.hdc, (xpos + BrickWidth), ypos, ByVal 0&)
            End If
        
            If bbGID = GID Then
                Call LineTo(frmEditor.picBackBuffer.hdc, (xpos + BrickWidth), (ypos + GridHeight + 1))
            Else
                Call LineTo(frmEditor.picBackBuffer.hdc, (xpos + BrickWidth), (ypos + BrickHeight + 1))
            End If
        End If
        
        
        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '/// Horizontal lines
        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        
        'Top horizontal lines
        If tbGID <> GID Then
            If lbGID = GID Then
                Call MoveToEx(frmEditor.picBackBuffer.hdc, (xpos - BrickThickness), ypos, ByVal 0&)
            Else
                Call MoveToEx(frmEditor.picBackBuffer.hdc, xpos, ypos, ByVal 0&)
            End If
            
            If rbGID = GID Then
                Call LineTo(frmEditor.picBackBuffer.hdc, (xpos + GridWidth + 1), ypos)
            Else
                Call LineTo(frmEditor.picBackBuffer.hdc, (xpos + BrickWidth), ypos)
            End If
        End If
        
        'Bottom horizontal lines
        If bbGID <> GID Then
            If lbGID = GID Then
                Call MoveToEx(frmEditor.picBackBuffer.hdc, (xpos - BrickThickness), (ypos + BrickHeight), ByVal 0&)
            Else
                Call MoveToEx(frmEditor.picBackBuffer.hdc, xpos, (ypos + BrickHeight), ByVal 0&)
            End If
            
            If rbGID = GID Then
                Call LineTo(frmEditor.picBackBuffer.hdc, (xpos + GridWidth + 1), (ypos + BrickHeight))
            Else
                Call LineTo(frmEditor.picBackBuffer.hdc, (xpos + BrickWidth), (ypos + BrickHeight))
            End If
        End If
        
        Lists(GID).MoveNext
    Next j
    
    Set tmpBrick = Nothing
    
End Sub

