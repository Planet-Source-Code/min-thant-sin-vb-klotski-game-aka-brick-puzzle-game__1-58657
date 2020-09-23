Attribute VB_Name = "basGraphicsFunctions"
Option Explicit

Public Sub DrawBrick(ByRef obj As PictureBox, ByRef Brick As CBrick)
    On Local Error GoTo ErrorHandler
    
    Dim col As Integer, row As Integer  'Brick's XCoord and YCoord
    Dim bw As Integer, bh As Integer    'Brick's width and height
    Dim bx As Integer, by As Integer     'Brick's x and y positions
    
    Dim lbType As Integer, lbGID As Integer
    Dim tbType As Integer, tbGID As Integer
    Dim ltbType As Integer, ltbGID As Integer
    Dim BrickGID As Integer
    
    Dim lb As New CBrick
    
    BrickGID = Brick.GroupID
    col = Brick.XCoord
    row = Brick.YCoord
    
    Call GetBrickInfo(col - 1, row, lbType, lbGID)            'Left block
    Call GetBrickInfo(col, row - 1, tbType, tbGID)            'Top block
    Call GetBrickInfo(col - 1, row - 1, ltbType, ltbGID)    'Left-top block
                
    If lbGID = BrickGID And tbGID = BrickGID Then
        If ltbGID = BrickGID Then
            bx = (col * GridWidth) - BrickThickness
            by = (row * GridHeight) - BrickThickness
            bw = GridWidth
            bh = GridHeight
            DrawIt obj, bx, by, bw, bh, Brick
        Else
            bx = (col * GridWidth) - BrickThickness
            by = (row * GridHeight)
            bw = GridWidth
            bh = BrickHeight
            DrawIt obj, bx, by, bw, bh, Brick
            
            
        '///////////////////////////////////////////////////////
            'For outlining region smoothly
            Set lb = Board(col - 1, row).Brick
            
            If Not (lb Is Nothing) Then
                lb.Width = GridWidth + BrickThickness
            End If
            
            Set lb = Nothing
        '///////////////////////////////////////////////////////
            
            bx = (col * GridWidth)
            by = (row * GridHeight) - BrickThickness
            bw = BrickWidth
            bh = GridHeight
            DrawIt obj, bx, by, bw, bh, Brick
        End If
    Else
    
        If lbGID = BrickGID Then
            bx = (col * GridWidth) - BrickThickness
            by = (row * GridHeight)
            bw = GridWidth
            bh = BrickHeight
            DrawIt obj, bx, by, bw, bh, Brick
        Else
            If tbGID = BrickGID Then
                bx = (col * GridWidth)
                by = (row * GridHeight) - BrickThickness
                bw = BrickWidth
                bh = GridHeight
                DrawIt obj, bx, by, bw, bh, Brick
            Else
                bx = (col * GridWidth)
                by = (row * GridHeight)
                bw = BrickWidth
                bh = BrickHeight
                
                DrawIt obj, bx, by, bw, bh, Brick
            End If
        End If
    End If
    
    Brick.Left = bx
    Brick.Top = by
    Brick.Width = bw
    Brick.Height = bh
        
    Exit Sub
ErrorHandler:
    MsgBox "Sub DrawBrick() error", vbInformation, Err.Description
        
End Sub

Sub DrawIt(ByRef obj As PictureBox, _
                ByVal bx As Integer, ByVal by As Integer, _
                ByVal bw As Integer, ByVal bh As Integer, ByVal Brick As CBrick)
                        
    Dim pts(0 To 3) As POINTAPI
    
    obj.FillColor = ColorTable(Brick.BrickType)
    
    If Brick.BrickType = BARRIER_BRICK Then
        If Brick.Locked Then
            obj.FillColor = ColorTable(Brick.BrickType)
        Else
            obj.FillColor = RGB(0, 255, 255)
        End If
    End If
    
    obj.FillStyle = vbSolid
    obj.ForeColor = obj.FillColor
        
    'The brick's surface
    Call Rectangle(obj.hdc, bx, by, (bx + bw), (by + bh))
    

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////
    'Horizontal shade
    obj.ForeColor = RGB(0, 0, 0)
    obj.FillColor = obj.ForeColor
    
    pts(0).X = bx
    pts(0).Y = (by + bh)
    pts(1).X = pts(0).X + bw
    pts(1).Y = pts(0).Y
    pts(2).X = pts(1).X + (BrickThickness - 1)
    pts(2).Y = pts(1).Y + (BrickThickness - 1)
    pts(3).X = pts(0).X + (BrickThickness - 1)
    pts(3).Y = pts(0).Y + (BrickThickness - 1)
    
    Polygon obj.hdc, pts(0), 4
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////
        
        
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////
    'Vertical shade
    obj.ForeColor = RGB(190, 190, 190)
    obj.FillColor = obj.ForeColor
        
    pts(0).X = (bx + bw)
    pts(0).Y = (by + bh)
    pts(1).X = pts(0).X + (BrickThickness - 1)
    pts(1).Y = pts(0).Y + (BrickThickness - 1)
    pts(2).X = (bx + bw) + (BrickThickness - 1)
    pts(2).Y = by + (BrickThickness - 1)
    pts(3).X = (bx + bw)
    pts(3).Y = by
    
    Polygon obj.hdc, pts(0), 4
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////
End Sub

Public Sub DrawDestination(ByRef obj As PictureBox, ByVal XCoord As Integer, ByVal YCoord As Integer)
    Dim bx As Integer, by As Integer, bw As Integer, bh As Integer
    Dim hd As Integer, wd As Integer
    
    hd = GridHeight - BrickHeight
    wd = GridWidth - BrickWidth
        
    bx = (XCoord * GridWidth) + wd     'brick x
    by = (YCoord * GridHeight) + hd     'brick y
    bw = BrickWidth - wd                    'brick width
    bh = BrickHeight - hd                    'brick height
    
    obj.FillColor = RGB(128, 0, 0)
    obj.FillStyle = vbSolid
    obj.ForeColor = obj.FillColor
    
    Rectangle obj.hdc, bx, by, (bx + bw), (by + bh)
End Sub

Public Sub DisplayGroup(ByVal GID As Integer)
    Dim i As Integer
    Dim tmpBrick As CBrick
    Dim tmpList As Collection
    
    Set tmpList = BrickLists(GID)
    
    For i = 1 To tmpList.Count
        Set tmpBrick = tmpList.Item(i)
        DrawBrick frmKlotski.picBackBuffer, tmpBrick
    Next i
End Sub
