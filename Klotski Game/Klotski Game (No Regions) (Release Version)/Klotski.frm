VERSION 5.00
Begin VB.Form frmKlotski 
   Caption         =   "Microsoft(R) Klotski Game - Copyright ZH Computer, 1991 (Recreated by Min Thant Sin, 2005) That's a long time!"
   ClientHeight    =   8745
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBoard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2790
      Left            =   150
      ScaleHeight     =   186
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   2
      Top             =   150
      Width           =   2415
   End
   Begin VB.PictureBox picBlank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0080FF80&
      Height          =   2790
      Left            =   2700
      ScaleHeight     =   186
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   1
      Top             =   150
      Width           =   2415
   End
   Begin VB.PictureBox picBackBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H0080FF80&
      Height          =   2790
      Left            =   5250
      ScaleHeight     =   186
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   171
      TabIndex        =   0
      Top             =   150
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuLoadLevel 
         Caption         =   "&Load Level..."
         Shortcut        =   ^L
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHowToPlay 
         Caption         =   "Ho&w to play..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu sepAbout 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "frmKlotski"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/// ********************************************************************************
'/// *** I have included levels from Microsoft(R) Klotski as well as from Brick Puzzle Jr.
'/// *** Brick Puzzle Jr. is also a very cool game with many special objects such as...
'/// *** Magnetic Bricks, Anti-Magnetic Bricks, Magic Bricks, Black Hole, Traps, Sliders
'/// *** Eliminators, KeyStones, the list goes on...
'/// *** If you're interested, check it out at http://www.bricks-game.de
'/// ********************************************************************************
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/// Microsoft(R) Klotski Game - Recreated by Min Thant Sin, February 2005
'/// I'm sure someone asked me about this game several months ago (I don't remember who).
'/// I've been struggling for admission to Singapore's Republic Polytechnic.
'/// For a few weeks, I was working on this game intermittently.
'/// Eventually, several sleepless nights and hard work culminated in this game.
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/// Feel free to modify this program to suit your needs.
'/// This is by ALL means a slapdash version of the original game.
'/// Many features of the original game have been left out (like keeping track of "Steps" and "Moves")
'/// If possible, please report any bugs found in this program to me.
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/// NO REGIONS are used in outlining bricks in this version.
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Private CurGID As Long  'Current group ID
Private OldMouseX As Single, OldMouseY As Single
Private NewMouseX As Single, NewMouseY As Single

Private boolCanProceed As Boolean

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuHowToPlay_Click()
    frmHowToPlay.Show vbModal
End Sub

Private Sub mnuLoadLevel_Click()
      frmLevels.Show vbModal
End Sub

Public Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If Not boolGameStarted Then Exit Sub
    
    Call ReDimensionBoard
    Call UpdateGameBoard
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CleanItUp
    Unload frmLevels
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    Dim BackgroundColor As Long
            
    BackgroundColor = RGB(127, 127, 127)
    
    'We're going to work in pixels
    picBoard.ScaleMode = vbPixels
    picBoard.BackColor = BackgroundColor
    picBoard.Visible = False
    
    picBackBuffer.ScaleMode = vbPixels
    picBackBuffer.BackColor = BackgroundColor
    picBackBuffer.Visible = False
    
    picBlank.ScaleMode = vbPixels
    picBlank.BackColor = BackgroundColor
    picBlank.Visible = False
        
    ColorTable(FRAME_BRICK) = RGB(0, 0, 255)
    ColorTable(MASTER_BRICK) = RGB(128, 0, 0)
    ColorTable(NORMAL_BRICK) = RGB(128, 128, 0)
    ColorTable(BARRIER_BRICK) = RGB(0, 150, 150)
    ColorTable(DEST_SQUARE) = RGB(128, 0, 0)
        
    boolGameStarted = False
    boolLoadSuccessful = False
          
    If Dir(App.Path & "\Levels", vbDirectory) <> "" Then
          frmLevels.Dir1.Path = App.Path & "\Levels"
    Else
          frmLevels.Dir1.Path = App.Path
    End If
    
    If frmLevels.File1.ListCount = 0 Then
        MsgBox "No level files found in the current directory." & vbCrLf & _
                     "You will have to manually search for the level files", vbInformation, App.Title
    End If
    
    NumGameFiles = frmLevels.File1.ListCount
    CurrentLevel = GetSetting(MY_APP, MY_SECTION, MY_KEY, 0)
    
    If CurrentLevel >= NumGameFiles Then
          CurrentLevel = 0
    End If
    
    If frmLevels.File1.ListCount > 0 Then
          frmLevels.File1.ListIndex = CurrentLevel
          LoadLevel AddASlash(frmLevels.File1.Path) & frmLevels.File1.FileName
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Sub Form_Load() error", vbInformation, Err.Description
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub

Private Sub picBoard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not boolGameStarted Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    Dim col As Integer, row As Integer
    Dim thisType As Integer, thisGID As Integer
    Dim lbType As Integer, lbGID As Integer
    Dim tbType As Integer, tbGID As Integer
    Dim rbType As Integer, rbGID As Integer
    Dim bbType As Integer, bbGID As Integer
    
    Dim NewBrick As New CBrick
    Dim tmpBrick As CBrick
    
    CurGID = 0
    boolCanProceed = True
    
    Call GetCoordFromCursorPos(X, Y, col, row)
    Call GetBrickInfo(col, row, thisType, thisGID)
    
    CurrentBrickType = thisType
    
    Select Case thisType
    Case FRAME_BRICK, DEST_SQUARE, EMPTY_SQUARE
        boolCanProceed = False
        Exit Sub
        
    Case MASTER_BRICK, NORMAL_BRICK
        OldMouseX = X
        OldMouseY = Y
        
        'Current group ID
        CurGID = Board(col, row).GroupID
        Set CurrentList = BrickLists(CurGID)
        
    Case BARRIER_BRICK
        Dim i As Integer
        Dim boolRemoveIt As Boolean
        
        boolCanProceed = False
        boolRemoveIt = True
        
        Set CurrentList = BrickLists(Board(col, row).GroupID)
        
        For i = 1 To CurrentList.Count
            Set tmpBrick = CurrentList.Item(i)
            If tmpBrick.Locked = True Then
                boolRemoveIt = False
                Exit For
            End If
        Next i
        
        If boolRemoveIt Then
            For i = 1 To CurrentList.Count
                Set tmpBrick = CurrentList.Item(i)
                
                Board(tmpBrick.XCoord, tmpBrick.YCoord).GroupID = -1
                Board(tmpBrick.XCoord, tmpBrick.YCoord).DestGID = -1
                Board(tmpBrick.XCoord, tmpBrick.YCoord).BrickType = EMPTY_SQUARE
                Set Board(tmpBrick.XCoord, tmpBrick.YCoord).Brick = Nothing
                
                'Clear old position
                BitBlt picBackBuffer.hdc, _
                       (tmpBrick.XCoord * GridWidth), (tmpBrick.YCoord * GridHeight), _
                        GridWidth, GridHeight, picBlank.hdc, 0, 0, vbSrcCopy
            Next i
            
            Do Until CurrentList.Count <= 0
                CurrentList.Remove 1
            Loop
            
            picBoard_Paint
        End If
    
    Case Else
        boolCanProceed = False
    End Select
        
    Exit Sub
ErrorHandler:
    MsgBox "Sub MouseDown() error", vbInformation, Err.Description
        
End Sub

Private Sub picBoard_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    If X < 0 Or X > BoardWidth Then Exit Sub
    If Y < 0 Or Y > BoardHeight Then Exit Sub
    
    If (boolGameStarted = False) Or (boolCanProceed = False) Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    Dim col As Integer, row As Integer, i As Integer
    Dim DiffX As Integer, DiffY As Integer
    Dim xpos As Integer, ypos As Integer
    Dim bType As Integer, bGID As Integer
    
    Dim hBrush As Long
    Dim boolCanMove As Boolean
        
    Dim tmpBrick As New CBrick
    Dim pt As POINTAPI
    
    Call GetCoordFromCursorPos(X, Y, col, row)
    
    NewMouseX = X
    NewMouseY = Y
    
    DiffX = Abs(OldMouseX - NewMouseX)
    DiffY = Abs(OldMouseY - NewMouseY)
        
    If DiffX >= (BrickWidth \ MOUSE_SENSITIVITY) Then
        'Right direction
        boolCanMove = True
        If NewMouseX > OldMouseX Then
            For i = 1 To CurrentList.Count
                Set tmpBrick = CurrentList.Item(i)
                GetBrickInfo tmpBrick.XCoord + 1, tmpBrick.YCoord, bType, bGID
                
                If bGID <> tmpBrick.GroupID Then
                    If bType <> EMPTY_SQUARE And bType <> DEST_SQUARE Then
                        boolCanMove = False
                        Exit For
                    End If
                End If
                
                If tmpBrick.XCoord >= (BoardDimX - 1) Then
                    boolCanMove = False
                    Exit For
                End If
            Next i
            
            If boolCanMove Then
                Call MoveBrick(RightDirection)
                
                pt.X = OldMouseX + GridWidth
                pt.Y = NewMouseY
                
                'The following three lines could be placed in a sub.
                '/////////////////////////////////////////////////////////////////////////////
                Call ClientToScreen(picBoard.hwnd, pt)
                Call SetCursorPos(pt.X, pt.Y)
                Call ScreenToClient(picBoard.hwnd, pt)
                '/////////////////////////////////////////////////////////////////////////////
                
                OldMouseX = pt.X
                
                Call OutlineGroup(CurGID)
            End If
            
        Else    'Left direction
            
            boolCanMove = True
            For i = 1 To CurrentList.Count
                Set tmpBrick = CurrentList.Item(i)
                GetBrickInfo tmpBrick.XCoord - 1, tmpBrick.YCoord, bType, bGID
                
                If bGID <> tmpBrick.GroupID Then
                    If bType <> EMPTY_SQUARE And bType <> DEST_SQUARE Then
                        boolCanMove = False
                        Exit For
                    End If
                End If
                
                If tmpBrick.XCoord <= 0 Then
                    boolCanMove = False
                    Exit For
                End If
            Next i
            
            If boolCanMove Then
                Call MoveBrick(LeftDirection)
                
                pt.X = OldMouseX - GridWidth
                pt.Y = NewMouseY
                
                'The following three lines could be placed in a sub.
                '/////////////////////////////////////////////////////////////////////////////
                Call ClientToScreen(picBoard.hwnd, pt)
                Call SetCursorPos(pt.X, pt.Y)
                Call ScreenToClient(picBoard.hwnd, pt)
                '/////////////////////////////////////////////////////////////////////////////
                
                OldMouseX = pt.X
                
                Call OutlineGroup(CurGID)
            End If
        End If
    End If
    
    
    If DiffY >= (BrickHeight \ MOUSE_SENSITIVITY) Then
        'Down direction
        boolCanMove = True
        If NewMouseY > OldMouseY Then
            For i = 1 To CurrentList.Count
                Set tmpBrick = CurrentList.Item(i)
                GetBrickInfo tmpBrick.XCoord, tmpBrick.YCoord + 1, bType, bGID
                
                If bGID <> tmpBrick.GroupID Then
                    If bType <> EMPTY_SQUARE And bType <> DEST_SQUARE Then
                        boolCanMove = False
                        Exit For
                    End If
                End If
                
                If tmpBrick.YCoord >= (BoardDimY - 1) Then
                    boolCanMove = False
                    Exit For
                End If
            Next i
            
            If boolCanMove Then
                Call MoveBrick(DownDirection)
                
                pt.X = NewMouseX
                pt.Y = OldMouseY + GridWidth
                
                'The following three lines could be placed in a sub.
                '/////////////////////////////////////////////////////////////////////////////
                Call ClientToScreen(picBoard.hwnd, pt)
                Call SetCursorPos(pt.X, pt.Y)
                Call ScreenToClient(picBoard.hwnd, pt)
                '/////////////////////////////////////////////////////////////////////////////
                
                OldMouseY = pt.Y
                
                Call OutlineGroup(CurGID)
            End If
            
        Else    'UP direction
            
            boolCanMove = True
            For i = 1 To CurrentList.Count
                Set tmpBrick = CurrentList.Item(i)
                GetBrickInfo tmpBrick.XCoord, tmpBrick.YCoord - 1, bType, bGID
                
                If bGID <> tmpBrick.GroupID Then
                    If bType <> EMPTY_SQUARE And bType <> DEST_SQUARE Then
                        boolCanMove = False
                        Exit For
                    End If
                End If
                
                If tmpBrick.YCoord <= 0 Then
                    boolCanMove = False
                    Exit For
                End If
            Next i
            
            If boolCanMove Then
                Call MoveBrick(UpDirection)
                
                pt.X = NewMouseX
                pt.Y = OldMouseY - GridWidth
                
                'The following three lines could be placed in a sub.
                '/////////////////////////////////////////////////////////////////////////////
                Call ClientToScreen(picBoard.hwnd, pt)
                Call SetCursorPos(pt.X, pt.Y)
                Call ScreenToClient(picBoard.hwnd, pt)
                '/////////////////////////////////////////////////////////////////////////////
                
                OldMouseY = pt.Y
                
                Call OutlineGroup(CurGID)
            End If
            
        End If
    End If
    
    picBoard_Paint
    
    If PuzzleSolved Then
        boolGameStarted = False
        frmSolved.Show vbModal
    End If
        
    Exit Sub
ErrorHandler:
    'MsgBox "Sub MouseMove() error", vbInformation, Err.Description
    
End Sub

Private Sub picBoard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
End Sub

Public Sub picBoard_Paint()
    If Not boolLoadSuccessful Then Exit Sub
    Call BitBlt(picBoard.hdc, 0, 0, BoardWidth, BoardHeight, picBackBuffer.hdc, 0, 0, vbSrcCopy)
End Sub

Sub MoveBrick(ByVal d As ENUM_DIRECTION)
    On Local Error GoTo ErrorHandler
    
    Dim i As Integer, j As Integer
    Dim tmpList As New Collection
    Dim tmpBrick As CBrick
    
    For i = 1 To CurrentList.Count
        Set tmpBrick = CurrentList.Item(i)
        
        'Clear old position
        BitBlt picBackBuffer.hdc, _
                tmpBrick.XCoord * GridWidth, tmpBrick.YCoord * GridHeight, _
                GridWidth, GridHeight, picBlank.hdc, 0, 0, vbSrcCopy
                
        If Board(tmpBrick.XCoord, tmpBrick.YCoord).DestGID > 0 Then
            DrawDestination picBackBuffer, tmpBrick.XCoord, tmpBrick.YCoord
            Board(tmpBrick.XCoord, tmpBrick.YCoord).BrickType = DEST_SQUARE
        Else
            Board(tmpBrick.XCoord, tmpBrick.YCoord).BrickType = EMPTY_SQUARE
        End If
        
        Board(tmpBrick.XCoord, tmpBrick.YCoord).GroupID = -1
        Set Board(tmpBrick.XCoord, tmpBrick.YCoord).Brick = Nothing
    Next i
        
    'Update data
    For i = 1 To CurrentList.Count
        Set tmpBrick = CurrentList.Item(i)
        
        If d = RightDirection Then
            tmpBrick.XCoord = tmpBrick.XCoord + 1
        Else
            If d = LeftDirection Then
                tmpBrick.XCoord = tmpBrick.XCoord - 1
            Else
                If d = UpDirection Then
                    tmpBrick.YCoord = tmpBrick.YCoord - 1
                Else
                    tmpBrick.YCoord = tmpBrick.YCoord + 1
                End If
            End If
        End If
        
        Board(tmpBrick.XCoord, tmpBrick.YCoord).GroupID = tmpBrick.GroupID
        Board(tmpBrick.XCoord, tmpBrick.YCoord).BrickType = tmpBrick.BrickType
        Set Board(tmpBrick.XCoord, tmpBrick.YCoord).Brick = tmpBrick
    Next i
        
    For i = 1 To CurrentList.Count
        Set tmpBrick = CurrentList.Item(i)
        DrawBrick picBackBuffer, tmpBrick
    Next i
    
    
    
    'If all the Barrier Bricks in a particular group have been touched by
    'Master Brick, the group can be removed by clicking on it.
    If Not (CurrentBrickType = MASTER_BRICK) Then Exit Sub
    
    Dim xc As Integer, yc As Integer
    Dim lb As CBrick, rb As CBrick, tb As CBrick, bb As CBrick
    Dim ltb As CBrick, rtb As CBrick, lbb As CBrick, rbb As CBrick
    Dim Brick As CBrick
    
    For i = 1 To CurrentList.Count
        Set tmpBrick = CurrentList.Item(i)
        xc = tmpBrick.XCoord
        yc = tmpBrick.YCoord
        
        Set lb = GetBrick(xc - 1, yc)           'Left brick
        Set rb = GetBrick(xc + 1, yc)          'Right brick
        Set tb = GetBrick(xc, yc - 1)          'Top brick
        Set bb = GetBrick(xc, yc + 1)          'Bottom brick
        
        Set ltb = GetBrick(xc - 1, yc - 1)      'Left-top brick
        Set rtb = GetBrick(xc + 1, yc - 1)     'Right-top brick
        Set lbb = GetBrick(xc - 1, yc + 1)     'Left-bottom brick
        Set rbb = GetBrick(xc + 1, yc + 1)    'Right-bottom brick
        
        'If a Barrier Brick is touched by Master Brick, set its Locked flag to False.
        If Not (lb Is Nothing) Then
            If lb.BrickType = BARRIER_BRICK Then
                lb.Locked = False
                Call DisplayGroup(lb.GroupID)
                'Call OutlineGroup(lb.GroupID)
            End If
        End If
        
        If Not (rb Is Nothing) Then
            If rb.BrickType = BARRIER_BRICK Then
                rb.Locked = False
                Call DisplayGroup(rb.GroupID)
                'Call OutlineGroup(rb.GroupID)
            End If
        End If
        
        If Not (tb Is Nothing) Then
            If tb.BrickType = BARRIER_BRICK Then
                tb.Locked = False
                Call DisplayGroup(tb.GroupID)
                'Call OutlineGroup(tb.GroupID)
            End If
        End If
        
        If Not (bb Is Nothing) Then
            If bb.BrickType = BARRIER_BRICK Then
                bb.Locked = False
                Call DisplayGroup(bb.GroupID)
                'Call OutlineGroup(bb.GroupID)
            End If
        End If
                
        'Diagonal bricks
        If Not (ltb Is Nothing) Then
            If ltb.BrickType = BARRIER_BRICK Then
                If Not (tb Is Nothing) And Not (lb Is Nothing) Then
                    If tb.GroupID = ltb.GroupID And lb.GroupID = ltb.GroupID Then
                        ltb.Locked = False
                        Call DisplayGroup(ltb.GroupID)
                        'Call OutlineGroup(ltb.GroupID)
                    End If
                End If
            End If
        End If
        
        If Not (rtb Is Nothing) Then
            If rtb.BrickType = BARRIER_BRICK Then
                If Not (tb Is Nothing) And Not (rb Is Nothing) Then
                    If tb.GroupID = rtb.GroupID And rb.GroupID = rtb.GroupID Then
                        rtb.Locked = False
                        Call DisplayGroup(rtb.GroupID)
                        'Call OutlineGroup(rtb.GroupID)
                    End If
                End If
            End If
        End If
        
        If Not (lbb Is Nothing) Then
            If lbb.BrickType = BARRIER_BRICK Then
                If Not (lb Is Nothing) And Not (bb Is Nothing) Then
                    If lb.GroupID = lbb.GroupID And bb.GroupID = lbb.GroupID Then
                        lbb.Locked = False
                        Call DisplayGroup(lbb.GroupID)
                        'Call OutlineGroup(lbb.GroupID)
                    End If
                End If
            End If
        End If
        
        If Not (rbb Is Nothing) Then
            If rbb.BrickType = BARRIER_BRICK Then
                If Not (rb Is Nothing) And Not (bb Is Nothing) Then
                    If rb.GroupID = rbb.GroupID And bb.GroupID = rbb.GroupID Then
                        rbb.Locked = False
                        Call DisplayGroup(rbb.GroupID)
                        'Call OutlineGroup(rbb.GroupID)
                    End If
                End If
            End If
        End If
        
    Next i
    
    Exit Sub
ErrorHandler:
    'MsgBox "Sub MoveBrick() error", vbInformation, Err.Description
    
End Sub
