VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEditor 
   Caption         =   "Klotski Level Editor  (by Min Thant Sin, February 2005)"
   ClientHeight    =   9240
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBlank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   465
      Left            =   8100
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   10
      Top             =   225
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Frame fraBricks 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6990
      Left            =   75
      TabIndex        =   0
      Top             =   150
      Width           =   1965
      Begin VB.Frame Frame1 
         Height          =   90
         Left            =   75
         TabIndex        =   18
         Top             =   6075
         Width           =   1815
      End
      Begin VB.CommandButton cmdNewBrick 
         Caption         =   "&New Brick"
         Height          =   540
         Left            =   150
         TabIndex        =   17
         ToolTipText     =   "Click here or press F2 or N key to place a new brick on the game board"
         Top             =   6300
         Width           =   1665
      End
      Begin VB.OptionButton optBrick 
         Caption         =   "Empty Grid"
         Height          =   915
         Index           =   5
         Left            =   150
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   $"Klotski Level Editor.frx":0000
         Top             =   5100
         UseMaskColor    =   -1  'True
         Width           =   1665
      End
      Begin VB.OptionButton optBrick 
         Caption         =   "Frame Brick"
         Height          =   915
         Index           =   0
         Left            =   150
         MaskColor       =   &H0000FF00&
         Picture         =   "Klotski Level Editor.frx":009F
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "(BrickType = 0) Used to draw bricks that are NOT movable"
         Top             =   225
         UseMaskColor    =   -1  'True
         Value           =   -1  'True
         Width           =   1665
      End
      Begin VB.OptionButton optBrick 
         Caption         =   "Master Brick"
         Height          =   915
         Index           =   1
         Left            =   150
         MaskColor       =   &H0000FF00&
         Picture         =   "Klotski Level Editor.frx":0CE1
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "(BrickType = 1) This is the Master Brick which the user has to move to its destination"
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   1665
      End
      Begin VB.OptionButton optBrick 
         Caption         =   "Normal Brick"
         Height          =   915
         Index           =   2
         Left            =   150
         MaskColor       =   &H0000FF00&
         Picture         =   "Klotski Level Editor.frx":1923
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "(BrickType = 2) Normal movable brick"
         Top             =   2175
         UseMaskColor    =   -1  'True
         Width           =   1665
      End
      Begin VB.OptionButton optBrick 
         Caption         =   "Barrier Brick"
         Height          =   915
         Index           =   3
         Left            =   150
         MaskColor       =   &H0000FF00&
         Picture         =   "Klotski Level Editor.frx":2565
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "(BrickType = 3) This is NOT movable but removable if touched by master brick completely"
         Top             =   3150
         UseMaskColor    =   -1  'True
         Width           =   1665
      End
      Begin VB.OptionButton optBrick 
         Caption         =   "Destination Square"
         Height          =   915
         Index           =   4
         Left            =   150
         MaskColor       =   &H00808080&
         Picture         =   "Klotski Level Editor.frx":31A7
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "(BrickType = 4) Used to mark Master Brick's destination"
         Top             =   4125
         UseMaskColor    =   -1  'True
         Width           =   1665
      End
   End
   Begin VB.Frame fraCurrentStatus 
      Caption         =   "Status"
      Height          =   2340
      Left            =   75
      TabIndex        =   7
      Top             =   7275
      Width           =   1965
      Begin VB.Label lblNumGroups 
         AutoSize        =   -1  'True
         Caption         =   "Num Groups :"
         Height          =   240
         Left            =   150
         TabIndex        =   16
         ToolTipText     =   "The number of brick groups currently on the board"
         Top             =   1350
         Width           =   1275
      End
      Begin VB.Label lblNumBricks 
         AutoSize        =   -1  'True
         Caption         =   "Num Bricks :"
         Height          =   240
         Left            =   255
         TabIndex        =   15
         ToolTipText     =   "Number of bricks currently on the board"
         Top             =   1050
         Width           =   1170
      End
      Begin VB.Label lblBoardDimX 
         AutoSize        =   -1  'True
         Caption         =   "BoardDimX :"
         Height          =   240
         Left            =   240
         TabIndex        =   14
         ToolTipText     =   "The number of grids on the board horizontally"
         Top             =   1650
         Width           =   1185
      End
      Begin VB.Label lblBoardDimY 
         AutoSize        =   -1  'True
         Caption         =   "BoardDimY :"
         Height          =   240
         Left            =   225
         TabIndex        =   13
         ToolTipText     =   "The number of grids on the board vertically"
         Top             =   1950
         Width           =   1185
      End
      Begin VB.Label lblCurrentGID 
         AutoSize        =   -1  'True
         Caption         =   "Group ID :"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   450
         TabIndex        =   12
         ToolTipText     =   "Currently selected group ID"
         Top             =   375
         Width           =   975
      End
      Begin VB.Label lblCurrentBrickType 
         AutoSize        =   -1  'True
         Caption         =   "Brick Type :"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   300
         TabIndex        =   11
         ToolTipText     =   "Currently selected brick type"
         Top             =   675
         Width           =   1125
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8100
      Top             =   750
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picBoard 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   2190
      Left            =   2175
      ScaleHeight     =   146
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   166
      TabIndex        =   8
      Top             =   225
      Width           =   2490
   End
   Begin VB.PictureBox picBackBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   2265
      Left            =   5100
      ScaleHeight     =   151
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   166
      TabIndex        =   9
      Top             =   225
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenLevel 
         Caption         =   "&Open Level..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSaveLevelAs 
         Caption         =   "&Save Level As..."
         Shortcut        =   ^S
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewBrick 
         Caption         =   "&New Brick"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuNewBoard 
         Caption         =   "New &Board"
         Shortcut        =   ^B
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHowToEdit 
         Caption         =   "How to edit..."
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
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/// Klotski Level Editor by Min Thant Sin, February 2005
'/// Feel free to e-mail me any bugs found in this program.
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub cmdNewBrick_Click()
    mnuNewBrick_Click
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyN Then
        mnuNewBrick_Click
    End If
End Sub

Private Sub Form_Resize()
    Call ReDimensionBoard
    Call DisplayBoard
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmBoardDim
End Sub

Private Sub mnuAbout_Click()
    MsgBox "Create by Min Thant Sin on February 2005", vbInformation, "About Klotski Level Editor..."
End Sub

Private Sub mnuExit_Click()
    Call CleanItUp
    Unload Me
    Unload frmBoardDim
End Sub

Private Sub mnuHowToEdit_Click()
    frmHelp.Show vbModal
End Sub

Private Sub optBrick_Click(Index As Integer)
    CurrentBrickType = Index
    Call DisplayInfo
End Sub

Private Sub mnuNewBoard_Click()
    frmBoardDim.Show vbModal
End Sub

Private Sub mnuNewBrick_Click()
    'Indicate that the first brick of the new group has NOT been placed.
    boolFirstBrickPlaced = False
    
    'Indicate that the brick can be placed right now.
    boolCanPlaceBrick = True
End Sub

Private Sub mnuOpenLevel_Click()
       With CommonDialog1
        .FileName = ""
        .Filter = "Klotski Game File (*.ksk)|*.*"
        .Flags = cdlOFNPathMustExist Or cdlOFNFileMustExist
        .ShowOpen
      End With
      
      If CommonDialog1.FileName = "" Then Exit Sub
      Call LoadLevel(CommonDialog1.FileName)
      Call DisplayInfo
End Sub

Private Sub mnuSaveLevelAs_Click()
    With CommonDialog1
        .FileName = ""
        .Filter = "Klotski Game File (*.ksk)|*.ksk"
        .Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt
        .ShowSave
    End With
      
    If CommonDialog1.FileName = "" Then Exit Sub
      
    Call SaveLevel(CommonDialog1.FileName)
End Sub

Private Sub DoOneTimeInitialization()
    On Local Error GoTo ErrorHandler
    
    Dim Backgroundcolor As Long
    Dim FrameRect As RECT
    
    Call GetWindowRect(fraBricks.hwnd, FrameRect)
    FrameWidth = (FrameRect.Right - FrameRect.Left)
            
    Backgroundcolor = RGB(127, 127, 127)
    
    ColorTable(FRAME_BRICK) = RGB(0, 0, 200)
    ColorTable(MASTER_BRICK) = RGB(160, 0, 0)
    ColorTable(NORMAL_BRICK) = RGB(128, 128, 0)
    ColorTable(BARRIER_BRICK) = RGB(0, 200, 200)
    ColorTable(DEST_SQUARE) = RGB(160, 0, 0)
    ColorTable(EMPTY_GRID) = Backgroundcolor
    
    'We're going to work in pixels
    picBoard.ScaleMode = vbPixels
    picBoard.BackColor = Backgroundcolor
    
    picBackBuffer.ScaleMode = vbPixels
    picBackBuffer.BackColor = Backgroundcolor
    picBackBuffer.Visible = False
    
    picBlank.ScaleMode = vbPixels
    picBlank.BackColor = Backgroundcolor
    picBlank.Visible = False
    
    'No bricks and no groups yet.
    NumBricks = 0
    NumGroups = 0
    
    'Initialize board dimensions
    BoardDimX = (MAX_BOARD_DIM - MIN_BOARD_DIM) \ 2
    BoardDimY = (MAX_BOARD_DIM - MIN_BOARD_DIM) \ 2
        
    Exit Sub
ErrorHandler:
    MsgBox "Unexpected and unwanted error has occurred!", vbInformation, Err.Description
        
End Sub

Private Sub Form_Load()
    On Local Error GoTo ErrorHandler
    
    Call DoOneTimeInitialization
    Call InitializeBoard
    Call ReDimensionBoard
    Call DisplayBoard
    Call DisplayInfo
    
    Exit Sub
ErrorHandler:
    MsgBox "Unexpected and unwanted error has occurred!", vbInformation, Err.Description
    
End Sub

Private Sub picBoard_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    'This means that when empty-square is selected (when CurrentBrickType = EMPTY_GRID),
    'the user does not have to click "New Brick" to place a new brick on the game board.
    If CurrentBrickType <> EMPTY_GRID Then
        'boolCanPlaceBrick is set to True when the user clicks "New Brick" menu or button.
        If boolCanPlaceBrick = False Then
            '*******
            Exit Sub
            '*******
        End If
    End If
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    On Local Error GoTo ErrorHandler
    
    Dim i As Integer, j As Integer
    Dim col As Integer, row As Integer
        
    Dim lb As New CBrick, rb As New CBrick   'Left brick and right brick
    Dim tb As New CBrick, bb As New CBrick  'Top brick and bottom brick
    
    Dim lbType As Integer, lbGID As Integer
    Dim tbType As Integer, tbGID As Integer
    Dim rbType As Integer, rbGID As Integer
    Dim bbType As Integer, bbGID As Integer
    
    Dim boolCreateNewBrick As Boolean
    
    Dim tmpBrick As New CBrick  'Used to store a brick temporarily
    Dim NewBrick As New CBrick  'Used to create new brick and add it to the Lists()
            
    Call GetCoordFromCursorPos(x, y, col, row)
    
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '/// Removing a block from a group
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    'User has selected Empty Square brick, this is used to clear an existing brick from the board
    If CurrentBrickType = EMPTY_GRID Then
        'No point in clearing an already-empty grid, so exit sub.
        If Board(col, row).BrickType = EMPTY_GRID Then
            '*******
            Exit Sub
            '*******
        End If
        
        'Now we've reached this point, we're going to
        'remove the brick from the grid under the cursor position.
        'So far, we have not yet checked what kind of brick we are removing.
        'What we know is that the grid under the cursor position is not empty,
        'in other words, there's some kind of brick in the grid, and we need to clear it.
        
        '****************************************************************************************************
        '**** ALGORITHM
        '****************************************************************************************************
        '(1) Get the brick's group ID from the game board.
        '(2) Find the brick's in the list and store its position.
        '(3) Make a copy of that going-to-be-removed brick.
        '(4) Clear it from the game board graphically.
        '(5) Remove it from the current list's data.
        '(6) Update game board data appropriately.
        '(7) Update graphical display of the current group's bricks.
        '(8) Remove the current group if there is no brick left in it, and
        '     update the data structures accordingly.
        '****************************************************************************************************
        
        'BrickPos stores the position of the brick to be removed from the list
        Dim BrickPos As Integer
                        
        'Get current group ID from the board
        CurrentGID = Board(col, row).GID
        
        'Find the brick in the "main list" and store its position.
        Lists(CurrentGID).MoveFirst     'Move to the beginning of the list
        For i = 1 To Lists(CurrentGID).NumBricks
        
            Set tmpBrick = Lists(CurrentGID).CurrentBrick
            'Found it?
            If tmpBrick.XCoord = col And tmpBrick.YCoord = row Then
                'Store its position in the list
                BrickPos = i
                Exit For
            End If

            Lists(CurrentGID).MoveNext  'Go to next brick in the list
        Next i
        
        'Get a copy of this brick.
        Set tmpBrick = Lists(CurrentGID).GetBrick(BrickPos)
        
        'Clear the brick from the board graphically
        Call ClearBrick(tmpBrick)
        
        'Remove the brick from the list's data structure
        Lists(CurrentGID).Removebrick BrickPos
        
        '****************************************************************************************************
        '*** Removing dest square
        '****************************************************************************************************
        'If the brick being removed is dest square, there can be no brick beneath it,
        'so we mark the grid as an empty square and modify its data appropriately
        If tmpBrick.BrickType = DEST_SQUARE Then
        
            'Remove from DestList
            For i = 1 To DestList.Count
                Set tmpBrick = DestList.Item(i)
                If tmpBrick.XCoord = col And tmpBrick.YCoord = row Then
                    DestList.Remove i
                    Exit For
                End If
            Next i
        
            'Indicate there's no brick in this grid now
            Board(tmpBrick.XCoord, tmpBrick.YCoord).GID = -1
            
            'Indicate there's no dest square in this grid now
            Board(tmpBrick.XCoord, tmpBrick.YCoord).DestGID = -1
            
            'Flag this grid as an empty square
            Board(tmpBrick.XCoord, tmpBrick.YCoord).BrickType = EMPTY_GRID
            Set Board(tmpBrick.XCoord, tmpBrick.YCoord).Brick = Nothing
        '****************************************************************************************************
        
        Else
                
            '****************************************************************************************************
            '*** Retrieving the obscured dest square
            '****************************************************************************************************
            'DestGID > 0 means there's a dest square obscured by the
            'current brick which is being removed. We need to get it back.
            If Board(tmpBrick.XCoord, tmpBrick.YCoord).DestGID > 0 Then
                'Restore the obscured dest square's GID into the game board
                Board(tmpBrick.XCoord, tmpBrick.YCoord).GID = Board(tmpBrick.XCoord, tmpBrick.YCoord).DestGID
                
                'This grid now contains dest square
                Board(tmpBrick.XCoord, tmpBrick.YCoord).BrickType = DEST_SQUARE
                Set Board(tmpBrick.XCoord, tmpBrick.YCoord).Brick = Nothing
                
                'Redraw dest square that was beneath a brick
                Call DrawDestination(picBackBuffer, tmpBrick.XCoord, tmpBrick.YCoord)
            '****************************************************************************************************
            
            Else
                
                '****************************************************************************************************
                '*** Removing other bricks
                '****************************************************************************************************
                Board(tmpBrick.XCoord, tmpBrick.YCoord).GID = -1
                Board(tmpBrick.XCoord, tmpBrick.YCoord).BrickType = EMPTY_GRID
                Set Board(tmpBrick.XCoord, tmpBrick.YCoord).Brick = Nothing
                '****************************************************************************************************
                
            End If
        End If
        
        'Get four adjacent bricks (left, right, top, bottom) and clear them
        'if they are in the same group as the brick being removed.
        'This is for graphical esthetics only.
        If tmpBrick.BrickType <> DEST_SQUARE Then
            Set lb = GetBrick(tmpBrick.XCoord - 1, tmpBrick.YCoord)
            Set rb = GetBrick(tmpBrick.XCoord + 1, tmpBrick.YCoord)
            Set tb = GetBrick(tmpBrick.XCoord, tmpBrick.YCoord - 1)
            Set bb = GetBrick(tmpBrick.XCoord, tmpBrick.YCoord + 1)
            
            'Check for group ID
            If lb.GID = tmpBrick.GID Then Call ClearBrick(lb)
            If rb.GID = tmpBrick.GID Then Call ClearBrick(rb)
            If tb.GID = tmpBrick.GID Then Call ClearBrick(tb)
            If bb.GID = tmpBrick.GID Then Call ClearBrick(bb)
            
            'Redraw the current group of bricks
            Call DisplayGroup(CurrentGID)
            Call OutlineGroup(CurrentGID)
        End If
        
        'A brick has been removed, so decrement the brick count
        NumBricks = NumBricks - 1
        
        'Check if the current group's list has any bricks in it.
        'If there isn't any bricks left in this list, we have to remove it from
        'the Lists() array and shift all the lists down which are above it.
        If Lists(CurrentGID).NumBricks = 0 Then
            'Decrement the group id of each brick
            'starting from the current list's next list
            For i = (CurrentGID + 1) To NumGroups
            
                Lists(i).MoveFirst
                For j = 1 To Lists(i).NumBricks
                    Set tmpBrick = Lists(i).CurrentBrick
                    
                    tmpBrick.GID = tmpBrick.GID - 1
                    'Update game board data too
                    Board(tmpBrick.XCoord, tmpBrick.YCoord).GID = tmpBrick.GID
                    
                    If tmpBrick.BrickType = DEST_SQUARE Then
                        Board(tmpBrick.XCoord, tmpBrick.YCoord).DestGID = tmpBrick.GID
                    End If
                    
                    Lists(i).MoveNext
                Next j
            Next i
            
            'Bubble down
            For i = CurrentGID To (NumGroups - 1)
                Set Lists(i) = Lists(i + 1)
            Next i
                            
            'A group has been removed
            NumGroups = NumGroups - 1
            
            If NumGroups > 0 Then
                ReDim Preserve Lists(1 To NumGroups)
            End If
        End If
        
        Call DisplayInfo
        
        picBoard_Paint
        
        '*******
        Exit Sub
        '*******
    End If
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    
    '*************************************************************************
    'The scenario in which the user selects an empty square to
    'clear a brick from the game board has been handled above.
    
    'The following code deals with the process of creating new brick and
    'placing it on the game board as well as into the data structure.
    '*************************************************************************
        
    If CurrentBrickType = DEST_SQUARE Then
        If Board(col, row).BrickType = DEST_SQUARE Then
            CurrentGID = Board(col, row).GID
            boolFirstBrickPlaced = True
            Call DisplayInfo
            '*******
            Exit Sub
            '*******
        Else
            If Board(col, row).BrickType <> EMPTY_GRID Then
                '*******
                Exit Sub
                '*******
            End If
        End If
        
    Else
            
        'The user can place a new brick only if the grid is empty OR there is a dest square in it
        If Board(col, row).BrickType <> EMPTY_GRID And _
           Board(col, row).BrickType <> DEST_SQUARE Then
           'This generally means selecting an existing group to add new brick to it.
            If CurrentBrickType = Board(col, row).BrickType Then
                'Store the group ID so we can use it later to add new brick to the current group.
                CurrentGID = Board(col, row).GID
                'As the user is selecting a group, there is already some brick(s) placed.
                boolFirstBrickPlaced = True
                Call DisplayInfo
            End If
            
            'Just selecting a group, don't do anything right now.
            Exit Sub
        End If
    End If
    
    
    Call GetBrickInfo(col - 1, row, lbType, lbGID)        'Left block
    Call GetBrickInfo(col + 1, row, rbType, rbGID)      'Right block
    Call GetBrickInfo(col, row - 1, tbType, tbGID)        'Top block
    Call GetBrickInfo(col, row + 1, bbType, bbGID)      'Bottom block
                    
    'This is the first brick placed in a new group.
    If boolFirstBrickPlaced = False Then    'This means this is a NEW group.
        NumBricks = NumBricks + 1
        NumGroups = NumGroups + 1
        CurrentGID = NumGroups
        
        ReDim Preserve Lists(1 To NumGroups)
        
        NewBrick.XCoord = col
        NewBrick.YCoord = row
        NewBrick.BrickType = CurrentBrickType
        NewBrick.GID = NumGroups
        
        Lists(CurrentGID).AddBrick NewBrick
        
        'Dest square is erased when the user places a brick on it.
        'We store dest square's group ID so that later we can redraw it
        'when the brick above it has been removed.
        If Board(col, row).BrickType = DEST_SQUARE Then
            Board(col, row).DestGID = Board(col, row).GID
        End If
        
        Board(col, row).BrickType = CurrentBrickType
        Board(col, row).GID = NumGroups
        Set Board(col, row).Brick = NewBrick
        
        'Just adding dest squares to the collection DestList.
        'This is to check for "puzzle solved" situation.
        If CurrentBrickType = DEST_SQUARE Then
            DestList.Add NewBrick
            Call DrawDestination(picBackBuffer, NewBrick.XCoord, NewBrick.YCoord)
        Else
            Call DrawBrick(picBackBuffer, NewBrick)
            Call OutlineGroup(NewBrick.GID)
        End If
        
        Set NewBrick = Nothing
         
        boolFirstBrickPlaced = True
        picBoard_Paint
        
        Call DisplayInfo
        '***************
        Exit Sub
        '***************
    End If
    
    
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////
    'Adding a new brick to an EXISTING group.
    boolCreateNewBrick = False
    
    'Check left brick
    If lbType = CurrentBrickType Then
        If lbGID = CurrentGID Then
            boolCreateNewBrick = True
        End If
    End If
        
    'Check top brick
    If tbType = CurrentBrickType Then
        If tbGID = CurrentGID Then
            boolCreateNewBrick = True
        End If
    End If
    
    'Check right brick
    If rbType = CurrentBrickType Then
        If rbGID = CurrentGID Then
            boolCreateNewBrick = True
        End If
    End If
        
    'Check bottom brick
    If bbType = CurrentBrickType Then
        If bbGID = CurrentGID Then
            boolCreateNewBrick = True
        End If
    End If
            
    If boolCreateNewBrick Then
        NumBricks = NumBricks + 1
            
        Set NewBrick = New CBrick
        
        NewBrick.XCoord = col
        NewBrick.YCoord = row
        NewBrick.BrickType = CurrentBrickType
        NewBrick.GID = CurrentGID
        
        Lists(CurrentGID).AddBrick NewBrick
        
        'Dest square is erased when the user places a brick on it.
        'We store dest square's group ID so that later we can redraw it
        'when the brick above it has been removed.
        If Board(col, row).BrickType = DEST_SQUARE Then
            Board(col, row).DestGID = Board(col, row).GID
        End If
        
        Board(col, row).BrickType = CurrentBrickType
        Board(col, row).GID = CurrentGID
        Set Board(col, row).Brick = NewBrick
        
        'Used to check if the puzzle has been solved.
        If CurrentBrickType = DEST_SQUARE Then
            DestList.Add NewBrick
            Call DrawDestination(picBackBuffer, NewBrick.XCoord, NewBrick.YCoord)
        Else
            
            Call DisplayGroup(CurrentGID)
            Call OutlineGroup(CurrentGID)
        End If
        
    End If
    
    Set NewBrick = Nothing
    Set tmpBrick = Nothing
        
    Call DisplayInfo
    
    picBoard_Paint
    
    Exit Sub
ErrorHandler:
    MsgBox "Unexpected and unwanted error has occurred!", vbInformation, Err.Description
    
End Sub

'Draw a group's bricks to the game board
Sub DisplayGroup(ByVal GID As Integer)
    On Local Error GoTo ErrorHandler
    
    Dim i As Integer
    Dim tmpBrick As CBrick
    
    Lists(GID).MoveFirst
    For i = 1 To Lists(GID).NumBricks
        Set tmpBrick = Lists(GID).CurrentBrick
        DrawBrick picBackBuffer, tmpBrick
        Lists(GID).MoveNext
    Next i
    
    Set tmpBrick = Nothing
    
    Exit Sub
ErrorHandler:
    MsgBox "Unexpected and unwanted error has occurred!", vbInformation, Err.Description
    
End Sub

Public Sub picBoard_Paint()
    'Blit the contents of picBackBuffer to picBoard
    Call BitBlt(picBoard.hdc, 0, 0, BoardWidth, BoardHeight, picBackBuffer.hdc, 0, 0, vbSrcCopy)
End Sub

Sub ClearBrick(ByRef b As CBrick)
    If (b Is Nothing) Then Exit Sub
    
    Dim bx As Integer, by As Integer
    
    bx = b.XCoord * GridWidth   'Brick's left
    by = b.YCoord * GridHeight  'Brick's top
    
    'Clear it graphically
    BitBlt picBackBuffer.hdc, bx, by, GridWidth, GridHeight, _
            picBlank.hdc, 0, 0, vbSrcCopy
                    
    'Redraw the grid that got erased when clearing a brick.
    picBackBuffer.ForeColor = RGB(200, 200, 200)
    MoveToEx picBackBuffer.hdc, bx, by, ByVal 0
    
    LineTo picBackBuffer.hdc, (bx + GridWidth), by                       'Top line
    LineTo picBackBuffer.hdc, (bx + GridWidth), (by + GridHeight)   'Right line
    LineTo picBackBuffer.hdc, bx, (by + GridHeight)                      'Bottom line
    LineTo picBackBuffer.hdc, bx, by                                          'Left line
End Sub
