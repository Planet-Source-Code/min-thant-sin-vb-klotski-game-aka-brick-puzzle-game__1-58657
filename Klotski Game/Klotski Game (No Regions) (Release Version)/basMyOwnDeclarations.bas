Attribute VB_Name = "basMyOwnDeclarations"
Option Explicit

Public Const MOUSE_SENSITIVITY As Integer = 3

Public Const MIN_GRID_SIZE As Integer = 10      'Minimum grid size

Public Const BRICK_WIDTH_PERCENT As Single = 0.8    'BrickWidth based on GridWidth  (GridWidth * 0.8)
Public Const BRICK_HEIGHT_PERCENT As Single = 0.8   'BrickHeight based on GridHeight (GridHeight * 0.8)

Public Const NUM_BRICK_TYPES As Integer = 6     'Number of brick types

'Brick types
Public Const FRAME_BRICK As Integer = 0      'Non-movable
Public Const MASTER_BRICK As Integer = 1    'Movable, the user has to maneuver this brick to its destination marked in red squares
Public Const NORMAL_BRICK As Integer = 2    'Movable, used as obstacles
Public Const BARRIER_BRICK As Integer = 3    'Non-movable but removable if touched by Master Brick completely (wholly)
Public Const DEST_SQUARE As Integer = 4      'Marks the Master Brick's destination
Public Const EMPTY_SQUARE As Integer = 5    'Used to clear a brick on the game board

'Used to save current game level
Public Const MY_APP As String = "Min Thant Sin's Klotski Game"
Public Const MY_SECTION As String = "Game Levels Info"
Public Const MY_KEY As String = "Current Level"

Public Enum ENUM_DIRECTION
    LeftDirection = 0
    RightDirection = 2
    UpDirection = 3
    DownDirection = 4
End Enum

Public Type BOARD_INFO
    Brick As CBrick
    GroupID As Integer
    BrickType As Integer
    
    DestGID As Integer
    DestBrick As CBrick
End Type

Public NumBricks As Integer
Public NumGroups As Integer

Public CurrentBrickType As Integer

Public BoardWidth As Integer, BoardHeight As Integer    'Board's width and height in pixels
Public BoardDimX As Integer, BoardDimY As Integer       'The number of grids on the board horizontally and vertically
Public GridWidth As Integer, GridHeight As Integer         'A grid's width and height
Public BrickWidth As Integer, BrickHeight As Integer       'A brick's width and height
Public BrickThickness As Integer

Public NumGameFiles As Integer
Public CurrentLevel As Integer

Public boolGameStarted As Boolean
Public boolLoadSuccessful As Boolean

Public direction As ENUM_DIRECTION

Public CurrentList As New Collection
Public DestList As New Collection
Public BrickLists()  As New Collection

Public ColorTable(0 To NUM_BRICK_TYPES - 1) As Long
Public Board() As BOARD_INFO
