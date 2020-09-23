Attribute VB_Name = "basMyOwnDeclarations"
Option Explicit

Public Const NUM_BRICK_TYPES As Integer = 6     'The number of brick types
Public Const MIN_GRID_SIZE As Integer = 10        'Minimum grid size

Public Const MIN_BOARD_DIM As Integer = 10      'Minimum board dimension
Public Const MAX_BOARD_DIM As Integer = 50      'Maximum board dimension

'Brick types
Public Const FRAME_BRICK As Integer = 0       'Non-movable brick, usually representing a case or a border
Public Const MASTER_BRICK As Integer = 1     'Get this brick to its destination
Public Const NORMAL_BRICK As Integer = 2     'Normal movable brick
Public Const BARRIER_BRICK As Integer = 3     'Non-movable brick, BUT removable if touched completely by Master Brick
Public Const DEST_SQUARE As Integer = 4       'Used to mark the Master Brick's destination
Public Const EMPTY_GRID As Integer = 5     'Used to clear and remove a brick from the game board
'/////////////////////////////////////////////////////////////////////////////////////////////////////////

Public Const BRICK_WIDTH_PERCENT As Single = 0.8    'Used to calculate BrickWidth (based on grid width)
Public Const BRICK_HEIGHT_PERCENT As Single = 0.8   'Used to calculate BrickHeight (based on grid height)

'Self-explanatory
Public GridWidth As Integer, GridHeight As Integer
Public BrickWidth As Integer, BrickHeight As Integer
Public BrickThickness As Integer    'The height of horizontal shade and the width of vertical shade

Public BoardDimX As Integer, BoardDimY As Integer
Public BoardWidth As Integer, BoardHeight As Integer      'Board's width and height in pixels

Public Type BOARD_INFO
    Brick As CBrick
    GID As Integer           'Group ID
    BrickType As Integer   'What kind of brick?
    
    DestGID As Integer     'Destionation Square's group ID
End Type

Public NumBricks As Integer       'Number of bricks on the game board
Public NumGroups As Integer     'Number of brick-groups on the game board

Public CurrentGID As Integer
Public CurrentBrickType As Integer  'Currently selected brick type

Public FrameWidth As Integer   'The frame control which contains graphical OptionButtons

Public boolFirstBrickPlaced As Boolean
Public boolCanPlaceBrick As Boolean

Public DestList As New Collection

'Used to fill the surface of the bricks based on their types
Public ColorTable(0 To NUM_BRICK_TYPES - 1) As Long

Public Board() As BOARD_INFO    'Stores game board data
Public Lists() As New CBrickList    'Stores the brick objects
