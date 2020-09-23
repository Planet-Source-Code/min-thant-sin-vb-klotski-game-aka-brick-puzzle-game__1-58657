VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBrick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Brick's data
Public XCoord As Integer    'Brick's column  [ 2 chars ]
Public YCoord As Integer    'Brick's row [ 2 chars ]
Public BrickType As Integer '[ 2 chars ]
Public GID As Integer         '[ 3 chars ]

'These don't get saved to the game file
Public Left As Integer
Public Top As Integer
Public Width As Integer
Public Height As Integer

Public PrevBrick As CBrick     'Pointer to previous brick in the linked list
Public NextBrick As CBrick     'Pointer to next brick in the linked list
