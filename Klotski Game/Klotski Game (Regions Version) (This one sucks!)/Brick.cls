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

'Cell's data
Public XCoord As Integer    'X coordinate (cell's column)
Public YCoord As Integer    'Y coordinate (cell's row)
Public BrickType As Integer
Public GroupID As Integer

Public Left As Integer
Public Top As Integer
Public Width As Integer
Public Height As Integer
Public Locked As Boolean

Public PrevBrick As CBrick     'Pointer to previous cell in the linked list
Public NextBrick As CBrick     'Pointer to next cell in the linked list
