VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBrickList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ************************************************
' Copyright (C) 1997 John Wiley & Sons, Inc.
' All rights reserved. See additional copyright
' information in RIGHTS.TXT.
' ************************************************

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/// Based on part of the source code from the book:
'/// [ Ready-to-Run Visual Basic(R) Algorithms ] Second Edition by Rod Stephens
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Private m_NumBricks As Long                  'Number of bricks in the list
Private m_CurrentBrick As New CBrick      'Pointer to current brick

Private m_TopSentinel As New CBrick         'Top sentinel of the list
Private m_BottomSentinel As New CBrick    'Bottom sentinel of the list

Private Sub Class_Initialize()
    DoInitialization
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Add a new brick to the list, and sort the bricks according to their XCoord and YCoord values.
'The brick with the smallest XCoord AND smallest YCoord gets the front seats in the list.
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub AddBrick(ByVal NewBrick As CBrick)
    On Error GoTo ErrorHandler
    
    Dim BeforeBrick As CBrick  'Pointer to the brick BEFORE the new brick
    Dim AfterBrick As CBrick    'Pointer to the brick AFTER the new brick
    Dim ptr As CBrick
    
    'ptr will point to the first brick in the list (if there are some bricks already in the list), OR
    'ptr will point to BottomSentinel (if this is the first time the new brick is added)
    Set ptr = m_TopSentinel.NextBrick
    
    'Loop to the end of list
    Do Until (ptr Is m_BottomSentinel)
        
        'Find the brick with the XCoord greater than that of new brick.
        If NewBrick.XCoord < ptr.XCoord Then
            '(If the new brick's XCoord is greater than those of all the existing bricks,
            'ptr will finally reach BottomSentinel, the loop ends, and the new brick
            'is added at the end of the list just before BottomSentinel)
            Exit Do
            
        Else
            'Same XCoord means there are two or more objects in the same column
            If NewBrick.XCoord = ptr.XCoord Then
                'XCoords are the same, compare the YCoords.
                'The smaller the new brick's YCoord, the greater
                'chance it has getting the front seat in the list.
                'Note the use of '<='
                If NewBrick.YCoord <= ptr.YCoord Then
                    Exit Do
                End If
            End If
        End If
        
        'Check next brick in the list
        Set ptr = ptr.NextBrick
    Loop
    
    'Add the new brick and set the pointers appropriately
    Set BeforeBrick = ptr.PrevBrick
    
    Set BeforeBrick.NextBrick = NewBrick
    Set NewBrick.NextBrick = ptr
    Set ptr.PrevBrick = NewBrick
    Set NewBrick.PrevBrick = BeforeBrick
    
    'Increment the brick counter
    m_NumBricks = m_NumBricks + 1
    
    Exit Sub
ErrorHandler:
    MsgBox "Unexpected and unwanted error has occurred!", vbInformation, Err.Description
    
End Sub

Public Sub Removebrick(ByVal pos As Long)
    On Error GoTo ErrorHandler
    
    If pos < 1 Or pos > m_NumBricks Then Exit Sub
    
    Dim ptr As CBrick
    Dim BeforeBrick As CBrick
    Dim AfterBrick As CBrick
    
    Set ptr = m_TopSentinel
    
    Do Until (pos <= 0)
        Set ptr = ptr.NextBrick
        pos = pos - 1
    Loop
    
    Set BeforeBrick = ptr.PrevBrick
    Set AfterBrick = ptr.NextBrick
    
    Set BeforeBrick.NextBrick = AfterBrick
    Set AfterBrick.PrevBrick = BeforeBrick
    
    'Decrement the brick counter
    m_NumBricks = m_NumBricks - 1
    
    Exit Sub
ErrorHandler:
    MsgBox "Unexpected and unwanted error has occurred!", vbInformation, Err.Description
    
End Sub

Function GetBrick(ByVal pos As Long) As CBrick
    On Error GoTo ErrorHandler
    
    If pos < 1 Or pos > m_NumBricks Then Exit Function
    
    Dim ptr As CBrick
    
    Set ptr = m_TopSentinel
    Do Until (pos <= 0)
        Set ptr = ptr.NextBrick
        pos = pos - 1
    Loop
    
    Set GetBrick = ptr
    
    Exit Function
ErrorHandler:
    MsgBox "Unexpected and unwanted error has occurred!", vbInformation, Err.Description
    
End Function

Public Sub MoveFirst()
    Set m_CurrentBrick = m_TopSentinel.NextBrick
End Sub

Public Sub MoveNext()
    If Not (m_CurrentBrick Is Nothing) Then
        Set m_CurrentBrick = m_CurrentBrick.NextBrick
    End If
End Sub

Public Function EndOfList() As Boolean
    EndOfList = CBool(m_CurrentBrick Is m_BottomSentinel)
End Function

Public Function CurrentBrick() As CBrick
    Set CurrentBrick = m_CurrentBrick
End Function

Public Function NumBricks() As Long
    NumBricks = m_NumBricks
End Function

Public Sub ClearList()
    CleanUp
    DoInitialization
End Sub

Private Sub CleanUp()
    On Error GoTo ErrorHandler
    
    Dim ptr As CBrick
    
    Set ptr = m_TopSentinel.NextBrick
    
    Do Until (ptr Is m_BottomSentinel)
        Set ptr.PrevBrick = Nothing
        Set ptr = ptr.NextBrick
    Loop
    
    Set m_TopSentinel.NextBrick = Nothing
    Set m_BottomSentinel.PrevBrick = Nothing
    
    Exit Sub
ErrorHandler:
    MsgBox "Unexpected and unwanted error has occurred!", vbInformation, Err.Description
    
End Sub

Sub DoInitialization()
    'No bricks yet
    m_NumBricks = 0
    
    'Initialize top and bottom sentinels
    Set m_TopSentinel.PrevBrick = Nothing
    Set m_TopSentinel.NextBrick = m_BottomSentinel
    
    Set m_BottomSentinel.PrevBrick = m_TopSentinel
    Set m_BottomSentinel.NextBrick = Nothing
End Sub

Private Sub Class_Terminate()
    CleanUp
End Sub
