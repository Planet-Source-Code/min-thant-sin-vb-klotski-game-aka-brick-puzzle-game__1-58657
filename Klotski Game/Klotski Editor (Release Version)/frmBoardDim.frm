VERSION 5.00
Begin VB.Form frmBoardDim 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Board Dimensions"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1290
      Left            =   75
      TabIndex        =   2
      Top             =   1500
      Width           =   3990
      Begin VB.ComboBox cboBoardDimX 
         Height          =   390
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   675
         Width           =   1440
      End
      Begin VB.ComboBox cboBoardDimY 
         Height          =   390
         Left            =   2325
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   675
         Width           =   1440
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "X"
         Height          =   270
         Left            =   1875
         TabIndex        =   7
         Top             =   750
         Width           =   150
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "BoardDimX :"
         Height          =   270
         Left            =   150
         TabIndex        =   6
         Top             =   300
         Width           =   1380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "BoardDimY :"
         Height          =   270
         Left            =   2325
         TabIndex        =   5
         Top             =   300
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   465
      Left            =   2175
      TabIndex        =   1
      Top             =   2925
      Width           =   1890
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Okay"
      Default         =   -1  'True
      Height          =   465
      Left            =   75
      TabIndex        =   0
      Top             =   2925
      Width           =   1890
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BoardDimX is the number of grids horizontally on the board. BoardDimY is the number of grids vertically on the board."
      Height          =   1365
      Left            =   75
      TabIndex        =   8
      Top             =   75
      Width           =   3990
   End
End
Attribute VB_Name = "frmBoardDim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    On Local Error GoTo ErrorHandler
    
    BoardDimX = Val(cboBoardDimX.Text)
    BoardDimY = Val(cboBoardDimY.Text)
        
    NumGroups = 0
    NumBricks = 0
    
    Call CleanItUp
    Call ReDimensionBoard
    Call InitializeBoard
    Call DisplayBoard
    Call DisplayInfo
    
    Me.Hide
    
    Exit Sub
ErrorHandler:
    MsgBox "Unexpected and unwanted error has occurred!", vbInformation, Err.Description
    
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    For i = MIN_BOARD_DIM To MAX_BOARD_DIM Step 2
        cboBoardDimX.AddItem i
        cboBoardDimY.AddItem i
    Next i
    
    cboBoardDimX.ListIndex = 0
    cboBoardDimY.ListIndex = 0
    
End Sub
