VERSION 5.00
Begin VB.Form frmSolved 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Puzzle Solved!"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
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
   ScaleHeight     =   4350
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGiveMeABreak 
      Cancel          =   -1  'True
      Caption         =   "Give me a break!"
      Height          =   615
      Left            =   375
      TabIndex        =   5
      Top             =   3600
      Width           =   5640
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Go to the next level (if any)"
      Default         =   -1  'True
      Height          =   615
      Left            =   375
      TabIndex        =   3
      Top             =   2850
      Width           =   5640
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "You deserve a pat on the back."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   450
      TabIndex        =   4
      Top             =   2025
      Width           =   5505
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Congratulations!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   870
      Left            =   75
      TabIndex        =   2
      Top             =   75
      Width           =   6045
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "The puzzle has been solved!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   600
      TabIndex        =   1
      Top             =   1425
      Width           =   5115
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Congratulations!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   6045
   End
End
Attribute VB_Name = "frmSolved"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
          
    CurrentLevel = CurrentLevel + 1
    If CurrentLevel >= NumGameFiles Then
          CurrentLevel = 0
    End If
    
    If frmLevels.File1.ListCount > 0 Then
          frmLevels.File1.ListIndex = CurrentLevel
          Call LoadLevel(AddASlash(frmLevels.File1.Path) & frmLevels.File1.FileName)
          Call SaveLevelStatus
    End If
    
    Unload Me
End Sub

Private Sub cmdGiveMeABreak_Click()
    Unload Me
End Sub
