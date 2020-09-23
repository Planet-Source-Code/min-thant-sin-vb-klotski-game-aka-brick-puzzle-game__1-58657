VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help is on the way, dude..."
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Okay, thanks!!"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   75
      TabIndex        =   18
      Top             =   8475
      Width           =   8265
   End
   Begin VB.Frame Frame4 
      Caption         =   "If the above information does NOT help:"
      ForeColor       =   &H000000C0&
      Height          =   990
      Left            =   75
      TabIndex        =   16
      Top             =   7350
      Width           =   8490
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Play around with it and you'll get it!!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   525
         TabIndex        =   17
         Top             =   300
         Width           =   7395
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Placing a New Brick to a New Group:"
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   75
      TabIndex        =   11
      Top             =   5400
      Width           =   8490
      Begin VB.Label Label12 
         Caption         =   "(1) Select the brick type that you want to place."
         Height          =   315
         Left            =   225
         TabIndex        =   14
         Top             =   375
         Width           =   7215
      End
      Begin VB.Label Label8 
         Caption         =   "(2) Click ""New Brick"" button or ""New Brick"" menu (shortcut keys = 'N' key or F2)."
         Height          =   315
         Left            =   225
         TabIndex        =   13
         Top             =   750
         Width           =   8115
      End
      Begin VB.Label Label1 
         Caption         =   "(3) Place the new brick on any grids that are either EMPTY or occupied by dest       squares."
         Height          =   540
         Left            =   225
         TabIndex        =   12
         Top             =   1125
         Width           =   7740
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Placing a New Brick to an Existing Group:"
      ForeColor       =   &H00C00000&
      Height          =   2190
      Left            =   75
      TabIndex        =   7
      Top             =   3075
      Width           =   8490
      Begin VB.Label Label13 
         Caption         =   "(4) Repeat step (3) until you want to add a new brick to a NEW group."
         Height          =   315
         Left            =   225
         TabIndex        =   15
         Top             =   1725
         Width           =   7140
      End
      Begin VB.Label Label11 
         Caption         =   "(3) Place the new brick on any grids that are horizontally or vertically adjacent       to the group."
         Height          =   540
         Left            =   225
         TabIndex        =   10
         Top             =   1125
         Width           =   7740
      End
      Begin VB.Label Label10 
         Caption         =   "(2) Click on the group."
         Height          =   315
         Left            =   225
         TabIndex        =   9
         Top             =   750
         Width           =   6390
      End
      Begin VB.Label Label9 
         Caption         =   "(1) Select the same brick type as the group."
         Height          =   315
         Left            =   225
         TabIndex        =   8
         Top             =   375
         Width           =   7215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Six Kinds of Objects You Can Place on the Game Board:"
      ForeColor       =   &H00C00000&
      Height          =   2865
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   8490
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Frame Brick - Used to draw bricks that are NOT movable."
         Height          =   240
         Left            =   225
         TabIndex        =   6
         Top             =   375
         Width           =   5460
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Master Brick - This is the brick the user has to maneuver to its destination."
         Height          =   240
         Left            =   225
         TabIndex        =   5
         Top             =   750
         Width           =   7305
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Normal Brick - This is ordinary movable brick which serves as obstacle."
         Height          =   240
         Left            =   225
         TabIndex        =   4
         Top             =   1125
         Width           =   6825
      End
      Begin VB.Label Label5 
         Caption         =   "Barrier Brick - This brick is NOT movable, but removable when touched completely                          by Master Brick."
         Height          =   540
         Left            =   225
         TabIndex        =   3
         Top             =   1500
         Width           =   7950
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Dest Square - Used to mark the Master Brick's destination."
         Height          =   240
         Left            =   225
         TabIndex        =   2
         Top             =   2100
         Width           =   5730
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Empty Grid - This is used to remove a brick from the game board."
         Height          =   240
         Left            =   225
         TabIndex        =   1
         Top             =   2475
         Width           =   6300
      End
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

