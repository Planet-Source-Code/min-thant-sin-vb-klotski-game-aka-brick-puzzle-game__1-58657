VERSION 5.00
Begin VB.Form frmHowToPlay 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help is on the way..."
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
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
   ScaleHeight     =   6195
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "The Object of the Game:"
      ForeColor       =   &H00C00000&
      Height          =   1365
      Left            =   150
      TabIndex        =   8
      Top             =   75
      Width           =   7740
      Begin VB.Label Label1 
         Caption         =   $"frmHowToPlay.frx":0000
         Height          =   765
         Left            =   225
         TabIndex        =   9
         Top             =   375
         Width           =   7215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Five Kinds of Objects You May Find on the Game Board:"
      ForeColor       =   &H00C00000&
      Height          =   2490
      Left            =   150
      TabIndex        =   3
      Top             =   1650
      Width           =   7740
      Begin VB.Label Label5 
         Caption         =   "Barrier Brick - Although non-movable, the brick is removable when touched                          completely by Master Brick."
         Height          =   540
         Left            =   225
         TabIndex        =   10
         Top             =   1500
         Width           =   7425
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Dest Square - Indicate the Master Brick's destination.."
         Height          =   240
         Left            =   225
         TabIndex        =   7
         Top             =   2100
         Width           =   5310
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Normal Brick - This is ordinary movable brick which serves as obstacle."
         Height          =   240
         Left            =   225
         TabIndex        =   6
         Top             =   1125
         Width           =   6825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Master Brick - You have to maneuver this brick to its destination."
         Height          =   240
         Left            =   225
         TabIndex        =   5
         Top             =   750
         Width           =   6375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Frame Brick - This is non-movable brick."
         Height          =   240
         Left            =   225
         TabIndex        =   4
         Top             =   375
         Width           =   3810
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "If the above information does NOT help:"
      ForeColor       =   &H000000C0&
      Height          =   990
      Left            =   150
      TabIndex        =   1
      Top             =   4350
      Width           =   7740
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Have a go at it and find out!!"
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
         Left            =   1050
         TabIndex        =   2
         Top             =   375
         Width           =   5790
      End
   End
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
      Left            =   150
      TabIndex        =   0
      Top             =   5550
      Width           =   7740
   End
End
Attribute VB_Name = "frmHowToPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub
