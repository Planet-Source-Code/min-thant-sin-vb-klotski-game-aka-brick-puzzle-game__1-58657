VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Klotski"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4020
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
   ScaleHeight     =   3945
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Okay"
      Default         =   -1  'True
      Height          =   465
      Left            =   1200
      TabIndex        =   0
      Top             =   3300
      Width           =   1590
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   225
      Picture         =   "frmAbout.frx":0000
      Top             =   1350
      Width           =   480
   End
   Begin VB.Label Label4 
      Caption         =   "(Recreated by Min Thant Sin )"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   675
      TabIndex        =   4
      Top             =   2550
      Width           =   2910
   End
   Begin VB.Label Label3 
      Caption         =   "Copyright ZH Computer, 1991"
      Height          =   240
      Left            =   750
      TabIndex        =   3
      Top             =   2175
      Width           =   2850
   End
   Begin VB.Label Label2 
      Caption         =   "Minneapolis - Warsaw"
      Height          =   240
      Left            =   1050
      TabIndex        =   2
      Top             =   1800
      Width           =   2130
   End
   Begin VB.Label Label1 
      Caption         =   "Klotski"
      Height          =   240
      Left            =   1725
      TabIndex        =   1
      Top             =   1425
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   75
      Picture         =   "frmAbout.frx":0C42
      Top             =   75
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub
