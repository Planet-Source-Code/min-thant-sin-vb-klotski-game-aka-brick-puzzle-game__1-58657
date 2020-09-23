VERSION 5.00
Begin VB.Form frmLevels 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Game Levels"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   75
      TabIndex        =   3
      Top             =   75
      Width           =   4290
   End
   Begin VB.DirListBox Dir1 
      Height          =   5490
      Left            =   75
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   450
      Width           =   4290
   End
   Begin VB.FileListBox File1 
      Height          =   5160
      Left            =   4425
      Pattern         =   "*.bxw"
      TabIndex        =   0
      Top             =   75
      Width           =   4440
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7200
      TabIndex        =   2
      Top             =   5400
      Width           =   1590
   End
   Begin VB.CommandButton cmdLoadLevel 
      Caption         =   "&Load Level"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4500
      TabIndex        =   1
      Top             =   5400
      Width           =   2415
   End
End
Attribute VB_Name = "frmLevels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
      Me.Hide
End Sub

Private Sub cmdLoadLevel_Click()
    On Local Error GoTo ErrorHandler
    
    Dim FileName As String
    
    FileName = AddASlash(File1.Path) & File1.FileName
    
    If Trim(File1.FileName) = "" Then Exit Sub
    
    frmKlotski.picBoard.Visible = False
    CurrentLevel = File1.ListIndex
    
    Call LoadLevel(FileName)
    Call SaveLevelStatus
    
    frmKlotski.picBoard.Visible = True
    Me.Hide
    
    Exit Sub
ErrorHandler:
    MsgBox "cmdLoadLevel_Click() error", vbInformation, Err.Description
    
End Sub

Private Sub Drive1_Change()
      On Error GoTo ErrorHandler
      Dir1.Path = Drive1.Drive
      Exit Sub
ErrorHandler:
      On Error Resume Next
      Drive1.Drive = Dir1.Path
End Sub

Private Sub Dir1_Change()
      On Error GoTo ErrorHandler
      File1.Path = Dir1.Path
      
      Exit Sub
ErrorHandler:
      On Error Resume Next
      Dir1.Path = File1.Path
End Sub

Private Sub Form_Load()
    File1.Pattern = "*.ksk"
End Sub
