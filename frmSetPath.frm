VERSION 5.00
Begin VB.Form frmSetPath 
   Caption         =   "Choose Folder for Slideshow"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6120
   Icon            =   "frmSetPath.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   2970
      Left            =   3600
      Pattern         =   "*.bmp;*.jpg"
      TabIndex        =   4
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Images Located Here:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmSetPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    frmMain.filWallPaper.path = Dir1.path
    Unload Me
End Sub

Private Sub Dir1_Change()
    File1.path = Dir1.path
End Sub

Private Sub Drive1_Change()
    Dir1.path = Drive1.Drive
End Sub



