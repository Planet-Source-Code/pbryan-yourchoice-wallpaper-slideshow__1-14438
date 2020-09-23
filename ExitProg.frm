VERSION 5.00
Begin VB.Form ExitProg 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WinCryptic Wallpaper Slidshow"
   ClientHeight    =   1620
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   7320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&No Nevermind"
      Height          =   372
      Left            =   4080
      TabIndex        =   2
      Top             =   960
      Width           =   1932
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Yes Exit"
      Height          =   372
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   1932
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Are You Sure you want to Exit WallSlider?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   492
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   7332
   End
End
Attribute VB_Name = "ExitProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    End
End Sub
    
Private Sub Command2_Click()

Unload Me
End Sub
