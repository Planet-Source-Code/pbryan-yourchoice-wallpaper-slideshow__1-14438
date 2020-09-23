VERSION 5.00
Begin VB.Form Splash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Done"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Height          =   855
      Left            =   2880
      Shape           =   2  'Oval
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   3360
      Picture         =   "Splash.frx":0000
      Top             =   480
      Width           =   1050
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "By Paul Bryan Bensch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   -120
      TabIndex        =   5
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label Label5 
      Caption         =   "Version:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Application"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "3teq@softhome.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5160
      TabIndex        =   1
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "http://www.3teq.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   2520
      Width           =   2175
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FadeSpeed As Integer

Private Sub Command1_Click()
    Call BeTop(Me, False)
    Unload Me
End Sub

Private Sub Form_Load()
        Call BeTop(Splash, True)
    Label4.Caption = App.EXEName: Label5.Caption = "Version " & App.Major & "." & App.Minor
        FadeSpeed = 20
     Dim ps As Integer ' Transparantcy
    
    If Not EnableTransparanty(Me.hWnd, 0) = 0 Then
'
' Show Form, No Effects
'
     Me.Show
Else
    Me.Enabled = False
    Me.Show
    DoEvents
    
    For ps = 0 To 255 Step FadeSpeed
        DoEvents
        Call EnableTransparanty(Me.hWnd, ps)
        DoEvents
    Next
    Me.Enabled = True
    
    ' if you use disableTransparanty the the form will filcker is het transparanty is enabled again!
    ' Call EnableTransparanty(Me.hWnd)
    '
    
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim ps As Integer

If Not EnableTransparanty(Me.hWnd, 0) = 0 Then
    Exit Sub
Else
    For ps = 255 To 0 Step -FadeSpeed
        DoEvents
        Call EnableTransparanty(Me.hWnd, ps)
        DoEvents
    Next
    
End If

End Sub

Private Sub Image1_Click()
On Error Resume Next
Shell "c:\progra~1\intern~1\iexplore.exe http://www.3teq.com"
End Sub

Private Sub Label2_Click()
On Error Resume Next
Shell "c:\progra~1\intern~1\iexplore.exe http://www.3teq.com"
End Sub

Private Sub Label3_Click()
On Error Resume Next
Shell "c:\progra~1\intern~1\iexplore.exe mailto:3teq@softhome.net?Subject=" & App.EXEName & "---ICode=" & Icode & "---UserName=" & RegU & ".?Body=Order from WinCryptic Software!"
End Sub

