VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WallPaper Slideshow"
   ClientHeight    =   4755
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   8085
   Icon            =   "frmChanger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Change &Image Source Folder"
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
      Left            =   240
      TabIndex        =   12
      Top             =   840
      Width           =   3375
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Change To Selected Now ->"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2760
      Top             =   1200
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Change &Now"
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
      Left            =   1920
      TabIndex        =   14
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Startup with &Windows, in the System Tray."
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
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   7935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sto&p"
      Default         =   -1  'True
      Enabled         =   0   'False
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
      Left            =   1920
      TabIndex        =   8
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start"
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
      Left            =   1920
      TabIndex        =   7
      Top             =   2280
      Width           =   1695
   End
   Begin VB.FileListBox filWallPaper 
      Height          =   2430
      Left            =   3840
      Pattern         =   "*.bmp;*.jpg"
      TabIndex        =   1
      Top             =   840
      Width           =   4095
   End
   Begin VB.Timer Timer10 
      Enabled         =   0   'False
      Interval        =   1111
      Left            =   3480
      Top             =   1080
   End
   Begin VB.TextBox txtCallbackMessage 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Text            =   "Hmmm"
      Top             =   1680
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   375
      Left            =   1200
      Max             =   0
      Min             =   1441
      TabIndex        =   2
      Top             =   2760
      Value           =   61
      Width           =   255
   End
   Begin VB.TextBox txtCounter 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "10"
      Top             =   2760
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   6120
      ScaleHeight     =   915
      ScaleWidth      =   1995
      TabIndex        =   11
      Top             =   3480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "(c)2000 by Paul Bryan Bensch"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4320
      Width           =   7935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Source Folder="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblPath 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   5895
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   6120
      Picture         =   "frmChanger.frx":08CA
      Top             =   960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5040
      Picture         =   "frmChanger.frx":1194
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblMinutes 
      Caption         =   "Minutes Left:"
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
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "&MenuPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUpShowWallPaperChanger 
         Caption         =   "&Open Wallpaper Slideshow"
      End
      Begin VB.Menu mnuPopUpChangeNow 
         Caption         =   "&Change Now"
      End
      Begin VB.Menu mnuPopUpActivateTimer 
         Caption         =   "&Slideshow Activated"
      End
      Begin VB.Menu mnuPopUp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopUpAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuPopUp2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopUpExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsChangePath 
         Caption         =   "&Change Path"
      End
      Begin VB.Menu mnuOptions1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsTileWallpaper 
         Caption         =   "&Tile WallPaper"
      End
      Begin VB.Menu mnuOptionsWallpaperStye 
         Caption         =   "&Wallpaper Style Centered or Stretched"
      End
      Begin VB.Menu mnuOptionsShowOnStartup 
         Caption         =   "&Start-up With Windows"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsRandomSelection 
         Caption         =   "&Random Selection"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "&Help"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Win32 API Call
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
'Used to change various Control Panel Settings.
Private Const SPIF_UPDATEINIFILE = &H1
Private Const SPIF_SENDWININICHANGE = &H2
Private Const SPI_SETDESKWALLPAPER = 20

' Form_Move Event - For maximizing out of the system tray.
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDBLCLK = &H209
Public chkActivate As Integer
'Classes
'Class that allows an Icon to be built in the system try
Private WithEvents SysIcon As CSystrayIcon
Attribute SysIcon.VB_VarHelpID = -1
'Class that performs basic Reading and Writing to the Registry
Private REG As clsRegEdit

'Local Variables
Private x As Long 'Used when setting the Wallpaper
Private y As Long
Private sNewWallPaper As String 'Indicates the pathname of the Image
Private iTimeCounter As Integer 'Total minutes to count down
Private iSeconds As Integer 'Used by counting routine (timer event)
Private iFileIndex As Integer 'Used to indicate which file to display
Public clz As Integer
Private Sub change()
'Purpose: Change the wallpaper
'Two steps:
'   1.  ensure registry is set up properly (win95/98 only)
'   2.  change the wallpaper
    Dim nfile As String, kfile As String
    Dim b As Boolean

    On Error GoTo EH

    'setup the Tilewallpaper
    REG.ValueName = "TileWallpaper" 'Setup the Name of the registry value
    b = REG.RegistryReadValue()  'Read the registry value into property RegCurrentValue
    If mnuOptionsTileWallpaper.Checked = False Then
        If Left(REG.RegCurrentValue, 1) = "1" Then
            b = REG.RegistryWriteString("0")  'Change registry value
        End If
    Else
        If Left(REG.RegCurrentValue, 1) = "0" Then
            b = REG.RegistryWriteString("1")  'Change registry value
        End If
    End If
    If b = False Then MsgBox ("Problem Setting Registry (TileWallPaper)")
    DoEvents

    
    'setup the WallpaperStyle
    REG.ValueName = "WallpaperStyle"  'Setup the Name of the registry value
    b = REG.RegistryReadValue()  'Read the registry value into property RegCurrentValue
    If mnuOptionsWallpaperStye.Checked = False Then
        If Left(REG.RegCurrentValue, 1) = "0" Then
            b = REG.RegistryWriteString("2")  'Change registry value
        End If
    Else
        If Left(REG.RegCurrentValue, 1) = "2" Then
            b = REG.RegistryWriteString("0")  'Change registry value
        End If
    End If
    If b = False Then MsgBox ("Problem Setting Registry (WallpaperStyle Centered)")
    DoEvents
    
    'Change the Wallpaper 'Api Call (See Declares Module)
    nfile = filWallPaper.path & "\" & sNewWallPaper
    Picture1.Picture = LoadPicture(nfile)
    Call SavePicture(Picture1, App.path & "\WinCrypt.bmp")
    kfile = App.path & "\WinCrypt.bmp"
    x = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, kfile, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)

    Exit Sub
    
EH:
    MsgBox "Error in Wall.Change: " & Err.Number & " - " & Err.Description, vbCritical, "WinCryptic WallPaper Slideshow"
End Sub

Private Sub Check1_Click()
Dim path As String
If Check1.Value = 1 Then
mnuOptionsShowOnStartup.Checked = True
If getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "WallSlide") = "" Then path = App.path & "\" & App.EXEName & ".exe": Call savestring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "WallSlide", path)
Exit Sub
Else
Call DeleteValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "WallSlide")
mnuOptionsShowOnStartup.Checked = False
End If

End Sub

Private Sub cmdChange_Click()
'Purpose: allow the desktop wallpaper to be changed immediatly
    If filWallPaper.ListIndex > (-1) Then
        iFileIndex = filWallPaper.ListIndex
        sNewWallPaper = filWallPaper.List(iFileIndex)
        change
    Else
        MsgBox "Please select an Image to change the Wallpaper to!", vbExclamation, "WinCryptic Software"
    End If
    
End Sub


Private Sub Command1_Click()
          chkActivate = 1
        Command1.Enabled = False
        Command2.Enabled = True
        Command3.Enabled = False
        Timer10.Enabled = True
        VScroll1.Visible = False
        cmdChange.Enabled = False
        SysIcon.TipText = "Wallpaper Slideshow (" & CStr(iTimeCounter) & ":00)"
        mnuPopUpActivateTimer.Checked = True
        SysIcon.IconHandle = Image1.Picture
        mnuOptionsChangePath.Enabled = False
        mnuOptionsRandomSelection.Enabled = False
        mnuOptionsTileWallpaper.Enabled = False
        mnuOptionsWallpaperStye.Enabled = False
         Me.Caption = "Wallpaper Slideshow V" & App.Major & "." & App.Minor
    iSeconds = 0  'initilize the time counter
    txtCounter.Text = iTimeCounter  'Countdown is performed in the text box
                                    'the variable holds the full amount
 
    End Sub
    
Private Sub Command2_Click()
        
        chkActivate = 0
        Command1.Enabled = True
        Command2.Enabled = False
        Command3.Enabled = True
        Timer10.Enabled = False
        VScroll1.Visible = True
        cmdChange.Enabled = True
        SysIcon.TipText = "Wallpaper Slideshow De-Activated"
        mnuPopUpActivateTimer.Checked = False
        SysIcon.IconHandle = Image2.Picture
        mnuOptionsChangePath.Enabled = True
        mnuOptionsRandomSelection.Enabled = True
        mnuOptionsTileWallpaper.Enabled = True
        mnuOptionsWallpaperStye.Enabled = True
         Me.Caption = "Wallpaper Slideshow V" & App.Major & "." & App.Minor
    iSeconds = 0  'initilize the time counter
    txtCounter.Text = iTimeCounter  'Countdown is performed in the text box
                                    'the variable holds the full amount
End Sub
                                    
Private Sub Command3_Click()
Call BeTop(Me, False)
    mnuOptionsChangePath_Click
End Sub


Private Sub Command4_Click()
mnuPopUpChangeNow_Click
End Sub

Private Sub filWallPaper_DblClick()
    'Purpose: allow the desktop wallpaer to be changed immediatly
    If filWallPaper.ListIndex > (-1) Then
        iFileIndex = filWallPaper.ListIndex
        sNewWallPaper = filWallPaper.List(iFileIndex)
        change
    Else
        MsgBox "Please select an Image to change the Wallpaper to!", vbExclamation, "WinCryptic Software"
    End If
    
End Sub

Private Sub Form_GotFocus()
Call BeTop(Me, True)
End Sub

Private Sub Form_Load()


'MnuPopUp is used to for the System Tray Icon Right Click Popup menu.

    Dim s As String
    Dim b As Boolean
     Me.Caption = "Wallpaper Slideshow v" & App.Major & "." & App.Minor
    On Error GoTo EH

    'setup registry editor
    Set REG = New clsRegEdit
    REG.TopHKey = HKEY_CURRENT_USER 'Establish Registry HKey
    REG.KeyName = "Control Panel\desktop" 'Establish Registry Key (Path)

    'SetUp the timer
    
    iTimeCounter = GetSetting(App.Title, "Presets", "Time", iTimeCounter)
    If iTimeCounter <= 0 Then iTimeCounter = 5
    VScroll1.Value = iTimeCounter
    
    'Set up the path
    filWallPaper.path = GetSetting(App.Title, "Presets", "Path", App.path)
    lblPath.Caption = filWallPaper.path

    'setup the Systray Icon
    Set SysIcon = New CSystrayIcon 'Set the new instance
    SysIcon.Initialize hWnd, Image1.Picture, "Wallpaper Slideshow De-actived"
    SysIcon.ShowIcon
    
    'Setup Timer
    chkActivate = CInt(GetSetting(App.Title, "Presets", "TimerOn", CStr(chkActivate)))
    
    'Setup Random Selection
    mnuOptionsRandomSelection.Checked = GetSetting(App.Title, "Presets", "Random", mnuOptionsRandomSelection.Checked)
    
    'Set up Tile Wallpaper
    REG.ValueName = "TileWallpaper" 'Set Registry Value to read
    b = REG.RegistryReadValue() 'Read the value into the property RegCurrentValue
    mnuOptionsTileWallpaper.Checked = IIf(Left(REG.RegCurrentValue, 1) = "1", 1, 0)

    'Set up Wallpaper Style
    REG.ValueName = "WallpaperStyle" 'Set Registry Value to read
    b = REG.RegistryReadValue() 'Read the value into the property RegCurrentValue
    mnuOptionsWallpaperStye.Checked = IIf(Left(REG.RegCurrentValue, 1) = "2", 0, 1)

    'check for errors
    If b = False Then MsgBox "Form.Load - Error Reading Registry!"
    
    'Set up iFileIndex
    iFileIndex = filWallPaper.ListCount - 1

    'Setup form Visibliity
    mnuOptionsShowOnStartup.Checked = GetSetting(App.Title, "Presets", "Startup_with_windows", mnuOptionsShowOnStartup.Checked)
    If mnuOptionsShowOnStartup.Checked = False Then frmMain.Visible = True: Check1.Value = 0: Exit Sub
    

    Check1.Value = 1
    Command1_Click
    chkActivate = 1
       Me.Visible = False
    Timer1.Enabled = True
    Exit Sub
    
EH:
    MsgBox Err.Description
    
End Sub


Private Sub Form_LostFocus()
 Call BeTop(Me, False)
 mnuPopUp.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim msgCallBackMessage As Long
  
  msgCallBackMessage = x / Screen.TwipsPerPixelX 'must be done to determine the actual callback message
   
   'Make the form bigger to see the Text box refered to below
   'the text box is invisible.  I used it only to test and determine
   'the call backs I need.  This way I never forget what call backs are available.
   
  Select Case msgCallBackMessage
'    Case WM_MOUSEMOVE
'      txtCallbackMessage.Text = "Mouse is moving"
'    Case WM_LBUTTONDOWN
'      txtCallbackMessage.Text = "Left button went down"
'        frmChanger.Visible = True
'    Case WM_LBUTTONUP
'      txtCallbackMessage.Text = "Left button came up"
    Case WM_LBUTTONDBLCLK
    mnuPopUpShowWallPaperChanger_Click
'      txtCallbackMessage.Text = "Double click catched from left button"
'    Case WM_RBUTTONDOWN
'      txtCallbackMessage.Text = "Right button went down"
    Case WM_RBUTTONUP
        PopupMenu mnuPopUp
'      txtCallbackMessage.Text = "Right button came up"
'    Case WM_RBUTTONDBLCLK
'      txtCallbackMessage.Text = "Double click catched from right button"
'   Case WM_MBUTTONDOWN
'      txtCallbackMessage.Text = "Middle button went down"
'    Case WM_MBUTTONUP
'      txtCallbackMessage.Text = "Middle button came up"
'    Case WM_MBUTTONDBLCLK
'      txtCallbackMessage.Text = "Double click catched from middle button"
  End Select

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

ExitProg.Show
Me.WindowState = vbMinimized: Cancel = 1
End Sub

Private Sub Form_Resize()
    'hide the form if minimed
    If Me.WindowState = vbMinimized Then
        Me.WindowState = vbNormal
        Me.Visible = False
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
'save basic registry settings upon exit
    Call BeTop(Me, False)
    
        SaveSetting App.Title, "Presets", "Time", iTimeCounter
        SaveSetting App.Title, "Presets", "Startup_with_windows", mnuOptionsShowOnStartup.Checked
        SaveSetting App.Title, "Presets", "Path", filWallPaper.path
        SaveSetting App.Title, "Presets", "TimerOn", CStr(chkActivate)
        SaveSetting App.Title, "Presets", "Random", mnuOptionsRandomSelection.Checked
        SysIcon.HideIcon
        DoEvents
        End
   
End Sub


Private Sub Image3_Click()
On Error Resume Next
Shell "c:\progra~1\intern~1\iexplore.exe http://www.3teq.com"
End Sub

Private Sub Label3_Click()
On Error Resume Next
Shell "c:\progra~1\intern~1\iexplore.exe mailto:3teq@softhome.net"
End Sub

Private Sub mnuFileClose_Click()
    frmMain.Visible = False
End Sub

Private Sub mnuFileExit_Click()
    clz = 1
    Unload Me
End Sub


Private Sub mnuHelpAbout_Click()
    Call BeTop(Me, False)
    Splash.Show
End Sub

Private Sub mnuHelpHelp_Click()
' sorry for the Lame Help...
    Call BeTop(Me, False)
    Dim sMsg As String
    
    sMsg = "The Wallpaper Slideshow is an Application that sits in the System Tray, and changes your desktop wallpaper," & vbCrLf & vbCrLf & "at your will, or automatically at the end of a timed interval." & vbCrLf & vbCrLf
    sMsg = sMsg & "Activate by Double-Clicking the Icon in the System Tray, or Left-Click for other options." & vbCrLf & vbCrLf
    sMsg = sMsg & "The program changes the wallpaper to one of the Images (*.bmp & *.jpg) in the Source folder, set by you." & vbCrLf & vbCrLf
    sMsg = sMsg & "The Application will automattically change the wallpaper sequentially (in alpha order) or in random order." & vbCrLf & vbCrLf
    sMsg = sMsg & "Various other options are available, including Start with Windows, and Tile Options." & vbCrLf & vbCrLf
    sMsg = sMsg & "When The Slideshow is activated, most options are not available for modification. Stop The Slideshow to modify the options." & vbCrLf & vbCrLf
    sMsg = sMsg & "All Settings are automatically saved to the regisrty, upon exiting the Application."

    MsgBox sMsg, vbInformation, "Help For the WinCryptic WallSlider Wallpaper Slideshow."
    Call BeTop(Me, True)
End Sub

Private Sub mnuOptionsChangePath_Click()
    Load frmSetPath
    frmSetPath.Drive1.Drive = Left(filWallPaper.path, 1)
    frmSetPath.Dir1.path = filWallPaper.path
    frmSetPath.Show vbModal
    lblPath.Caption = filWallPaper.path
    SaveSetting App.Title, "Presets", "Path", filWallPaper.path
End Sub

Private Sub mnuOptionsRandomSelection_Click()

    mnuOptionsRandomSelection.Checked = Not (mnuOptionsRandomSelection.Checked)

End Sub

Private Sub mnuOptionsShowOnStartup_Click()
    Dim path As String
    mnuOptionsShowOnStartup.Checked = Not (mnuOptionsShowOnStartup.Checked)
    If Not mnuOptionsShowOnStartup.Checked Then
    Call DeleteValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "WallSlide")
    
    Check1.Value = 0
    Exit Sub
    Else
    Check1.Value = 1
    path = App.path & "\" & App.EXEName & ".exe": Call savestring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "WallSlide", path)
    End If
End Sub

Private Sub mnuOptionsTileWallpaper_Click()
    mnuOptionsTileWallpaper.Checked = Not (mnuOptionsTileWallpaper.Checked)
End Sub

Private Sub mnuOptionsWallpaperStye_Click()
    mnuOptionsWallpaperStye.Checked = Not (mnuOptionsWallpaperStye.Checked)
End Sub

Private Sub mnuPopUpAbout_Click()
    Splash.Show
End Sub

Private Sub mnuPopUpActivateTimer_Click()
    If chkActivate = 1 Then Command2_Click: mnuPopUpActivateTimer.Checked = False: Label1.Caption = "": Exit Sub
    If chkActivate = 0 Then Command1_Click: mnuPopUpActivateTimer.Checked = True: Label1.Caption = ""
End Sub

Private Sub mnuPopUpExit_Click()
    clz = 1
    Unload Me
End Sub

Private Sub mnuPopUpShowWallPaperChanger_Click()
    frmMain.Visible = True
    Call BeTop(Me, True)
End Sub
Private Sub mnuPopUpChangeNow_Click()
    Dim xtemd As Integer, xtempi As Integer
    xtemd = Val(txtCounter.Text): xtempi = iSeconds
    iSeconds = 1: txtCounter.Text = "0": timer10_Timer
    txtCounter.Text = xtemd: iSeconds = xtempi
End Sub

Private Sub SysIcon_NIError(ByVal ErrorNumber As Long)

End Sub

Private Sub Timer1_Timer()
    Me.Visible = False
    Timer1.Enabled = False
End Sub

Private Sub timer10_Timer()
    Dim i As Integer
    
    iSeconds = iSeconds - 1
    Label1.Caption = Label1.Caption + "*"
    If Len(Label1.Caption) > 10 Then Label1.Caption = ""
    If iSeconds <= 0 Then
        iSeconds = 59
        txtCounter.Text = txtCounter.Text - 1
        If txtCounter.Text < 0 Then
            txtCounter.Text = iTimeCounter
            iSeconds = 0
stp:
            If filWallPaper.ListCount > 0 Then
                If mnuOptionsRandomSelection.Checked = False Then
                    iFileIndex = iFileIndex + 1
                    If iFileIndex >= filWallPaper.ListCount Then iFileIndex = 0
                    sNewWallPaper = filWallPaper.List(iFileIndex)
                    If sNewWallPaper = "WinCrypt.bmp" Then GoTo stp
                    filWallPaper.Selected(iFileIndex) = True
                    change
                    i = iFileIndex - 2
                    If i < 0 Then i = filWallPaper.ListCount - 1
                    filWallPaper.Selected(i) = False
                Else
                    For iFileIndex = 0 To filWallPaper.ListCount - 1
                        filWallPaper.Selected(iFileIndex) = False
                    Next iFileIndex
                    iFileIndex = Int(Rnd * filWallPaper.ListCount)
                    sNewWallPaper = filWallPaper.List(iFileIndex)
                    filWallPaper.Selected(iFileIndex) = True
                    change
                End If
            End If
        End If
    End If
    Me.Caption = "Countdown to Next Image: " & txtCounter.Text & ":" & Format(iSeconds, "00")
    SysIcon.TipText = "WinCryptic Wallpaper SlideShow (" & txtCounter.Text & ":" & Format(iSeconds, "00") & ") Ver." & App.Major & "." & App.Minor

End Sub

Private Sub VScroll1_Change()

    If VScroll1.Value = 0 Then VScroll1.Value = 1440
    If VScroll1.Value = 1441 Then VScroll1.Value = 1
    txtCounter.Text = VScroll1.Value
    iTimeCounter = VScroll1.Value
End Sub


