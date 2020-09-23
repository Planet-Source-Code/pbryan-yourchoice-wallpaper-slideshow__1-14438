VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSystrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Purpose:  To place an Icon into the System Tray and maintain it

Option Explicit

'Win32 API declarations
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function GetLastError Lib "kernel32" () As Long
'GetLastError returns
'the code number of the error when the systray access fail

'Constants
Private Const NIM_ADD = &H0 'Add a new icon to the Systray (dwMessage)
Private Const NIM_MODIFY = &H1 'Modify an existing icon (dwMessage)
Private Const NIM_DELETE = &H2 'Delete an existing icon (dwMessage)
Private Const NIF_MESSAGE = &H1 'Message has been updated (uCallbackMessage)
Private Const NIF_ICON = &H2 'Icon picture has been changed (uCallbackMessage)
Private Const NIF_TIP = &H4 'New ToopTip for the icon is set (uCallbackMessage)
Private Const WM_MOUSEMOVE = &H200 'Used as the ID of the callback message
Private Const MAX_TIP_LENGTH As Long = 64 'This is the max length of a ToolTip.

Private Type NOTIFYICONDATA
    cbSize As Long 'The size of this type
    hWnd As Long 'The hWnd that will receive the CallBack message
    uId As Long 'The ID of the application. Zero represent this application
    uFlags As Long 'The flags. Look at the constants beginning by NIF_ for the flags and their definition
    uCallbackMessage As Long 'This is the callback message
    hIcon As Long 'An handle to the icon that will be displayed
    szTip As String * MAX_TIP_LENGTH 'The string of the ToopTip. Must be terminating by a null zero (chr(0))
End Type

Private nidTrayIcon As NOTIFYICONDATA 'Response Type Variable
Private bIconDisplayed As Boolean 'The status of the icon. True=Displayed
Private bUpdateOnChange As Boolean 'If True, Shell_NotifyIcon is call
                                  'whenever a change to a properties
                                  'of nidTrayIcon is made
                                  'True when class is initialized.

Public Event CSysTrayError(ByVal ErrorNumber As Long)

Public Function Initialize(ByVal hWnd As Long, ByVal hIcon As Long, ByVal sTip As String, Optional ByVal uCallbackMessage As Long = WM_MOUSEMOVE) As Long
'Initialize the systray icon (The icon won't show)

    With nidTrayIcon
        .cbSize = Len(nidTrayIcon)
        .hIcon = hIcon 'Handle of the icon
        .hWnd = hWnd 'Handle of the window that receives notification
        .szTip = Left(sTip, MAX_TIP_LENGTH - 1) & vbNullChar
        .uCallbackMessage = uCallbackMessage 'Default Event: &H200 (512, WM_MOUSEMOVE)
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uId = vbNull
    End With
    bIconDisplayed = False
    bUpdateOnChange = True
    
'Note: hIcon
'   This is not directly the bitmap. It is a handle of the Bitmap or Icon.
'   You may refer to it by refering to the picture property or icon property
'   of a form, picturebox or image control etc...
'   you can Also use LoadPicture() or LoadResPicture()
    
End Function

Public Function ShowIcon() As Boolean
'Purpose: To show the icon specified by hIcon in the systray.
    If Not bIconDisplayed Then                      'If the icon is not already displayed...
        ShowIcon = Shell_NotifyIcon(NIM_ADD, nidTrayIcon)
        If ShowIcon = False Then                    'Check for Error
            RaiseEvent CSysTrayError(GetLastError)  'Yes there is an error
        Else
            bIconDisplayed = True                   'No it worked OK
        End If
    End If
End Function

Public Function HideIcon() As Boolean
'Purpose:  Remove the icon from the systray.
    If bIconDisplayed Then                          'If the icon is displayed...
        HideIcon = Shell_NotifyIcon(NIM_DELETE, nidTrayIcon)
        If HideIcon = False Then                    'Check for Error
            RaiseEvent CSysTrayError(GetLastError)  'Yes there is an error
        Else
            bIconDisplayed = False                  'No it worked OK
        End If
    End If
    
End Function

Public Property Let IconHandle(ByVal hIcon As Long)
'Purose:  Change the icon displayed in the systray for the icon handled by hIcon.
    nidTrayIcon.hIcon = hIcon                       'Set the Icon Handle
    If bUpdateOnChange Then
        nidTrayIcon.uFlags = NIF_ICON               'Change only the Icon
        Update                                      'Make the change
        nidTrayIcon.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP 'set back to normal
    End If
End Property

Public Property Let TipText(ByVal sTip As String)
'Purpose:  Change the tooltip text of the icon by sTip
    nidTrayIcon.szTip = Left(sTip, MAX_TIP_LENGTH - 1) & vbNullChar 'Set the ToolTip
    If bUpdateOnChange Then
        nidTrayIcon.uFlags = NIF_TIP                'Change only the Tooltip
        Update                                      'Make the Change
        nidTrayIcon.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP 'Set back to normal
    End If
End Property

Public Property Let CallbackMessage(ByVal uCallbackMessage As Long)
'Purpose:  Change the callback message that is sent to the form.
    nidTrayIcon.uCallbackMessage = uCallbackMessage
    If bUpdateOnChange Then
        nidTrayIcon.uFlags = NIF_MESSAGE            'Change only the Call back message
        Update                                      'Make the Change
        nidTrayIcon.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP  'Set back to normal
    End If
End Property

Public Function Update() As Boolean
'Update the icon on the systray.
    If bIconDisplayed Then                  'Perform only if Icon is already Showing
        Update = Shell_NotifyIcon(NIM_MODIFY, nidTrayIcon)
        If Update = False Then              'Check For Error
            RaiseEvent CSysTrayError(GetLastError)
        End If
    End If
End Function

Public Property Get IconHandle() As Long
    IconHandle = nidTrayIcon.hIcon
End Property

Public Property Get TipText() As String
    TipText = Left(nidTrayIcon.szTip, Len(nidTrayIcon.szTip) - 1)
End Property

Public Property Get CallbackMessage() As Long
    CallbackMessage = nidTrayIcon.uCallbackMessage
End Property

Public Property Let UpdateOnChange(bUpdate As Boolean)
    bUpdateOnChange = bUpdate
End Property

Private Property Get UpdateOnChange() As Boolean
    UpdateOnChange = bUpdateOnChange
End Property

Private Sub Class_Terminate()
'Purpose:  Removes the icon from the Systray
    HideIcon
End Sub

Public Property Get Visible() As Boolean
'Purpose:  Indicate if the icon in the systray is visible
    If bIconDisplayed Then
        Visible = True
    End If
End Property

Public Property Let Visible(ByVal bVisible As Boolean)
'Purpose: Sets Icon Visiblility (Property instead of methods)
    If bVisible Then
        ShowIcon
    Else
        HideIcon
    End If
End Property
