Attribute VB_Name = "Secure"
' This Application was Developed by Paul Bryan (C)2000
'               Registry and Form Handler Module
'
'
Option Explicit
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CharUpper Lib "user32" Alias "CharUpperA" (ByVal lpsz As String) As String
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&
Public Const REG_SZ = 1 ' Unicode nul terminated String
Public Const REG_DWORD = 4 ' 32-bit number


Dim x1a0(9) As Long
Dim cle(200) As Long
Dim x1a2 As Long
Dim inter As Long, res As Long, ax As Long, bx As Long
Dim cx As Long, dx As Long, si As Long, tmp As Long
Dim i As Long, c As Byte

Public Function getstring(Hkey As Long, strPath As String, strValue As String)
    Dim r As Variant
    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long, lvaluetype As Long
    Dim intZeroPos As Integer
    r = RegOpenKey(Hkey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lvaluetype, ByVal 0&, lDataBufSize)
    If lvaluetype = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                getstring = Left$(strBuf, intZeroPos - 1)
            Else
                getstring = strBuf
            End If
        End If
    End If
End Function
Public Sub savestring(Hkey As Long, strPath As String, strValue As String, strdata As String)
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(Hkey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)
End Sub
Function getdword(ByVal Hkey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
    Dim lResult As Long
    Dim lvaluetype As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long
    Dim r As Long
    Dim keyhand As Long
    r = RegOpenKey(Hkey, strPath, keyhand)
    lDataBufSize = 4
    lResult = RegQueryValueEx(keyhand, strValueName, 0&, lvaluetype, lBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        If lvaluetype = REG_DWORD Then
            getdword = lBuf
        End If
    End If
    r = RegCloseKey(keyhand)
End Function
Function SaveDword(ByVal Hkey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(Hkey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    r = RegCloseKey(keyhand)
End Function
Public Function DeleteKey(ByVal Hkey As Long, ByVal strKey As String)
    Dim r As Long
    r = RegDeleteKey(Hkey, strKey)
End Function
Public Function DeleteValue(ByVal Hkey As Long, ByVal strPath As String, ByVal strValue As String)
    Dim keyhand As Long, r As Variant
    r = RegOpenKey(Hkey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function
Public Sub SaveKey(Hkey As Long, strPath As String)
    Dim keyhand&
    Dim r As Variant
    r = RegCreateKey(Hkey, strPath, keyhand&)
    r = RegCloseKey(keyhand&)
End Sub

Public Sub BeTop(frmForm As Form, fOnTop As Boolean)
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    Dim lState As Long
    Dim iLeft As Integer, iTop As Integer, iWidth As Integer, iHeight As Integer
    
    With frmForm
        iLeft = .Left / Screen.TwipsPerPixelX
        iTop = .Top / Screen.TwipsPerPixelY
        iWidth = .Width / Screen.TwipsPerPixelX
        iHeight = .Height / Screen.TwipsPerPixelY
    End With
         If fOnTop Then
        lState = HWND_TOPMOST
    Else
        lState = HWND_NOTOPMOST
    End If
    Call SetWindowPos(frmForm.hWnd, lState, iLeft, iTop, iWidth, iHeight, 0)
End Sub

Public Function isTransparent(ByVal hWnd As Long) As Boolean
'
' Check for Windows 2000 Transparent Form OS.
'
On Error Resume Next
Dim msg As Long

msg = GetWindowLong(hWnd, GWL_EXSTYLE)

If (msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
    isTransparent = True
Else
    isTransparent = False
End If

If Err Then
    isTransparent = False
End If

End Function

Public Function EnableTransparanty(ByVal hWnd As Long, Perc As Integer) As Long
'
' Win2K Only !!
'
' Makes a window transparent (Win2K Only)
'
Dim msg As Long

On Error Resume Next

'Perc must be between 0 and 255

If Perc < 0 Or Perc > 255 Then
    EnableTransparanty = 1 'Invalid Percentage
Else

    'Makes transparent
    msg = GetWindowLong(hWnd, GWL_EXSTYLE)
    
    msg = msg Or WS_EX_LAYERED
    
    SetWindowLong hWnd, GWL_EXSTYLE, msg
    
    SetLayeredWindowAttributes hWnd, 0, Perc, LWA_ALPHA
    
    EnableTransparanty = 0 'OK
End If

If Err Then
    EnableTransparanty = 2 'Error: Unsupported function
End If
    
End Function

Public Function DisableTransparanty(ByVal hWnd As Long) As Long
'
' Win2K Only !!
'
' Disable the transparenty (form will flicker if enabled again) use
'
' ---> EnableTransparanty me.hwnd, 255 <---
'
' if you want to use the transparenty again (RECOMMEND!)
'

Dim msg As Long

On Error Resume Next

' Disables transparent
 msg = GetWindowLong(hWnd, GWL_EXSTYLE)

 msg = msg And Not WS_EX_LAYERED

 SetWindowLong hWnd, GWL_EXSTYLE, msg

 SetLayeredWindowAttributes hWnd, 0, 0, LWA_ALPHA
 DisableTransparanty = 0 'OK

If Err Then
    DisableTransparanty = 2 'Error: Unsupported function
    
End If
   
End Function


