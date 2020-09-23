VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Enum for Registry HKeys
Public Enum REG_TOPLEVEL_KEYS
 HKEY_CLASSES_ROOT = &H80000000
 HKEY_CURRENT_CONFIG = &H80000005
 HKEY_CURRENT_USER = &H80000001
 HKEY_DYN_DATA = &H80000006
 HKEY_PERFORMANCE_DATA = &H80000004
 HKEY_USERS = &H80000003
End Enum


Public Enum RegDataType
    REG_NONE = 0
    REG_SZ = 1                      'Strings:  We use this
    REG_EXPAND_SZ = 2
    REG_BINARY = 3
    REG_DWORD_LITTLE_ENDIAN = 4     'Long value:  We use this also
    REG_DWORD_BIG_ENDIAN = 5
    REG_LINK = 6
    REG_MULTI_SZ = 7
    REG_RESOURCE_LIST = 8
    REG_FULL_RESOURCE_DESCRIPTOR = 9
    REG_RESOURCE_REQUIREMENTS_LIST = 10
End Enum

'Enum to simplify Registry Errors
'Use primarily by ISReturnCodeError Function
Public Enum Reg_Errors
    ERROR_NONE = 0
    ERROR_BADDB = 1
    ERROR_BADKEY = 2
    ERROR_CANTOPEN = 3
    ERROR_CANTREAD = 4
    ERROR_CANTWRITE = 5
    ERROR_OUTOFMEMORY = 6
    ERROR_INVALID_PARAMETER = 7
    ERROR_ACCESS_DENIED = 8
    ERROR_INVALID_PARAMETERS = 87
    ERROR_NO_MORE_ITEMS = 259
End Enum

'Main constants used by API Call RegOpenKeyEx
Private Const KEY_ALL_ACCESS = &H3F
Private Const REG_OPTION_NON_VOLATILE = 0

'Win32 API Calls
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
    'Create new Registry Key, If the Registry Key exists then just open.
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal Hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
    'Open an existing Registry Key
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
    'Retrieves the Registry ValueKey Data Type (and Length if String)
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
    'Retrieves a string value from the registry
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
    'Retrieves a Long value from the registry
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    'Writes a value to a Value key.
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
    'Closes the Registry key (must be done after reading a value)

'local variable(s) to hold property value(s)
Private mvarTopHKey As REG_TOPLEVEL_KEYS
Private mvarKeyName As String
Private mvarValueName As String
Private mvarRegCurrentValue As Variant
Private mvarValueDataType As RegDataType


Public Function RegistryWriteString(Optional sData As String) As Boolean
'Purpose: Write string data to the registry.  Maybe later I'll add additional
'data types.  But for now this will suffice for my purposes

    Dim Hkey As Long
    Dim lrc As Long
    Dim b As Boolean
    
    On Error GoTo ErrorHandler
        
    If IsMissing(sData) Or Len(sData) = 0 Then
        sData = mvarRegCurrentValue
    End If
    
    lrc = RegCreateKey(mvarTopHKey, mvarKeyName, Hkey) 'create or Open the Key
    If IsReturnCodeError(lrc) = False Then
        'Write the value to the Registry
        lrc = RegSetValueEx(Hkey, mvarValueName, 0, REG_SZ, ByVal sData, Len(sData))
        RegistryWriteString = Not (IsReturnCodeError(lrc))
        lrc = RegCloseKey(Hkey)  'always close a key when done working with it.
        RegistryWriteString = Not (IsReturnCodeError(lrc))
    End If

Exit Function

ErrorHandler:
    RegistryWriteString = False
    Exit Function
    
End Function


Function RegistryReadValue() As Boolean
'Purpose:  Read either String or Long values from the registry.

    Dim lpcbData As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String
    Dim Hkey As Long

    On Error GoTo RegistryReadValueError
    
    'open the key first
    lrc = RegOpenKeyEx(mvarTopHKey, mvarKeyName, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, Hkey)
    If IsReturnCodeError(lrc) Then
        RegistryReadValue = False
        Exit Function
    End If

    ' Determine the size and type of data to be read
    lrc = RegQueryValueExNULL(Hkey, mvarValueName, 0&, mvarValueDataType, 0&, lpcbData)
    If IsReturnCodeError(lrc) Then
        RegistryReadValue = False
        Exit Function
    End If

    'read the value
    Select Case mvarValueDataType
        Case REG_SZ  ' For Strings
            sValue = String(lpcbData, 0)
            lrc = RegQueryValueExString(Hkey, mvarValueName, 0&, mvarValueDataType, sValue, lpcbData)
            If IsReturnCodeError(lrc) = False Then
                mvarRegCurrentValue = Left$(sValue, lpcbData)
            Else
                mvarRegCurrentValue = Empty
            End If
        
        Case REG_DWORD_LITTLE_ENDIAN  ' For Longs
            lrc = RegQueryValueExLong(Hkey, mvarValueName, 0&, lType, lValue, lpcbData)
            If IsReturnCodeError(lrc) = False Then
                mvarRegCurrentValue = lValue
            Else
                mvarRegCurrentValue = 0
            End If
            
        Case Else
            'all other data types not supported
            mvarRegCurrentValue = "Data Type Not Supported"
            lrc = RegCloseKey(Hkey)
            RegistryReadValue = False
            Exit Function
    End Select
    
    lrc = RegCloseKey(Hkey) 'Always close the registry key when done using it
    RegistryReadValue = Not (IsReturnCodeError(lrc))

    Exit Function

RegistryReadValueError:
    RegistryReadValue = False
    
End Function


Private Function IsReturnCodeError(r As Reg_Errors) As Boolean
'Purpose: Check to see if the there is an error and if there is then
'Raise an error

    Dim sMsg As String
    
    IsReturnCodeError = True
    Select Case r
        Case ERROR_NONE
            sMsg = "No Error"
            IsReturnCodeError = False
            Exit Function
        Case ERROR_BADDB
            sMsg = "Bad Or Missing Registry"
        Case ERROR_BADKEY
            sMsg = "Bad Or Missing Registry Key"
        Case ERROR_CANTOPEN
            sMsg = "Unable To open Registry Key"
        Case ERROR_CANTREAD
            sMsg = "Unable To Read Registry Key"
        Case ERROR_CANTWRITE
            sMsg = "Unable To Write Value To Registry Key"
        Case ERROR_OUTOFMEMORY
            sMsg = "Out Of Memory"
        Case ERROR_INVALID_PARAMETER
            sMsg = "Invalid Parameter"
        Case ERROR_ACCESS_DENIED
            sMsg = "Access To Registry Or Registry Key Has Been Denied"
        Case ERROR_INVALID_PARAMETERS
            sMsg = "Invalid Parameters"
        Case ERROR_NO_MORE_ITEMS
            sMsg = "No More Items"
    End Select
    
    Err.Raise -10000, "clsRegEdit", sMsg
    
End Function
Public Property Let ValueDataType(ByVal vData As RegDataType)
    mvarValueDataType = vData
End Property
Public Property Get ValueDataType() As RegDataType
    ValueDataType = mvarValueDataType
End Property
Public Property Let RegCurrentValue(ByVal vData As Variant)
    mvarRegCurrentValue = vData
End Property
Public Property Get RegCurrentValue() As Variant
    If IsObject(mvarRegCurrentValue) Then
        Set RegCurrentValue = mvarRegCurrentValue
    Else
        RegCurrentValue = mvarRegCurrentValue
    End If
End Property
Public Property Let ValueName(ByVal vData As String)
    mvarValueName = vData
End Property
Public Property Get ValueName() As String
    ValueName = mvarValueName
End Property
Public Property Let KeyName(ByVal vData As String)
    mvarKeyName = vData
End Property
Public Property Get KeyName() As String
    KeyName = mvarKeyName
End Property
Public Property Let TopHKey(ByVal vData As REG_TOPLEVEL_KEYS)
    mvarTopHKey = vData
End Property
Public Property Get TopHKey() As REG_TOPLEVEL_KEYS
    TopHKey = mvarTopHKey
End Property




