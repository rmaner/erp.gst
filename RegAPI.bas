Attribute VB_Name = "RegAPI"
Option Explicit
Public Const ERROR_SUCCESS = 0&
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_CREATE_LINK = &H20
Public Const SYNCHRONIZE = &H100000
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const REG_DWORD = 4
Public Const REG_OPTION_NON_VOLATILE = 0
Public Const REG_SZ = 1

Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long

Public Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
    ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    ByRef lpType As Long, _
    ByRef lpData As Long, _
    ByRef lpcbData As Long) As Long
Public Declare Function RegQueryValueExLONG Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
    ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    ByRef lpType As Long, _
    ByRef lpData As Long, _
    ByRef lpcbData As Long) As Long
Public Declare Function RegQueryValueExSTRING Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
    ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    ByRef lpType As Long, _
    ByVal lpData As String, _
    ByRef lpcbData As Long) As Long

Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type

Public Sub CreateNewKey(lRoot As Long, sSubKey As String)
    Dim lResult, hKeyHandle, lType, lValue As Long
    Dim sMsg As String
    Dim sAttr As SECURITY_ATTRIBUTES
    
    sAttr.nLength = Len(sAttr)
    sAttr.lpSecurityDescriptor = 0
    sAttr.bInheritHandle = True
    
    If Len(sSubKey) = 0 Then
        MsgBox "Invalid input to create key" & vbCrLf & "Key=" & sSubKey
        GoTo Exit_subCreateKey
    End If
    
    lResult = RegCreateKeyEx(lRoot, sSubKey, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, sAttr, hKeyHandle, lResult)
    If lResult <> ERROR_SUCCESS Then
        MsgBox "Could not create key"
        GoTo Exit_subCreateKey
    End If
    lResult = RegCloseKey(hKeyHandle)
    If lResult <> ERROR_SUCCESS Then
        MsgBox "Could not close key"
        GoTo Exit_subCreateKey
    End If
    
Exit_subCreateKey:
    Exit Sub
    
Error_subCreateKey:
    MsgBox Error
    Resume Exit_subCreateKey
End Sub


Public Sub SetNewValue(lRoot As Long, sSubKey As String, sType As String, sValueName As String, sValue As String)
    Dim lResult, hKeyHandle, lType, lValue As Long
    Dim sMsg As String
    Dim sAttr As SECURITY_ATTRIBUTES
    
    sAttr.nLength = Len(sAttr)
    sAttr.lpSecurityDescriptor = 0
    sAttr.bInheritHandle = True
   
    If Len(sSubKey) = 0 Then
        MsgBox "Invalid input to create value" & vbCrLf & "Value=" & sValue
        GoTo Exit_SetNewValue
    End If
    
    lResult = RegCreateKeyEx(lRoot, sSubKey, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, sAttr, hKeyHandle, lResult)
    
    If lResult <> ERROR_SUCCESS Then
        MsgBox "Could not open key"
        GoTo Exit_SetNewValue
    End If
    
    Select Case sType
        Case "String"
            lType = REG_SZ
            sValue = sValue
            lResult = RegSetValueEx(hKeyHandle, sValueName, 0&, lType, ByVal sValue, Len(sValue))
        Case "Number"
            lType = REG_DWORD
            lValue = Val(sValue)
            lResult = RegSetValueEx(hKeyHandle, sValueName, 0&, lType, lValue, Len(lValue))
        Case Else
            MsgBox "Bad input to create value" & vbCrLf & "Value=" & sValue
            GoTo Exit_SetNewValue
    End Select
    
    If lResult <> ERROR_SUCCESS Then
        MsgBox "Could not set value"
        GoTo Exit_SetNewValue
    Else
        lResult = RegCloseKey(hKeyHandle)
    End If
    
    lResult = RegCloseKey(hKeyHandle)

Exit_SetNewValue:
    Exit Sub
    
Error_SetNewValue:
    MsgBox Error
    Resume Exit_SetNewValue
End Sub


Public Function ReadRegValue(lRoot As Long, sSubKey As String, sValueName As String) As String
    Dim lResult, hKeyHandle, lType, lValue, lcch, lrc As Long
    Dim sMsg, sValue As String
    
    On Error GoTo Exit_ReadRegValue
    
    lResult = RegOpenKeyEx(lRoot, sSubKey, 0, KEY_QUERY_VALUE, hKeyHandle)
    If lResult <> ERROR_SUCCESS Then
        MsgBox "Could not open key"
        GoTo Exit_ReadRegValue
    End If
    
    lrc = RegQueryValueExNULL(hKeyHandle, sValueName, 0&, lType, 0&, lcch)
    
    Select Case lType
        Case REG_SZ
            sValue = String(lcch, 0)
            lrc = RegQueryValueExSTRING(hKeyHandle, sValueName, 0&, lType, sValue, lcch)
            ReadRegValue = Left(sValue, lcch - 1)
        
        Case REG_DWORD
            lrc = RegQueryValueExLONG(hKeyHandle, sValueName, 0&, lType, lValue, lcch)
            If lrc = ERROR_SUCCESS Then ReadRegValue = str(lValue)
        
        Case Else
            lrc = -1
    End Select
    
    If lrc <> ERROR_SUCCESS Then
        bLoad = False 'DO NOT LOAD MDI FORM
        MsgBox "Could not read key"
        lResult = RegCloseKey(hKeyHandle)
        GoTo Exit_ReadRegValue
    End If
    
Exit_ReadRegValue:
    Exit Function
    
Error_ReadRegValue:
    MsgBox Error
    Resume Exit_ReadRegValue
    
End Function


Public Sub DeleteRegValue(lRoot As Long, sSubKey As String, sValueName As String)
    Dim lResult, hKeyHandle As Long
    Dim sMsg As String
    
    On Error GoTo Error_DeleteRegValue
    
    lResult = RegOpenKeyEx(lRoot, sSubKey, 0, KEY_SET_VALUE, hKeyHandle)
    If lResult <> ERROR_SUCCESS Then
        MsgBox "Could not open key"
        GoTo Exit_DeleteRegValue
    End If
    
    lResult = RegDeleteValue(hKeyHandle, sValueName)
    If lResult <> ERROR_SUCCESS Then
        MsgBox "Could not delete key value"
        lResult = RegCloseKey(hKeyHandle)
        GoTo Exit_DeleteRegValue
    End If
    
    lResult = lResult = RegCloseKey(hKeyHandle)
    
Exit_DeleteRegValue:
    Exit Sub
    
Error_DeleteRegValue:
    MsgBox Error
    Resume Exit_DeleteRegValue
End Sub


Public Sub DeleteRegKey(lRoot As Long, sSubKey As String)
    Dim lResult, hKeyHandle As Long
    Dim sMsg As String
    
    On Error GoTo Error_DeleteRegKey
    
    lResult = RegOpenKeyEx(lRoot, sSubKey, 0, KEY_QUERY_VALUE, hKeyHandle)
    If lResult <> ERROR_SUCCESS Then
        MsgBox "Could not open key"
        GoTo Exit_DeleteRegKey
    End If
    
    lResult = RegDeleteKey(lRoot, sSubKey)
    
    If lResult <> ERROR_SUCCESS Then
        MsgBox "Could not delete key"
        GoTo Exit_DeleteRegKey
    End If
    
    lResult = lResult = RegCloseKey(hKeyHandle)
    
Exit_DeleteRegKey:
    Exit Sub
    
Error_DeleteRegKey:
    MsgBox Error
    Resume Exit_DeleteRegKey
End Sub

