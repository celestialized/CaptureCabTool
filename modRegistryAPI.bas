Attribute VB_Name = "modRegistryAPI"
Option Explicit

'This module contains routines often used with regards to the Registry.

Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const HKEY_LOCAL_MACHINE = &H80000002

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String

  Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
  Dim strData As Integer
  'retrieve information about the key

    lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = REG_SZ Then
            'Create a buffer
            strBuf = String$(lDataBufSize, Chr$(0))
            'retrieve the key's content
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then
                'Remove the unnecessary chr$(0)'s
                RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If
          ElseIf lValueType = REG_BINARY Then 'NOT LVALUETYPE...
            'retrieve the key's value
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strData, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = strData
            End If
        End If
    End If

End Function

Public Function GetRegString(hKey As Long, strPath As String, strValue As String) As String

  Dim Ret As Long
  'Open the key

    RegOpenKey hKey, strPath, Ret
    'Get the key's content
    GetRegString = RegQueryStringValue(Ret, strValue)
    'Close the key
    RegCloseKey Ret

End Function
