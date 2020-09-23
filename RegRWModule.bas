Attribute VB_Name = "RegRWModule"
' Module to read from and write to the registry

Option Explicit

Private r As Long
Private keyhand As Long
Private lResult As Long
Private lValueType As Long
Private lDataBufSize As Long

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&
Public Const REG_SZ = 1
Public Const REG_DWORD = 4

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Sub savekey(Hkey As Long, strPath As String)
    Dim keyhand&
    r = RegCreateKey(Hkey, strPath, keyhand&)
    r = RegCloseKey(keyhand&)
End Sub

Public Function getstring(Hkey As Long, strPath As String, strValue As String)
    Dim datatype As Long
    Dim strBuf As String
    Dim intZeroPos As Integer
    
    ' Open the key
    r = RegOpenKey(Hkey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)

    If lValueType = REG_SZ Then
        ' Create a buffer
        strBuf = String(lDataBufSize, " ")
        ' Retrieve the key's content
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
    ' Create a new key
    r = RegCreateKey(Hkey, strPath, keyhand)
    ' Save a string to the key
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    ' Close the key
    r = RegCloseKey(keyhand)
End Sub

Function getdword(ByVal Hkey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
    Dim lBuf As Long
    
    r = RegOpenKey(Hkey, strPath, keyhand)
    ' Get length/data type
    lDataBufSize = 4
    lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)

    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            getdword = lBuf
        End If
    End If
    r = RegCloseKey(keyhand)
End Function

Function SaveDword(ByVal Hkey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    r = RegCreateKey(Hkey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    r = RegCloseKey(keyhand)
End Function

Public Function DeleteKey(ByVal Hkey As Long, ByVal strKey As String)
    r = RegDeleteKey(Hkey, strKey)
End Function

Public Function DeleteValue(ByVal Hkey As Long, ByVal strPath As String, ByVal strValue As String)
    r = RegOpenKey(Hkey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function

