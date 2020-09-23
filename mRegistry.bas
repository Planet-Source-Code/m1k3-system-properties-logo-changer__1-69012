Attribute VB_Name = "mRegistry"
Option Explicit

Public Enum KeyRoot
  [HKEY_CLASSES_ROOT] = &H80000000
  [HKEY_CURRENT_CONFIG] = &H80000005
  [HKEY_CURRENT_USER] = &H80000001
  [HKEY_LOCAL_MACHINE] = &H80000002
  [HKEY_USERS] = &H80000003
End Enum

Public Const ERROR_SUCCESS = 0&

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
(ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
(ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
(ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
(ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
(ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" _
(ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long

Private Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" _
(ByVal hKey As Long, ByVal lpFile As String, lpSecurityAttributes As Any) As Long

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" _
(ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public Enum KeyType
  [REG_BINARY] = 3
  [REG_DWORD] = 4
  [REG_SZ] = 1
End Enum

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Public Function rREG(hKey As KeyRoot, sKey As String, sVAL As String)
    
    Dim rLNG, rKEY, rDAT As Long
    Dim sBuf As String
    Dim nBUF, nVAL, nRST As Long
    Dim nPos As Long
    
    rLNG = RegOpenKey(hKey, sKey, rKEY)
    nRST = RegQueryValueEx(rKEY, sVAL, 0&, nVAL, ByVal 0&, nBUF)
    
    If nVAL = REG_SZ Then
        sBuf = String(nBUF, Chr(32))
        nRST = RegQueryValueEx(rKEY, sVAL, 0&, 0&, ByVal sBuf, nBUF)
        If nRST = ERROR_SUCCESS Then
            nPos = InStr(sBuf, Chr$(0))
            If nPos > 0 Then
                rREG = Left$(sBuf, nPos - 1)
            Else
                rREG = sBuf
            End If
        End If
    End If
    
End Function
    
Public Sub wREG(hKey As KeyRoot, sKey As String, sVAL As String, sDAT As String)
    
    Dim rKEY As Long
    Dim rLNG As Long
    
    rLNG = RegCreateKey(hKey, sKey, rKEY)
    rLNG = RegSetValueEx(rKEY, sVAL, 0, REG_SZ, ByVal sDAT, Len(sDAT))
    rLNG = RegCloseKey(rKEY)
    
End Sub


