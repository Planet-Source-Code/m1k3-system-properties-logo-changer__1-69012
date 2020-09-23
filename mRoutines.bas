Attribute VB_Name = "mRoutines"
Global Const WM_USER = &H400
Global Const EM_GETLINECOUNT = WM_USER + 10

#If Win32 Then

Public Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" _
    (ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Integer, _
    lParam As Any) As Long
#Else

Public Declare Function SendMessage Lib "user" _
    (ByVal hWnd As Integer, _
    ByVal wMsg As Integer, _
    ByVal wParam As Integer, _
    lParam As Any) As Long
#End If

Public Declare Function GetSystemDirectory Lib "kernel32" Alias _
 "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
 
Public Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" _
(ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long

Public Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" _
    (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long

Public Declare Function GetFileAttributes Lib "kernel32" _
 Alias "GetFileAttributesA" (ByVal lpSpec As String) As Long

Public Const INVALID_FILE_ATTRIBUTES As Long = -1

Function FileExists(sFileSpec As String) As Boolean

 Dim Attribs As Long
 
 Attribs = GetFileAttributes(sFileSpec)
 
  If (Attribs <> INVALID_FILE_ATTRIBUTES) Then
   FileExists = ((Attribs And vbDirectory) <> vbDirectory)
  End If
  
End Function

Public Sub File_MOVE(sFile As String, dFile As String)
  Call MoveFile(sFile, dFile)
End Sub

Private Function GetLongFilename(ByVal sShortFilename As String) As String
    
    Dim lRet As Long
    Dim sLongFilename As String
    sLongFilename = String$(1024, Chr(32))
    lRet = GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))

    If lRet > Len(sLongFilename) Then
        sLongFilename = String$(lRet + 1, Chr(32))
        lRet = GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))
    End If

    If lRet > 0 Then
        GetLongFilename = Left$(sLongFilename, lRet) & Chr(92)
    End If
    
End Function

Public Function sSystemDirectory() As String

    Dim sOut As String
    
    sOut = Space(260)
    GetSystemDirectory sOut, 260
    sOut = Left(sOut, InStr(sOut, Chr(0)) - 1)
    sSystemDirectory = GetLongFilename(sOut)
    
    If Right(sSystemDirectory, 1) <> Chr(92) Then sSystemDirectory = sSystemDirectory & Chr(92)
    
End Function

Public Function FILE_COPY(sFile As String, dFile As String)

Dim FILE_LEN As Long
Dim nSF, nDF As Long
Dim STR_BYTE As String
Dim bGet As Long
Dim bCopy As Long

FILE_LEN = FileLen(sFile)

nSF = 1
nDF = 2

On Error GoTo ERR_COPY

Open sFile For Binary As nSF
Open dFile For Binary As nDF

bGet = 4096
bCopy = 0

Do While bCopy < FILE_LEN

    If bGet < (FILE_LEN - bCopy) Then
        STR_BYTE = Space(bGet)
        Get #nSF, , STR_BYTE
    Else
        STR_BYTE = Space(FILE_LEN - bCopy)
        Get #nSF, , STR_BYTE
    End If

    bCopy = bCopy + Len(STR_BYTE)
    
    DoEvents

    Put #nDF, , STR_BYTE
    
Loop

Close #nSF
Close #nDF

Exit Function

ERR_COPY:
MsgBox Err.Description, vbCritical, "File Transfer Error"

End Function

