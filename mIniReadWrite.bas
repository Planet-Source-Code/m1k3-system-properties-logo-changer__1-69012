Attribute VB_Name = "mIniReadWrite"
'[General]
 'Manufacturer=Manufacturer
 'Model=Model
'[OEMSpecific]
 'SubModel=SubModel
 'SerialNo=SerialNo
 'OEM1=OEM1
 'OEM2=OEM2
'[Support Information]
 'Line1=Line1
 'Line2=Line2
 'Line3=Line3

Option Explicit

Public Declare Function WritePrivateProfileString& Lib "kernel32" Alias _
    "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, _
     ByVal keydefault$, ByVal FileName$)

Public Declare Function GetPrivateProfileString& Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, _
     ByVal keydefault$, ByVal ReturnedString$, ByVal RSSize&, ByVal FileName$)

Public Sub WriteINI(INISection As String, INIKey As String, INIValue As String, INIFile As String)
    Call WritePrivateProfileString(INISection, INIKey, INIValue, INIFile)
End Sub

Public Function ReadINI(INISection As String, INIKey As String, INIFile As String) As String

    Dim StringBuffer As String
    Dim StringBufferSize As Long
    
    StringBuffer = Space$(255)
    StringBufferSize = Len(StringBuffer)
    StringBufferSize = GetPrivateProfileString(INISection, INIKey, "", StringBuffer, StringBufferSize, INIFile)

    If StringBufferSize > 0 Then
        ReadINI = Left$(StringBuffer, StringBufferSize)
    Else
        ReadINI = vbNullString
    End If
    
End Function



