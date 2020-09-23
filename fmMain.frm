VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Windows System Logo Utility"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   Icon            =   "fmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pDest 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   240
      ScaleHeight     =   1800
      ScaleWidth      =   2700
      TabIndex        =   26
      Top             =   360
      Width           =   2700
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   75
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   7230
      TabIndex        =   25
      Top             =   6225
      Width           =   7290
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   24
      Top             =   6300
      Width           =   7290
      _ExtentX        =   12859
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5080
            MinWidth        =   5080
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5080
            MinWidth        =   5080
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "User Information"
      Height          =   5895
      Left            =   3180
      TabIndex        =   5
      Top             =   240
      Width           =   4035
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   180
         Top             =   4680
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmMain.frx":0CCA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmMain.frx":0EE0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmMain.frx":1128
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Timer Timer1 
         Interval        =   1500
         Left            =   180
         Top             =   5340
      End
      Begin VB.TextBox txInf 
         Height          =   3195
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   2520
         Width           =   2715
      End
      Begin VB.TextBox txFld 
         Height          =   285
         Index           =   0
         Left            =   1200
         MaxLength       =   25
         TabIndex        =   6
         Top             =   240
         Width           =   2715
      End
      Begin VB.TextBox txFld 
         Height          =   285
         Index           =   1
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   8
         Top             =   540
         Width           =   2715
      End
      Begin VB.TextBox txFld 
         Height          =   285
         Index           =   2
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   10
         Top             =   840
         Width           =   2715
      End
      Begin VB.TextBox txFld 
         Height          =   285
         Index           =   3
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   12
         Top             =   1140
         Width           =   2715
      End
      Begin VB.TextBox txFld 
         Height          =   285
         Index           =   4
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   14
         Top             =   1440
         Width           =   2715
      End
      Begin VB.TextBox txFld 
         Height          =   285
         Index           =   5
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   16
         Top             =   1740
         Width           =   2715
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[Will Not Be Displayed ]"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   9
         Left            =   1200
         TabIndex        =   23
         Top             =   2280
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[Will Be Displayed ]"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   8
         Left            =   1200
         TabIndex        =   22
         Top             =   2040
         Width           =   1365
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OEM2"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Information"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   19
         Top             =   2700
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Manufacturer"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   300
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Model"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   900
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serial"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OEM1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   1500
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   60
      TabIndex        =   2
      Top             =   2520
      Width           =   3015
      Begin VB.CommandButton cmShow 
         Caption         =   "Show System Properties"
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   2400
         Width           =   1995
      End
      Begin VB.PictureBox Picture4 
         Height          =   75
         Left            =   120
         ScaleHeight     =   15
         ScaleWidth      =   2715
         TabIndex        =   33
         Top             =   2160
         Width           =   2775
      End
      Begin VB.PictureBox Picture3 
         Height          =   1815
         Left            =   1440
         ScaleHeight     =   1755
         ScaleWidth      =   15
         TabIndex        =   32
         Top             =   240
         Width           =   75
      End
      Begin VB.CommandButton cmExample 
         Caption         =   "Example"
         Height          =   315
         Left            =   1680
         TabIndex        =   31
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton cmClear 
         Caption         =   "Clear Fields"
         Height          =   315
         Left            =   1680
         TabIndex        =   30
         Top             =   600
         Width           =   1155
      End
      Begin VB.CommandButton cmRetrieve 
         Caption         =   "Retrieve"
         Height          =   315
         Left            =   1680
         TabIndex        =   29
         Top             =   960
         Width           =   1155
      End
      Begin VB.CommandButton cmCreate 
         Caption         =   "Save INI"
         Height          =   315
         Left            =   1680
         TabIndex        =   28
         Top             =   1440
         Width           =   1155
      End
      Begin VB.CommandButton cmKill 
         Caption         =   "Reset"
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   3180
         Width           =   975
      End
      Begin VB.CommandButton cmSavePic 
         Caption         =   "Save Logo"
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   1155
      End
      Begin VB.CommandButton cmLoadPic 
         Caption         =   "Load Logo"
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1155
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   2340
         Picture         =   "fmMain.frx":1361
         Top             =   3000
         Width           =   480
      End
   End
   Begin VB.PictureBox pSrce 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   60
      ScaleHeight     =   600
      ScaleWidth      =   780
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   75
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   7230
      TabIndex        =   0
      Top             =   0
      Width           =   7290
   End
   Begin MSComDlg.CommonDialog Cdl1 
      Left            =   180
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   2055
      Left            =   120
      Top             =   240
      Width           =   2955
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Actual Size [180 x 120] Logo Picture"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   2340
      Width           =   2580
   End
End
Attribute VB_Name = "fmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim srcPic$
Dim oemIni$
Dim oemBmp$

Private Sub Load_Picture()
 
 With Cdl1
  .Filter = "BMP Files ONLY (*.bmp)|*.bmp"
  .FilterIndex = 1
  .ShowOpen
 End With
 
 If Cdl1.FileName <> Empty Then
  srcPic = Cdl1.FileName
  pSrce.Picture = LoadPicture(srcPic)
  pSrce.PaintPicture pSrce.Picture, 0, 0, pSrce.Width, pSrce.Height
 End If
 
End Sub

Private Sub cmClear_Click()

 Dim x%
 
 For x = 0 To 5
  txFld(x) = Empty
 Next x
 
 txInf = Empty
 
End Sub

Private Sub cmExample_Click()
 objWMI
'***********************************************
'EXAMPLE TEXT FOR [CONTACT INFORMATION] FILE
 txInf = "Description or Builder" & vbCrLf & _
         "Address Line 1" & vbCrLf & _
         "Address Line 2" & vbCrLf & _
         "City, State     Zip" & vbCrLf & _
          vbCrLf & _
         "Phone: 800-555-5555" & vbCrLf & _
          vbCrLf & _
         "Customer Support:" & vbCrLf & _
         "Someones Real Name" & vbCrLf & _
         "800-555-5555 Ext: 123" & vbCrLf & _
         vbCrLf & _
         "email@some.com"
'***********************************************

End Sub

Private Sub cmKill_Click()

  If FileExists(oemBmp) Then
   If rREG(HKEY_LOCAL_MACHINE, "Software\SL", "Run") = Empty Then
    Kill oemBmp: StatusBar1.Panels(1) = "BMP File Destroyed"
    Call FILE_COPY(sSystemDirectory & "oemlogo.bak", oemBmp)
   Else
    Kill oemBmp: StatusBar1.Panels(1) = "BMP File Recovered"
     If Not FileExists(oemIni) Then
      Call wREG(HKEY_LOCAL_MACHINE, "Software\SL", "Run", vbNullString)
     End If
    Call FILE_COPY(sSystemDirectory & "oemlogo.bak", oemBmp)
   End If
  Else
    StatusBar1.Panels(1) = "BMP File Not Present"
  End If

  If FileExists(oemIni) Then
   If rREG(HKEY_LOCAL_MACHINE, "Software\SL", "Run") = Empty Then
    Kill oemIni: StatusBar1.Panels(2) = "INI File Destroyed"
    Call FILE_COPY(sSystemDirectory & "oeminfo.bak", oemIni)
   Else
    Kill oemIni: StatusBar1.Panels(2) = "INI File Recovered"
    Call wREG(HKEY_LOCAL_MACHINE, "Software\SL", "Run", vbNullString)
    Call FILE_COPY(sSystemDirectory & "oeminfo.bak", oemIni)
   End If
  Else
    StatusBar1.Panels(2) = "INI File Not Present"
  End If
  
End Sub

Private Sub cmLoadPic_Click()
  Call Load_Picture
  Call Set_Small(180, 120, pSrce, pDest)
End Sub

Private Function Set_Small(pWidth As Integer, pHeight As Integer, pSource As PictureBox, pDestination As PictureBox)
    
    Dim ix As Single, iy As Single
    Dim nx As Single, ny As Single, xcounter As Integer, ycounter As Integer
    
    pSource.Parent.ScaleMode = vbPixels
    pDestination.Parent.ScaleMode = vbPixels
    pSource.ScaleMode = vbPixels
    pDestination.ScaleMode = vbPixels
    
    ix = pSource.ScaleWidth / pWidth
    iy = pSource.ScaleHeight / pHeight
    
    pDestination.Height = pSource.ScaleHeight / iy
    pDestination.Width = pSource.ScaleWidth / ix

    For ny = 0 To pSource.ScaleHeight - 1 Step iy
      For nx = 0 To pSource.ScaleWidth - 1 Step ix
       pDestination.PSet (xcounter, ycounter), pSource.Point(nx, ny)
       xcounter = xcounter + 1
      Next
     ycounter = ycounter + 1
     xcounter = 0
    Next
    
    Set pDestination.Picture = pDestination.Image

End Function

Private Sub cmRetrieve_Click()
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
 'Line3=Line3... etc...
  
  txInf = Empty
  
  If FileExists(oemIni) Then
   txFld(0) = ReadINI("General", "Manufacturer", oemIni)
   txFld(1) = ReadINI("General", "Model", oemIni)
   txFld(2) = ReadINI("OEMSpecific", "SubModel", oemIni)
   txFld(3) = ReadINI("OEMSpecific", "SerialNo", oemIni)
   txFld(4) = ReadINI("OEMSpecific", "OEM1", oemIni)
   txFld(5) = ReadINI("OEMSpecific", "OEM2", oemIni)
   Call ReadLineValues
  End If


End Sub

Private Sub cmCreate_Click()
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
 'Line3=Line3... etc...
  
  If FileExists(oemIni) Then
   Kill oemIni: StatusBar1.Panels(2) = "INI File Recreated"
  Else
   StatusBar1.Panels(2) = "INI File Created"
  End If
  
  Call WriteINI("General", "Manufacturer", txFld(0), oemIni)
  Call WriteINI("General", "Model", txFld(1), oemIni)
  Call WriteINI("OEMSpecific", "SubModel", txFld(2), oemIni)
  Call WriteINI("OEMSpecific", "SerialNo", txFld(3), oemIni)
  Call WriteINI("OEMSpecific", "OEM1", txFld(4), oemIni)
  Call WriteINI("OEMSpecific", "OEM2", txFld(5), oemIni)

  Call WriteLineValues(txInf)
  
End Sub

Private Sub cmSavePic_Click()
   
  If FileExists(oemBmp) Then
   Kill oemIni: StatusBar1.Panels(1) = "BMP File Recreated"
  Else
   StatusBar1.Panels(1) = "BMP File Created"
  End If
 
 SavePicture pDest.Picture, oemBmp
 
  If FileExists(oemBmp) Then
   StatusBar1.Panels(1) = "BMP File Saved"
  End If
 
  pSrce.Picture = LoadPicture(srcPic)
  pSrce.PaintPicture pSrce.Picture, 0, 0, pSrce.Width, pSrce.Height

End Sub

Private Sub cmShow_Click()
 Call Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,0", 5)
End Sub

Private Sub Form_Load()
  
 oemIni = sSystemDirectory & "oeminfo.ini"
 oemBmp = sSystemDirectory & "oemlogo.bmp"
 
 If rREG(HKEY_LOCAL_MACHINE, "Software\SL", "Run") = Empty Then
  Call wREG(HKEY_LOCAL_MACHINE, "Software\SL", "Run", 1)
  Call FILE_COPY(oemIni, sSystemDirectory & "oeminfo.bak")
  Call FILE_COPY(oemBmp, sSystemDirectory & "oemlogo.bak")
 End If
 
End Sub

Function WriteLineValues(objTB As TextBox)

 Dim sLines() As String
 
  sLines = Split(objTB.Text, vbCrLf)
  
  For i = 0 To UBound(sLines)
    Call WriteINI("Support Information", "Line" & i + 1, sLines(i), oemIni)
  Next

End Function

Private Function ReadLineValues()

Dim iFile%
Dim sLine$
Dim rLine$
Dim iCnt%
Dim nCnt%

 iFile = FreeFile

 Open oemIni For Input As iFile
 iCnt = 1
 
  Do While Not EOF(iFile)
   Line Input #iFile, sLine
    If Left(sLine, 4) = "Line" Then
     iCnt = iCnt + 1
    End If
  Loop
  
  For nCnt = 0 To iCnt - 1
   rLine = ReadINI("Support Information", "Line" & nCnt + 1, oemIni)
    If nCnt < iCnt - 1 Then
     txInf = Trim(txInf & ReadINI("Support Information", "Line" & nCnt + 1, oemIni)) & vbCrLf
    Else
     txInf = Trim(txInf & ReadINI("Support Information", "Line" & nCnt + 1, oemIni))
    End If
  Next nCnt
  
  Close iFile

End Function

Private Sub objWMI()

Dim objs
Dim obj
Dim WMI
Dim strMBD$

Set WMI = GetObject("WinMgmts:")
Set objs = WMI.InstancesOf("Win32_BaseBoard")

 For Each obj In objs
  txFld(0) = obj.Manufacturer
  txFld(1) = obj.Product
  txFld(3) = obj.Serialnumber
 Next
 
End Sub

Private Sub Timer1_Timer()

  If FileExists(oemBmp) And FileExists(oemIni) Then
   Image1.Picture = ImageList1.ListImages(3).Picture
  ElseIf FileExists(oemBmp) And Not FileExists(oemIni) Then
   Image1.Picture = ImageList1.ListImages(2).Picture
  ElseIf FileExists(oemIni) And Not FileExists(oemBmp) Then
   Image1.Picture = ImageList1.ListImages(2).Picture
  ElseIf Not FileExists(oemIni) And Not FileExists(oemBmp) Then
   Image1.Picture = ImageList1.ListImages(1).Picture
  End If

End Sub
