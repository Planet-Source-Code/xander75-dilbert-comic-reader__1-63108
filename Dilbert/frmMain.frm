VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00663300&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dilbert Comic Reader"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9720
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   9720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDesktop 
      Caption         =   "Set as Wallpaper"
      Height          =   300
      Left            =   8160
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker dtpArchive 
      Height          =   300
      Left            =   8160
      TabIndex        =   4
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   16515073
      CurrentDate     =   39562
   End
   Begin VB.PictureBox picTemp 
      Height          =   615
      Left            =   4320
      ScaleHeight     =   555
      ScaleWidth      =   1155
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin SHDocVwCtl.WebBrowser wbrDilbert 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   885
      Width           =   9720
      ExtentX         =   17145
      ExtentY         =   7223
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin InetCtlsObjects.Inet InetDilbert 
      Left            =   9000
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   0
      Picture         =   "frmMain.frx":08CA
      ScaleHeight     =   885
      ScaleWidth      =   3360
      TabIndex        =   1
      Top             =   0
      Width           =   3360
   End
   Begin VB.Label lblChoose 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Date:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7080
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim dtDate As String
    Dim iPos As Long
    Dim WebText As String
    Dim WebURL As String
    Dim sDay As String

    Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
    Private Const SPI_SETDESKWALLPAPER = 20
    Private Const SPIF_SENDWININICHANGE = &H2
    Private Const SPIF_UPDATEINIFILE = &H1

Private Sub dtpArchive_Change()
    Call dtpArchive_Click
End Sub

Private Sub dtpArchive_Click()
On Error GoTo Errhandler
    Screen.MousePointer = vbHourglass
    dtDate = Format(dtpArchive.Value, "yyyy-mm-dd")
    WebURL = "http://dilbert.com/strips/comic/" & dtDate & "/"
    Call GetHTML(WebURL)

    iPos = InStr(WebText, "/dyn/str_strip/")
    If iPos = 0 Then
        Screen.MousePointer = vbNormal
        MsgBox "There is no comic strip for the date selected!", vbInformation + vbOKOnly, "No Comic Strip"
        Exit Sub
    End If
    
    Dim iPos2 As Integer
    WebText = "http://dilbert.com" & Mid(WebText, iPos, 100)
    iPos2 = InStr(WebText, ".gif")
    If iPos2 <> 0 Then
        WebText = Left(WebText, iPos2 + 3)
    Else
        iPos2 = InStr(WebText, ".jpg")
        WebText = Left(WebText, iPos2 + 3)
    End If

    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2, 8800, 4500
    
    wbrDilbert.Navigate WebText
    Call SaveDilbertPicture
    Screen.MousePointer = vbNormal
    Exit Sub

Errhandler:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbInformation + vbOKOnly, "Load Error"
End Sub

Private Sub Form_Load()
On Error GoTo Errhandler
    
    dtpArchive.Value = Date
    
    dtDate = Format(Date, "yyyy-mm-dd")
    WebURL = "http://dilbert.com/strips/comic/" & dtDate & "/"
    Call GetHTML(WebURL)

    iPos = InStr(WebText, "/dyn/str_strip/")
    If iPos = 0 Then
        Screen.MousePointer = vbNormal
        MsgBox "There is no comic strip for the date selected!", vbInformation + vbOKOnly, "No Comic Strip"
        Exit Sub
    End If
    
    Dim iPos2 As Integer
    WebText = "http://dilbert.com/" & Mid(WebText, iPos, 100)
    iPos2 = InStr(WebText, ".gif")
    If iPos2 <> 0 Then
        WebText = Left(WebText, iPos2 + 3)
    Else
        iPos2 = InStr(WebText, ".jpg")
        WebText = Left(WebText, iPos2 + 3)
    End If
       
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2, 8800, 4500
        
    wbrDilbert.Navigate WebText
    Call SaveDilbertPicture
    Exit Sub

Errhandler:
    MsgBox Err.Description, vbInformation + vbOKOnly, "Load Error"
End Sub

Private Function GetHTML(url$) As String
    Dim response$
    Dim vData As Variant
    
    InetDilbert.Cancel
    response = InetDilbert.OpenURL(url)
    If response <> "" Then
        Do
            vData = InetDilbert.GetChunk(1024, icString)
            DoEvents
            If Len(vData) Then
                response = response & vData
            End If
        Loop While Len(vData)
    End If
    WebText = response
End Function

Private Sub cmdDesktop_Click()
    Dim WallPaper As Long
    SavePicture picTemp.Picture, "c:\dilbert desktop.bmp"
    WallPaper = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, "c:\dilbert desktop.bmp", SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE)   'Change the wallpaper.
End Sub

Private Sub SaveDilbertPicture()
    Dim Dilbert() As Byte
    Dilbert() = InetDilbert.OpenURL(WebText, icByteArray) ' Download picture.
    Open "C:\dilbert.gif" For Binary Access Write As #1 ' Save the file.
    Put #1, , Dilbert()
    Close #1

    picTemp.Picture = LoadPicture("c:\dilbert.gif") 'Reload it To PictureBox
    SavePicture picTemp.Picture, "c:\dilbert.bmp" 'Converted To bmp..
    Kill "c:\dilbert.gif"
End Sub

Private Sub Form_Resize()
    cmdDesktop.Left = Me.Width - 1650
    dtpArchive.Left = Me.Width - 1650
    lblChoose.Left = Me.Width - 2730
End Sub
