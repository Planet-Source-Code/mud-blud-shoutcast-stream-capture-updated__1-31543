VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Capture SHOUTcast"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3390
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   3390
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox holder 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      ScaleHeight     =   975
      ScaleWidth      =   3135
      TabIndex        =   10
      Top             =   1680
      Width           =   3135
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   855
      End
      Begin VB.Label lServerName 
         Height          =   230
         Left            =   960
         TabIndex        =   17
         Top             =   0
         Width           =   2175
      End
      Begin VB.Label lGenre 
         Height          =   230
         Left            =   960
         TabIndex        =   16
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Genre:"
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lURL 
         Height          =   230
         Left            =   960
         TabIndex        =   14
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "URL:"
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lBitRate 
         Height          =   230
         Left            =   960
         TabIndex        =   12
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Bit Rate:"
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2520
      Top             =   1920
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Disconnect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Connect"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   ".."
      Height          =   285
      Left            =   2880
      TabIndex        =   7
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox tPath 
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   840
      Width           =   1815
   End
   Begin MSComctlLib.StatusBar STB 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2670
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   185
            MinWidth        =   176
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5715
         EndProperty
      EndProperty
   End
   Begin VB.TextBox tPort 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "8000"
      Top             =   480
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2640
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Sck 
      Left            =   2280
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox tHost 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "By MudBlud"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   19
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "File:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   870
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Port:"
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   510
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "IP/Host:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private fullsize As Integer
Private totalbytes As Long
Private clearf As Integer, mp3f As Integer

Private Sub Command1_Click()
CD1.DialogTitle = "Save MP3"
CD1.Filter = "MP3 Files (*.mp3)|*.mp3"
mm:
CD1.ShowSave
If CD1.FileName = "" Then Exit Sub
If Dir$(CD1.FileName) <> "" Then
    Dim ans As Integer
    ans = MsgBox("Are you sure? if click Yes " & Dir$(CD1.FileName) & " will be deleted!", vbCritical + vbYesNo)
    If ans = vbNo Then GoTo mm
End If
tPath.Text = CD1.FileName
End Sub

Private Sub Command2_Click()
If tHost = "" Or tPort = "" Or Not IsNumeric(tPort) Or tPath = "" Then MsgBox "Please fill in all 3 text boxes!", vbCritical + vbOKOnly: Exit Sub
Sck.Connect tHost.Text, CInt(tPort.Text)
End Sub

Private Sub Command3_Click()
Sck.Close
Close #1
End Sub

Private Sub Form_Load()
clearf = FreeFile()
mp3f = FreeFile
fullsize = Me.Height
STB.Panels(1).Text = "Not Connected"
STB.Panels(2).Text = "-"
holder.Visible = False
Command2.Enabled = True
Command3.Enabled = False
Me.Height = holder.Top + (STB.Height * 3)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Close #mp3f
Sck.Close
End Sub

Private Sub Sck_Close()
Close #mp3f
Sck.Close
End Sub

Private Sub Sck_Connect()
Open tPath.Text For Output As #clearf
Print #clearf, , "";
Close #clearf
Open tPath.Text For Binary Access Write As #mp3f
Sck.SendData "GET / HTTP/1.0" & vbCrLf & vbCrLf
End Sub

Public Sub Sck_DataArrival(ByVal bytesTotal As Long)
Static sendnumber As Integer
If sendnumber = 0 Then
    Dim Dta As String
    Sck.GetData Dta, vbString
    If Mid$(Dta, 5, 3) = "200" Then
        sendnumber = sendnumber + 1
    Else
        MsgBox "Server isn't avalible", vbOKOnly + vbCritical
        Close #mp3f
        Sck.Close
    End If
ElseIf sendnumber = 1 Then
    Dim Dta1 As String
    Sck.GetData Dta1, vbString
    If Mid$(Dta1, 10, InStr(Dta1, vbCrLf) - 10) = "" Then lServerName.Caption = "Not Specified" Else lServerName.Caption = Mid$(Dta1, 10, InStr(Dta1, vbCrLf) - 10)
    If Mid$(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), 11, InStr(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), vbCrLf) - 11) = "" Then lGenre.Caption = "Not Specified" Else lGenre.Caption = Mid$(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), 11, InStr(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), vbCrLf) - 11)
    If Mid$(Mid$(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), InStr(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), vbCrLf) + 2), 9, InStr(Mid$(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), InStr(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), vbCrLf) + 2), vbCrLf) - 9) = "" Then lURL.Caption = "Not Specified" Else lURL.Caption = Mid$(Mid$(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), InStr(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), vbCrLf) + 2), 9, InStr(Mid$(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), InStr(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), vbCrLf) + 2), vbCrLf) - 9)
    If Mid$(Mid$(Mid$(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), InStr(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), vbCrLf) + 2), InStr(Mid$(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), InStr(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), vbCrLf) + 2), vbCrLf) + 2), InStr(Mid$(Mid$(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), InStr(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), vbCrLf) + 2), InStr(Mid$(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), InStr(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), vbCrLf) + 2), vbCrLf) + 2), vbCrLf) + 2) = "" Then lBitRate.Caption = "Not Specified" Else lBitRate.Caption = Mid$(Mid$(Mid$(Mid$(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), InStr(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), vbCrLf) + 2), InStr(Mid$(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), InStr(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), vbCrLf) + 2), vbCrLf) + 2), InStr(Mid$(Mid$(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), InStr(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), vbCrLf) + 2), InStr(Mid$(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), InStr(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), vbCrLf) + 2), vbCrLf) + 2), vbCrLf) + 2) _
, 8, InStr(Mid$(Mid$(Mid$(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), InStr(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), vbCrLf) + 2), InStr(Mid$(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), InStr(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), vbCrLf) + 2), vbCrLf) + 2), InStr(Mid$(Mid$(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), InStr(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), vbCrLf) + 2), InStr(Mid$(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), InStr(Mid$(Dta1, InStr(Dta1, vbCrLf) + 2), vbCrLf) + 2), vbCrLf) + 2), vbCrLf) + 2), vbCrLf) - 8) & " kbps"
    sendnumber = sendnumber + 1
ElseIf sendnumber = 2 Then
    Dim Data() As Byte
    Sck.GetData Data(), vbByte
    Put #mp3f, , Data()
    STB.Panels(2).Text = "Recieved: " & DoConv(totalbytes + bytesTotal)
    totalbytes = totalbytes + bytesTotal
End If
End Sub

Private Sub Sck_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Close #mp3f
Sck.Close
End Sub

Private Sub Timer1_Timer()
Static olds As Integer
If olds = Sck.State Then Exit Sub
Select Case Sck.State
Case sckConnected: STB.Panels(1).Text = "Connected": Command3.Enabled = True: Command2.Enabled = False: holder.Visible = True: Me.Height = fullsize: tHost.Enabled = False: tPort.Enabled = False: tPath.Enabled = False: Command1.Enabled = False
Case sckConnecting: STB.Panels(1).Text = "Connecting...": Command3.Enabled = False: Command2.Enabled = False: holder.Visible = False: STB.Panels(2).Text = "-": Me.Height = holder.Top + (STB.Height * 3): tHost.Enabled = False: tPort.Enabled = False: tPath.Enabled = False: Command1.Enabled = False
Case sckClosed: STB.Panels(1).Text = "Not Connected": Command3.Enabled = False: Command2.Enabled = True: holder.Visible = False: STB.Panels(2).Text = "-": Me.Height = holder.Top + (STB.Height * 3): totalbytes = 0: tHost.Enabled = True: tPort.Enabled = True: tPath.Enabled = True: Command1.Enabled = True
End Select
olds = Sck.State
End Sub

Private Function DoConv(Number As Long) As String
If Number < 1024 Then DoConv = CStr(Number) & " B"
If Number >= 1024 And Number < (1024& * 1024&) Then DoConv = CStr(Round(Number / 1024, 2)) & " KB"
If Number >= (1024& * 1024&) And Number < (1024& * 1024& * 1024&) Then DoConv = CStr(Round(Number / 1024 / 1024, 2)) & " MB"
If Number >= (1024& * 1024& * 1024&) Then DoConv = CStr(Round(Number / 1024 / 1024 / 1024, 2)) & " GB"
End Function
