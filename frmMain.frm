VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Internet Web Server"
   ClientHeight    =   6975
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.CmdButton vote 
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   6480
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Click here to Vote"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      FCOL            =   0
      FCOLO           =   0
      MPTR            =   0
      MICON           =   "frmMain.frx":0000
   End
   Begin Project1.CmdButton cmdStart 
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   6000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Start Services"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      FCOL            =   0
      FCOLO           =   0
      MPTR            =   0
      MICON           =   "frmMain.frx":001C
   End
   Begin VB.CommandButton CmdSet 
      Caption         =   "*"
      Height          =   315
      Left            =   5520
      TabIndex        =   19
      Top             =   6000
      Width           =   255
   End
   Begin VB.TextBox ButtonType 
      Height          =   285
      Left            =   5280
      TabIndex        =   18
      Text            =   "8"
      Top             =   6000
      Width           =   255
   End
   Begin Project1.CmdButton cmdPreview 
      Height          =   375
      Left            =   3000
      TabIndex        =   17
      Top             =   6000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Preview Site"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      FCOL            =   0
      FCOLO           =   0
      MPTR            =   0
      MICON           =   "frmMain.frx":0038
   End
   Begin Project1.CmdButton cmdStop 
      Height          =   375
      Left            =   1560
      TabIndex        =   16
      Top             =   6000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Stop Services"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      FCOL            =   0
      FCOLO           =   0
      MPTR            =   0
      MICON           =   "frmMain.frx":0054
   End
   Begin VB.Frame Frame4 
      Caption         =   "Full Client Log"
      Height          =   2295
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   5655
      Begin VB.TextBox txtLog 
         Height          =   1935
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Connection Log"
      Height          =   2055
      Left            =   3000
      TabIndex        =   8
      Top             =   1560
      Width           =   2775
      Begin VB.ListBox lstLog 
         Height          =   1425
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblhits 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   480
         TabIndex        =   25
         Top             =   1680
         Width           =   90
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hits:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   315
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Setup"
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   2775
      Begin VB.TextBox txtRoot 
         Height          =   285
         Left            =   600
         TabIndex        =   15
         Text            =   "Index.html"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtRootDirectory 
         Height          =   325
         Left            =   1200
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtPort 
         Height          =   325
         Left            =   600
         TabIndex        =   5
         Text            =   "80"
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Root"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label lblsite 
         Alignment       =   2  'Center
         Caption         =   "Your Sites IP"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Root Directory"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Port"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   285
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5655
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   4080
         Top             =   600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.PictureBox picConnected 
         Height          =   135
         Left            =   600
         ScaleHeight     =   75
         ScaleWidth      =   4395
         TabIndex        =   1
         Top             =   480
         Visible         =   0   'False
         Width           =   4455
         Begin VB.PictureBox picProgress 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   135
            TabIndex        =   2
            Top             =   0
            Width           =   135
         End
      End
      Begin VB.Image imgClient 
         Height          =   480
         Left            =   5040
         Picture         =   "frmMain.frx":0070
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmMain.frx":04B2
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HTTP Internet Web Server v1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   5415
   End
   Begin VB.Label cmdMinimize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      TabIndex        =   23
      Top             =   0
      Width           =   255
   End
   Begin VB.Label cmdClose 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5640
      TabIndex        =   22
      Top             =   0
      Width           =   255
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   0
      Y1              =   120
      Y2              =   6960
   End
   Begin VB.Line Line2 
      X1              =   5880
      X2              =   5880
      Y1              =   240
      Y2              =   6960
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   5880
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Button Type 1-8"
      Height          =   375
      Left            =   4320
      TabIndex        =   20
      Top             =   6000
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal iparam As Long) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()

Private Sub Form_Load()
txtRootDirectory = App.Path & "\Root"
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Drag(Me)
End Sub

Private Sub cmdClose_Click()
End
End Sub

Private Sub cmdMinimize_Click()
'Window state has three possible values.
'
'vbMinimized - Minimize the form
'vbMaximize - Maximize the form
'vbNormal - Normalize the form

WindowState = vbMinimized
End Sub

Private Function Drag(TheForm As Form)
ReleaseCapture
Call SendMessage(TheForm.HWnd, &HA1, 2, 0&)
End Function

Private Sub CmdSet_Click()
Dim Num(1 To 8) As Integer

Num1 = [Flat Highlight]
Num2 = [Java metal]
Num3 = Mac
Num4 = [Netscape 6]
Num5 = [Simple Flat]
Num6 = [Windows 16-bit]
Num7 = [Windows 32-bit]
Num8 = [Windows XP]

cmdStart.ButtonType = ButtonType.Text
cmdStop.ButtonType = ButtonType.Text
cmdPreview.ButtonType = ButtonType.Text
vote.ButtonType = ButtonType.Text
End Sub

Public Function FileExists(FullFileName As String) As Boolean
    On Error Resume Next
    
    Open FullFileName For Input As #1
    Close #1
    
    If Err = 0 Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function

Public Function RemoveIpDuplicates()
Dim IP As Integer
IP = 0

Do While IP < lstLog.ListCount
    
    lstLog.Text = lstLog.List(IP)


    If lstLog.ListIndex <> IP Then
        lstLog.RemoveItem IP
    Else
        IP = IP + 1
    End If
Loop
End Function

Private Sub cmdStart_Click()
cmdStart.Enabled = False
cmdStop.Enabled = True
cmdPreview.Enabled = True
Frame2.Enabled = False
txtPort.Enabled = False
txtRootDirectory.Enabled = False
txtRoot.Enabled = False
'Start the services
Winsock1.LocalPort = txtPort.Text
Winsock1.Listen
Call checkport
End Sub

Private Sub cmdStop_Click()
cmdStart.Enabled = True
cmdStop.Enabled = False
cmdPreview.Enabled = False
Frame2.Enabled = True
txtPort.Enabled = True
txtRootDirectory.Enabled = True
txtRoot.Enabled = True
'Disable Services
Winsock1.Close
Call checkport
End Sub

Private Sub cmdPreview_Click()
If txtPort.Text = "80" Then
URL "http://" & Winsock1.LocalIP
Else
URL "http://" & Winsock1.LocalIP & ":" & txtPort.Text
End If
End Sub

Private Sub vote_Click()
URL "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=49928&lngWId=1"
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
If Winsock1.State <> sckckosed Then Winsock1.Close
Winsock1.Accept requestID
imgClient.Visible = True
lstLog.AddItem Winsock1.RemoteHostIP & " Connected"
Call RemoveIpDuplicates
lblhits.Caption = lblhits + 1
txtLog.Text = ""
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error GoTo Err
Dim strRequest As String
Dim strPath As String
Dim FileData As String
Dim ZapData As String

Winsock1.GetData strRequest
Debug.Print strRequest
    If strRequest = "" Then
        Winsock1.Close
        Exit Sub
    End If
    If Left(strRequest, 3) <> "GET" Then
        FileData = "<body bgcolor=#ffffff text=#000000 scroll=no><font size=1 face=tahoma><center>Sorry But Only GET Requests<br><br>---------------------------------------------------------------------------<br>" + Text2.Text
        GoTo SendFile
    End If
txtLog.Text = txtLog.Text & strRequest
 
strPath = Mid(strRequest, 5, InStr(5, strRequest, " ") - 5)

If Right(strPath, 1) = "/" Then 'User wants homepage
strPath = strPath & txtRoot
End If

    If FileExists(txtRootDirectory.Text & strPath) = True Then
        Open txtRootDirectory.Text & strPath For Binary Access Read As #1
        FileData = Input(LOF(1), 1)
        Close #1
    Else
        Err.Raise 53
    End If
    
SendFile:
    
    ZapData = _
    "HTTP/1.1 200 OK" & vbCrLf & _
    "Server: HTTP Internet WebServer" & vbCrLf & _
    "Connection: close" & vbCrLf & _
    "Content-Type: application/x-msdownload" & vbCrLf & _
    vbCrLf & FileData
    Winsock1.SendData ZapData
    'Enable Progress Counter
    imgClient.Visible = True
    picConnected.Visible = True
    picProgress.Visible = True
    picProgress.Width = 1
   Exit Sub
    
Err:
    FileData = "<body bgcolor=#ffffff text=#000000 scroll=no><font size=1 face=tahoma><center>You Have Come To A <b>404 Error</b><br>Please Contact The Admin Of This Site<br>And Report The Page You Were Trying To Acces At:<br><b>"
    GoTo SendFile
End Sub

Private Sub Winsock1_SendComplete()
Winsock1.Close
Winsock1.Listen
imgClient.Visible = False
picProgress.Width = 1
picConnected.Visible = False

End Sub

Private Sub Winsock1_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
Progress = bytesSent * picConnected.Width / (bytesSent + bytesRemaining)
picProgress.Width = Progress

End Sub

Sub URL(URL As String)
'this opens a website in IE
On Error GoTo someerror
Shell ("C:\Program Files\Internet Explorer\IEXPLORE.EXE " + URL), vbMaximizedFocus
Exit Sub
someerror:
Beep
Exit Sub
End Sub

Sub checkport()
If txtPort.Text = "80" Then
lblsite.Caption = "http://" & Winsock1.LocalIP
Else
lblsite.Caption = "http://" & Winsock1.LocalIP & ":" & txtPort.Text
End If
End Sub
