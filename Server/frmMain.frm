VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personal Webserver"
   ClientHeight    =   4410
   ClientLeft      =   4875
   ClientTop       =   4155
   ClientWidth     =   4260
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   4260
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   240
      ScaleHeight     =   3135
      ScaleWidth      =   3735
      TabIndex        =   21
      Top             =   480
      Visible         =   0   'False
      Width           =   3735
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Open/Closed"
         Height          =   790
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   3255
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Temporarily Closed"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   330
            Width           =   2175
         End
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00000000&
         X1              =   3720
         X2              =   3720
         Y1              =   3120
         Y2              =   -120
      End
      Begin VB.Line Line16 
         BorderColor     =   &H00000000&
         X1              =   0
         X2              =   3720
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00FFFFFF&
         X1              =   3720
         X2              =   0
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3120
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C0C0&
      Caption         =   "OK"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   240
      ScaleHeight     =   3135
      ScaleWidth      =   3735
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   3735
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Connections/Requests:"
         Height          =   1695
         Left            =   240
         TabIndex        =   19
         Top             =   1200
         Width           =   3255
         Begin VB.ListBox List1 
            Height          =   1035
            ItemData        =   "frmMain.frx":0E42
            Left            =   120
            List            =   "frmMain.frx":0E44
            TabIndex        =   20
            Top             =   480
            Width           =   3015
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Logging:"
         Height          =   735
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   3255
         Begin VB.CheckBox cheLogging 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Logging"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   330
            Width           =   975
         End
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00000000&
         X1              =   3720
         X2              =   3720
         Y1              =   3120
         Y2              =   -120
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3120
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00FFFFFF&
         X1              =   3720
         X2              =   0
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00000000&
         X1              =   0
         X2              =   3720
         Y1              =   3120
         Y2              =   3120
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   240
      ScaleHeight     =   3135
      ScaleWidth      =   3735
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   3735
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Active Objects:"
         Height          =   1095
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   3255
         Begin VB.CheckBox cheCounter 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Enable counter"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   650
            Width           =   1695
         End
         Begin VB.CheckBox cheGuest 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Enable guestbook"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   330
            Width           =   1935
         End
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000000&
         X1              =   3720
         X2              =   3720
         Y1              =   3120
         Y2              =   -120
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3120
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         X1              =   3720
         X2              =   0
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00000000&
         X1              =   0
         X2              =   3720
         Y1              =   3120
         Y2              =   3120
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   240
      ScaleHeight     =   3135
      ScaleWidth      =   3735
      TabIndex        =   2
      Top             =   480
      Width           =   3735
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Start"
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Server Options:"
         Height          =   1095
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   3255
         Begin VB.CheckBox cheMinimized 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Start minimized"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   330
            Width           =   1815
         End
         Begin VB.CheckBox cheActivate 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Activate server on start"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   650
            Width           =   2535
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Server Directory:"
         Height          =   855
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   3255
         Begin VB.TextBox txtRoot 
            Height          =   285
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton cmdDirChoose 
            Caption         =   "..."
            Height          =   285
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Stop"
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2640
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Server 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   240
         MouseIcon       =   "frmMain.frx":0E46
         TabIndex        =   18
         Top             =   2640
         Width           =   60
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000000&
         X1              =   0
         X2              =   3720
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         X1              =   3720
         X2              =   3720
         Y1              =   3120
         Y2              =   -120
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   3720
         X2              =   0
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3120
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3840
      Left            =   20
      TabIndex        =   0
      Top             =   20
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   6773
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Server"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Active Objects"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Security"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Access"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock sckWS 
      Index           =   0
      Left            =   120
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image ServerOff 
      Height          =   240
      Left            =   2280
      Picture         =   "frmMain.frx":1150
      Top             =   3480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ServerOn 
      Height          =   240
      Left            =   2520
      Picture         =   "frmMain.frx":129A
      Top             =   3480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu mnuTray 
      Caption         =   "&Tray"
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Show S&erver"
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStart 
         Caption         =   "&Start"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private requestedPage As String
Private strdata As String
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub cmdDirChoose_Click()
frmDirChoose.Show ownerform:=Me
frmMain.Enabled = False
End Sub
Private Sub cmdOK_Click()
If FileExists(AddASlash(txtRoot.Text)) = False Then
MsgBox "Please enter a valid path for Server Directory.", vbMsgBoxSetForeground + vbInformation
Exit Sub
End If
htmlPageDir = txtRoot.Text
Me.Hide
End Sub
Private Sub Command1_Click()
load_defaults
Command2.Visible = True
Command1.Visible = False
End Sub
Private Sub Command2_Click()
stop_server
Command1.Visible = True
Command2.Visible = False
End Sub


Private Sub Form_Load()

SendMessage Command1.hWnd, &HF4&, &H0&, 0&
SendMessage Command2.hWnd, &HF4&, &H0&, 0&
SendMessage cmdOK.hWnd, &HF4&, &H0&, 0&
SendMessage cmdDirChoose.hWnd, &HF4&, &H0&, 0&
Dim OS As OSVERSIONINFO
OS.dwOSVersionInfoSize = Len(OS)
GetVersionEx OS
If OS.dwMajorVersion < 4 Then
MsgBox "Sorry. You must have Windows 95, Windows 98, NT4 or later!", vbInformation, "Program closed!"
End
End If
If App.PrevInstance Then 'This checks if webserver is allready started
MsgBox "Sorry, but you have Webserver allready started.", vbMsgBoxSetForeground + vbInformation
End
End If
Left = Screen.Width \ 2 - Width \ 2
Top = Screen.Height \ 2 - Height \ 2
TakeOutMenu Me, SC_CLOSE ', SC_MOVE
gHW = Me.hWnd
myNID.cbSize = Len(myNID)
myNID.hWnd = gHW
myNID.uID = uID
myNID.uFlags = NIF_MESSAGE Or NIF_TIP Or NIF_ICON
myNID.uCallbackMessage = cbNotify
myNID.hIcon = ServerOff
myNID.szTip = "Server Inactive" & Chr(0)
ShellNotifyIcon NIM_ADD, myNID
Hook
SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
ttlConnections = 0 'Set the ttlConnections varible to zero. :)
Server.Caption = "Inactive"
If FileExists(AddASlash(App.Path) & "Webserver.ini") = True Then
Dim Cache As String
Files = FreeFile
Open AddASlash(App.Path) & "Webserver.ini" For Input As #Files
Do While Not EOF(Files)
Line Input #Files, Cache
If Mid(Chache, 1, 1) <> "[" Then
If Mid(Cache, 1, 10) = "ServerRoot" Then
If FileExists(AddASlash(Mid(Cache, 12, Len(Cache)))) = True Then
txtRoot.Text = Mid(Cache, 12, Len(Cache))
Else
txtRoot.Text = App.Path
End If
ElseIf Mid(Cache, 1, 7) = "Logging" Then
If Mid(Cache, 9, 1) = "1" Then
cheLogging.Value = 1
End If
ElseIf Mid(Cache, 1, 9) = "Guestbook" Then
If Mid(Cache, 11, 1) = "1" Then
cheGuest.Value = 1
End If
ElseIf Mid(Cache, 1, 7) = "Counter" Then
If Mid(Cache, 9, 1) = "1" Then
cheCounter.Value = 1
End If
ElseIf Mid(Cache, 1, 9) = "Minimized" Then
If Mid(Cache, 11, 1) = "1" Then
cheMinimized = 1
Me.Hide
End If
ElseIf Mid(Cache, 1, 11) = "TempOffline" Then
If Mid(Cache, 13, 1) = "1" Then
Check1.Value = 1
End If
ElseIf Mid(Cache, 1, 15) = "ActivateOnStart" Then
If Mid(Cache, 17, 1) = "1" Then
cheActivate.Value = 1
load_defaults
Command2.Visible = True
Command1.Visible = False
End If
End If
End If
Loop
Close #Files
Else
txtRoot.Text = App.Path
cheGuest.Value = 1
cheCounter.Value = 1
cheLogging.Value = 1
cheMinimized.Value = 0
cheActivate.Value = 0
End If
htmlPageDir = txtRoot.Text
End Sub
Private Sub Form_Unload(Cancel As Integer)
Call stop_server
Files = FreeFile
Open AddASlash(App.Path) & "Webserver.ini" For Output As Files
Buffer = ""
Buffer = "[Webserver Options]" & vbCrLf
Buffer = Buffer & "ServerRoot=" & txtRoot.Text & vbCrLf
Buffer = Buffer & "Logging=" & cheLogging.Value & vbCrLf
Buffer = Buffer & "Guestbook=" & cheGuest.Value & vbCrLf
Buffer = Buffer & "Counter=" & cheCounter.Value & vbCrLf
Buffer = Buffer & "Minimized=" & cheMinimized & vbCrLf
Buffer = Buffer & "TempOffline=" & Check1.Value & vbCrLf
Buffer = Buffer & "ActivateOnStart=" & cheActivate.Value & vbCrLf
Print #Files, Buffer
Close #Files
SetWindowPos Me.hWnd, -2, 0, 0, 0, 0, 3
Unhook
ShellNotifyIcon NIM_DELETE, myNID
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show ownerform:=Me
frmMain.Enabled = False
End Sub
Private Sub mnuExit_Click()
Unload Me
End Sub
Private Sub mnuFileExit_Click()
Unload Me
End Sub
Private Sub mnuHelpAbout_Click()
frmAbout.Show ownerform:=Me
frmMain.Enabled = False
End Sub
Private Sub mnuOptions_Click()
frmMain.Visible = True
AppActivate frmMain.Caption
End Sub
Private Sub mnuStart_Click()
If mnuStart.Caption = "&Start" Then
load_defaults
Command1.Visible = False
Command2.Visible = True
Else
stop_server
End If
End Sub
Private Sub sckWS_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
If Index = 0 Then
If Check1.Value = 1 Then Exit Sub
If sckWS(ttlConnections).RemoteHostIP = "192.168.0.5" Then Exit Sub
ttlConnections = ttlConnections + 1  'add 1 to the total # of connections
numConnections = numConnections + 1 'number of connected clients + 1
If numConnections = maxConnections Then GoTo done 'if we've reached the max # of connections, exit sub.
Load sckWS(ttlConnections) 'load a new instance of sckWS.
sckWS(ttlConnections).LocalPort = 0 'set its local port to 0
sckWS(ttlConnections).Accept requestID 'Accept the connection request.
List1.AddItem sckWS(ttlConnections).RemoteHostIP & " Connected"
StartOver:
DoEvents 'DoEvents so it doesn't freeze while we wait.
If requestedPage$ = "" Then GoTo StartOver 'if we havent gotten the page request yet, go back to startOver.
List1.AddItem "Requested: " & requestedPage$
If cheLogging.Value = 1 Then
Logging = FreeFile      'This is for the logging function
Open AddASlash(App.Path) & "Log.log" For Append As #Logging
Print #Logging, Format(Date, "Long Date") & " " & Format(Time, "Long Time") & " ; " & sckWS(ttlConnections).RemoteHostIP & "; " & Mid(strdata$, InStr(1, UCase(strdata$), "USER-AGENT:") + 12, InStr(InStr(1, UCase(strdata$), "USER-AGENT:") + 12, UCase(strdata$), vbCrLf) - InStr(1, UCase(strdata$), "USER-AGENT:") - 12) & "; requested Language: " & Mid(strdata$, InStr(1, UCase(strdata$), "ACCEPT-LANGUAGE:") + 17, InStr(InStr(1, UCase(strdata$), "ACCEPT-LANGUAGE:") + 17, UCase(strdata$), vbCrLf) - InStr(1, UCase(strdata$), "ACCEPT-LANGUAGE:") - 17) & "; requested page: " & requestedPage$
Close #Logging
End If
If requestedPage$ = "/" Then
requestedPage$ = htmlIndexPage$ ' if the page '/' was requested, set requested page to the index html page.
Else
requestedPage$ = Mid(requestedPage$, 2, Len(requestedPage$) - 1)
End If
If cheGuest.Value = 1 Then
If UCase(requestedPage$) = "GUESTBOOK.CGI" Then 'This is check if the Guestbook.cgi is requested
NameStart = InStr(UCase(strdata$), "NAME=")
NameEnd = InStr(NameStart + 5, strdata$, "&")
NameValue = Mid$(strdata$, NameStart + 5, NameEnd - (NameStart + 5))
MailStart = InStr(UCase(strdata$), "E-MAIL=")
MailEnd = InStr(MailStart + 7, strdata$, "&")
MailValue = Mid$(strdata$, MailStart + 7, MailEnd - (MailStart + 7))
CommentStart = InStr(UCase(strdata$), "COMMENT=")
CommentEnd = InStr(CommentStart + 8, strdata$, "&")
CommentValue = Mid$(strdata$, CommentStart + 8, CommentEnd - (CommentStart + 8))
CommentValue = ReplaceStr(CommentValue, "+", " ")
CommentValue = ReplaceStr(CommentValue, "%0D%0A", "<br>")
CommentValue = ReplaceStr(CommentValue, "%21", "!")
CommentValue = ReplaceStr(CommentValue, "%22", "&quot;")
CommentValue = ReplaceStr(CommentValue, "%A7", "§")
CommentValue = ReplaceStr(CommentValue, "%24", "$")
CommentValue = ReplaceStr(CommentValue, "%25", "%")
CommentValue = ReplaceStr(CommentValue, "%26", "&")
CommentValue = ReplaceStr(CommentValue, "%2F", "/")
CommentValue = ReplaceStr(CommentValue, "%28", "(")
CommentValue = ReplaceStr(CommentValue, "%29", ")")
CommentValue = ReplaceStr(CommentValue, "%3D", "=")
CommentValue = ReplaceStr(CommentValue, "%3F", "?")
CommentValue = ReplaceStr(CommentValue, "%B2", "²")
CommentValue = ReplaceStr(CommentValue, "%B3", "³")
CommentValue = ReplaceStr(CommentValue, "%7B", "{")
CommentValue = ReplaceStr(CommentValue, "%5B", "[")
CommentValue = ReplaceStr(CommentValue, "%5D", "]")
CommentValue = ReplaceStr(CommentValue, "%7D", "}")
CommentValue = ReplaceStr(CommentValue, "%5C", "\")
CommentValue = ReplaceStr(CommentValue, "%DF", "ß")
CommentValue = ReplaceStr(CommentValue, "%23", "#")
CommentValue = ReplaceStr(CommentValue, "%27", "'")
CommentValue = ReplaceStr(CommentValue, "%3A", ":")
CommentValue = ReplaceStr(CommentValue, "%2C", ",")
CommentValue = ReplaceStr(CommentValue, "%3B", ";")
CommentValue = ReplaceStr(CommentValue, "%60", "`")
CommentValue = ReplaceStr(CommentValue, "%7E", "~")
CommentValue = ReplaceStr(CommentValue, "%2B", "+")
CommentValue = ReplaceStr(CommentValue, "%B4", "´")
MailValue = ReplaceStr(MailValue, "%21", "!")
MailValue = ReplaceStr(MailValue, "%22", "&quot;")
MailValue = ReplaceStr(MailValue, "%A7", "§")
MailValue = ReplaceStr(MailValue, "%24", "$")
MailValue = ReplaceStr(MailValue, "%25", "%")
MailValue = ReplaceStr(MailValue, "%26", "&")
MailValue = ReplaceStr(MailValue, "%2F", "/")
MailValue = ReplaceStr(MailValue, "%28", "(")
MailValue = ReplaceStr(MailValue, "%29", ")")
MailValue = ReplaceStr(MailValue, "%3D", "=")
MailValue = ReplaceStr(MailValue, "%3F", "?")
MailValue = ReplaceStr(MailValue, "%B2", "²")
MailValue = ReplaceStr(MailValue, "%B3", "³")
MailValue = ReplaceStr(MailValue, "%7B", "{")
MailValue = ReplaceStr(MailValue, "%5B", "[")
MailValue = ReplaceStr(MailValue, "%5D", "]")
MailValue = ReplaceStr(MailValue, "%7D", "}")
MailValue = ReplaceStr(MailValue, "%5C", "\")
MailValue = ReplaceStr(MailValue, "%DF", "ß")
MailValue = ReplaceStr(MailValue, "%23", "#")
MailValue = ReplaceStr(MailValue, "%27", "'")
MailValue = ReplaceStr(MailValue, "%3A", ":")
MailValue = ReplaceStr(MailValue, "%2C", ",")
MailValue = ReplaceStr(MailValue, "%3B", ";")
MailValue = ReplaceStr(MailValue, "%60", "`")
MailValue = ReplaceStr(MailValue, "%7E", "~")
MailValue = ReplaceStr(MailValue, "%2B", "+")
MailValue = ReplaceStr(MailValue, "%B4", "´")
NameValue = ReplaceStr(NameValue, "%21", "!")
NameValue = ReplaceStr(NameValue, "%22", "&quot;")
NameValue = ReplaceStr(NameValue, "%A7", "§")
NameValue = ReplaceStr(NameValue, "%24", "$")
NameValue = ReplaceStr(NameValue, "%25", "%")
NameValue = ReplaceStr(NameValue, "%26", "&")
NameValue = ReplaceStr(NameValue, "%2F", "/")
NameValue = ReplaceStr(NameValue, "%28", "(")
NameValue = ReplaceStr(NameValue, "%29", ")")
NameValue = ReplaceStr(NameValue, "%3D", "=")
NameValue = ReplaceStr(NameValue, "%3F", "?")
NameValue = ReplaceStr(NameValue, "%B2", "²")
NameValue = ReplaceStr(NameValue, "%B3", "³")
NameValue = ReplaceStr(NameValue, "%7B", "{")
NameValue = ReplaceStr(NameValue, "%5B", "[")
NameValue = ReplaceStr(NameValue, "%5D", "]")
NameValue = ReplaceStr(NameValue, "%7D", "}")
NameValue = ReplaceStr(NameValue, "%5C", "\")
NameValue = ReplaceStr(NameValue, "%DF", "ß")
NameValue = ReplaceStr(NameValue, "%23", "#")
NameValue = ReplaceStr(NameValue, "%27", "'")
NameValue = ReplaceStr(NameValue, "%3A", ":")
NameValue = ReplaceStr(NameValue, "%2C", ",")
NameValue = ReplaceStr(NameValue, "%3B", ";")
NameValue = ReplaceStr(NameValue, "%60", "`")
NameValue = ReplaceStr(NameValue, "%7E", "~")
NameValue = ReplaceStr(NameValue, "%2B", "+")
NameValue = ReplaceStr(NameValue, "%B4", "´")
NameValue = ReplaceStr(NameValue, "+", " ")
Guestbook = FreeFile
Open AddASlash(App.Path) & "guestbook.ini" For Append As #Guestbook
datastr = "<b><u>Name:</u></b>&nbsp;&nbsp;" & NameValue
datastr = datastr & "&nbsp;&nbsp;&nbsp;<b><u>E-Mail:</u></b>&nbsp;&nbsp;<a href=mailto:" & MailValue
datastr = datastr & ">" & MailValue
datastr = datastr & "</a><br><br><b><u>Comment:</u></b><br>" & CommentValue
datastr = datastr & "<br><br><br><br>"
Print #Guestbook, datastr
Close #Guestbook
strdata$ = ""
requestedPage$ = "guestbook.html"
End If
If UCase(requestedPage$) = "GUESTBOOK.HTML" Then
htmldata$ = html_guestbookstart & vbCrLf & text_read(AddASlash(App.Path) & "guestbook.ini") & vbCrLf & html_guestbookend & vbCrLf
sckWS(ttlConnections).SendData ReplaceStr(htmldata$, "$ip", sckWS(0).LocalIP)
GoTo done
End If
End If
If FileExists(AddASlash(htmlPageDir) & requestedPage$) Then 'if the requested page exists, then..
htmldata$ = text_read(AddASlash(htmlPageDir) & requestedPage$) 'This reads the file and stores it's contents in htmldata$
If cheCounter.Value = 1 Then
If InStr(1, htmldata$, "$counter") <> 0 Then 'Checks if $counter is in the html page
If FileExists(AddASlash(App.Path) & "counter.ini") Then  ' if true the counter will count one up
CountValue = text_read(AddASlash(App.Path) & "counter.ini")
Else
CountValue = "0"
End If
CountValue = CountValue + 1
Counter = FreeFile
Open AddASlash(App.Path) & "counter.ini" For Output As #Counter
Print #Counter, CountValue
Close #Counter
htmldata$ = ReplaceStr(htmldata$, "$counter", Str(CountValue))
End If
End If
htmldata$ = ReplaceStr(htmldata$, "$ip", sckWS(0).LocalIP) 'Oops, i didn't use the replace function right.  Now it's fixed at replaces $ip with your IP.
sckWS(ttlConnections).SendData htmldata$ & vbCrLf  'open and read the requested HTML page.
Else 'if it doesn't exist, then...
If requestedPage$ = htmlIndexPage$ Then 'If the requested page is the index page and it doesn't exist, print this.
sckWS(ttlConnections).SendData "<html><font face=""Verdana, Arial, Helvetica, sans-serif"" size=""1""><b>Please create an index html page.  It was not found.</font></html>" & vbCrLf ' If the requested page is the index and it doesn't exist, it tells you.
requestedPage$ = ""
End If
requestedPage$ = "/a"
sckWS(ttlConnections).SendData html_404$ & vbCrLf 'Send the 404 Error HTML
End If
End If
done:
numConnections = numConnections - 1 'number of connections at the moment - 1
End Sub
Private Sub sckWS_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
sckWS(Index).GetData strdata$ 'Get any data sent to us
'MsgBox strdata$ ' I used this for debugging
If Mid$(strdata$, 1, 3) = "GET" Then 'If it is trying to get a site, find out
findget = InStr(strdata$, "GET ")      ' the site they want then set requestedPage$
spc2 = InStr(findget + 5, strdata$, " ") ' to it.
pagetoget$ = Mid$(strdata$, findget + 4, spc2 - (findget + 4))
requestedPage$ = pagetoget$
ElseIf Mid$(strdata$, 1, 4) = "POST" Then 'This is the code when it is trying to post something!
findpost = InStr(strdata$, "POST ")        'the data where filtered in the ConnectionRequest
spc2 = InStr(findpost + 5, strdata$, " ")   'Function of the winsock control
pagetopost$ = Mid$(strdata$, findpost + 5, spc2 - (findpost + 5))
requestedPage$ = pagetopost$
End If
End Sub
Private Sub sckWS_SendComplete(Index As Integer)
'This was a bug that was fixed from v.2a.
If requestedPage$ <> "" Then 'f the requested page doesn't = nothing then...
requestedPage$ = "" 'clear the requestedPage varible.
sckWS(ttlConnections).Close 'Close the connection.
End If
End Sub
Private Sub Server_Click()
If Mid(Server.Caption, 1, 7) = "http://" Then
Call ShellExecute(Me.hWnd, "Open", Server.Caption, "", "", 1)
End If
End Sub
Private Sub TabStrip1_Click()
If TabStrip1.SelectedItem = "Server" Then Picture1.Visible = True: Picture2.Visible = False: Picture3.Visible = False: Picture4.Visible = False
If TabStrip1.SelectedItem = "Active Objects" Then Picture2.Visible = True: Picture1.Visible = False: Picture3.Visible = False: Picture4.Visible = False
If TabStrip1.SelectedItem = "Security" Then Picture3.Visible = True: Picture1.Visible = False: Picture2.Visible = False: Picture4.Visible = False
If TabStrip1.SelectedItem = "Access" Then Picture4.Visible = True: Picture1.Visible = False: Picture2.Visible = False: Picture3.Visible = False
End Sub
