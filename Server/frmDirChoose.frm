VERSION 5.00
Begin VB.Form frmDirChoose 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose Directory"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2685
   Icon            =   "frmDirChoose.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   2685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C0C0&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Left            =   80
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   80
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmDirChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Sub cmdOK_Click()
frmMain.txtRoot.Text = Dir1.Path
frmMain.Enabled = True
AppActivate frmMain.Caption
Unload Me
End Sub
Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub
Private Sub Form_Load()
On Error Resume Next
SendMessage cmdOK.hWnd, &HF4&, &H0&, 0&
Drive1.Drive = Mid(frmMain.txtRoot.Text, 1, 2)
Dir1.Path = frmMain.txtRoot.Text
TakeOutMenu Me, SC_CLOSE
Left = Screen.Width \ 2 - Width \ 2
Top = Screen.Height \ 2 - Height \ 2
SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
End Sub
