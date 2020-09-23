VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3000
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":058A
   ScaleHeight     =   1500
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
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
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sim Software"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   120
      MouseIcon       =   "frmAbout.frx":4918
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1200
      Width           =   1125
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":4C22
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Personal Webserver"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Sub Command1_Click()
frmMain.Enabled = True
AppActivate frmMain.Caption
Unload Me
End Sub
Private Sub Form_Load()
SendMessage Command1.hWnd, &HF4&, &H0&, 0&
Left = Screen.Width \ 2 - Width \ 2
Top = Screen.Height \ 2 - Height \ 2
SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
TakeOutMenu Me, SC_CLOSE
End Sub
Private Sub Label3_Click()
Call ShellExecute(Me.hWnd, "Open", "mailto:simsoftware@hotmail.com", "", "", 1)
End Sub

