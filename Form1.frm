VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Winsock Information"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3285
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   3285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2040
      TabIndex        =   14
      Top             =   2160
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1440
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   1001
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Remote Port:"
      Height          =   255
      Left            =   2040
      TabIndex        =   13
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "State:"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Host:"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Remote Host:"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Socket:"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Port:"
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "IP Address:"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
Text1.Text = Winsock1.LocalIP
Text2.Text = Winsock1.LocalHostName
Text3.Text = Winsock1.SocketHandle
Text4.Text = Winsock1.State
Text5.Text = Winsock1.LocalPort
Text6.Text = Winsock1.RemoteHost
Text7.Text = Winsock1.RemotePort
End Sub
