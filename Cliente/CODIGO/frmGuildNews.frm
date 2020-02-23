VERSION 5.00
Begin VB.Form frmGuildNews 
   BorderStyle     =   0  'None
   Caption         =   "GuildNews"
   ClientHeight    =   6630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5040
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox news 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   680
      Width           =   4335
   End
   Begin VB.ListBox guerra 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1005
      ItemData        =   "frmGuildNews.frx":0000
      Left            =   360
      List            =   "frmGuildNews.frx":0002
      TabIndex        =   1
      Top             =   3340
      Width           =   4335
   End
   Begin VB.ListBox aliados 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1005
      ItemData        =   "frmGuildNews.frx":0004
      Left            =   360
      List            =   "frmGuildNews.frx":0006
      TabIndex        =   0
      Top             =   4895
      Width           =   4335
   End
   Begin VB.Image Command1 
      Height          =   255
      Left            =   1800
      Top             =   6120
      Width           =   1455
   End
End
Attribute VB_Name = "frmGuildNews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("53.gif")
End Sub
Private Sub Command1_Click()
On Error Resume Next
Call General_Set_Wav(SND_CLICK)
Unload Me
frmMain.SetFocus
End Sub
