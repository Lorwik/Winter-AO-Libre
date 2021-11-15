VERSION 5.00
Begin VB.Form frmGuildNews 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "GuildNews"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Clanes aliados"
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   4575
      Begin VB.ListBox aliados 
         Height          =   1035
         ItemData        =   "frmGuildNews.frx":0000
         Left            =   120
         List            =   "frmGuildNews.frx":0002
         TabIndex        =   6
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clanes con los que estamos en guerra"
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   4575
      Begin VB.ListBox guerra 
         Height          =   1035
         ItemData        =   "frmGuildNews.frx":0004
         Left            =   120
         List            =   "frmGuildNews.frx":0006
         TabIndex        =   4
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "GuildNews"
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
      Begin VB.TextBox news 
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Height          =   195
      Left            =   240
      MouseIcon       =   "frmGuildNews.frx":0008
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   5880
      Width           =   4335
   End
End
Attribute VB_Name = "frmGuildNews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
On Error Resume Next
Unload Me
frmMain.SetFocus
End Sub

Public Sub ParseGuildNews(ByVal S As String)
news = Replace(ReadField(1, S, Asc("¬")), "º", vbCrLf)

Dim h%, j%

h% = Val(ReadField(2, S, Asc("¬")))

For j% = 1 To h%
    
    guerra.AddItem ReadField(j% + 2, S, Asc("¬"))
    
Next j%

j% = j% + 2

h% = Val(ReadField(j%, S, Asc("¬")))

For j% = j% + 1 To j% + h%
    
    aliados.AddItem ReadField(j%, S, Asc("¬"))
    
Next j%

Me.Show , frmMain

End Sub
