VERSION 5.00
Begin VB.Form frmGuildAdm 
   BorderStyle     =   0  'None
   Caption         =   "Lista de Clanes Registrados"
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   4155
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox GuildsList 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2370
      ItemData        =   "frmGuildAdm.frx":0000
      Left            =   390
      List            =   "frmGuildAdm.frx":0002
      TabIndex        =   0
      Top             =   690
      Width           =   3255
   End
   Begin VB.Image Command1 
      Height          =   255
      Left            =   600
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Image Command3 
      Height          =   255
      Left            =   2040
      Top             =   3285
      Width           =   1455
   End
End
Attribute VB_Name = "frmGuildAdm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Call General_Set_Wav(SND_CLICK)
    frmGuildBrief.EsLeader = False
    Call WriteGuildRequestDetails(guildslist.List(guildslist.ListIndex))
End Sub

Private Sub Command3_Click()
Call General_Set_Wav(SND_CLICK)
    Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("48.gif")
End Sub
