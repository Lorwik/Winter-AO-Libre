VERSION 5.00
Begin VB.Form frmGuildSol 
   BorderStyle     =   0  'None
   Caption         =   "Ingreso"
   ClientHeight    =   3420
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   4785
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
   ScaleHeight     =   3420
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   480
      MaxLength       =   400
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1960
      Width           =   3735
   End
   Begin VB.Image Command2 
      Height          =   255
      Left            =   3120
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Image Command1 
      Height          =   255
      Left            =   240
      Top             =   3000
      Width           =   1455
   End
End
Attribute VB_Name = "frmGuildSol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CName As String

Private Sub Command1_Click()
Call General_Set_Wav(SND_CLICK)
    Call WriteGuildRequestMembership(CName, Replace(Replace(Text1.Text, ",", ";"), vbCrLf, "º"))

    Unload Me
End Sub

Private Sub Command2_Click()
Call General_Set_Wav(SND_CLICK)
    Unload Me
End Sub

Public Sub RecieveSolicitud(ByVal GuildName As String)
    CName = GuildName
End Sub

Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("54.gif")
End Sub
