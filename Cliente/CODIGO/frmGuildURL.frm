VERSION 5.00
Begin VB.Form frmGuildURL 
   BorderStyle     =   0  'None
   Caption         =   "Oficial Web Site"
   ClientHeight    =   1545
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   6285
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
   ScaleHeight     =   1545
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   280
      TabIndex        =   0
      Top             =   630
      Width           =   5655
   End
   Begin VB.Image Command1 
      Height          =   255
      Left            =   2400
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "frmGuildURL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Call General_Set_Wav(SND_CLICK)
    If Text1 <> "" Then _
        Call WriteGuildNewWebsite(Text1)
    
    Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("55.gif")
End Sub
