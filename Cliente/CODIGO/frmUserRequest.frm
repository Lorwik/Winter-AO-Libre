VERSION 5.00
Begin VB.Form frmUserRequest 
   BorderStyle     =   0  'None
   Caption         =   "Peticion"
   ClientHeight    =   3150
   ClientLeft      =   -60
   ClientTop       =   -105
   ClientWidth     =   5010
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
   ScaleHeight     =   210
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   334
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   315
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   765
      Width           =   4335
   End
   Begin VB.Image Command1 
      Height          =   255
      Left            =   1800
      Top             =   2520
      Width           =   1335
   End
End
Attribute VB_Name = "frmUserRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Public Sub recievePeticion(ByVal p As String)

Text1 = Replace$(p, "º", vbCrLf)
Me.Show vbModeless, frmMain

End Sub

Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("88.gif")
End Sub
