VERSION 5.00
Begin VB.Form frmMensaje 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   3990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMensaje.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   217
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   266
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Command1 
      Height          =   285
      Left            =   1320
      Top             =   2790
      Width           =   1365
   End
   Begin VB.Label msg 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3495
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Deactivate()
Me.SetFocus
End Sub
Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("58.gif")
Call Make_Transparent_Form(Me.hwnd, 210)
End Sub
Private Sub command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Command1.Picture = General_Load_Picture_From_Resource("29.gif")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Command1.Picture = LoadPicture("")
End Sub
