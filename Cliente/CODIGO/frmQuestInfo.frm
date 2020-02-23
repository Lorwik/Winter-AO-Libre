VERSION 5.00
Begin VB.Form frmQuestInfo 
   BorderStyle     =   0  'None
   Caption         =   "Información de la misión"
   ClientHeight    =   5985
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label txtInfo 
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
      Height          =   4575
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   3495
   End
   Begin VB.Image cmdAceptar 
      Height          =   255
      Left            =   360
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Image cmdRechazar 
      Height          =   255
      Left            =   2760
      Top             =   5520
      Width           =   1455
   End
End
Attribute VB_Name = "frmQuestInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Maneja el click del CommandButton.
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Call WriteQuestAccept
    Unload Me
    
    Call General_Set_Wav("255.wav")
End Sub

Private Sub cmdRechazar_Click()
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Maneja el click del CommandButton.
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("71.gif")
Call General_Set_Wav("257.wav")
End Sub

