VERSION 5.00
Begin VB.Form frmGuildURL 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Oficial Web Site"
   ClientHeight    =   1035
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   6135
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
   ScaleHeight     =   1035
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   255
      Left            =   120
      MouseIcon       =   "frmGuildURL.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   720
      Width           =   5895
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese la direccion del site:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmGuildURL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
If text1 <> "" Then _
    Call SendData("NEWWEBSI" & text1)
Unload Me
End Sub
