VERSION 5.00
Begin VB.Form frmSpawnList 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Invocar"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   2775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Spawn"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   105
      MouseIcon       =   "frmSpawnList.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2760
      Width           =   1650
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      MouseIcon       =   "frmSpawnList.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2760
      Width           =   810
   End
   Begin VB.ListBox lstCriaturas 
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2490
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Selecciona la criatura:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   405
      TabIndex        =   3
      Top             =   75
      Width           =   1935
   End
End
Attribute VB_Name = "frmSpawnList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
Call SendData("SPA" & lstCriaturas.listIndex + 1)
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Deactivate()
'Me.SetFocus
End Sub

