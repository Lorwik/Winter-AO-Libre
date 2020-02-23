VERSION 5.00
Begin VB.Form frmSpawnList 
   BorderStyle     =   0  'None
   Caption         =   "Invocar"
   ClientHeight    =   4170
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   4305
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   278
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   287
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstCriaturas 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2370
      Left            =   840
      TabIndex        =   0
      Top             =   930
      Width           =   2445
   End
   Begin VB.Image Command1 
      Height          =   255
      Left            =   720
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Image Command2 
      Height          =   255
      Left            =   2280
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "frmSpawnList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Call WriteSpawnCreature(lstCriaturas.ListIndex + 1)
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Deactivate()
    'Me.SetFocus
End Sub
Private Sub form_load()
Me.Picture = General_Load_Picture_From_Resource("66.gif")
End Sub
