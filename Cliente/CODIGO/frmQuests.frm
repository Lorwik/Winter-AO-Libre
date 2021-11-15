VERSION 5.00
Begin VB.Form frmQuests 
   BorderStyle     =   0  'None
   Caption         =   "Misiones"
   ClientHeight    =   6345
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5025
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
   Picture         =   "frmQuests.frx":0000
   ScaleHeight     =   6345
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstQuests 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1200
      ItemData        =   "frmQuests.frx":15F5A
      Left            =   120
      List            =   "frmQuests.frx":15F5C
      TabIndex        =   0
      Top             =   600
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   3720
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Image cmdAbandonar 
      Height          =   255
      Left            =   120
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label lblCriaturas 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   5640
      Width           =   3375
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Criaturas matadas:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   5640
      Width           =   3855
   End
   Begin VB.Label lblDescripcion 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   4815
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   2040
      Width           =   3855
   End
End
Attribute VB_Name = "frmQuests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAbandonar_Click()
    If lstQuests.listIndex < 1 Or lstQuests.List(lstQuests.listIndex) = "-" Then Exit Sub
    
    If MsgBox("¿Estás seguro que deseas abandonar la misión " & Chr(34) & lstQuests.List(lstQuests.listIndex) & Chr(34) & "?", vbCritical + vbYesNo, "Argentum Online") = vbYes Then
        Call SendData("QA" & lstQuests.listIndex + 1)
    End If
End Sub
Private Sub cmdCerrar_Click()
    Unload Me
End Sub
Private Sub Form_Load()


                 Call Audio.PlayWave("187.wav")

    If lstQuests.List(0) <> "-" Then
        Call SendData("QIR1")
    End If
End Sub
Private Sub Image1_Click()

                 Call Audio.PlayWave("188.wav")

Unload Me
End Sub
Private Sub lstQuests_Click()
    If lstQuests.listIndex < 1 Then Exit Sub
    
    If lstQuests.List(lstQuests.listIndex) = "-" Then
        lblCriaturas.Caption = "-"
        lblDescripcion.Caption = "-"
        lblNombre.Caption = "-"
    Else
        Call SendData("QIR" & lstQuests.listIndex + 1)

    End If
End Sub
