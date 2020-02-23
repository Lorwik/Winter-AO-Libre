VERSION 5.00
Begin VB.Form frmGM 
   BorderStyle     =   0  'None
   Caption         =   "Formulario de mensaje a administradores"
   ClientHeight    =   6060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optConsulta 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   3350
      TabIndex        =   3
      Top             =   2640
      Width           =   255
   End
   Begin VB.OptionButton optConsulta 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1840
      TabIndex        =   2
      Top             =   2640
      Width           =   255
   End
   Begin VB.OptionButton optConsulta 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   2640
      Value           =   -1  'True
      Width           =   195
   End
   Begin VB.TextBox TXTMessage 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   270
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2910
      Width           =   4215
   End
   Begin VB.Image CMDSalir 
      Height          =   255
      Left            =   2400
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Image ADV 
      Height          =   1905
      Left            =   240
      Top             =   510
      Width           =   4275
   End
   Begin VB.Image CMDEnviar 
      Height          =   255
      Left            =   960
      Top             =   5400
      Width           =   1455
   End
End
Attribute VB_Name = "frmGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'By Lorwik
'Form encargado de enviar el mensaje al GM :)
'************************************************************

Private Sub CMDEnviar_Click()
Call General_Set_Wav(SND_CLICK)
    If TXTMessage.Text = "" Then
        MsgBox "Debes de escribir el motivo de tu consulta."
        Exit Sub
    End If
    
    If optConsulta(0).value = True Then
        Call WriteGMRequest(0, TXTMessage.Text)
        Unload Me
        Exit Sub
    ElseIf optConsulta(1).value = True Then
        Call WriteGMRequest(1, TXTMessage.Text)
        Unload Me
        Exit Sub
    ElseIf optConsulta(2).value = True Then
        Call WriteGMRequest(2, TXTMessage.Text)
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub CMDSalir_Click()
Call General_Set_Wav(SND_CLICK)
Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("93.gif")
ADV.Picture = General_Load_Picture_From_Resource("16.gif")
End Sub

Private Sub optConsulta_Click(Index As Integer)
Select Case Index
    Case 0
        ADV.Picture = General_Load_Picture_From_Resource("16.gif")
    Case 1
        ADV.Picture = General_Load_Picture_From_Resource("94.gif")
    Case 2
        ADV.Picture = General_Load_Picture_From_Resource("15.gif")
End Select
End Sub

Private Sub CMDVolver_Click()
    Unload Me
End Sub
