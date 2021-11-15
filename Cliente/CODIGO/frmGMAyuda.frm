VERSION 5.00
Begin VB.Form frmGMAyuda 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Formulario de mensaje a administradores"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar sin Enviar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   4680
      Width           =   1695
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Consulta regular"
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
      TabIndex        =   7
      Top             =   1200
      Width           =   1515
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Denuncia"
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
      Left            =   2880
      TabIndex        =   6
      Top             =   1200
      Width           =   1575
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Reporte de bug"
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
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Sugerencias"
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
      Index           =   3
      Left            =   2880
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtMotivo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGMAyuda.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   3900
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGMAyuda.frx":009D
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmGMAyuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Ot As String * 3
Private Sub Command1_Click()
If Ot <> "Con" And Ot <> "Den" And Ot <> "Bug" And Ot <> "Sug" Then
    MsgBox "Debes especificar el tipo de mensaje", vbCritical
    Exit Sub
ElseIf Ot = "Con" Then
    Call SendData("/GX " & txtMotivo.Text)
ElseIf Ot = "Den" Then
    Call SendData("/DENUNCIAR " & txtMotivo.Text)
ElseIf Ot = "Bug" Then
    Call SendData("/BUG " & txtMotivo.Text)
ElseIf Ot = "Sug" Then
    Call SendData("/SUG " & txtMotivo.Text)
End If

    MsgBox "Tu mensaje será guardado en nuestra base de datos. A su vez sera enviado ahora a los GMs Online.", vbInformation
        Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Label1.Caption = "Por favor, escribe el mensaje para el administrador. Un llamado indebido será gravemente penado. Si tu consulta no está bien formulada, no se recibirá el mensaje. Recuerda, antes de llamar, leer la ayuda que se encuentra en opciones junto con el manual del juego."
Label2.Caption = "Tu mensaje será guardado en nuestra base de datos. Si no hay ningún administrador conectado, quedará almacenado. Todo mensaje mal formado será eliminado."
txtMotivo.Text = "Escriba Aqui Su Mensaje."
End Sub
Private Sub optConsulta_Click(Index As Integer)
Select Case Index
    Case 0
         frmGMAyuda.txtMotivo.Visible = True
        Label2.Caption = "¡Por favor explique correctamente el motivo de su consulta!"
        Ot = "Con"
    Case 1
    frmGMAyuda.txtMotivo.Visible = True
        Label2.Caption = "¡Por favor explique correctamente el motivo de su Denuncia lo mas detalladamente posible!"
        Ot = "Den"
    Case 2
    frmGMAyuda.txtMotivo.Visible = True
        Label2.Caption = "Se dará prioridad a su consulta enviando un mensaje a los administradores conectados, por favor utilize ésta opción responsablemente."
        Ot = "Bug"
    Case 3
    frmGMAyuda.txtMotivo.Visible = True
        Label2.Caption = "Su sugerencia SERÁ leída por un miembro del staff, y será tomada en cuenta para futuros cambios."
        Ot = "Sug"
End Select
End Sub

