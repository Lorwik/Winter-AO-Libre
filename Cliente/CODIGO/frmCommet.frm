VERSION 5.00
Begin VB.Form frmCommet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Oferta de paz o alianza"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4590
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
   ScaleHeight     =   2505
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   120
      MouseIcon       =   "frmCommet.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      Height          =   255
      Left            =   2400
      MouseIcon       =   "frmCommet.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmCommet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Nombre As String
Public T As TIPO
Public Enum TIPO
    ALIANZA = 1
    PAZ = 2
    RECHAZOPJ = 3
End Enum

Public Sub SetTipo(ByVal T As TIPO)
    Select Case T
        Case TIPO.ALIANZA
            Me.Caption = "Detalle de solicitud de alianza"
            Me.text1.MaxLength = 200
        Case TIPO.PAZ
            Me.Caption = "Detalle de solicitud de Paz"
            Me.text1.MaxLength = 200
        Case TIPO.RECHAZOPJ
            Me.Caption = "Detalle de rechazo de membresía"
            Me.text1.MaxLength = 50
    End Select
End Sub


Private Sub Command1_Click()


If text1 = "" Then
    If T = PAZ Or T = ALIANZA Then
        MsgBox "Debes redactar un mensaje solicitando la paz o alianza al líder de " & Nombre
    Else
        MsgBox "Debes indicar el motivo por el cual rechazas la membresía de " & Nombre
    End If
    Exit Sub
End If

If T = PAZ Then
    Call SendData("PEACEOFF" & Nombre & "," & Replace(text1, vbCrLf, "º"))
ElseIf T = ALIANZA Then
    Call SendData("ALLIEOFF" & Nombre & "," & Replace(text1, vbCrLf, "º"))
ElseIf T = RECHAZOPJ Then
    Call SendData("RECHAZAR" & Nombre & "," & Replace(Replace(text1.Text, ",", " "), vbCrLf, " "))
    'Sacamos el char de la lista de aspirantes
    Dim i As Long
    For i = 0 To frmGuildLeader.solicitudes.ListCount - 1
        If frmGuildLeader.solicitudes.List(i) = Nombre Then
            frmGuildLeader.solicitudes.RemoveItem i
            Exit For
        End If
    Next i
    
    Me.Hide
    Unload frmCharInfo
    'Call SendData("GLINFO")
End If
Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

