VERSION 5.00
Begin VB.Form frmCrearAccount 
   BorderStyle     =   0  'None
   ClientHeight    =   4665
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   8055
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCrearAccount.frx":0000
   ScaleHeight     =   4665
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox pregunta 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      ItemData        =   "frmCrearAccount.frx":18DE8
      Left            =   720
      List            =   "frmCrearAccount.frx":18DFB
      TabIndex        =   11
      Text            =   "¿Lugar de Nacimiento?"
      Top             =   3240
      Width           =   3015
   End
   Begin VB.TextBox respuesta 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   10
      Top             =   3780
      Width           =   3015
   End
   Begin VB.TextBox CodeKey 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4950
      TabIndex        =   8
      Top             =   3980
      Width           =   2320
   End
   Begin VB.CheckBox Check1 
      Height          =   200
      Left            =   480
      TabIndex        =   6
      Top             =   2760
      Width           =   200
   End
   Begin VB.TextBox Reglamento 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "frmCrearAccount.frx":18E71
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox Mail 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4950
      TabIndex        =   3
      Top             =   2910
      Width           =   2320
   End
   Begin VB.TextBox RePass 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   4950
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1800
      Width           =   2320
   End
   Begin VB.TextBox Pass 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   4950
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1260
      Width           =   2320
   End
   Begin VB.TextBox Nombre 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4950
      MaxLength       =   25
      TabIndex        =   0
      Top             =   750
      Width           =   2320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Respuesta Secreta:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1560
      TabIndex        =   13
      Top             =   3550
      Width           =   1410
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pregunta Secreta:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1560
      TabIndex        =   12
      Top             =   3000
      Width           =   1290
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Acepto los Términos y el reglamento del Juego."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   2760
      Width           =   3495
   End
   Begin VB.Label lblCode 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   7
      Top             =   3530
      Width           =   2415
   End
   Begin VB.Image Command2 
      Height          =   255
      Left            =   360
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Image Command1 
      Height          =   255
      Left            =   2880
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Estado 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Esperando..."
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4950
      TabIndex        =   4
      Top             =   2350
      Width           =   2325
   End
End
Attribute VB_Name = "frmCrearAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Len(Nombre.Text) < 5 Then
    MsgBox "El nombre de la cuenta debe de tener mas de 4 caracteres.", vbCritical
    Exit Sub
End If
        
If Len(Pass.Text) < 6 Then
    MsgBox "El password de la cuenta debe de tener mas de 6 caracteres.", vbCritical
    Exit Sub
End If

If Pass <> RePass Then
    MsgBox "Las passwords que tipeo no coinciden", , "Winter AO"
    Exit Sub
End If

If Not CheckMailString(Mail) Then
    MsgBox "Direccion de mail invalida."
    Exit Sub
End If

If Nombre = "" Or Pass = "" Or RePass = "" Or Mail = "" Or pregunta = "" Or respuesta = "" Then
    MsgBox "Completa todo!"
    Exit Sub
End If

If lblCode.Caption <> CodeKey.Text Then
    MsgBox "El Codigo ingresado es Invalido.", vbCritical
    lblCode.Caption = GenerateKey
    Exit Sub
End If

If Not Check1.value = Checked Then
    MsgBox "Debe Aceptar los términos y Reglamento para poder crear la cuenta.", vbCritical
    Exit Sub
End If

        UserName = Nombre.Text
        
        If Right$(UserName, 1) = " " Then
                UserName = RTrim$(UserName)
                MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
                Exit Sub
        End If

Call SendData("INIFED" & Nombre & "," & Pass & "," & Mail & "," & pregunta & "," & respuesta)

Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("CrearCuenta.gif")
lblCode.Caption = GenerateKey
End Sub

Private Sub Mail_Click()
If Idioma = 1 Then
If Not RePass.Text = "" Then
If Pass.Text = RePass.Text Then
    Estado.ForeColor = vbGreen
    Estado.Caption = "Password Correcto."
        Else
    Estado.ForeColor = vbRed
    Estado.Caption = "Password Incorrecto."
    End If
Else
    Estado.ForeColor = vbWhite
    Estado.Caption = "Esperando..."
End If
Else
If Not RePass.Text = "" Then
If Pass.Text = RePass.Text Then
    Estado.ForeColor = vbGreen
    Estado.Caption = "Password Correcto."
        Else
    Estado.ForeColor = vbRed
    Estado.Caption = "Password Incorrecto."
    End If
Else
    Estado.ForeColor = vbWhite
    Estado.Caption = "Comprobando..."
End If
End If
End Sub

Private Sub nombre_Click()
If Idioma = 1 Then
If Pass.Text = "" Then
    Estado.ForeColor = vbWhite
    Estado.Caption = "Esperando..."
Else
If Pass.Text = RePass.Text Then
    Estado.ForeColor = vbGreen
    Estado.Caption = "Password Correcto."
        Else
    Estado.ForeColor = vbRed
    Estado.Caption = "Password Incorrecto."
    End If
End If
Else
If Pass.Text = "" Then
    Estado.ForeColor = vbWhite
    Estado.Caption = "Comprobando..."
Else
If Pass.Text = RePass.Text Then
    Estado.ForeColor = vbGreen
    Estado.Caption = "Password Correcto."
        Else
    Estado.ForeColor = vbRed
    Estado.Caption = "Password Incorrecto."
    End If
End If
End If
End Sub

Private Sub RePass_Click()
If Idioma = 1 Then
    Estado.ForeColor = vbWhite
    Estado.Caption = "Esperando..."
Else
    Estado.ForeColor = vbWhite
    Estado.Caption = "Comprobando..."
End If
End Sub
