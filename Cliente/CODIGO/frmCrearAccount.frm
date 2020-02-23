VERSION 5.00
Begin VB.Form frmCrearAccount 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox CodeKey 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4950
      TabIndex        =   10
      Top             =   3980
      Width           =   2340
   End
   Begin VB.TextBox RTPass 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   4950
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   1800
      Width           =   2340
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Height          =   200
      Left            =   360
      MaskColor       =   &H00808080&
      TabIndex        =   7
      Top             =   2640
      Width           =   200
   End
   Begin VB.TextBox Reglamento 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "frmCrearAccount.frx":0000
      Top             =   1080
      Width           =   3615
   End
   Begin VB.TextBox respuesta 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1070
      TabIndex        =   4
      Top             =   3810
      Width           =   2340
   End
   Begin VB.ComboBox pregunta 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      ItemData        =   "frmCrearAccount.frx":1804
      Left            =   1050
      List            =   "frmCrearAccount.frx":1817
      TabIndex        =   3
      Text            =   "¿Lugar de Nacimiento?"
      Top             =   3120
      Width           =   2340
   End
   Begin VB.TextBox TName 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4930
      TabIndex        =   2
      Top             =   750
      Width           =   2340
   End
   Begin VB.TextBox TPass 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   4930
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1260
      Width           =   2340
   End
   Begin VB.TextBox TMail 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4950
      TabIndex        =   0
      Top             =   2910
      Width           =   2340
   End
   Begin VB.Image Command1 
      Height          =   285
      Left            =   2700
      Top             =   4175
      Width           =   2115
   End
   Begin VB.Image Command2 
      Height          =   285
      Left            =   460
      Top             =   4175
      Width           =   1365
   End
   Begin VB.Label Estado 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Esperando..."
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4950
      TabIndex        =   11
      Top             =   2360
      Width           =   2340
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
      TabIndex        =   9
      Top             =   3540
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Acepto los Términos y el reglamento del Juego."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   2640
      Width           =   3495
   End
End
Attribute VB_Name = "frmCrearAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call General_Set_Wav(SND_CLICK)

'Lorwik> Primero tiene que pasar la revision
If Len(TName.Text) < 5 Then
    MsgBox "El nombre de la cuenta debe de tener mas de 4 caracteres.", vbCritical
    Exit Sub
End If

If Len(TPass.Text) < 6 Then
    MsgBox "El password de la cuenta debe de tener mas de 6 caracteres.", vbCritical
    Exit Sub
End If

If TPass <> RTPass Then
    MsgBox "Las passwords que tipeo no coinciden", , "Winter AO"
    Exit Sub
End If

If Not CheckMailString(TMail) Then
    MsgBox "Direccion de mail invalida."
    Exit Sub
End If

If TName = "" Or TPass = "" Or RTPass = "" Or TMail = "" Or pregunta = "" Or respuesta = "" Then
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

'Lorwik> Ahora si le dejamos crear
        Cuenta.name = UCase(LTrim(TName.Text))
        Cuenta.Pass = TPass.Text
        Cuenta.Email = TMail.Text
        Cuenta.preg = pregunta
        Cuenta.resp = respuesta.Text
        
        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
            DoEvents
        End If
   
            EstadoLogin = CrearCuenta
            frmMain.Winsock1.Connect CurServerIp, CurServerPort
            
            Unload Me
        Exit Sub
End Sub
Private Sub command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command1.Picture = General_Load_Picture_From_Resource("20.gif")
End Sub
Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command2.Picture = General_Load_Picture_From_Resource("25.gif")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command1.Picture = LoadPicture("")
    Command2.Picture = LoadPicture("")
End Sub
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
lblCode.Caption = GenerateKey
Me.Picture = General_Load_Picture_From_Resource("41.gif")
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then Call Auto_Drag(Me.hWnd)
End Sub
Private Sub tMail_Click()
If Idioma = 1 Then
If Not RTPass.Text = "" Then
If TPass.Text = RTPass.Text Then
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
If Not RTPass.Text = "" Then
If TPass.Text = RTPass.Text Then
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

Private Sub TName_Click()
If Idioma = 1 Then
If TPass.Text = "" Then
    Estado.ForeColor = vbWhite
    Estado.Caption = "Esperando..."
Else
If TPass.Text = RTPass.Text Then
    Estado.ForeColor = vbGreen
    Estado.Caption = "Password Correcto."
        Else
    Estado.ForeColor = vbRed
    Estado.Caption = "Password Incorrecto."
    End If
End If
Else
If TPass.Text = "" Then
    Estado.ForeColor = vbWhite
    Estado.Caption = "Comprobando..."
Else
If TPass.Text = RTPass.Text Then
    Estado.ForeColor = vbGreen
    Estado.Caption = "Password Correcto."
        Else
    Estado.ForeColor = vbRed
    Estado.Caption = "Password Incorrecto."
    End If
End If
End If
End Sub

Private Sub RTPass_click()
If Idioma = 1 Then
    Estado.ForeColor = vbWhite
    Estado.Caption = "Esperando..."
Else
    Estado.ForeColor = vbWhite
    Estado.Caption = "Comprobando..."
End If
End Sub
