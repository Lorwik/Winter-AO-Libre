VERSION 5.00
Begin VB.Form frmRecuperarAccount 
   BorderStyle     =   0  'None
   Caption         =   "frmrecuperarcuenta"
   ClientHeight    =   4515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3570
   LinkTopic       =   "Form1"
   ScaleHeight     =   301
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   238
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox pregunta 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      ItemData        =   "frmRecuperarAccount.frx":0000
      Left            =   1905
      List            =   "frmRecuperarAccount.frx":0013
      TabIndex        =   6
      Text            =   "¿Lugar de Nacimiento?"
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox respuesta 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   2340
      Width           =   1455
   End
   Begin VB.TextBox CodeKey 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   645
      MaxLength       =   6
      TabIndex        =   3
      Top             =   3600
      Width           =   2250
   End
   Begin VB.TextBox ReNewPassword 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1530
      Width           =   1455
   End
   Begin VB.TextBox NewPassword 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1140
      Width           =   1455
   End
   Begin VB.TextBox CuentaName 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin VB.Image Command2 
      Height          =   255
      Left            =   360
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Image Command1 
      Height          =   255
      Left            =   1920
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label lblCode 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   3150
      Width           =   1815
   End
End
Attribute VB_Name = "frmRecuperarAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
        If Len(NewPassword.Text) < 6 Then
            MsgBox "El password de la cuenta debe de tener mas de 6 caracteres.", vbCritical
            Exit Sub
        End If
        
        If NewPassword <> ReNewPassword Then
            MsgBox "Las passwords que tipeo no coinciden", , "Winter AO"
            Exit Sub
        End If
        
        If lblCode.Caption <> CodeKey.Text Then
            MsgBox "El Codigo ingresado es Invalido.", vbCritical
            lblCode.Caption = GenerateKey
            Exit Sub
        End If
        
        NameAccount = CuentaName.Text
        NWPasswd = NewPassword.Text
        PRGScrta = pregunta
        Repstscrta = respuesta.Text
        
        EstadoLogin = RecuperarAccount
        
        frmMain.Winsock1.Connect CurServerIp, CurServerPort
        
        Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("72.gif")
lblCode.Caption = GenerateKey
End Sub
