VERSION 5.00
Begin VB.Form frmPasswdSinPadrinos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5160
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
   Moveable        =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   300
      Left            =   3120
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   960
      Width           =   1800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Volver"
      Height          =   300
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   960
      Width           =   1680
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Crear el PJ:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmPasswdSinPadrinos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 On Error Resume Next
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    
    Me.MousePointer = 11
    EstadoLogin = CrearNuevoPj

    If Not frmMain.Socket1.Connected Then
        MsgBox "Error: Se ha perdido la conexion con el server."
        Unload Me
    Else
        Call login(ValidarLoginMSG(CInt(bRK)))
    End If
MP3P.stopMP3
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Label1 = "Crear el PJ: " & vbNewLine & UserName & " ?"
End Sub

