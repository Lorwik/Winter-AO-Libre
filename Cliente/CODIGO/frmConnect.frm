VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtPasswd 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   4830
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4800
      Width           =   2340
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   4830
      TabIndex        =   0
      Top             =   3870
      Width           =   2340
   End
   Begin VB.Image Command1 
      Height          =   285
      Left            =   10395
      Top             =   7380
      Width           =   1365
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   4800
      Top             =   5280
      Width           =   225
   End
   Begin VB.Image imgrecuperar 
      Height          =   285
      Left            =   10395
      Top             =   7035
      Width           =   1365
   End
   Begin VB.Image imgConectarse 
      Height          =   285
      Left            =   4920
      Top             =   5700
      Width           =   2115
   End
   Begin VB.Image imgSalir 
      Height          =   285
      Left            =   10410
      Top             =   7725
      Width           =   1365
   End
   Begin VB.Image imgCrearPj 
      Height          =   285
      Left            =   4935
      Top             =   6180
      Width           =   2115
   End
   Begin VB.Image imgServArgentina 
      Height          =   795
      Left            =   360
      MousePointer    =   99  'Custom
      Top             =   9240
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Label version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   45
      TabIndex        =   2
      Top             =   7995
      Width           =   510
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
frmCreditos.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        prgRun = False
    End If
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then Call Auto_Drag(Me.hWnd)
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'Make Server IP and Port box visible
If KeyCode = vbKeyI And Shift = vbCtrlMask Then
    KeyCode = 0
    Exit Sub
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    Me.Caption = Form_Caption
    EngineRun = False
    version.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision & " GNU/GPL"
    Me.Picture = General_Load_Picture_From_Resource("39.gif")
    
    If GetVar(App.Path & "\init\config.cfg", "Cuenta", "Recordar") = 1 Then
        txtNombre.Text = GetVar(App.Path & "\init\config.cfg", "Cuenta", "Name")
        Image1.Picture = General_Load_Picture_From_Resource("80.gif")
    Else
        txtNombre.Text = ""
        Image1.Picture = LoadPicture("")
    End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgCrearPj.Picture = LoadPicture("")
    imgConectarse.Picture = LoadPicture("")
    imgSalir.Picture = LoadPicture("")
    imgrecuperar.Picture = LoadPicture("")
    Command1.Picture = LoadPicture("")
End Sub

Private Sub CheckServers()
    CurServer = 0
    IPdelServidor = CurServerIp
    PuertoDelServidor = CurServerPort
End Sub

Private Sub recuperar_Click()
Call General_Set_Wav(SND_CLICK)
frmRecuperarAccount.Show
End Sub

Private Sub Image1_Click()
    If GetVar(App.Path & "\init\config.cfg", "Cuenta", "Recordar") = 0 Then
        Call WriteVar(App.Path & "\init\config.cfg", "Cuenta", "Name", txtNombre.Text)
        Call WriteVar(App.Path & "\init\config.cfg", "Cuenta", "Recordar", "1")
        Image1.Picture = General_Load_Picture_From_Resource("80.gif")
    Else
        Call WriteVar(App.Path & "\init\config.cfg", "Cuenta", "Name", "")
        Call WriteVar(App.Path & "\init\config.cfg", "Cuenta", "Recordar", "0")
        Image1.Picture = LoadPicture("")
    End If
End Sub

Private Sub imgConectarse_Click()
Call General_Set_Wav(SND_CLICK)

    If txtNombre.Text = "" Or txtPasswd.Text = "" Then
        frmMensaje.msg.Caption = "Escriba su nombre de cuenta y contraseña para logear."
        frmMensaje.Show
        Exit Sub
    End If

    If GetVar(App.Path & "\init\config.cfg", "Cuenta", "Recordar") = 1 Then
        Call WriteVar(App.Path & "\init\config.cfg", "Cuenta", "Name", txtNombre.Text)
        Call WriteVar(App.Path & "\init\config.cfg", "Cuenta", "Recordar", "1")
    End If

    If frmMain.Winsock1.State <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
   
    Cuenta.name = txtNombre.Text
    Cuenta.Pass = txtPasswd.Text
 
    'If CheckAccData(False, False) = True Then
        EstadoLogin = LoginCuenta
 
        frmMain.Winsock1.Connect CurServerIp, CurServerPort
    'End If
End Sub
Private Sub imgConectarse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgConectarse.Picture = General_Load_Picture_From_Resource("18.gif")
End Sub
Private Sub imgCrearPj_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgCrearPj.Picture = General_Load_Picture_From_Resource("19.gif")
End Sub
Private Sub imgSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgSalir.Picture = General_Load_Picture_From_Resource("22.gif")
End Sub
Private Sub command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command1.Picture = General_Load_Picture_From_Resource("86.gif")
End Sub
Private Sub imgcrearpj_click()
Call General_Set_Wav(SND_CLICK)
frmCrearAccount.Show
End Sub
Private Sub imgrecuperar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgrecuperar.Picture = General_Load_Picture_From_Resource("21.gif")
End Sub
Private Sub imgSalir_Click()
    Call General_Set_Wav(SND_CLICK)
    prgRun = False
End Sub

Private Sub imgrecuperar_click()
frmRecuperarAccount.Show
End Sub
Private Sub txtPasswd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then imgConectarse_Click
End Sub
