VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "Msinet.ocx"
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Winter AO Return"
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
   MouseIcon       =   "frmConnect.frx":000C
   Moveable        =   0   'False
   Picture         =   "frmConnect.frx":1CD6
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.ListBox LstServidores 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2175
      ItemData        =   "frmConnect.frx":45A0F
      Left            =   6300
      List            =   "frmConnect.frx":45A16
      TabIndex        =   6
      Top             =   5460
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   -4080
      TabIndex        =   5
      Top             =   8880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   -1440
      TabIndex        =   4
      Top             =   8880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox Recuerda 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1320
      Picture         =   "frmConnect.frx":45A29
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   4140
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.TextBox PasswordTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
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
      Left            =   3075
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3570
      Width           =   2340
   End
   Begin VB.TextBox NameTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2130
      TabIndex        =   1
      Top             =   3105
      Width           =   2340
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Image Command8 
      Height          =   375
      Left            =   2040
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   9240
      Top             =   5040
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   11760
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   1260
      Top             =   4125
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   11400
      Top             =   0
      Width           =   375
   End
   Begin VB.Label version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   75
      Left            =   -360
      TabIndex        =   0
      Top             =   9000
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   225
      Index           =   0
      Left            =   3240
      MouseIcon       =   "frmConnect.frx":45C9D
      Top             =   4680
      Width           =   2010
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   1
      Left            =   840
      MouseIcon       =   "frmConnect.frx":470E7
      Top             =   4680
      Width           =   2085
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command8_Click()
frmCreditos.Show
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 27 Then
        frmCargando.Show
        frmCargando.Refresh
        frmConnect.MousePointer = 1
        frmMain.MousePointer = 1
        prgRun = False
        frmCargando.Refresh
        LiberarObjetosDX
        frmCargando.Refresh
        Call UnloadAllForms
        MP3P.stopMP3
End If
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

Call ListarServidores

    EngineRun = False
    
 Dim j
 For Each j In Image1()
    j.Tag = "0"
 Next
 
Me.Picture = General_Load_Picture_From_Resource("Conectar.gif")

'RECUERDA PASS Y NOMBRE
If GetVar(App.Path & "\init\config.ini", "check", "A") = 1 Then
Recuerda.Visible = True
NameTxt.Text = GetVar(App.Path & "\init\config.ini", "Nick", "Name")
PasswordTxt.Text = GetVar(App.Path & "\init\config.ini", "Passwd", "Pass")
Else
Recuerda.Visible = False
End If
'RECUERDA PASS Y NOMBRE
End Sub
Private Sub Image1_Click(Index As Integer)
Call Audio.PlayWave(SND_CLICK)

Select Case Index
    Case 0
        
       EstadoLogin = CrearAccount
       If frmMain.Socket1.Connected Then
                frmMain.Socket1.Disconnect
                frmMain.Socket1.Cleanup
            End If
                frmMain.Socket1.HostName = CurServerIp
                frmMain.Socket1.RemotePort = CurServerPort
                frmMain.Socket1.Connect

        
Case 1
         If GetVar(App.Path & "\init\config.ini", "check", "A") = 1 Then
Call WriteVar(App.Path & "\init\config.ini", "Nick", "Name", NameTxt.Text)
Call WriteVar(App.Path & "\init\config.ini", "Passwd", "Pass", PasswordTxt.Text)
Call WriteVar(App.Path & "\init\config.ini", "Check", "A", "1")
End If
#If UsarWrench = 1 Then
       If frmMain.Socket1.Connected Then
                frmMain.Socket1.Disconnect
                frmMain.Socket1.Cleanup
        DoEvents
        End If
#Else
        If frmMain.Socket1.State <> sckClosed Then _
            frmMain.Socket1.Cleanup
#End If
        'update user info
        UserName = NameTxt.Text
        Dim aux As String
        aux = PasswordTxt.Text
        UserPassword = aux

        If CheckUserData(False) = True Then
            EstadoLogin = loginaccount
            Me.MousePointer = 99
#If UsarWrench = 1 Then
                frmMain.Socket1.HostName = CurServerIp
                frmMain.Socket1.RemotePort = CurServerPort
                frmMain.Socket1.Connect
#Else

            frmMain.Socket1.Connect CurServerIp, CurServerPort
#End If
        End If
        
End Select
Exit Sub
End Sub
Private Sub Image2_Click()
Recuerda.Visible = True
Call WriteVar(App.Path & "\init\config.ini", "check", "A", "1")
          If GetVar(App.Path & "\init\config.ini", "check", "A") = 1 Then
Call WriteVar(App.Path & "\init\config.ini", "Nick", "Name", NameTxt.Text)
Call WriteVar(App.Path & "\init\config.ini", "Passwd", "Pass", PasswordTxt.Text)
End If
End Sub

Private Sub Image3_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub Image4_Click()
On Error Resume Next
If MsgBox("¿Está seguro que desea cerrar Winter-AO Return?", vbYesNo + vbQuestion, "Winter-AO Return") = vbYes Then

        frmCargando.Show
        frmCargando.Refresh
        AddtoRichTextBox frmCargando.status, "Cerrando Winter-AO Return.", 0, 0, 0, 1, 0, 1
        frmConnect.MousePointer = 1
        frmMain.MousePointer = 1
        prgRun = False
        
        AddtoRichTextBox frmCargando.status, "Liberando recursos..."
        frmCargando.Refresh
        LiberarObjetosDX
        AddtoRichTextBox frmCargando.status, "Hecho", 0, 0, 0, 1, 0, 1
        AddtoRichTextBox frmCargando.status, "¡¡Gracias por jugar Winter-AO Return!!", 0, 0, 0, 1, 0, 1
        frmCargando.Refresh
        Call UnloadAllForms
        MP3P.stopMP3
End If
End Sub

Private Sub LstServidores_Click()
    Text1.Text = Servidor(LstServidores.listIndex + 1).IP
    Text2.Text = Servidor(LstServidores.listIndex + 1).Puerto
End Sub

Private Sub Recuerda_Click()
'Lorwik
If MsgBox("¿Está seguro que desea dejar de recordar su personaje?", vbYesNo + vbQuestion, "Winter-AO Return") = vbYes Then
Call WriteVar(App.Path & "\init\config.ini", "Nick", "Name", "")
Call WriteVar(App.Path & "\init\config.ini", "Passwd", "Pass", "")
Call WriteVar(App.Path & "\init\config.ini", "Check", "A", "0")
Recuerda.Visible = False
MsgBox "Su Personaje no esta guardado"
End If
'Lorwik
End Sub
Private Sub Image5_Click()
Call ListarServidores
End Sub
