VERSION 5.00
Begin VB.Form frmNewPassword 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   ScaleHeight     =   175
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   302
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   735
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1140
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1575
      Width           =   2295
   End
   Begin VB.Image Command2 
      Height          =   255
      Left            =   840
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Image Command1 
      Height          =   255
      Left            =   2280
      Top             =   2160
      Width           =   1455
   End
End
Attribute VB_Name = "frmNewPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
        If Text2.Text <> Text3.Text Then
            MsgBox "Las contraseñas no coinciden"
            Unload Me
            Exit Sub
        End If

       PsswdAnte = Text1.Text
       PasswdNew = Text2.Text

        EstadoLogin = CambiarPass
        
        frmMain.Winsock1.Connect CurServerIp, CurServerPort
        
            Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("89.gif")
End Sub
