VERSION 5.00
Begin VB.Form frmCuent 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11970
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmCuent.frx":0000
   Picture         =   "frmCuent.frx":1CCA
   ScaleHeight     =   9000
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1150
      Index           =   5
      Left            =   2490
      MouseIcon       =   "frmCuent.frx":4533E
      Picture         =   "frmCuent.frx":46788
      ScaleHeight     =   1155
      ScaleWidth      =   735
      TabIndex        =   32
      Top             =   5520
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1150
      Index           =   6
      Left            =   4090
      MouseIcon       =   "frmCuent.frx":46A23
      Picture         =   "frmCuent.frx":47E6D
      ScaleHeight     =   1155
      ScaleWidth      =   735
      TabIndex        =   31
      Top             =   5500
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1150
      Index           =   7
      Left            =   5690
      MouseIcon       =   "frmCuent.frx":48108
      Picture         =   "frmCuent.frx":49552
      ScaleHeight     =   1155
      ScaleWidth      =   735
      TabIndex        =   30
      Top             =   5500
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1150
      Index           =   4
      Left            =   890
      MouseIcon       =   "frmCuent.frx":497ED
      Picture         =   "frmCuent.frx":4AC37
      ScaleHeight     =   1155
      ScaleWidth      =   735
      TabIndex        =   4
      Top             =   5500
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1150
      Index           =   3
      Left            =   5690
      MouseIcon       =   "frmCuent.frx":4AED2
      Picture         =   "frmCuent.frx":4C31C
      ScaleHeight     =   1155
      ScaleWidth      =   735
      TabIndex        =   3
      Top             =   3120
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1150
      Index           =   2
      Left            =   4110
      MouseIcon       =   "frmCuent.frx":4C5B7
      Picture         =   "frmCuent.frx":4DA01
      ScaleHeight     =   1093.21
      ScaleMode       =   0  'User
      ScaleWidth      =   735
      TabIndex        =   2
      Top             =   3120
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1150
      Index           =   1
      Left            =   2480
      MouseIcon       =   "frmCuent.frx":4DC9C
      Picture         =   "frmCuent.frx":4F0E6
      ScaleHeight     =   1093.21
      ScaleMode       =   0  'User
      ScaleWidth      =   735
      TabIndex        =   1
      Top             =   3120
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1150
      Index           =   0
      Left            =   850
      MouseIcon       =   "frmCuent.frx":4F381
      Picture         =   "frmCuent.frx":507CB
      ScaleHeight     =   1155
      ScaleWidth      =   735
      TabIndex        =   0
      Top             =   3120
      Width           =   735
   End
   Begin VB.Image Command5 
      Height          =   390
      Left            =   5520
      Picture         =   "frmCuent.frx":50A66
      Top             =   7570
      Width           =   1605
   End
   Begin VB.Image Pj8 
      Height          =   1650
      Left            =   5560
      Picture         =   "frmCuent.frx":52B90
      Top             =   5260
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image Pj7 
      Height          =   1650
      Left            =   3960
      Picture         =   "frmCuent.frx":53A44
      Top             =   5260
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image Pj6 
      Height          =   1650
      Left            =   2340
      Picture         =   "frmCuent.frx":548F8
      Top             =   5260
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image Pj5 
      Height          =   1650
      Left            =   720
      Picture         =   "frmCuent.frx":557AC
      Top             =   5260
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image Pj4 
      Height          =   1650
      Left            =   5560
      Picture         =   "frmCuent.frx":56660
      Top             =   2880
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image Pj3 
      Height          =   1650
      Left            =   3960
      Picture         =   "frmCuent.frx":57514
      Top             =   2880
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image Pj2 
      Height          =   1650
      Left            =   2340
      Picture         =   "frmCuent.frx":583C8
      Top             =   2880
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image Pj1 
      Height          =   1650
      Left            =   720
      Picture         =   "frmCuent.frx":5927C
      Top             =   2880
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image Command1 
      Height          =   255
      Left            =   9720
      MouseIcon       =   "frmCuent.frx":5A130
      Top             =   7680
      Width           =   2055
   End
   Begin VB.Image Command3 
      Height          =   255
      Left            =   7200
      MouseIcon       =   "frmCuent.frx":5B57A
      Top             =   7680
      Width           =   2055
   End
   Begin VB.Image Command4 
      Height          =   255
      Left            =   3360
      MouseIcon       =   "frmCuent.frx":5C9C4
      Top             =   7680
      Width           =   2055
   End
   Begin VB.Image Command2 
      Height          =   255
      Left            =   240
      MouseIcon       =   "frmCuent.frx":5DE0E
      Top             =   7680
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de la Cuenta:"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   720
      TabIndex        =   33
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   7
      Left            =   6090
      TabIndex        =   29
      Top             =   7440
      Width           =   75
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   6
      Left            =   4470
      TabIndex        =   28
      Top             =   7440
      Width           =   75
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   5
      Left            =   2835
      TabIndex        =   27
      Top             =   7440
      Width           =   75
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Height          =   195
      Index           =   7
      Left            =   6090
      TabIndex        =   26
      Top             =   7200
      Width           =   75
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Height          =   195
      Index           =   6
      Left            =   4470
      TabIndex        =   25
      Top             =   7200
      Width           =   75
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Height          =   195
      Index           =   5
      Left            =   2835
      TabIndex        =   24
      Top             =   7200
      Width           =   75
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   7
      Left            =   5865
      TabIndex        =   23
      Top             =   6960
      Width           =   525
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   6
      Left            =   4245
      TabIndex        =   22
      Top             =   6960
      Width           =   525
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   5
      Left            =   2610
      TabIndex        =   21
      Top             =   6960
      Width           =   525
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PJClick"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   2880
      TabIndex        =   20
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Height          =   195
      Index           =   4
      Left            =   1230
      TabIndex        =   19
      Top             =   7200
      Width           =   75
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Height          =   195
      Index           =   3
      Left            =   6060
      TabIndex        =   18
      Top             =   4800
      Width           =   75
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Height          =   195
      Index           =   2
      Left            =   4470
      TabIndex        =   17
      Top             =   4800
      Width           =   75
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Height          =   195
      Index           =   1
      Left            =   2835
      TabIndex        =   16
      Top             =   4800
      Width           =   75
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   4
      Left            =   1245
      TabIndex        =   15
      Top             =   7440
      Width           =   75
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   3
      Left            =   6045
      TabIndex        =   14
      Top             =   5040
      Width           =   105
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   2
      Left            =   4470
      TabIndex        =   13
      Top             =   5040
      Width           =   75
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   1
      Left            =   2835
      TabIndex        =   12
      Top             =   5040
      Width           =   75
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   0
      Left            =   1230
      TabIndex        =   11
      Top             =   5040
      Width           =   75
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Height          =   195
      Index           =   0
      Left            =   1230
      TabIndex        =   10
      Top             =   4800
      Width           =   75
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   4
      Left            =   1005
      TabIndex        =   9
      Top             =   6960
      Width           =   525
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   3
      Left            =   5840
      TabIndex        =   8
      Top             =   4560
      Width           =   525
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   2
      Left            =   4245
      TabIndex        =   7
      Top             =   4560
      Width           =   525
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   1
      Left            =   2605
      TabIndex        =   6
      Top             =   4560
      Width           =   525
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   0
      Left            =   1005
      TabIndex        =   5
      Top             =   4560
      Width           =   525
   End
End
Attribute VB_Name = "frmCuent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub Command1_Click()
On Error Resume Next
If PJClickeado = "Nada" Then
MsgBox "Seleccione un pj"
End If
Call Audio.PlayWave(SND_CLICK)
UserName = PJClickeado
SendData ("PUNMAK" & UserName)
MP3P.stopMP3
Unload Me
End Sub

Private Sub Command4_Click()
Dim PjABorrar As String

PjABorrar = PJClickeado

Debug.Print PjABorrar
Debug.Print UserName

If MsgBox("Seguro que desea borrar el pj " & Chr(34) & PjABorrar & Chr(34) & " ?", vbYesNo, "Confirmar borrado de pj.") = vbNo Then Exit Sub
Call SendData("DELETE" & PjABorrar & "," & UserName)
End Sub

Private Sub Command5_Click()
frmcambiarpass.Show
End Sub

Private Sub Form_Load()
Dim i As Integer
Label3.Caption = UserName
Me.Picture = General_Load_Picture_From_Resource("cuenta.gif")
End Sub
Private Sub Command2_Click()
frmMain.Socket1.Disconnect
Unload Me
frmConnect.Show
End Sub

Private Sub Command3_Click()
Call Audio.PlayWave(SND_CLICK)

If Nombre(7).Caption <> "Nada" Then
    MsgBox "Tu cuenta ha llegado al máximo de personajes."
    Exit Sub
End If

    EstadoLogin = Dados
    frmCrearPersonaje.Show vbModal
    Me.MousePointer = 11
    
End Sub


Private Sub nombre_dblClick(Index As Integer)
If PJClickeado = "Nada" Then Exit Sub
Call Audio.PlayWave(SND_CLICK)
UserName = PJClickeado
SendData ("PUNMAK" & UserName)
Unload Me
End Sub
Private Sub nombre_Click(Index As Integer)
PJClickeado = frmCuent.Nombre(Index).Caption
End Sub
Private Sub PJ_Click(Index As Integer)
PJClickeado = frmCuent.Nombre(Index).Caption
Select Case Index
    Case 0
    Pj1.Visible = True
Pj2.Visible = False
Pj3.Visible = False
Pj4.Visible = False
Pj5.Visible = False
Pj6.Visible = False
Pj7.Visible = False
Pj8.Visible = False
Case 1
Pj1.Visible = False
Pj2.Visible = True
Pj3.Visible = False
Pj4.Visible = False
Pj5.Visible = False
Pj6.Visible = False
Pj7.Visible = False
Pj8.Visible = False
Case 2
Pj1.Visible = False
Pj2.Visible = False
Pj3.Visible = True
Pj4.Visible = False
Pj5.Visible = False
Pj6.Visible = False
Pj7.Visible = False
Pj8.Visible = False

Case 3
Pj1.Visible = False
Pj2.Visible = False
Pj3.Visible = False
Pj4.Visible = True
Pj5.Visible = False
Pj6.Visible = False
Pj7.Visible = False
Pj8.Visible = False

Case 4
Pj1.Visible = False
Pj2.Visible = False
Pj3.Visible = False
Pj4.Visible = False
Pj5.Visible = True
Pj6.Visible = False
Pj7.Visible = False
Pj8.Visible = False

Case 5
Pj1.Visible = False
Pj2.Visible = False
Pj3.Visible = False
Pj4.Visible = False
Pj5.Visible = False
Pj6.Visible = True
Pj7.Visible = False
Pj8.Visible = False

Case 6
Pj1.Visible = False
Pj2.Visible = False
Pj3.Visible = False
Pj4.Visible = False
Pj5.Visible = False
Pj6.Visible = False
Pj7.Visible = True
Pj8.Visible = False

Case 7
Pj1.Visible = False
Pj2.Visible = False
Pj3.Visible = False
Pj4.Visible = False
Pj5.Visible = False
Pj6.Visible = False
Pj7.Visible = False
Pj8.Visible = True

End Select
End Sub

Private Sub PJ_dblClick(Index As Integer)
On Error Resume Next
If PJClickeado = "Nada" Then Exit Sub
Call Audio.PlayWave(SND_CLICK)
UserName = PJClickeado
SendData ("PUNMAK" & UserName)
Unload Me
MP3P.stopMP3
End Sub
