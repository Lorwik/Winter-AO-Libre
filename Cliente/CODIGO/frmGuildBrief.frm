VERSION 5.00
Begin VB.Form frmGuildBrief 
   BorderStyle     =   0  'None
   Caption         =   "Detalles del Clan"
   ClientHeight    =   8130
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   7590
   ClipControls    =   0   'False
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
   ScaleHeight     =   8130
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Command2 
      Height          =   375
      Left            =   4560
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Image Guerra 
      Height          =   300
      Left            =   3120
      Top             =   7690
      Width           =   1425
   End
   Begin VB.Image aliado 
      Height          =   300
      Left            =   1680
      Top             =   7690
      Width           =   1425
   End
   Begin VB.Image Command3 
      Height          =   300
      Left            =   240
      Top             =   7695
      Width           =   1425
   End
   Begin VB.Image Command1 
      Height          =   375
      Left            =   6000
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label Desc 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   480
      TabIndex        =   19
      Top             =   6240
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   18
      Top             =   3840
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   17
      Top             =   4080
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   16
      Top             =   4320
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   15
      Top             =   4560
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   14
      Top             =   4800
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   13
      Top             =   5040
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   12
      Top             =   5280
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   480
      TabIndex        =   11
      Top             =   5520
      Width           =   6735
   End
   Begin VB.Label antifaccion 
      BackStyle       =   0  'Transparent
      Caption         =   "Puntos Antifaccion:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2595
      TabIndex        =   10
      Top             =   3120
      Width           =   4650
   End
   Begin VB.Label Aliados 
      BackStyle       =   0  'Transparent
      Caption         =   "Clanes Aliados:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   2880
      Width           =   5295
   End
   Begin VB.Label Enemigos 
      BackStyle       =   0  'Transparent
      Caption         =   "Clanes Enemigos:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   2640
      Width           =   4815
   End
   Begin VB.Label lblAlineacion 
      BackStyle       =   0  'Transparent
      Caption         =   "Alineacion:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   2400
      Width           =   5535
   End
   Begin VB.Label eleccion 
      BackStyle       =   0  'Transparent
      Caption         =   "Elecciones:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   2160
      Width           =   5655
   End
   Begin VB.Label Miembros 
      BackStyle       =   0  'Transparent
      Caption         =   "Miembros:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   1920
      Width           =   5535
   End
   Begin VB.Label web 
      BackStyle       =   0  'Transparent
      Caption         =   "Web site:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   1680
      Width           =   5775
   End
   Begin VB.Label lider 
      BackStyle       =   0  'Transparent
      Caption         =   "Lider:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   1440
      Width           =   6135
   End
   Begin VB.Label creacion 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de creacion:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Label fundador 
      BackStyle       =   0  'Transparent
      Caption         =   "Fundador:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   5535
   End
   Begin VB.Label nombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   720
      Width           =   5775
   End
End
Attribute VB_Name = "frmGuildBrief"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public EsLeader As Boolean

Private Sub aliado_Click()
Call General_Set_Wav(SND_CLICK)
    frmCommet.Nombre = Right(Nombre.Caption, Len(Nombre.Caption) - 7)
    frmCommet.t = TIPO.ALIANZA
    frmCommet.Caption = "Ingrese propuesta de alianza"
    Call frmCommet.Show(vbModal, frmGuildBrief)
End Sub

Private Sub Command1_Click()
Call General_Set_Wav(SND_CLICK)
    Unload Me
End Sub

Private Sub Command2_Click()
Call General_Set_Wav(SND_CLICK)
    Call frmGuildSol.RecieveSolicitud(Nombre)
    Call frmGuildSol.Show(vbModal, frmGuildBrief)
End Sub

Private Sub Command3_Click()
Call General_Set_Wav(SND_CLICK)
    frmCommet.Nombre = Right(Nombre.Caption, Len(Nombre.Caption) - 7)
    frmCommet.t = TIPO.PAZ
    frmCommet.Caption = "Ingrese propuesta de paz"
    Call frmCommet.Show(vbModal, frmGuildBrief)
End Sub

Private Sub Guerra_Click()
Call General_Set_Wav(SND_CLICK)
    Call WriteGuildDeclareWar(Right(Nombre.Caption, Len(Nombre.Caption) - 7))
    Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("49.gif")
guerra.Picture = General_Load_Picture_From_Resource("90.gif")
aliado.Picture = General_Load_Picture_From_Resource("91.gif")
Command3.Picture = General_Load_Picture_From_Resource("92.gif")
End Sub
