VERSION 5.00
Begin VB.Form frmCharInfo 
   BorderStyle     =   0  'None
   Caption         =   "Información del personaje"
   ClientHeight    =   6705
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   6795
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
   ScaleHeight     =   6705
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPeticiones 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1110
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3128
      Width           =   5790
   End
   Begin VB.TextBox txtMiembro 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1110
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4630
      Width           =   5790
   End
   Begin VB.Label status 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   14
      Top             =   2230
      Width           =   1560
   End
   Begin VB.Label reputacion 
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   13
      Top             =   2000
      Width           =   1605
   End
   Begin VB.Label criminales 
      BackStyle       =   0  'Transparent
      Caption         =   "Criminales asesinados:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   12
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Ciudadanos 
      BackStyle       =   0  'Transparent
      Caption         =   "Ciudadanos asesinados:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   11
      Top             =   1560
      Width           =   450
   End
   Begin VB.Label ejercito 
      BackStyle       =   0  'Transparent
      Caption         =   "Faccion:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   1305
      Width           =   2160
   End
   Begin VB.Label guildactual 
      BackStyle       =   0  'Transparent
      Caption         =   "Clan Actual:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   1080
      Width           =   1560
   End
   Begin VB.Label Banco 
      BackStyle       =   0  'Transparent
      Caption         =   "Banco:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   2250
      Width           =   2145
   End
   Begin VB.Label Oro 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1000
      TabIndex        =   7
      Top             =   1995
      Width           =   2325
   End
   Begin VB.Label Nivel 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   1785
      Width           =   2265
   End
   Begin VB.Label Genero 
      BackStyle       =   0  'Transparent
      Caption         =   "Genero:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1290
      TabIndex        =   5
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Clase 
      BackStyle       =   0  'Transparent
      Caption         =   "Clase:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1180
      TabIndex        =   4
      Top             =   1305
      Width           =   2175
   End
   Begin VB.Label Raza 
      BackStyle       =   0  'Transparent
      Caption         =   "Raza:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1120
      TabIndex        =   3
      Top             =   1050
      Width           =   2160
   End
   Begin VB.Label Nombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   820
      Width           =   5040
   End
   Begin VB.Image Echar 
      Height          =   255
      Left            =   360
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Image desc 
      Height          =   255
      Left            =   1800
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Image Rechazar 
      Height          =   255
      Left            =   360
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Image Command1 
      Height          =   255
      Left            =   4920
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Image Aceptar 
      Height          =   255
      Left            =   1800
      Top             =   6000
      Width           =   1455
   End
End
Attribute VB_Name = "frmCharInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum CharInfoFrmType
    frmMembers
    frmMembershipRequests
End Enum

Public frmType As CharInfoFrmType

Private Sub Aceptar_Click()
    Call WriteGuildAcceptNewMember(Trim$(Right$(Nombre, Len(Nombre) - 8)))
    Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
    Unload Me
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub desc_Click()
    Call WriteGuildRequestJoinerInfo(Right$(Nombre, Len(Nombre) - 8))
End Sub

Private Sub Echar_Click()
    Call WriteGuildKickMember(Right$(Nombre, Len(Nombre) - 8))
    Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
    Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("43.gif")
End Sub

Private Sub Rechazar_Click()
    Load frmCommet
    frmCommet.t = RECHAZOPJ
    frmCommet.Nombre = Right$(Nombre, Len(Nombre) - 8)
    frmCommet.Caption = "Ingrese motivo para rechazo"
    frmCommet.Show vbModeless, frmCharInfo
End Sub
