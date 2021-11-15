VERSION 5.00
Begin VB.Form frmGuildLeader 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administración del Clan"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5880
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
   ScaleHeight     =   5700
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Caption         =   "Propuestas de alianzas"
      Height          =   255
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   4440
      Width           =   2775
   End
   Begin VB.CommandButton Command8 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   195
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   4920
      Width           =   2775
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Propuestas de paz"
      Height          =   255
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Editar URL de la web del clan"
      Height          =   255
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   3960
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Editar Codex o Descripcion"
      Height          =   255
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":0548
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   3690
      Width           =   2775
   End
   Begin VB.Frame Frame3 
      Caption         =   "Clanes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   2895
      Begin VB.ListBox guildslist 
         Height          =   1425
         ItemData        =   "frmGuildLeader.frx":069A
         Left            =   120
         List            =   "frmGuildLeader.frx":069C
         TabIndex        =   11
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Detalles"
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":069E
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   1800
         Width           =   2655
      End
   End
   Begin VB.Frame txtnews 
      Caption         =   "GuildNews"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   6
      Top             =   2040
      Width           =   5775
      Begin VB.CommandButton Command3 
         Caption         =   "Actualizar"
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":07F0
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   1080
         Width           =   5535
      End
      Begin VB.TextBox txtguildnews 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Miembros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2880
      TabIndex        =   3
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton Command2 
         Caption         =   "Detalles"
         Height          =   195
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":0942
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   1800
         Width           =   2655
      End
      Begin VB.ListBox members 
         Height          =   1425
         ItemData        =   "frmGuildLeader.frx":0A94
         Left            =   120
         List            =   "frmGuildLeader.frx":0A96
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Solicitudes de ingreso"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   3720
      Width           =   2895
      Begin VB.CommandButton cmdElecciones 
         Caption         =   "Abrir elecciones"
         Height          =   195
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":0A98
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   1680
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Detalles"
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":0BEA
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   1170
         Width           =   2655
      End
      Begin VB.ListBox solicitudes 
         Height          =   840
         ItemData        =   "frmGuildLeader.frx":0D3C
         Left            =   120
         List            =   "frmGuildLeader.frx":0D3E
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Miembros 
         Caption         =   "El clan cuenta con x miembros"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmGuildLeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdElecciones_Click()
    Call SendData("ABREELEC")
    Unload Me
End Sub
Private Sub Command1_Click()
frmCharInfo.frmsolicitudes = True
Call SendData("1HRINFO<" & solicitudes.List(solicitudes.listIndex))
End Sub
Private Sub Command2_Click()
frmCharInfo.frmmiembros = True
Call SendData("1HRINFO<" & members.List(members.listIndex))
End Sub

Private Sub Command3_Click()
Dim k$
k$ = Replace(txtguildnews, vbCrLf, "º")
Call SendData("ACTGNEWS" & k$)
End Sub

Private Sub Command4_Click()
frmGuildBrief.EsLeader = True
Call SendData("CLANDETAILS" & guildslist.List(guildslist.listIndex))
End Sub
Private Sub Command5_Click()
Call frmGuildDetails.Show(vbModal, frmGuildLeader)
End Sub

Private Sub Command6_Click()
Call frmGuildURL.Show(vbModeless, frmGuildLeader)
End Sub
Private Sub Command7_Click()
Call SendData("ENVPROPP")
End Sub
Private Sub Command9_Click()
Call SendData("ENVALPRO")
End Sub
Private Sub Command8_Click()
Unload Me
frmMain.SetFocus
End Sub
Public Sub ParseLeaderInfo(ByVal data As String)
If Me.Visible Then Exit Sub

Dim r%, T%

r% = Val(ReadField(1, data, Asc("¬")))

For T% = 1 To r%
    guildslist.AddItem ReadField(1 + T%, data, Asc("¬"))
Next T%

r% = Val(ReadField(T% + 1, data, Asc("¬")))
Miembros.Caption = "El clan cuenta con " & r% & " miembros."

Dim k%

For k% = 1 To r%
    members.AddItem ReadField(T% + 1 + k%, data, Asc("¬"))
Next k%

txtguildnews = Replace(ReadField(T% + k% + 1, data, Asc("¬")), "º", vbCrLf)

T% = T% + k% + 2

r% = Val(ReadField(T%, data, Asc("¬")))

For k% = 1 To r%
    solicitudes.AddItem ReadField(T% + k%, data, Asc("¬"))
Next k%

Me.Show , frmMain

End Sub
