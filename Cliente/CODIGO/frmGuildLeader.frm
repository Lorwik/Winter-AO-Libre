VERSION 5.00
Begin VB.Form frmGuildLeader 
   BorderStyle     =   0  'None
   Caption         =   "Administración del Clan"
   ClientHeight    =   6060
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   6000
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
   ScaleHeight     =   6060
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox solicitudes 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      ItemData        =   "frmGuildLeader.frx":0000
      Left            =   170
      List            =   "frmGuildLeader.frx":0002
      TabIndex        =   3
      Top             =   4265
      Width           =   2655
   End
   Begin VB.TextBox txtguildnews 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   170
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2570
      Width           =   5535
   End
   Begin VB.ListBox members 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1395
      ItemData        =   "frmGuildLeader.frx":0004
      Left            =   3050
      List            =   "frmGuildLeader.frx":0006
      TabIndex        =   1
      Top             =   540
      Width           =   2655
   End
   Begin VB.ListBox guildslist 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1395
      ItemData        =   "frmGuildLeader.frx":0008
      Left            =   170
      List            =   "frmGuildLeader.frx":000A
      TabIndex        =   0
      Top             =   540
      Width           =   2655
   End
   Begin VB.Image Command8 
      Height          =   255
      Left            =   4320
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Image Command9 
      Height          =   255
      Left            =   3120
      Top             =   4920
      Width           =   2655
   End
   Begin VB.Image Command7 
      Height          =   255
      Left            =   3120
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   3120
      Top             =   4365
      Width           =   2655
   End
   Begin VB.Image Command5 
      Height          =   255
      Left            =   3120
      Top             =   4080
      Width           =   2655
   End
   Begin VB.Image cmdElecciones 
      Height          =   255
      Left            =   3120
      Top             =   5280
      Width           =   2655
   End
   Begin VB.Label Miembros 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "El clan cuenta con x miembros"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   5520
      Width           =   2535
   End
   Begin VB.Image Command1 
      Height          =   255
      Left            =   720
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Image Command3 
      Height          =   255
      Left            =   2160
      Top             =   3430
      Width           =   1455
   End
   Begin VB.Image Command2 
      Height          =   255
      Left            =   3600
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Image Command4 
      Height          =   255
      Left            =   720
      Top             =   2040
      Width           =   1335
   End
End
Attribute VB_Name = "frmGuildLeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_NEWS_LENGTH As Integer = 512

Private Sub cmdElecciones_Click()
Call General_Set_Wav(SND_CLICK)
    Call WriteGuildOpenElections
    Unload Me
End Sub

Private Sub Command1_Click()
Call General_Set_Wav(SND_CLICK)
    If solicitudes.ListIndex = -1 Then Exit Sub
    
    frmCharInfo.frmType = CharInfoFrmType.frmMembershipRequests
    Call WriteGuildMemberInfo(solicitudes.List(solicitudes.ListIndex))

    'Unload Me
End Sub

Private Sub Command2_Click()
Call General_Set_Wav(SND_CLICK)
    If members.ListIndex = -1 Then Exit Sub
    
    frmCharInfo.frmType = CharInfoFrmType.frmMembers
    Call WriteGuildMemberInfo(members.List(members.ListIndex))

    'Unload Me
End Sub

Private Sub Command3_Click()
Call General_Set_Wav(SND_CLICK)
    Dim k As String

    k = Replace(txtguildnews, vbCrLf, "º")
    
    Call WriteGuildUpdateNews(k)
End Sub

Private Sub Command4_Click()
Call General_Set_Wav(SND_CLICK)
    frmGuildBrief.EsLeader = True
    Call WriteGuildRequestDetails(guildslist.List(guildslist.ListIndex))

    'Unload Me
End Sub

Private Sub Command5_Click()
Call General_Set_Wav(SND_CLICK)
    Call frmGuildDetails.Show(vbModal, frmGuildLeader)
    
    'Unload Me
End Sub

Private Sub Command6_Click()
Call General_Set_Wav(SND_CLICK)
Call frmGuildURL.Show(vbModeless, frmGuildLeader)
'Unload Me
End Sub

Private Sub Command7_Click()
Call General_Set_Wav(SND_CLICK)
    Call WriteGuildPeacePropList
End Sub
Private Sub Command9_Click()
Call General_Set_Wav(SND_CLICK)
    Call WriteGuildAlliancePropList
End Sub

Private Sub Command8_Click()
Call General_Set_Wav(SND_CLICK)
    Unload Me
    frmMain.SetFocus
End Sub

Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("52.gif")
End Sub

Private Sub txtguildnews_Change()
    If Len(txtguildnews.Text) > MAX_NEWS_LENGTH Then _
        txtguildnews.Text = Left$(txtguildnews.Text, MAX_NEWS_LENGTH)
End Sub
