VERSION 5.00
Begin VB.Form frmPeaceProp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ofertas de paz"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
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
   ScaleHeight     =   2895
   ScaleWidth      =   4980
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Rechazar"
      Height          =   495
      Left            =   3720
      MouseIcon       =   "frmPeaceProp.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   2520
      MouseIcon       =   "frmPeaceProp.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Detalles"
      Height          =   495
      Left            =   1320
      MouseIcon       =   "frmPeaceProp.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MouseIcon       =   "frmPeaceProp.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
   Begin VB.ListBox lista 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      ItemData        =   "frmPeaceProp.frx":0548
      Left            =   120
      List            =   "frmPeaceProp.frx":054A
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmPeaceProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private tipoprop As TIPO_PROPUESTA

Public Enum TIPO_PROPUESTA
    ALIANZA = 1
    PAZ = 2
End Enum

Public Property Let ProposalType(ByVal nValue As TIPO_PROPUESTA)
    tipoprop = nValue
End Property

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
'Me.Visible = False
If tipoprop = PAZ Then
    Call WriteGuildPeaceDetails(lista.List(lista.ListIndex))
Else
    Call WriteGuildAllianceDetails(lista.List(lista.ListIndex))
End If
End Sub

Private Sub Command3_Click()
    'Me.Visible = False
    If tipoprop = PAZ Then
        Call WriteGuildAcceptPeace(lista.List(lista.ListIndex))
    Else
        Call WriteGuildAcceptAlliance(lista.List(lista.ListIndex))
    End If
    Me.Hide
    Unload Me
End Sub

Private Sub Command4_Click()
    If tipoprop = PAZ Then
        Call WriteGuildRejectPeace(lista.List(lista.ListIndex))
    Else
        Call WriteGuildRejectAlliance(lista.List(lista.ListIndex))
    End If
    Me.Hide
    Unload Me
End Sub
