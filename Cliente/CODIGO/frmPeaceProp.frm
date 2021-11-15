VERSION 5.00
Begin VB.Form frmPeaceProp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ofertas de paz"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4845
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
   ScaleHeight     =   2445
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Rechazar"
      Height          =   255
      Left            =   3720
      MouseIcon       =   "frmPeaceProp.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   2520
      MouseIcon       =   "frmPeaceProp.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Detalles"
      Height          =   255
      Left            =   1320
      MouseIcon       =   "frmPeaceProp.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2160
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
      Height          =   195
      Left            =   120
      MouseIcon       =   "frmPeaceProp.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2160
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
      Width           =   4575
   End
End
Attribute VB_Name = "frmPeaceProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private tipoprop As TIPO_PROPUESTA
Private Enum TIPO_PROPUESTA
    ALIANZA = 1
    PAZ = 2
End Enum



Private Sub Command1_Click()
Unload Me
End Sub

Public Sub ParsePeaceOffers(ByVal S As String)

Dim T%, r%

T% = Val(ReadField(1, S, 44))

For r% = 1 To T%
    Call lista.AddItem(ReadField(r% + 1, S, 44))
Next r%


tipoprop = PAZ

Me.Show vbModeless, frmMain

End Sub

Public Sub ParseAllieOffers(ByVal S As String)

Dim T%, r%

T% = Val(ReadField(1, S, 44))

For r% = 1 To T%
    Call lista.AddItem(ReadField(r% + 1, S, 44))
Next r%

tipoprop = ALIANZA
Me.Show vbModeless, frmMain

End Sub

Private Sub Command2_Click()
'Me.Visible = False
If tipoprop = PAZ Then
    Call SendData("PEACEDET" & lista.List(lista.listIndex))
Else
    Call SendData("ALLIEDET" & lista.List(lista.listIndex))
End If
End Sub

Private Sub Command3_Click()
'Me.Visible = False
If tipoprop = PAZ Then
    Call SendData("ACEPPEAT" & lista.List(lista.listIndex))
Else
    Call SendData("ACEPALIA" & lista.List(lista.listIndex))
End If
Me.Hide
Unload Me
End Sub

Private Sub Command4_Click()
If tipoprop = PAZ Then
    Call SendData("RECPPEAT" & lista.List(lista.listIndex))
Else
    Call SendData("RECPALIA" & lista.List(lista.listIndex))
End If
Me.Hide
Unload Me
End Sub
