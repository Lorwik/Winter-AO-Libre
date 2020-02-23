VERSION 5.00
Begin VB.Form frmParticle 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agregar Particulas"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   1815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin WorldEditor.lvButtons_H cmdAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "&Agregar"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.TextBox Life 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Text            =   "-1"
      Top             =   2400
      Width           =   390
   End
   Begin VB.ListBox lstParticle 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin WorldEditor.lvButtons_H cmdDel 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "&Quitar"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H cmdView 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "&Mostrar"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Label Label2 
      Caption         =   "LiveCounter:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
End
Attribute VB_Name = "frmParticle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
If cmdAddvalue = True Then
    cmdDel.Enabled = False
    Call modPaneles.EstSelectPanel(7, True)
Else
    cmdDel.Enabled = True
    Call modPaneles.EstSelectPanel(7, False)
End If
End Sub

Private Sub cmdDel_Click()
If cmdDel.value = True Then
    lstParticle.Enabled = False
    cmdAdd.Enabled = False
    Call modPaneles.EstSelectPanel(7, True)
Else
    lstParticle.Enabled = True
    cmdAdd.Enabled = True
    Call modPaneles.EstSelectPanel(7, False)
End If
End Sub

