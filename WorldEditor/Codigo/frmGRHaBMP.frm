VERSION 5.00
Begin VB.Form frmGRHaBMP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GRH => BMP"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3765
   Icon            =   "frmGRHaBMP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   3765
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtGRH 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin WorldEditor.lvButtons_H cmdCerrar 
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      Caption         =   "&Cerrar"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Label lblBMP 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Número de BMP:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Número de GRH:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   3600
      Y1              =   1440
      Y2              =   1440
   End
End
Attribute VB_Name = "frmGRHaBMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCerrar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
Unload Me
End Sub

Private Sub Form_Load()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
Me.Icon = frmMain.Icon
End Sub

Private Sub txtGRH_Change()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
If txtGRH.Text <> "" And IsNumeric(txtGRH.Text) = True Then
    If txtGRH.Text > MaxGrhs Then Exit Sub
    If txtGRH.Text < 1 Then Exit Sub
    lblBMP.Caption = GrhData(txtGRH.Text).FileNum
End If
End Sub

Private Sub txtGRH_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub
