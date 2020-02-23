VERSION 5.00
Begin VB.Form frmCreditos 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Abajo 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   12015
      TabIndex        =   16
      Top             =   8160
      Width           =   12015
   End
   Begin VB.PictureBox Arriba 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   12015
      TabIndex        =   15
      Top             =   0
      Width           =   12015
   End
   Begin VB.Timer Subepos 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.aowinter.com.ar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   4845
      TabIndex        =   14
      Top             =   17640
      Width           =   2505
   End
   Begin VB.Label mrt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Winter-AO Ultimate V4.0 GNU/GPL 2011 "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   3900
      TabIndex        =   13
      Top             =   17280
      Width           =   4665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A todos aquellos que nos ayudaron. A todos aquellos que nos apoyaron. Y a toda la comunidad de Argentum Online."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   960
      Left            =   4440
      TabIndex        =   12
      Top             =   16080
      Width           =   3165
   End
   Begin VB.Label gracias 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Agradecimientos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   4440
      TabIndex        =   11
      Top             =   15600
      Width           =   3015
   End
   Begin VB.Label cosas7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mapeo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5640
      TabIndex        =   10
      Top             =   14040
      Width           =   675
   End
   Begin VB.Label cosas6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sigfrido"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   5400
      TabIndex        =   9
      Top             =   13560
      Width           =   1155
   End
   Begin VB.Label cosas5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Director de Multimedia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   13080
      Width           =   2535
   End
   Begin VB.Label cosas4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Harry"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   12600
      Width           =   1215
   End
   Begin VB.Label staff 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Equipo de Winter-AO Ultimate"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   8160
      Width           =   7095
   End
   Begin VB.Label LwK 
      BackStyle       =   0  'Transparent
      Caption         =   "Lorwik"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   8880
      Width           =   1095
   End
   Begin VB.Label cosas1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Idea original Director Principal Programación Graficación"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   5160
      TabIndex        =   4
      Top             =   9360
      Width           =   1815
   End
   Begin VB.Label hnx 
      BackStyle       =   0  'Transparent
      Caption         =   "Hennox"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   10440
      Width           =   1215
   End
   Begin VB.Label cosas2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Director de Mapeo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   10920
      Width           =   1815
   End
   Begin VB.Label stk 
      BackStyle       =   0  'Transparent
      Caption         =   "Tefo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   11520
      Width           =   975
   End
   Begin VB.Label cosas3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Indexación"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   12000
      Width           =   1455
   End
End
Attribute VB_Name = "frmCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("83.gif")
Arriba.Picture = General_Load_Picture_From_Resource("84.gif")
Abajo.Picture = General_Load_Picture_From_Resource("85.gif")
End Sub
Private Sub form_click()
Unload Me
End Sub
Private Sub Subepos_Timer()
staff.Top = Val(staff.Top) - 1
LwK.Top = Val(LwK.Top) - 1
cosas1.Top = Val(cosas1.Top) - 1
hnx.Top = Val(hnx.Top) - 1
cosas2.Top = Val(cosas2.Top) - 1
stk.Top = Val(stk.Top) - 1
cosas3.Top = Val(cosas3.Top) - 1
mrt.Top = Val(mrt.Top) - 1
cosas4.Top = Val(cosas4.Top) - 1
cosas5.Top = Val(cosas5.Top) - 1
cosas6.Top = Val(cosas6.Top) - 1
cosas7.Top = Val(cosas7.Top) - 1
gracias.Top = Val(gracias.Top) - 1
Label2.Top = Val(Label2.Top) - 1
Label1.Top = Val(Label1.Top) - 1

If Label2.Top = 328 Then
    Subepos.Enabled = False
    staff.Visible = False
    LwK.Visible = False
    cosas1.Visible = False
    hnx.Visible = False
    cosas2.Visible = False
    stk.Visible = False
    cosas3.Visible = False
    cosas7.Visible = False
    cosas4.Visible = False
    cosas5.Visible = False
    cosas6.Visible = False
End If
End Sub
