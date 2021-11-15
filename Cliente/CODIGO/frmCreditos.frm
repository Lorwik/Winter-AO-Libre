VERSION 5.00
Begin VB.Form frmCreditos 
   BorderStyle     =   0  'None
   ClientHeight    =   8850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   Picture         =   "frmCreditos.frx":0000
   ScaleHeight     =   8850
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Subepos 
      Interval        =   1
      Left            =   120
      Top             =   120
   End
   Begin VB.Image Image1 
      Height          =   8895
      Left            =   0
      Top             =   0
      Width           =   12015
   End
   Begin VB.Label gracias 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GRACIAS"
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
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   15600
      Width           =   7095
   End
   Begin VB.Label cosas6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCreditos.frx":4D8F3
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
      Height          =   1695
      Left            =   4320
      TabIndex        =   11
      Top             =   13800
      Width           =   3375
   End
   Begin VB.Label cosas5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "   Gs-Zone                     LwK-Foros"
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
      Left            =   5280
      TabIndex        =   10
      Top             =   12960
      Width           =   1095
   End
   Begin VB.Label agradecimientos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Agradecimientos Especiales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   12600
      Width           =   7095
   End
   Begin VB.Label cosas4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Balance, Mapeo"
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
      Left            =   3480
      TabIndex        =   8
      Top             =   12120
      Width           =   4695
   End
   Begin VB.Label mrt 
      BackStyle       =   0  'Transparent
      Caption         =   "Mortis"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   7
      Top             =   11760
      Width           =   975
   End
   Begin VB.Label cosas3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Jefe de Balance, Dateo, Mapeo"
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
      Left            =   3480
      TabIndex        =   6
      Top             =   11400
      Width           =   4695
   End
   Begin VB.Label stk 
      BackStyle       =   0  'Transparent
      Caption         =   "Stick"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   5
      Top             =   11040
      Width           =   975
   End
   Begin VB.Label cosas2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Director de Game Master, Jefe de Mapeo"
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
      Left            =   3480
      TabIndex        =   4
      Top             =   10680
      Width           =   4695
   End
   Begin VB.Label hnx 
      BackStyle       =   0  'Transparent
      Caption         =   "Hennox"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   3
      Top             =   10320
      Width           =   975
   End
   Begin VB.Label cosas1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Director, idea original, Indexación, Graficación, Programación"
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
      Height          =   615
      Left            =   3480
      TabIndex        =   2
      Top             =   9720
      Width           =   4695
   End
   Begin VB.Label LwK 
      BackStyle       =   0  'Transparent
      Caption         =   "Lorwik"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   1
      Top             =   9360
      Width           =   975
   End
   Begin VB.Label staff 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Staff de Winter-AO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   8640
      Width           =   7095
   End
End
Attribute VB_Name = "frmCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
       Unload Me
End If
End Sub
Private Sub Form_Load()
On Error Resume Next
If MPTres Then
Windows_Temp_Dir = General_Get_Temp_Dir
 Set MP3P = New clsMP3Player
    Call Extract_File2(mp3, App.Path & "\ARCHIVOS\", "3.mp3", Windows_Temp_Dir, False)
    MP3P.stopMP3
    MP3P.mp3file = Windows_Temp_Dir & "3.mp3"
    MP3P.playMP3
    MP3P.Volume = 1000
    Delete_File (Windows_Temp_Dir & "3.mp3")
    End If
End Sub

Private Sub Image1_Click()
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
agradecimientos.Top = Val(agradecimientos.Top) - 1
cosas5.Top = Val(cosas5.Top) - 1
cosas6.Top = Val(cosas6.Top) - 1
gracias.Top = Val(gracias.Top) - 1
End Sub
