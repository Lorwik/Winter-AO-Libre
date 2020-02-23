VERSION 5.00
Begin VB.Form frmEstadisticas 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Estadisticas"
   ClientHeight    =   8040
   ClientLeft      =   75
   ClientTop       =   -105
   ClientWidth     =   7725
   ForeColor       =   &H00FFFFFF&
   Icon            =   "FrmEstadisticas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   536
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   4590
      TabIndex        =   39
      Top             =   7050
      Width           =   90
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   2280
      TabIndex        =   38
      Top             =   4560
      Width           =   90
   End
   Begin VB.Image Command3 
      Height          =   255
      Left            =   840
      Top             =   6570
      Width           =   1335
   End
   Begin VB.Image Command2 
      Height          =   255
      Left            =   780
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   41
      Left            =   6705
      Top             =   6750
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   40
      Left            =   6885
      Top             =   6750
      Width           =   195
   End
   Begin VB.Label text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   4440
      TabIndex        =   37
      Top             =   6450
      Width           =   90
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   39
      Left            =   6705
      Top             =   6450
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   38
      Left            =   6885
      Top             =   6450
      Width           =   195
   End
   Begin VB.Label text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   5445
      TabIndex        =   36
      Top             =   6150
      Width           =   90
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   37
      Left            =   6705
      Top             =   6150
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   36
      Left            =   6885
      Top             =   6150
      Width           =   195
   End
   Begin VB.Label text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   5100
      TabIndex        =   35
      Top             =   5850
      Width           =   90
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   1
      Left            =   6705
      Top             =   7410
      Width           =   195
   End
   Begin VB.Label text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   4410
      TabIndex        =   34
      Top             =   5550
      Width           =   90
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   35
      Left            =   6705
      Top             =   5850
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   34
      Left            =   6885
      Top             =   5850
      Width           =   195
   End
   Begin VB.Label text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   4320
      TabIndex        =   33
      Top             =   5250
      Width           =   90
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   33
      Left            =   6705
      Top             =   5550
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   32
      Left            =   6885
      Top             =   5550
      Width           =   195
   End
   Begin VB.Label text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   4230
      TabIndex        =   32
      Top             =   4950
      Width           =   90
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   31
      Left            =   6705
      Top             =   5250
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   30
      Left            =   6885
      Top             =   5250
      Width           =   195
   End
   Begin VB.Label text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   3975
      TabIndex        =   31
      Top             =   4650
      Width           =   90
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   29
      Left            =   6720
      Top             =   7050
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   28
      Left            =   6900
      Top             =   7050
      Width           =   195
   End
   Begin VB.Label text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   5385
      TabIndex        =   30
      Top             =   4350
      Width           =   90
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   27
      Left            =   6705
      Top             =   4950
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   26
      Left            =   6885
      Top             =   4950
      Width           =   195
   End
   Begin VB.Label text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   4395
      TabIndex        =   29
      Top             =   4050
      Width           =   90
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   25
      Left            =   6705
      Top             =   4650
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   24
      Left            =   6885
      Top             =   4650
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   23
      Left            =   6705
      Top             =   4350
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   22
      Left            =   6885
      Top             =   4350
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   21
      Left            =   6705
      Top             =   4125
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   20
      Left            =   6885
      Top             =   4050
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   19
      Left            =   6705
      Top             =   3750
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   18
      Left            =   6885
      Top             =   3750
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   17
      Left            =   6705
      Top             =   3450
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   16
      Left            =   6885
      Top             =   3450
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   15
      Left            =   6705
      Top             =   3150
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   14
      Left            =   6885
      Top             =   3150
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   13
      Left            =   6705
      Top             =   2850
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   12
      Left            =   6885
      Top             =   2850
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   11
      Left            =   6705
      Top             =   2550
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   10
      Left            =   6885
      Top             =   2550
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   9
      Left            =   6705
      Top             =   2250
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   8
      Left            =   6885
      Top             =   2250
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   7
      Left            =   6705
      Top             =   1950
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   6
      Left            =   6885
      Top             =   1950
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   5
      Left            =   6705
      Top             =   1650
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   4
      Left            =   6885
      Top             =   1650
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   3
      Left            =   6705
      Top             =   1350
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   2
      Left            =   6885
      Top             =   1350
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   0
      Left            =   6885
      Top             =   7410
      Width           =   195
   End
   Begin VB.Label text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   4860
      TabIndex        =   28
      Top             =   3750
      Width           =   90
   End
   Begin VB.Label text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   4770
      TabIndex        =   27
      Top             =   3450
      Width           =   90
   End
   Begin VB.Label text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   4455
      TabIndex        =   26
      Top             =   3150
      Width           =   90
   End
   Begin VB.Label text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   4440
      TabIndex        =   25
      Top             =   2850
      Width           =   90
   End
   Begin VB.Label text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   4320
      TabIndex        =   24
      Top             =   2550
      Width           =   105
   End
   Begin VB.Label text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   5385
      TabIndex        =   23
      Top             =   2250
      Width           =   90
   End
   Begin VB.Label text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   5280
      TabIndex        =   22
      Top             =   1950
      Width           =   105
   End
   Begin VB.Label text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   4170
      TabIndex        =   21
      Top             =   1650
      Width           =   90
   End
   Begin VB.Label text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   4125
      TabIndex        =   20
      Top             =   1350
      Width           =   105
   End
   Begin VB.Label text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   4560
      TabIndex        =   19
      Top             =   7350
      Width           =   90
   End
   Begin VB.Label text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   21
      Left            =   4560
      TabIndex        =   18
      Top             =   6750
      Width           =   90
   End
   Begin VB.Label Puntos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4620
      TabIndex        =   17
      Top             =   870
      Width           =   90
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   1950
      TabIndex        =   16
      Top             =   5775
      Width           =   90
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   1005
      TabIndex        =   15
      Top             =   5535
      Width           =   90
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   1680
      TabIndex        =   14
      Top             =   5265
      Width           =   90
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   2100
      TabIndex        =   13
      Top             =   5010
      Width           =   90
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   2325
      TabIndex        =   12
      Top             =   4785
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   1230
      TabIndex        =   11
      Top             =   3870
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   1005
      TabIndex        =   10
      Top             =   3660
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   1065
      TabIndex        =   9
      Top             =   3465
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   1200
      TabIndex        =   8
      Top             =   3225
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   1230
      TabIndex        =   7
      Top             =   3000
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   1230
      TabIndex        =   6
      Top             =   2775
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   1200
      TabIndex        =   5
      Top             =   2550
      Width           =   90
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   1575
      TabIndex        =   4
      Top             =   1830
      Width           =   90
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   1245
      TabIndex        =   3
      Top             =   1575
      Width           =   90
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   1455
      TabIndex        =   2
      Top             =   1320
      Width           =   90
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   1215
      TabIndex        =   1
      Top             =   1080
      Width           =   90
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   1200
      TabIndex        =   0
      Top             =   795
      Width           =   90
   End
End
Attribute VB_Name = "frmEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click(Index As Integer)

Call General_Set_Wav(SND_CLICK)

Dim indice
If (Index And &H1) = 0 Then
    If Alocados > 0 Then
        indice = Index \ 2 + 1
        If indice > NUMSKILLS Then indice = NUMSKILLS
        If Val(Text1(indice).Caption) < MAXSKILLPOINTS Then
            Text1(indice).Caption = Val(Text1(indice).Caption) + 1
            flags(indice) = flags(indice) + 1
            Alocados = Alocados - 1
        End If
            
    End If
Else
    If Alocados < SkillPoints Then
        
        indice = Index \ 2 + 1
        If Val(Text1(indice).Caption) > 0 And flags(indice) > 0 Then
            Text1(indice).Caption = Val(Text1(indice).Caption) - 1
            flags(indice) = flags(indice) - 1
            Alocados = Alocados + 1
        End If
    End If
End If

Puntos.Caption = Alocados
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Public Sub Iniciar_Labels()
'Iniciamos los labels con los valores de los atributos y los skills
Dim i As Integer
For i = 1 To NUMATRIBUTOS
    Atri(i).Caption = UserAtributos(i)
Next

For i = 1 To NUMSKILLS
    Text1(i).Caption = UserSkills(i)
Next i

'Flags para saber que skills se modificaron
ReDim flags(1 To NUMSKILLS)

Label4(1).Caption = UserReputacion.AsesinoRep
Label4(2).Caption = UserReputacion.BandidoRep
Label4(3).Caption = UserReputacion.BurguesRep
Label4(4).Caption = UserReputacion.LadronesRep
Label4(5).Caption = UserReputacion.NobleRep
Label4(6).Caption = UserReputacion.PlebeRep

If UserReputacion.promedio < 0 Then
    Label4(7).ForeColor = vbRed
    Label4(7).Caption = "CRIMINAL"
Else
    Label4(7).ForeColor = vbBlue
    Label4(7).Caption = "Ciudadano"
End If

With UserEstadisticas
    Label6(0).Caption = .CriminalesMatados
    Label6(1).Caption = .CiudadanosMatados
    Label6(2).Caption = .UsuariosMatados
    Label6(3).Caption = .NpcsMatados
    Label6(4).Caption = .Clase
    Label6(5).Caption = .PenaCarcel
End With

End Sub

Private Sub Command3_Click()
    Dim skillChanges(NUMSKILLS) As Byte
    Dim i As Long

    For i = 1 To NUMSKILLS
        skillChanges(i) = CByte(Text1(i).Caption) - UserSkills(i)
        'Actualizamos nuestros datos locales
        UserSkills(i) = Val(Text1(i).Caption)
    Next i
    
    Call WriteModifySkills(skillChanges())
    SkillPoints = Alocados
    MsgBox "Skills guardados.", vbInformation, "Winter AO Ultimate - Skills"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("47.gif")
End Sub

