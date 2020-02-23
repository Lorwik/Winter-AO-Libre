VERSION 5.00
Begin VB.Form frmUnionAdyacente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Union con Mapas Adyasentes"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14460
   Icon            =   "frmUnionAdyasente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   346
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   964
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5400
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MOSTRAR   MAPA   DEL   JUEGO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   6050
      TabIndex        =   43
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4635
      Left            =   6360
      Picture         =   "frmUnionAdyasente.frx":628A
      ScaleHeight     =   307
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   534
      TabIndex        =   41
      Top             =   480
      Width           =   8040
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         FillStyle       =   4  'Upward Diagonal
         Height          =   350
         Left            =   3360
         Top             =   3850
         Width           =   640
      End
   End
   Begin VB.CheckBox AutoMapeo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Auto-Mapeo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1200
      TabIndex        =   40
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox AutoMapeo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Auto-Mapeo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   38
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox AutoMapeo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Auto-Mapeo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   37
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin WorldEditor.lvButtons_H cmdAplicar 
      Height          =   375
      Left            =   3240
      TabIndex        =   29
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Aplicar"
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
   Begin VB.TextBox PosLim 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   360
      TabIndex        =   26
      Text            =   "89"
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox PosLim 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   5640
      TabIndex        =   24
      Text            =   "12"
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox PosLim 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   5520
      TabIndex        =   22
      Text            =   "11"
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox PosLim 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   480
      TabIndex        =   20
      Text            =   "90"
      Top             =   360
      Width           =   375
   End
   Begin VB.CheckBox Aplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aplicar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1200
      TabIndex        =   19
      Top             =   2160
      Width           =   975
   End
   Begin VB.CheckBox Aplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aplicar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   18
      Top             =   2760
      Width           =   975
   End
   Begin VB.CheckBox Aplicar 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aplicar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   17
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Mapa 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   16
      Text            =   "0"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Mapa 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2520
      TabIndex        =   15
      Text            =   "0"
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Mapa 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   4200
      TabIndex        =   14
      Text            =   "0"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox PosLim 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   840
      TabIndex        =   13
      Text            =   "11"
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox PosLim 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   4800
      TabIndex        =   12
      Text            =   "90"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox PosLim 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   5640
      TabIndex        =   11
      Text            =   "10"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox PosLim 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Text            =   "91"
      Top             =   3240
      Width           =   375
   End
   Begin VB.CheckBox Aplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aplicar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Mapa 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2520
      TabIndex        =   0
      Text            =   "0"
      Top             =   840
      Width           =   735
   End
   Begin WorldEditor.lvButtons_H cmdCancelar 
      Height          =   375
      Left            =   4680
      TabIndex        =   30
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Cancelar"
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
   Begin WorldEditor.lvButtons_H cmdDefault 
      Height          =   375
      Left            =   120
      TabIndex        =   31
      Top             =   4080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&Default"
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
   Begin VB.CheckBox AutoMapeo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Auto-Mapeo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   39
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   "Mapa de Winter-AO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   42
      Top             =   0
      Width           =   8055
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Leyenda sobre posiciones:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   36
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   256
      X2              =   256
      Y1              =   330
      Y2              =   338.667
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Posicion Y del mapa actual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   3960
      TabIndex        =   35
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   256
      X2              =   256
      Y1              =   313
      Y2              =   323
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Posicion Y del mapa destino"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   210
      Left            =   3960
      TabIndex        =   34
      Top             =   4680
      Width           =   2025
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   104
      X2              =   104
      Y1              =   330
      Y2              =   338.667
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Posicion X del mapa actual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   1680
      TabIndex        =   33
      Top             =   4920
      Width           =   1920
   End
   Begin VB.Line Line15 
      BorderColor     =   &H008080FF&
      BorderWidth     =   2
      X1              =   104
      X2              =   104
      Y1              =   313
      Y2              =   323
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Posicion X del mapa destino"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   210
      Left            =   1680
      TabIndex        =   32
      Top             =   4680
      Width           =   2010
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00008000&
      X1              =   8
      X2              =   400
      Y1              =   304
      Y2              =   304
   End
   Begin VB.Label Label13 
      Caption         =   "NOTA: Mapa 0, borra el translado de mapa."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1320
      TabIndex        =   28
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Line Line13 
      BorderColor     =   &H008080FF&
      BorderWidth     =   2
      X1              =   56
      X2              =   56
      Y1              =   56
      Y2              =   224
   End
   Begin VB.Label Label12 
      Caption         =   "X:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   720
      Width           =   255
   End
   Begin VB.Line Line12 
      BorderColor     =   &H008080FF&
      BorderWidth     =   2
      X1              =   352
      X2              =   352
      Y1              =   48
      Y2              =   216
   End
   Begin VB.Label Label11 
      Caption         =   "X:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   5400
      TabIndex        =   25
      Top             =   3120
      Width           =   255
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   72
      X2              =   352
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label10 
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   5280
      TabIndex        =   23
      Top             =   3480
      Width           =   255
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   56
      X2              =   336
      Y1              =   32
      Y2              =   32
   End
   Begin VB.Label Label9 
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   360
      Width           =   255
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00404040&
      X1              =   64
      X2              =   344
      Y1              =   232
      Y2              =   232
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00404040&
      X1              =   344
      X2              =   344
      Y1              =   232
      Y2              =   40
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00404040&
      X1              =   344
      X2              =   64
      Y1              =   40
      Y2              =   40
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00404040&
      X1              =   64
      X2              =   64
      Y1              =   232
      Y2              =   40
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00008000&
      X1              =   8
      X2              =   400
      Y1              =   264
      Y2              =   264
   End
   Begin VB.Label Label8 
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5400
      TabIndex        =   9
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "X:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "X:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   120
      Width           =   255
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   48
      X2              =   328
      Y1              =   224
      Y2              =   224
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   80
      X2              =   360
      Y1              =   48
      Y2              =   48
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mapa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mapa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mapa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mapa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   336
      X2              =   336
      Y1              =   24
      Y2              =   216
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   72
      X2              =   72
      Y1              =   56
      Y2              =   240
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   2895
      Left            =   960
      Top             =   600
      Width           =   4215
   End
   Begin VB.Menu mnuDefault 
      Caption         =   "mnuDefault"
      Visible         =   0   'False
      Begin VB.Menu mnuLegal 
         Caption         =   "Borde Legal Automatico"
      End
      Begin VB.Menu mnuBasica 
         Caption         =   "11,10 y 90,91 - Basica"
      End
      Begin VB.Menu mnuUlla 
         Caption         =   "9,7 y 92,94 - Ullathorpe"
      End
   End
End
Attribute VB_Name = "frmUnionAdyacente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************
Option Explicit
Private ParaX As Long
Private ParaY As Long

Private Sub Aplicar_Click(index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
cmdAplicar.Enabled = False
For i = 0 To 3
    If Aplicar(i).value = 1 Then cmdAplicar.Enabled = True
Next
End Sub

Private Sub cmdAplicar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
Dim y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If

modEdicion.Deshacer_Add "Insertar Translados a mapas Adyasentes" ' Hago deshacer

' ARRIBA
If Mapa(0).text > -1 And Aplicar(0).value = 1 Then
    y = PosLim(1).text
    For X = (PosLim(3).text + 1) To (PosLim(2).text - 1)
        If MapData(X, y).Blocked = 0 Then
            MapData(X, y).TileExit.Map = Mapa(0).text
            If Mapa(0).text = 0 Then
                MapData(X, y).TileExit.X = 0
                MapData(X, y).TileExit.y = 0
            Else
                MapData(X, y).TileExit.X = X
                MapData(X, y).TileExit.y = PosLim(4).text
            End If
        End If
    Next
End If

' DERECHA
If Mapa(1).text > -1 And Aplicar(1).value = 1 Then
    X = PosLim(2).text
    For y = (PosLim(1).text + 1) To (PosLim(0).text - 1)
        If MapData(X, y).Blocked = 0 Then
            MapData(X, y).TileExit.Map = Mapa(1).text
                If Mapa(1).text = 0 Then
                    MapData(X, y).TileExit.X = 0
                    MapData(X, y).TileExit.y = 0
                Else
                    MapData(X, y).TileExit.X = PosLim(6).text
                    MapData(X, y).TileExit.y = y
                End If
        End If
    Next
End If

' ABAJO
If Mapa(2).text > -1 And Aplicar(2).value = 1 Then
    y = PosLim(0).text
    For X = (PosLim(3).text + 1) To (PosLim(2).text - 1)
        If MapData(X, y).Blocked = 0 Then
            MapData(X, y).TileExit.Map = Mapa(2).text
                If Mapa(2).text = 0 Then
                    MapData(X, y).TileExit.X = 0
                    MapData(X, y).TileExit.y = 0
                Else
                    MapData(X, y).TileExit.X = X
                    MapData(X, y).TileExit.y = PosLim(5).text
                End If
        End If
    Next
End If

' IZQUIERDA
If Mapa(3).text > -1 And Aplicar(3).value = 1 Then
    X = PosLim(3).text
    For y = (PosLim(1).text + 1) To (PosLim(0).text - 1)
        If MapData(X, y).Blocked = 0 Then
            MapData(X, y).TileExit.Map = Mapa(3).text
                If Mapa(3).text = 0 Then
                    MapData(X, y).TileExit.X = 0
                    MapData(X, y).TileExit.y = 0
                Else
                    MapData(X, y).TileExit.X = PosLim(7).text
                    MapData(X, y).TileExit.y = y
                End If
        End If
    Next
End If

'Set changed flag
MapInfo.Changed = 1
DoEvents

Unload Me
End Sub

Private Sub cmdCancelar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Unload Me
End Sub

Private Sub cmdDefault_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Me.PopupMenu mnuDefault
End Sub

''
'   Lee los Translados existentes en lugares claves en el Mapa
'

Private Sub LeerMapaExit()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next

Dim X As Integer
Dim y As Integer

' ARRIBA
Mapa(0).text = 0
y = PosLim(1).text
For X = (PosLim(3).text + 1) To (PosLim(2).text - 1)
        If MapData(X, y).TileExit.Map > 0 Then
            Mapa(0).text = MapData(X, y).TileExit.Map
            Exit For
        End If
Next
Aplicar(0).value = 0

' DERECHA
Mapa(1).text = 0
X = PosLim(2).text
For y = (PosLim(1).text + 1) To (PosLim(0).text - 1)
        If MapData(X, y).TileExit.Map > 0 Then
            Mapa(1).text = MapData(X, y).TileExit.Map
            Exit For
        End If
Next
Aplicar(1).value = 0

' ABAJO
Mapa(2).text = 0
y = PosLim(0).text
For X = (PosLim(3).text + 1) To (PosLim(2).text - 1)
        If MapData(X, y).TileExit.Map > 0 Then
            Mapa(2).text = MapData(X, y).TileExit.Map
            Exit For
        End If
Next
Aplicar(2).value = 0

' IZQUIERDA
Mapa(3).text = 0
X = PosLim(3).text
For y = (PosLim(1).text + 1) To (PosLim(0).text - 1)
        If MapData(X, y).TileExit.Map > 0 Then
            Mapa(3).text = MapData(X, y).TileExit.Map
            Exit For
        End If
Next
Aplicar(3).value = 0


End Sub

Private Sub Command1_Click()
If Not Me.Width = 6405 Then
    Me.Width = 6405
Else
    Me.Width = 14490
End If
End Sub

Private Sub Form_Load()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call mnuBasica_Click
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
ParaX = X
ParaY = y
End Sub

Private Sub picture1_click()
Shape2.Left = ParaX
Shape2.Top = ParaY
End Sub
Private Sub Mapa_Change(index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Aplicar(index).value = 1
End Sub

Private Sub Mapa_KeyPress(index As Integer, KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then
    KeyAscii = 0
    Exit Sub
End If

End Sub

Private Sub Mapa_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/10/06
'*************************************************
If LenB(Mapa(index).text) = 0 Then Mapa(index).text = 0
If Mapa(index).text > 1024 Then Mapa(index).text = 1024
End Sub

Private Sub mnuBasica_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
PosLim(0).text = 91
PosLim(1).text = 10
PosLim(2).text = 90
PosLim(3).text = 11
PosLim(4).text = 90
PosLim(5).text = 11
PosLim(6).text = 12
PosLim(7).text = 89
Call LeerMapaExit
End Sub

Private Sub mnuLegal_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 02/10/06
'*************************************************
PosLim(0).text = MaxYBorder
PosLim(1).text = MinYBorder
PosLim(2).text = MaxXBorder
PosLim(3).text = MinXBorder
PosLim(4).text = MaxYBorder - 1
PosLim(5).text = MinYBorder + 1
PosLim(6).text = MinXBorder + 1
PosLim(7).text = MaxXBorder - 1
Call LeerMapaExit
End Sub

Private Sub mnuUlla_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
PosLim(0).text = 94
PosLim(1).text = 7
PosLim(2).text = 92
PosLim(3).text = 9
PosLim(4).text = 93
PosLim(5).text = 8
PosLim(6).text = 10
PosLim(7).text = 91
Call LeerMapaExit
End Sub

Private Sub PosLim_KeyPress(index As Integer, KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then
    KeyAscii = 0
    Exit Sub
End If

End Sub

Private Sub PosLim_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
On Error Resume Next
If LenB(PosLim(index).text) = 0 Then PosLim(index).text = 1
If PosLim(index).text > 99 Then PosLim(index) = 99
If PosLim(index).text < 1 Then PosLim(index) = 1

Dim y As Integer
Dim X As Integer

' ARRIBA
y = PosLim(1).text
For X = (PosLim(3).text + 1) To (PosLim(2).text - 1)
        If MapData(X, y).TileExit.Map > 0 Then
            Mapa(0).text = MapData(X, y).TileExit.Map
            Aplicar(0).value = 0
            Exit For
        End If
Next

' DERECHA
X = PosLim(2).text
For y = (PosLim(1).text + 1) To (PosLim(0).text - 1)
        If MapData(X, y).TileExit.Map > 0 Then
            Mapa(1).text = MapData(X, y).TileExit.Map
            Aplicar(1).value = 0
            Exit For
        End If
Next

' ABAJO
y = PosLim(0).text
For X = (PosLim(3).text + 1) To (PosLim(2).text - 1)
        If MapData(X, y).TileExit.Map > 0 Then
            Mapa(2).text = MapData(X, y).TileExit.Map
            Aplicar(2).value = 0
            Exit For
        End If
Next

' IZQUIERDA
X = PosLim(3).text
For y = (PosLim(1).text + 1) To (PosLim(0).text - 1)
        If MapData(X, y).TileExit.Map > 0 Then
            Mapa(3).text = MapData(X, y).TileExit.Map
            Aplicar(3).value = 0
            Exit For
        End If
Next

End Sub
