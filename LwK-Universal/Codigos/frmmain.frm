VERSION 5.00
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form FrmMain 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Compresor/descompresor LwK Universal"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   5670
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin vbalProgBarLib6.vbalProgressBar Barrita 
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   2640
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
      Picture         =   "frmmain.frx":000C
      ForeColor       =   0
      Appearance      =   2
      BarPicture      =   "frmmain.frx":0028
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Graficos"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      Caption         =   "MP3"
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   1080
      Width           =   855
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Wav"
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   1080
      Width           =   855
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Inits"
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   1080
      Width           =   855
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Interface"
      Height          =   255
      Left            =   3720
      TabIndex        =   10
      Top             =   1080
      Width           =   975
   End
   Begin VB.OptionButton Option6 
      Caption         =   "Map"
      Height          =   255
      Left            =   4680
      TabIndex        =   9
      Top             =   1080
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Caption         =   "Parches"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   5415
      Begin VB.CommandButton Command17 
         Caption         =   "Desencriptar Pach"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   8
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Crear Pach"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Descomprension"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   2880
      TabIndex        =   4
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton Command15 
         Caption         =   "Descomprimir"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Comprension"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FFFF&
         Caption         =   "Comprimir"
         Height          =   255
         Left            =   120
         MaskColor       =   &H0000FFFF&
         TabIndex        =   3
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Cerrar"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   3120
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Compresor LwK Universal"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   5415
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Tipo As resource_file_type

Private Sub Command1_Click()
Label1.Caption = "ESPERE!!Trabajando"
Label1.BackColor = &HFF&
Call Compress_Files(Tipo, App.Path, App.Path & "\COMPRIMIDOS\")
Label1.Caption = "Compresion terminada"
Label1.BackColor = &HFF00&
MsgBox "Compresion terminada"
End Sub
Private Sub Command11_Click()
End
End Sub
Private Sub Command15_Click()
Label1.Caption = "ESPERE!!Trabajando"
Label1.BackColor = &HFF&
Call Extract_All_Files(Tipo, App.Path & "\COMPRIMIDOS")
Label1.Caption = "Extracción terminada"
Label1.BackColor = &HFF00&
MsgBox "Extracción terminada"
End Sub

Private Sub Command16_Click()
Call Compress_Files(Patch, App.Path, App.Path & "\COMPRIMIDOS\")
End Sub

Private Sub Command17_Click()
    If General_File_Exists(App.Path & "\tmp.WAO", vbNormal) Then
    '    'Instalamos el Parche
        Extract_Patch App.Path & "\COMPRIMIDOS", App.Path & "\tmp.WAO"
    '
    '    'Esperamos a que termine
        DoEvents
    '
    '    'Borramos el Parche
        Kill App.Path & "\tmp.WAO"
    Else
        MsgBox "No se encuentro el archivo de parche."
        MsgBox App.Path & "\tmp.WAO"
    End If
    
End Sub

Public Function General_File_Exists(ByVal file_path As String, ByVal file_type As VbFileAttribute) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Checks to see if a file exists
'*****************************************************************
    If Dir(file_path, file_type) = "" Then
        General_File_Exists = False
    Else
        General_File_Exists = True
    End If
End Function

Private Sub Option1_Click()
Tipo = graphics
End Sub

Private Sub Option2_Click()
Tipo = MP3
End Sub

Private Sub Option3_Click()
Tipo = Wav
End Sub

Private Sub Option4_Click()
Tipo = Scripts
End Sub

Private Sub Option5_Click()
Tipo = Interface
End Sub

Private Sub Option6_Click()
Tipo = Map
End Sub

