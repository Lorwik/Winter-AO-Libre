VERSION 5.00
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Indexador RincondelAO - By Lorwik - www.RincondelAO.com.ar"
   ClientHeight    =   5520
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   7740
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   368
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   516
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ver cabezas, cascos, armas, cuerpos, escudos y FX."
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   4560
      Width           =   3975
   End
   Begin VB.Timer Anim 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox VisorGraf 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   1800
      MouseIcon       =   "frmmain.frx":0ECA
      Picture         =   "frmmain.frx":2314
      ScaleHeight     =   301
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   393
      TabIndex        =   4
      Top             =   0
      Width           =   5895
   End
   Begin VB.ListBox VisorGrh 
      Height          =   4545
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label cargados 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Seleccione un Grh para ver su información..."
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   5160
      Width           =   7695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Información del Grh Seleccionado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4920
      Width           =   7695
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Indexarmnu 
         Caption         =   "Indexar"
         Begin VB.Menu Indexar 
            Caption         =   "Indexar Graficos"
         End
         Begin VB.Menu IndexHead 
            Caption         =   "Indexar Cabezas"
         End
      End
      Begin VB.Menu Desindexarmnu 
         Caption         =   "Desindexar"
         Begin VB.Menu Desindexar 
            Caption         =   "Desindexar Graficos"
         End
         Begin VB.Menu DesindexHead 
            Caption         =   "Desindexar Cabezas"
         End
      End
      Begin VB.Menu Actualizar 
         Caption         =   "Actualizar"
      End
   End
   Begin VB.Menu AutoIndexar 
      Caption         =   "AutoIndexar"
      Begin VB.Menu GrafInv 
         Caption         =   "Graficos de inventario"
      End
      Begin VB.Menu Grfbig 
         Caption         =   "Graficos de 1 sola imagen (Arboles, carteles, etc...)"
      End
   End
   Begin VB.Menu Mas 
      Caption         =   "Mas"
      Begin VB.Menu About 
         Caption         =   "Sobre..."
      End
      Begin VB.Menu Help 
         Caption         =   "Ayuda"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Means no grh is being rendered.
Private Const INFINITE_LOOPS As Long = -1
Private currentGrh As Long
''
' The current frame of the grh being displayed
Private currentFrame As Long

Private Sub About_Click()
frmAbout.Show
End Sub

Private Sub Actualizar_Click()
Call LoadGrhData
End Sub
Private Sub Anim_Timer()
    
    'If an animated grh is chosen, animate!
    If currentGrh <> INFINITE_LOOPS Then
        If GrhData(currentGrh).NumFrames > 1 Then
            'Move to next animation frame!
            currentFrame = currentFrame + 1
            
            If currentFrame > GrhData(currentGrh).NumFrames Then
                currentFrame = 1
            End If

        'Limpio
        frmmain.VisorGraf.Cls
        'Dibujo
        Call DrawGrhtoHdc(frmmain.VisorGraf.hDC, GrhData(currentGrh).Frames(currentFrame), 0, 0, False)
        'Y actualizo
        frmmain.VisorGraf.Refresh
        End If
    End If
End Sub

Private Sub Command1_Click()
frmExtra.Show
End Sub

Private Sub Desindexar_Click()
Call Makeini
End Sub

Private Sub DesindexHead_Click()
Call MakeiniHead
End Sub

Private Sub GrafInv_Click()
Dim Numero As Long
Numero = CInt(Val(InputBox("Escriba el numero del grafico que desea indexar.", "Indexar Automatico 32x32", "1")))

Call WriteVar(App.path & "\Graficos.ini", "Init", "NumGrh", grhCount + 1)

Open "Graficos.ini" For Append Shared As #1
Print #1, "Grh" & grhCount + 1 & "=1-" & Numero & "-0-0-32-32"
Close #1

MsgBox "Se agrego la linea a Graficos.ini. Ahora indexe y actualice para ver cambios."
End Sub

Private Sub Grfbig_Click()
Dim Numero As Long
Dim Ancho As Integer
Dim Alto As Integer

Numero = CInt(Val(InputBox("Escriba el numero del grafico que desea indexar.", "Indexar Automatico", "1")))
Ancho = CInt(Val(InputBox("Escriba el ancho de la imagen (Dejar el cursor sobre la imagen)", "Indexar Automatico", "1")))
Alto = CInt(Val(InputBox("Escriba el alto de la imagen (Dejar el cursor sobre la imagen)", "Indexar Automatico", "1")))

Call WriteVar(App.path & "\Graficos.ini", "Init", "NumGrh", grhCount + 1)

Open "Graficos.ini" For Append Shared As #1
Print #1, "Grh" & grhCount + 1 & "=1-" & Numero & "-0-0-" & Ancho & "-" & Alto
Close #1

MsgBox "Se agrego la linea a Graficos.ini. Ahora indexe y actualice para ver cambios."
End Sub

Private Sub Help_Click()
MsgBox "Actualmente no disponible"
End Sub

Private Sub Indexar_Click()
    If Not LoadGrh.Indexar Then
        Call MsgBox("The file could not be saved. This could be caused due to lack of space on disk.")
    Else
        Call MsgBox("File succesfully written.")
    End If
End Sub

Private Sub IndexHead_Click()
    If Not LoadGrh.IndexarHead Then
        Call MsgBox("The file could not be saved. This could be caused due to lack of space on disk.")
    Else
        Call MsgBox("File succesfully written.")
    End If
End Sub

Private Sub VisorGrh_Click()
Dim Animation As String
Dim frame As Integer
Dim i As Byte

If VisorGrh.ListIndex < 0 Then Exit Sub

With GrhData(VisorGrh.ListIndex + 1)

    currentGrh = Val(VisorGrh.Text)

    If .NumFrames > 1 Then
        Anim.Interval = Round(GrhData(currentGrh).Speed / GrhData(currentGrh).NumFrames)
        Anim.Enabled = True
        
            Animation = ""
            For frame = 1 To .NumFrames
                Animation = Animation & .Frames(frame) & "-"
            Next frame
        
        Label2.Caption = "Grh" & i & "=" & .NumFrames & "-" & Animation & .Speed
    Else
        Anim.Enabled = False
        Label2.Caption = "Grh" & i & "=" & .NumFrames & "-" & .FileNum & "-" & .SX & "-" & .SY & "-" & GrhData(.Frames(1)).pixelWidth & "-" & GrhData(.Frames(1)).pixelHeight
        'Limpio
        frmmain.VisorGraf.Cls
        'Dibujo
        Call DrawGrhtoHdc(frmmain.VisorGraf.hDC, currentGrh + 1, 0, 0, False)
        'Y actualizo
        frmmain.VisorGraf.Refresh
    End If
    
End With
End Sub
