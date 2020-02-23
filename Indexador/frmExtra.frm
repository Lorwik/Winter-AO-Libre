VERSION 5.00
Begin VB.Form frmExtra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver cabezas, cascos, armas, cuerpos, escudos y FX."
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10410
   FillStyle       =   0  'Solid
   Icon            =   "frmExtra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   485
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   694
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Caption         =   "FX's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   6960
      TabIndex        =   35
      Top             =   3120
      Width           =   3375
      Begin VB.ListBox VisorFX 
         Height          =   2595
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   1215
      End
      Begin VB.PictureBox FX 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   2535
         Left            =   1440
         ScaleHeight     =   2475
         ScaleWidth      =   1755
         TabIndex        =   37
         Top             =   360
         Width           =   1815
      End
      Begin VB.ListBox VisorFXInformacion 
         Height          =   1035
         Left            =   120
         TabIndex        =   36
         Top             =   3000
         Width           =   3135
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Escudos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   6960
      TabIndex        =   28
      Top             =   0
      Width           =   3375
      Begin VB.ListBox VisorEscudos 
         Height          =   2595
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   1215
      End
      Begin VB.ListBox VisorEscudosInformacion 
         Height          =   1035
         Left            =   1440
         TabIndex        =   33
         Top             =   1800
         Width           =   1815
      End
      Begin VB.PictureBox Escudos 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   735
         Index           =   1
         Left            =   2040
         ScaleHeight     =   675
         ScaleWidth      =   555
         TabIndex        =   32
         Top             =   240
         Width           =   615
      End
      Begin VB.PictureBox Escudos 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   735
         Index           =   2
         Left            =   2640
         ScaleHeight     =   675
         ScaleWidth      =   555
         TabIndex        =   31
         Top             =   960
         Width           =   615
      End
      Begin VB.PictureBox Escudos 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   735
         Index           =   3
         Left            =   2040
         ScaleHeight     =   675
         ScaleWidth      =   555
         TabIndex        =   30
         Top             =   960
         Width           =   615
      End
      Begin VB.PictureBox Escudos 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   735
         Index           =   4
         Left            =   1440
         ScaleHeight     =   675
         ScaleWidth      =   555
         TabIndex        =   29
         Top             =   960
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Armas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   3480
      TabIndex        =   21
      Top             =   3000
      Width           =   3375
      Begin VB.ListBox VisorArmasInformacion 
         Height          =   1035
         Left            =   120
         TabIndex        =   27
         Top             =   3000
         Width           =   3135
      End
      Begin VB.PictureBox Armas 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1215
         Index           =   3
         Left            =   1440
         ScaleHeight     =   1155
         ScaleWidth      =   795
         TabIndex        =   26
         Top             =   1560
         Width           =   855
      End
      Begin VB.PictureBox Armas 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1215
         Index           =   4
         Left            =   2280
         ScaleHeight     =   1155
         ScaleWidth      =   795
         TabIndex        =   25
         Top             =   1560
         Width           =   855
      End
      Begin VB.PictureBox Armas 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1215
         Index           =   2
         Left            =   2280
         ScaleHeight     =   1155
         ScaleWidth      =   795
         TabIndex        =   24
         Top             =   360
         Width           =   855
      End
      Begin VB.PictureBox Armas 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1215
         Index           =   1
         Left            =   1440
         ScaleHeight     =   1155
         ScaleWidth      =   795
         TabIndex        =   23
         Top             =   360
         Width           =   855
      End
      Begin VB.ListBox VisorArmas 
         Height          =   2595
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cuerpos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   0
      TabIndex        =   14
      Top             =   3000
      Width           =   3375
      Begin VB.ListBox VisorCuerpos 
         Height          =   2595
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.PictureBox Cuerpos 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1215
         Index           =   1
         Left            =   1440
         ScaleHeight     =   1155
         ScaleWidth      =   795
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
      Begin VB.PictureBox Cuerpos 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1215
         Index           =   2
         Left            =   2280
         ScaleHeight     =   1155
         ScaleWidth      =   795
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
      Begin VB.PictureBox Cuerpos 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1215
         Index           =   4
         Left            =   2280
         ScaleHeight     =   1155
         ScaleWidth      =   795
         TabIndex        =   17
         Top             =   1560
         Width           =   855
      End
      Begin VB.PictureBox Cuerpos 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1215
         Index           =   3
         Left            =   1440
         ScaleHeight     =   1155
         ScaleWidth      =   795
         TabIndex        =   16
         Top             =   1560
         Width           =   855
      End
      Begin VB.ListBox VisorCuerposInformacion 
         Height          =   1035
         Left            =   120
         TabIndex        =   15
         Top             =   3000
         Width           =   3135
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cascos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   3480
      TabIndex        =   7
      Top             =   0
      Width           =   3375
      Begin VB.PictureBox Casco 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   495
         Index           =   4
         Left            =   1440
         ScaleHeight     =   435
         ScaleWidth      =   555
         TabIndex        =   13
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox Casco 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   495
         Index           =   3
         Left            =   2040
         ScaleHeight     =   435
         ScaleWidth      =   555
         TabIndex        =   12
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox Casco 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   495
         Index           =   2
         Left            =   2640
         ScaleHeight     =   435
         ScaleWidth      =   555
         TabIndex        =   11
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox Casco 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   495
         Index           =   1
         Left            =   2040
         ScaleHeight     =   435
         ScaleWidth      =   555
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
      Begin VB.ListBox VisorCascoInformacion 
         Height          =   1425
         Left            =   1440
         TabIndex        =   9
         Top             =   1440
         Width           =   1815
      End
      Begin VB.ListBox VisorCasco 
         Height          =   2595
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cabezas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.Timer AnimFX 
         Enabled         =   0   'False
         Left            =   0
         Top             =   0
      End
      Begin VB.ListBox VisorHeadInformacion 
         Height          =   1425
         Left            =   1440
         TabIndex        =   6
         Top             =   1440
         Width           =   1815
      End
      Begin VB.PictureBox Head 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   495
         Index           =   3
         Left            =   2040
         ScaleHeight     =   435
         ScaleWidth      =   555
         TabIndex        =   5
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox Head 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   495
         Index           =   2
         Left            =   2640
         ScaleHeight     =   435
         ScaleWidth      =   555
         TabIndex        =   4
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox Head 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   495
         Index           =   4
         Left            =   1440
         ScaleHeight     =   435
         ScaleWidth      =   555
         TabIndex        =   3
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox Head 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   495
         Index           =   1
         Left            =   2040
         ScaleHeight     =   435
         ScaleWidth      =   555
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.ListBox VisorHead 
         Height          =   2595
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmExtra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
' The current frame of the grh being displayed
Private currentFrame As Long
Public CurrentFX As Integer

Private Sub AnimFX_Timer()
    'If an animated grh is chosen, animate!
    If CurrentFX <> -1 Then
        If GrhData(CurrentFX).NumFrames > 1 Then
            'Move to next animation frame!
            currentFrame = currentFrame + 1
            
            If currentFrame > GrhData(CurrentFX).NumFrames Then
                currentFrame = 1
            End If

        'Limpio
        FX.Cls
        'Dibujo
        Call DrawGrhtoHdc(FX.hDC, GrhData(CurrentFX).Frames(currentFrame), 0, 0, False)
        'Y actualizo
        FX.Refresh
        End If
    End If
End Sub

Private Sub VisorHead_Click()
Dim i As Integer

With HeadData(VisorHead.ListIndex + 1)
    For i = 1 To 4
        'Limpio
        frmExtra.Head(i).Cls
        'Dibujo
        Call DrawGrhtoHdc(frmExtra.Head(i).hDC, .Head(i).GrhIndex, 10, 5, False)
        'Y actualizo
        frmExtra.Head(i).Refresh
    Next i
    
    VisorHeadInformacion.Clear
    VisorHeadInformacion.AddItem "[Head" & VisorHead.ListIndex + 1 & "]"
    VisorHeadInformacion.AddItem ""
    VisorHeadInformacion.AddItem "Head1=" & .Head(1).GrhIndex
    VisorHeadInformacion.AddItem "Head2=" & .Head(2).GrhIndex
    VisorHeadInformacion.AddItem "Head3=" & .Head(3).GrhIndex
    VisorHeadInformacion.AddItem "Head4=" & .Head(4).GrhIndex
    
End With
End Sub

Private Sub VisorCasco_Click()
Dim i As Integer

With CascoAnimData(VisorCasco.ListIndex + 1)
    For i = 1 To 4
        'Limpio
        frmExtra.Casco(i).Cls
        'Dibujo
        Call DrawGrhtoHdc(frmExtra.Casco(i).hDC, .Head(i).GrhIndex, 10, 5, False)
        'Y actualizo
        frmExtra.Casco(i).Refresh
    Next i
    
    VisorCascoInformacion.Clear
    VisorCascoInformacion.AddItem "[Casco" & VisorCasco.ListIndex + 1 & "]"
    VisorCascoInformacion.AddItem ""
    VisorCascoInformacion.AddItem "Casco1=" & .Head(1).GrhIndex
    VisorCascoInformacion.AddItem "Casco2=" & .Head(2).GrhIndex
    VisorCascoInformacion.AddItem "Casco3=" & .Head(3).GrhIndex
    VisorCascoInformacion.AddItem "Casco4=" & .Head(4).GrhIndex
    
End With
End Sub

Private Sub VisorCuerpos_Click()
Dim i As Integer

With BodyData(VisorCuerpos.ListIndex + 1)
    For i = 1 To 4
        'Limpio
        frmExtra.Cuerpos(i).Cls
        'Dibujo
        Call DrawGrhtoHdc(frmExtra.Cuerpos(i).hDC, .Walk(i).GrhIndex, 10, 5, False)
        'Y actualizo
        frmExtra.Cuerpos(i).Refresh
    Next i
    
    VisorCuerposInformacion.Clear
    VisorCuerposInformacion.AddItem "[Body" & VisorCuerpos.ListIndex + 1 & "]"
    VisorCuerposInformacion.AddItem ""
    VisorCuerposInformacion.AddItem "Body1=" & .Walk(1).GrhIndex
    VisorCuerposInformacion.AddItem "Body2=" & .Walk(2).GrhIndex
    VisorCuerposInformacion.AddItem "Body3=" & .Walk(3).GrhIndex
    VisorCuerposInformacion.AddItem "Body4=" & .Walk(4).GrhIndex
    
End With
End Sub

Private Sub VisorArmas_Click()
Dim i As Integer

With WeaponAnimData(VisorArmas.ListIndex + 1)
    For i = 1 To 4
        'Limpio
        frmExtra.Armas(i).Cls
        'Dibujo
        Call DrawGrhtoHdc(frmExtra.Armas(i).hDC, .WeaponWalk(i).GrhIndex, 10, 5, False)
        'Y actualizo
        frmExtra.Armas(i).Refresh
    Next i
    
    VisorArmasInformacion.Clear
    VisorArmasInformacion.AddItem "[Weapon" & VisorArmas.ListIndex + 1 & "]"
    VisorArmasInformacion.AddItem ""
    VisorArmasInformacion.AddItem "WeaponWalk1=" & .WeaponWalk(1).GrhIndex
    VisorArmasInformacion.AddItem "WeaponWalk2=" & .WeaponWalk(2).GrhIndex
    VisorArmasInformacion.AddItem "WeaponWalk3=" & .WeaponWalk(3).GrhIndex
    VisorArmasInformacion.AddItem "WeaponWalk4=" & .WeaponWalk(4).GrhIndex
    
End With
End Sub
Private Sub VisorEscudos_Click()
Dim i As Integer

With ShieldAnimData(VisorEscudos.ListIndex + 1)
    For i = 1 To 4
        'Limpio
        frmExtra.Escudos(i).Cls
        'Dibujo
        Call DrawGrhtoHdc(frmExtra.Escudos(i).hDC, .ShieldWalk(i).GrhIndex, 10, 5, False)
        'Y actualizo
        frmExtra.Escudos(i).Refresh
    Next i
    
    VisorEscudosInformacion.Clear
    VisorEscudosInformacion.AddItem "[Shield" & VisorEscudos.ListIndex + 1 & "]"
    VisorEscudosInformacion.AddItem ""
    VisorEscudosInformacion.AddItem "ShieldWalk1=" & .ShieldWalk(1).GrhIndex
    VisorEscudosInformacion.AddItem "ShieldWalk2=" & .ShieldWalk(2).GrhIndex
    VisorEscudosInformacion.AddItem "ShieldWalk3=" & .ShieldWalk(3).GrhIndex
    VisorEscudosInformacion.AddItem "ShieldWalk4=" & .ShieldWalk(4).GrhIndex
    
End With
End Sub

Private Sub VisorFX_Click()
Dim i As Integer

With FxData(VisorFX.ListIndex + 1)

    CurrentFX = .Animacion
    AnimFX.Interval = 100
    AnimFX.Enabled = True

    VisorFXInformacion.Clear
    VisorFXInformacion.AddItem "[FX" & VisorFX.ListIndex + 1 & "]"
    VisorFXInformacion.AddItem ""
    VisorFXInformacion.AddItem "Animacion=" & .Animacion
    VisorFXInformacion.AddItem "OffsetX=" & .OffsetX
    VisorFXInformacion.AddItem "OffsetY=" & .OffsetY
    
End With
End Sub
