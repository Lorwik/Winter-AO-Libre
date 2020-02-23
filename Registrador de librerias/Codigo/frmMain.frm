VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registrador de librerias - RincondelAO - By Lorwik"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3705
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
   ScaleHeight     =   157
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   247
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton End 
      Caption         =   "Salir"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton ActualizarDlls 
      Caption         =   "Registrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   2040
      Width           =   1695
   End
   Begin RichTextLib.RichTextBox Consola 
      Height          =   1620
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   240
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2858
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      OLEDropMode     =   0
      TextRTF         =   $"frmMain.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "www.RincondelAO.com.ar - By Lorwik"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActualizarDlls_Click()
General.ActualizarDlls
End Sub

Private Sub End_Click()
End
End Sub

Private Sub Form_Load()
Call AddtoRichTextBox(frmMain.Consola, "Si usas Windows Vista o 7 ejecuta el registrador de librerias como administrador.", 255, 0, 0, True, True, False)
Call AddtoRichTextBox(frmMain.Consola, "Pulse el botón ""Registrar"" para comenzar...", 255, 255, 255, True, False, False)
End Sub
