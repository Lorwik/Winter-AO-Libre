VERSION 5.00
Begin VB.Form frmQuests 
   BorderStyle     =   0  'None
   Caption         =   "Misiones"
   ClientHeight    =   4920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   328
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstQuests 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3870
      Left            =   405
      TabIndex        =   0
      Top             =   510
      Width           =   3435
   End
   Begin VB.Label QuestNum 
      BackStyle       =   0  'Transparent
      Caption         =   "0/15"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   4470
      Width           =   495
   End
   Begin VB.Image cmdOptions 
      Height          =   255
      Index           =   1
      Left            =   5880
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Image cmdOptions 
      Height          =   375
      Index           =   0
      Left            =   4320
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label txtInfo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3495
      Left            =   4080
      TabIndex        =   1
      Top             =   720
      Width           =   3135
   End
End
Attribute VB_Name = "frmQuests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOptions_Click(Index As Integer)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Maneja el click de los CommandButtons cmdOptions.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Select Case Index
        Case 0 'Botón ABANDONAR MISIÓN
            'Chequeamos si hay items.
            If lstQuests.ListCount = 0 Then
                MsgBox "¡No tienes ninguna misión!", vbOKOnly + vbExclamation
                Exit Sub
            End If
            
            'Chequeamos si tiene algun item seleccionado.
            If lstQuests.ListIndex < 0 Then
                MsgBox "¡Primero debes seleccionar una misión!", vbOKOnly + vbExclamation
                Exit Sub
            End If
            
            Select Case MsgBox("¿Estás seguro que deseas abandonar la misión?", vbYesNo + vbExclamation)
                Case vbYes  'Botón SÍ.
                    'Enviamos el paquete para abandonar la quest
                    Call WriteQuestAbandon(lstQuests.ListIndex + 1)
                    
                Case vbNo   'Botón NO.
                    'Como seleccionó que no, no hace nada.
                    Exit Sub
            End Select
            
        Case 1 'Botón VOLVER
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("70.gif")
Call General_Set_Wav("256.wav")
End Sub

Private Sub lstQuests_Click()
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Maneja el click del ListBox lstQuests.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If lstQuests.ListIndex < 0 Then Exit Sub
    
    Call WriteQuestDetailsRequest(lstQuests.ListIndex + 1)
End Sub
