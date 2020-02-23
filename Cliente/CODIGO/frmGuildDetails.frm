VERSION 5.00
Begin VB.Form frmGuildDetails 
   BorderStyle     =   0  'None
   Caption         =   "Detalles del Clan"
   ClientHeight    =   7125
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   6900
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   ScaleHeight     =   7125
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   640
      TabIndex        =   8
      Top             =   3590
      Width           =   5655
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   640
      TabIndex        =   7
      Top             =   3950
      Width           =   5655
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   640
      TabIndex        =   6
      Top             =   4310
      Width           =   5655
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   640
      TabIndex        =   5
      Top             =   4650
      Width           =   5655
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   640
      TabIndex        =   4
      Top             =   5030
      Width           =   5655
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   640
      TabIndex        =   3
      Top             =   5390
      Width           =   5655
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   640
      TabIndex        =   2
      Top             =   5750
      Width           =   5655
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   640
      TabIndex        =   1
      Top             =   6110
      Width           =   5655
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   390
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   690
      Width           =   6135
   End
   Begin VB.Image Command1 
      Height          =   255
      Index           =   1
      Left            =   5160
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Image Command1 
      Height          =   255
      Index           =   0
      Left            =   360
      Top             =   6600
      Width           =   1455
   End
End
Attribute VB_Name = "frmGuildDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_DESC_LENGTH As Integer = 520
Private Const MAX_CODEX_LENGTH As Integer = 100

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
        Call General_Set_Wav(SND_CLICK)
            Unload Me
        
        Case 1
        Call General_Set_Wav(SND_CLICK)
            Dim fdesc As String
            Dim Codex() As String
            Dim k As Byte
            Dim Cont As Byte
    
            fdesc = Replace(txtDesc, vbCrLf, "º", , , vbBinaryCompare)
    
            '    If Not AsciiValidos(fdesc) Then
            '        MsgBox "La descripcion contiene caracteres invalidos"
            '        Exit Sub
            '    End If

            Cont = 0
            For k = 0 To txtCodex1.UBound
            '    If Not AsciiValidos(txtCodex1(k)) Then
            '        MsgBox "El codex tiene invalidos"
            '        Exit Sub
            '    End If
                If LenB(txtCodex1(k).Text) <> 0 Then Cont = Cont + 1
            Next k
            If Cont < 4 Then
                MsgBox "Debes definir al menos cuatro mandamientos."
                Exit Sub
            End If
                        
            ReDim Codex(txtCodex1.UBound) As String
            For k = 0 To txtCodex1.UBound
                Codex(k) = txtCodex1(k)
            Next k
    
            If CreandoClan Then
                Call WriteCreateNewGuild(fdesc, ClanName, Site, Codex)
            Else
                Call WriteClanCodexUpdate(fdesc, Codex)
            End If

            CreandoClan = False
            Unload Me
            
    End Select
End Sub

Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("50.gif")
End Sub

Private Sub txtCodex1_Change(Index As Integer)
    If Len(txtCodex1.Item(Index).Text) > MAX_CODEX_LENGTH Then _
        txtCodex1.Item(Index).Text = Left$(txtCodex1.Item(Index).Text, MAX_CODEX_LENGTH)
End Sub

Private Sub txtDesc_Change()
    If Len(txtDesc.Text) > MAX_DESC_LENGTH Then _
        txtDesc.Text = Left$(txtDesc.Text, MAX_DESC_LENGTH)
End Sub
