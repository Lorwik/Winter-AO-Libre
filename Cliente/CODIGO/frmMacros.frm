VERSION 5.00
Begin VB.Form frmMacros 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurar Macros Internos"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Tecla 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      ItemData        =   "frmMacros.frx":0000
      Left            =   0
      List            =   "frmMacros.frx":0013
      TabIndex        =   0
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
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
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   3600
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar Tecla"
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
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   3015
      Begin VB.OptionButton Option1 
         Caption         =   "Lanzar Hechizo Seleccionado."
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
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   2870
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Equipar Item  Seleccionado."
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
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   2775
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Usar Item  Seleccionado."
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
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Enviar Comando:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox SendCommand 
         Height          =   285
         Left            =   600
         TabIndex        =   3
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "/"
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
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   255
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMacros.frx":002B
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione una Tecla para configurar."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   1920
      Width           =   3015
   End
End
Attribute VB_Name = "frmMacros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0 'Guardar Macros
    If Option1(0).value = True Then
        If SendCommand.Text = "" Then
            MsgBox "Debes Ingresar el Comando", vbCritical
            Exit Sub
        Else
            Call WriteVar(IniPath & "Macros.bin", Frame1.Caption, "Comando", SendCommand.Text)
            Call WriteVar(IniPath & "Macros.bin", Frame1.Caption, "Accion", "1")
        End If
    ElseIf Option1(1).value = True Then
        Call WriteVar(IniPath & "Macros.Bin", Frame1.Caption, "UsarItem", Inventario.SelectedItem)
        Call WriteVar(IniPath & "Macros.bin", Frame1.Caption, "Accion", "2")
    ElseIf Option1(2).value = True Then
        Call WriteVar(IniPath & "Macros.Bin", Frame1.Caption, "EquiparItem", Inventario.SelectedItem)
        Call WriteVar(IniPath & "Macros.bin", Frame1.Caption, "Accion", "3")
    ElseIf Option1(3).value = True Then
        Call WriteVar(IniPath & "Macros.Bin", Frame1.Caption, "LanzarHechizo", frmMain.hlst.listIndex + 1)
        Call WriteVar(IniPath & "Macros.bin", Frame1.Caption, "Accion", "4")
    End If
    
Case 1 'No Guardamos Nada
    Unload Me
    
End Select
End Sub
Private Sub Tecla_Click()
Frame1.Visible = True
Select Case (Tecla.List(Tecla.listIndex))

    Case Is = "F1"
        Frame1.Caption = "F1"
        Call CargarTecla
    Case Is = "F2"
        Frame1.Caption = "F2"
        Call CargarTecla
    Case Is = "F3"
        Frame1.Caption = "F3"
        Call CargarTecla
    Case Is = "F4"
        Frame1.Caption = "F4"
        Call CargarTecla
    Case Is = "F5"
        Frame1.Caption = "F5"
        Call CargarTecla

End Select

End Sub
Private Sub CargarTecla()
On Error Resume Next
Dim Accion As Byte
    Accion = GetVar(IniPath & "Macros.bin", Frame1.Caption, "Accion")
    
    If Accion = 1 Then
        Dim Comando As String
        Comando = GetVar(IniPath & "Macros.bin", Frame1.Caption, "Comando")
        Option1(0).value = True
        SendCommand.Text = Comando
    ElseIf Accion = 2 Then
        SendCommand.Text = ""
        Option1(1).value = True
    ElseIf Accion = 3 Then
        SendCommand.Text = ""
        Option1(2).value = True
    ElseIf Accion = 4 Then
        SendCommand.Text = ""
        Option1(3).value = True
    ElseIf Accion <> 1 Or 2 Or 3 Or 4 Then
        Exit Sub
    ElseIf Accion = "" Then
        Exit Sub
    End If
    
End Sub
