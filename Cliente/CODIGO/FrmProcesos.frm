VERSION 5.00
Begin VB.Form FrmProcesos 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Procesos"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sacar Foto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   1455
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      ItemData        =   "FrmProcesos.frx":0000
      Left            =   120
      List            =   "FrmProcesos.frx":0002
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Filtrar:"
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
      TabIndex        =   4
      Top             =   3120
      Width           =   735
   End
End
Attribute VB_Name = "FrmProcesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const lb_findstring = &H18F
Private Declare Function SendMessage Lib "user32" Alias "sendmessagea" (ByVal hwnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Private Sub Command1_Click()
Unload Me
End Sub
Sub FindWord(List As ListBox, aSeek As String, HSearch As Integer)
Dim aWord() As String
Dim Element As String
Dim NumberList As Integer
Dim WordFound As Integer
aWord = Split(aSeek, " ")
For NumberList = HSearch To List.ListCount - 1
WordFound = 0
Element = List.List(NumberList)
For nWord = LBound(aWord) To UBound(aWord)
If InStr(1, Element, aWord(nWord), vbTextCompare) Then
Element = Replace(Element, aWord(nWord), vbNullString, 1, 1, vbTextCompare)
WordFound = WordFound + 1
End If
Next
If WordFound = UBound(aWord) + 1 Then
List.listIndex = NumberList
Exit For
Else
List.listIndex = -1
End If
Next
End Sub
Private Sub Command2_Click()
 Dim i As Integer
 For i = 1 To 1000
If Not FileExist(App.Path & "\Fotos\Proc" & i & ".bmp", vbNormal) Then Exit For
  Next
  Call Capturar_Guardar(App.Path & "/Fotos/Proc" & i & ".bmp")
  Call AddtoRichTextBox(frmMain.RecTxt, "Proc" & i & ".bmp Guardada en la Carpeta Fotos. Puedes subirla a http://www.lwk-images.com.ar", 255, 255, 255, False, False, False)
End Sub
Private Sub Text1_Change()
Call FindWord(List1, Text1.Text, 0)
End Sub

