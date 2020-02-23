VERSION 5.00
Begin VB.Form frmMusica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Musica"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   Icon            =   "frmMusica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5595
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Música"
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.ListBox LstPlay 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1620
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2640
      End
      Begin VB.Timer TmrCheck 
         Interval        =   100
         Left            =   1440
         Top             =   960
      End
      Begin WorldEditor.lvButtons_H cmdCerrar1 
         Height          =   495
         Left            =   2880
         TabIndex        =   1
         Top             =   1440
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         Caption         =   "&Cerrar"
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
      Begin WorldEditor.lvButtons_H cmdAplicarYCerrar1 
         Height          =   495
         Left            =   2880
         TabIndex        =   2
         Top             =   840
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         Caption         =   "&Aplicar y Cerrar"
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
         Enabled         =   0   'False
         cBack           =   12648447
      End
      Begin WorldEditor.lvButtons_H CmdStop 
         Height          =   495
         Left            =   4080
         TabIndex        =   4
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "&Detener"
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
         Enabled         =   0   'False
         cBack           =   12632319
      End
      Begin WorldEditor.lvButtons_H CmdPlay 
         Height          =   495
         Left            =   2880
         TabIndex        =   5
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "&Escuchar"
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
         Enabled         =   0   'False
         cBack           =   12648384
      End
   End
End
Attribute VB_Name = "frmMusica"
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

Dim mFile As String
Public FileSize As Long
Dim mPos As Long
Dim iTo As Long
Dim Paused As Boolean
Dim CPos As Long


Private Enum State
    cInvalid = -1
    cPlaying = 0
    cStopped = 1
    cPaused = 2
End Enum

Dim cPic As New StdPicture

Dim CState As State
Private MidiActual As String
Private Mp3Actual As String

''
' Aplica la Musica seleccionada y oculta la ventana
'

Private Sub cmdAplicarYCerrar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
If Len(MidiActual) >= 5 Then
    MapInfo.Music = Left(MidiActual, Len(MidiActual) - 4)
    frmMapInfo.txtMapMusica.text = MapInfo.Music
    frmMain.lblMapMusica = MapInfo.Music
    MidiActual = Empty
End If
Me.Hide
End Sub


Private Sub cmdAplicarYCerrar1_Click()
On Error Resume Next

If Len(Mp3Actual) >= 5 Then
    MapInfo.Music = Left(Mp3Actual, Len(Mp3Actual) - 4)
    frmMapInfo.txtMapMusica.text = MapInfo.Music
    frmMain.lblMapMusica = MapInfo.Music
    Mp3Actual = Empty
End If
    mciSendString "stop music", 0, 0, 0
Me.Hide
End Sub

''
' Oculta la ventana
'

Private Sub cmdCerrar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Me.Hide
End Sub

Private Sub cmdCerrar1_Click()
Me.Hide
    mciSendString "stop music", 0, 0, 0
End Sub

Private Sub CmdPause_Click()
    mciSendString "stop music", 0, 0, 0
End Sub

Private Sub CmdPlay_Click()
If LstPlay.ListIndex = -1 And LstPlay.ListCount <> 0 Then LstPlay.ListIndex = 0

Paused = False

If LstPlay.ListCount = 0 Then Exit Sub
Me.MousePointer = 11

'The Current file
Dim cFile As String

'Get the file to play
cFile = App.Path & "\" & "MP3\" & LstPlay.List(LstPlay.ListIndex)
mPos = LstPlay.ListIndex

'Play the first item if none are selected
If LstPlay.ListIndex = -1 Then
    LstPlay.ListIndex = 0
    mPos = 0
End If

If mFile = cFile Then GoTo SkipLoad

'TmrCheck.Enabled = True

'Close the file first
mciSendString "stop music", 0, 0, 0
mciSendString "close music", 0, 0, 0

Dim Ret As Long

'"open the File"
Ret = mciSendString("open " & Chr(34) & cFile & Chr(34) & " alias music", 0, 0, 0)

'Error handler
If Not Ret = 0 Then
    MsgBox "Error " & Ret & vbNewLine & GetMciError(Val(Ret))
    Me.MousePointer = 0
    LstPlay.RemoveItem LstPlay.ListIndex
End If

'Set the time format
Call mciSendString("set music time format milliseconds", 0, 0, 0)

'Get the file time
Dim tmptime As String * 15
Call mciSendString("status music length", tmptime, 15, 0)
FileSize = Val(tmptime)


'Get the file size
mFile = cFile
SkipLoad:

' "Rewind" the file
Call mciSendString("seek music to start", 0, 0, 0)

'Dim TmpHnd As String
'TmpHnd = Str(box.hWnd)

'lstNew.ListIndex = LstPlay.ListIndex

Call mciSendString("setvideo music on", 0, 0, 0)

'Call mciSendString("window music handle " & Str(TmpHnd), 0, 0, 0)

Call mciSendString("setvideo music on", 0, 0, 0)

Call mciSendString("put music destination", 0, 0, 0)

Call mciSendString("setvideo music on", 0, 0, 0)

'Be sure to show the video
Call mciSendString("setvideo music on", 0, 0, 0)

'Play the file
mciSendString "play music ", 0, 0, 0

Call mciSendString("setvideo music on", 0, 0, 0)

'Set cPic = box.Picture

Me.MousePointer = 0
CmdStop.Enabled = True
CmdPlay.Enabled = False
End Sub

Private Sub CmdStop_Click()
mciSendString "stop music", 0, 0, 0
End Sub

Private Sub Form_Load()
On Error GoTo ErrHand
 Set FSO = New FileSystemObject
 
 FillListFiles LstPlay, "mp3", App.Path & "\" & "MP3\"
 Exit Sub
 
ErrHand:
If Err.Number = 53 Then Exit Sub
MsgBox "Error & err.Number & vbnewline & err.Description"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set FSO = Nothing
End Sub
Private Sub LstPlay_Click()
'CmdPlay.value = True
CmdPlay.Enabled = True
CmdStop.Enabled = False

Mp3Actual = LstPlay.List(LstPlay.ListIndex)

cmdAplicarYCerrar1.Enabled = True

End Sub
Private Function GetMciError(Errorno As Integer)
    Dim Ret As String * 250
    mciGetErrorString Errorno, Ret, 250

    GetMciError = Ret
End Function
