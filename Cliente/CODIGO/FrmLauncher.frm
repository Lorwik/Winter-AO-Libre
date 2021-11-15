VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Begin VB.Form FrmLauncher 
   BorderStyle     =   0  'None
   Caption         =   "NoctumAO Launcher"
   ClientHeight    =   6885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
   Icon            =   "FrmLauncher.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "FrmLauncher.frx":0ECA
   MousePointer    =   99  'Custom
   Picture         =   "FrmLauncher.frx":2314
   ScaleHeight     =   6885
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7080
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   735
      Left            =   4320
      TabIndex        =   4
      Top             =   1680
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1296
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      TextRTF         =   $"FrmLauncher.frx":20D9B
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   205
      Left            =   400
      MouseIcon       =   "FrmLauncher.frx":20E1D
      Picture         =   "FrmLauncher.frx":22267
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   4850
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.PictureBox mp3 
      BorderStyle     =   0  'None
      Height          =   205
      Left            =   400
      MouseIcon       =   "FrmLauncher.frx":224DB
      Picture         =   "FrmLauncher.frx":23925
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   5840
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.PictureBox efecto 
      BorderStyle     =   0  'None
      Height          =   205
      Left            =   400
      MouseIcon       =   "FrmLauncher.frx":23B99
      Picture         =   "FrmLauncher.frx":24FE3
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   5518
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.PictureBox ambiental 
      BorderStyle     =   0  'None
      Height          =   205
      Left            =   400
      MouseIcon       =   "FrmLauncher.frx":25257
      Picture         =   "FrmLauncher.frx":266A1
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   5185
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   360
      MouseIcon       =   "FrmLauncher.frx":26915
      Top             =   4800
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   7200
      MouseIcon       =   "FrmLauncher.frx":27D5F
      Top             =   840
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   7440
      MouseIcon       =   "FrmLauncher.frx":291A9
      Top             =   840
      Width           =   255
   End
   Begin VB.Image manual 
      Height          =   405
      Left            =   970
      MouseIcon       =   "FrmLauncher.frx":2A5F3
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   4130
      Width           =   2235
   End
   Begin VB.Image web 
      Height          =   405
      Left            =   960
      MouseIcon       =   "FrmLauncher.frx":2BA3D
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   3440
      Width           =   2235
   End
   Begin VB.Image jugar 
      Height          =   405
      Left            =   960
      MouseIcon       =   "FrmLauncher.frx":2CE87
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   2115
      Width           =   2235
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   360
      MouseIcon       =   "FrmLauncher.frx":2E2D1
      Top             =   5520
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   360
      MouseIcon       =   "FrmLauncher.frx":2F71B
      Top             =   5160
      Width           =   255
   End
   Begin VB.Image Mp33 
      Height          =   255
      Left            =   360
      MouseIcon       =   "FrmLauncher.frx":30B65
      Top             =   5800
      Width           =   255
   End
End
Attribute VB_Name = "FrmLauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub ambiental_Click()
ambiental.Visible = False
Musica = False
'modSound.Music_Stop
Call WriteVar(App.Path & "\init\config.ini", "Init", "Music", "0")
End Sub
Private Sub efecto_Click()
 
Call WriteVar(App.Path & "\init\config.ini", "Init", "Sound", "0")
efecto.Visible = False
End Sub
Private Sub Form_Load()
On Error Resume Next

               If mpp33 = True Then
Windows_Temp_Dir = General_Get_Temp_Dir
 Set MP3P = New clsMP3Player
    Call Extract_File2(mp3, App.Path & "\ARCHIVOS\", "1.mp3", Windows_Temp_Dir, False)
    MP3P.mp3file = Windows_Temp_Dir & "2.mp3"
    MP3P.stopMP3
    MP3P.playMP3
    MP3P.Volume = 1000
    End If

 'Call CheckUpdates
'CARGA DE OPCIONES

'Resolución
If GetVar(App.Path & "\Init\config.ini", "INIT", "Res") = 0 Then
Picture1.Visible = True
reso = True
End If

'Musica ambiental
If GetVar(App.Path & "\init\config.ini", "Init", "Music") = 1 Then
ambiental.Visible = True
Musica = True
'funcion
End If

'Sonido Wav
If GetVar(App.Path & "\init\config.ini", "Init", "Sound") = 1 Then
efecto.Visible = True
'funcion
End If

'Sonido MP3
If GetVar(App.Path & "\init\config.ini", "Init", "MP3") = 1 Then
mp3.Visible = True
mpp33 = True
'funcion
End If

jugar.Picture = General_Load_Picture_From_Resource("jugaron.gif")
web.Picture = General_Load_Picture_From_Resource("webon.gif")
manual.Picture = General_Load_Picture_From_Resource("manualon.gif")
Me.Picture = General_Load_Picture_From_Resource("Launcher.gif")

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If jugar.Tag = 0 Then
    jugar.Picture = General_Load_Picture_From_Resource("jugaron.gif")
    jugar.Tag = 1
End If
If web.Tag = 0 Then
web.Picture = General_Load_Picture_From_Resource("webon.gif")
    web.Tag = 1
End If
If manual.Tag = 0 Then
manual.Picture = General_Load_Picture_From_Resource("manualon.gif")
    manual.Tag = 1
End If
End Sub
Private Sub Image1_Click()
 
Call WriteVar(App.Path & "\init\config.ini", "Init", "Music", "1")
ambiental.Visible = True
Musica = True
End Sub

Private Sub Image2_Click()
 
Call WriteVar(App.Path & "\init\config.ini", "Init", "Sound", "1")
efecto.Visible = True
End Sub

Private Sub Image3_Click()
MP3P.stopMP3
Delete_File (Windows_Temp_Dir & "1.mp3")
End
End Sub

Private Sub Image4_Click()
Me.WindowState = vbMinimized
End Sub


Private Sub Image5_Click()
 
Call WriteVar(App.Path & "\Init\config.ini", "INIT", "Res", "0")
Picture1.Visible = True
reso = True
End Sub

Private Sub mp3_Click()
 
Call WriteVar(App.Path & "\init\config.ini", "Init", "MP3", "0")
mp3.Visible = False
mpp33 = False
End Sub

Private Sub Mp33_Click()
 
Call WriteVar(App.Path & "\init\config.ini", "Init", "MP3", "1")
mp3.Visible = True
mpp33 = True
End Sub
Private Sub jugar_click()
If EnProceso Then Exit Sub
Call addConsole("Conectando...", 0, 200, 0, False, False) '>> Informacion
EnProceso = True
Analizar 'Iniciamos la función Analizar =).
 
Call Shell(App.Path & "\Winter AO Return.EXE", vbNormalFocus)
End Sub

Private Sub Picture1_Click()
Call WriteVar(App.Path & "\Init\config.ini", "INIT", "Res", "1")
Picture1.Visible = False
reso = False
End Sub

Private Sub web_click()
Call ShellExecute(0, "Open", "http://noctumao.com.ar/", "", App.Path, 0)
End Sub
Private Sub manual_click()
Call ShellExecute(0, "Open", "http://noctumao.com.ar/wiki/", "", App.Path, 0)
End Sub
Private Sub Jugar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If jugar.Tag = 1 Then
                jugar.Picture = General_Load_Picture_From_Resource("jugar.gif")
                jugar.Tag = 0
        End If
                End Sub
             Private Sub web_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If web.Tag = 1 Then
                web.Picture = General_Load_Picture_From_Resource("web.gif")
                web.Tag = 0
        End If
End Sub

             Private Sub manual_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If manual.Tag = 1 Then
                manual.Picture = General_Load_Picture_From_Resource("manual.gif")
                manual.Tag = 0
        End If
End Sub
'>> Funciones/Subs
        Function Analizar()
            On Error Resume Next
           
            Dim iX As Integer
            Dim tX As Integer
            Dim DifX As Integer
            Dim strsX As String
           
'LINK1            'Variable que contiene el numero de actualización correcto del servidor
                iX = Inet1.OpenURL("http://winter-ao.com.ar/update/VEREXE.txt")
            'Variable que contiene el numero de actualización del cliente
                tX = GetVar(App.Path & "\INIT\Update.ini", "INIT", "X")
            'Variable con la diferencia de actualizaciones servidor-cliente
                DifX = iX - tX
           
            If Not (DifX = 0) Then 'Si la diferencia no es nula,
            Call addConsole("Iniciando, se descargarán " & DifX & " actualizaciones.", 200, 200, 200, True, False)   '>> Informacion
                For i = 1 To DifX 'Descargamos todas las versiones de diferencia
'LINK2
                    strURL = "http://winter-ao.com.ar/update/Parche" & CStr(i + tX) & ".zip" 'URL del parche .zip
                    Darchivo = App.Path & "\INIT\Parche" & i + tX & ".zip" 'Directorio del parche
                        Call addConsole("   Descargando parche nº " & i, 0, 0, 255, False, True)    '>> Informacion
                    Call AutoDownload(i + tX) 'Descargamos todas las versiones faltantes a partir de la nuestra
                        Call addConsole("   Parche nº " & i & " descargado satisfactoriamente.", 0, 0, 255, False, True)    '>> Informacion
               
                  Call addConsole(" Actualizaciones: " & i & "/" & DifX, 100, 100, 100, True, False)   '>> Informacion
                Next i
            Else
                Call addConsole("No hay actualizaciones pendientes", 200, 200, 200, True, False)    '>> Informacion
            End If
           
           
            Call WriteVar(App.Path & "\INIT\Update.ini", "INIT", "X", CStr(iX)) 'Avisamos al cliente que está actualizado
           
            EnProceso = False
           
            Call addConsole("El cliente ya está listo para jugar", 200, 200, 200, True, False)  '>> Informacion
     
           
            Me.Visible = False
Call Shell(App.Path & "\Winter AO Return.EXE", vbNormalFocus)
            End
        End Function
       
        Public Sub AutoDownload(Numero As Integer)
            On Error Resume Next
           
     
           
            Inet1.AccessType = icUseDefault
            Dim B() As Byte
           
           
            B() = Inet1.OpenURL(strURL, icByteArray)
           
            'Descargamos y guardamos el archivo
            Open Darchivo For Binary Access _
            Write As #1
            Put #1, , B()
            Close #1
           
            'Informacion
            Call addConsole("   Instalando actualización.", 0, 100, 255, False, False)    '>> Informacion
           
           
           
            'Unzipeamos
            UnZip Darchivo, App.Path & "\"
           
            'Borramos el zip
            Kill Darchivo
        End Sub
'<< End

