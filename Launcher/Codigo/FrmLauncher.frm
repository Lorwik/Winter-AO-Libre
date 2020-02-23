VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmLauncher 
   BorderStyle     =   0  'None
   Caption         =   "NoctumAO Launcher"
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6840
   Icon            =   "FrmLauncher.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   Picture         =   "FrmLauncher.frx":3AFA
   ScaleHeight     =   225
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   456
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1200
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   825
      Left            =   2850
      TabIndex        =   0
      Top             =   2040
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   1455
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"FrmLauncher.frx":D74A
   End
   Begin VB.Image Progres 
      Height          =   285
      Left            =   3675
      Picture         =   "FrmLauncher.frx":D7CC
      Top             =   2895
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Image CMDAutoUpdate 
      Height          =   1140
      Left            =   1785
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   615
      Width           =   1620
   End
   Begin VB.Image CMDResolucion 
      Height          =   195
      Left            =   2760
      Top             =   1800
      Width           =   285
   End
   Begin VB.Image CMDCerrar 
      Height          =   255
      Left            =   6240
      Top             =   360
      Width           =   255
   End
   Begin VB.Image CMDManual 
      Height          =   1140
      Left            =   4905
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   615
      Width           =   1620
   End
   Begin VB.Image CMDWeb 
      Height          =   1140
      Left            =   3315
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   615
      Width           =   1620
   End
   Begin VB.Image CMDJugar 
      Height          =   1140
      Left            =   210
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   615
      Width           =   1620
   End
   Begin VB.Image CMDEffectSound 
      Height          =   195
      Left            =   4125
      Top             =   1800
      Width           =   285
   End
   Begin VB.Image CMDAmbiental 
      Height          =   195
      Left            =   5010
      Top             =   1800
      Width           =   285
   End
End
Attribute VB_Name = "FrmLauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Const LW_KEY = &H1
Const G_E = (-20)
Const W_E = &H80000

Private Sub CMDAmbiental_Click()
    If GetVar(App.Path & "\INIT\Config.cfg", "Sound", "Wav") = 0 Then
        Call WriteVar(App.Path & "\INIT\Config.cfg", "Sound", "Wav", 1)
        CMDAmbiental.Picture = General_Load_Picture_From_Resource("101.gif")
    Else
        Call WriteVar(App.Path & "\INIT\Config.cfg", "Sound", "Wav", 0)
        CMDAmbiental.Picture = LoadPicture("")
    End If
End Sub

Private Sub CMDEffectSound_Click()
    If GetVar(App.Path & "\INIT\Config.cfg", "Sound", "MP3") = 0 Then
        Call WriteVar(App.Path & "\INIT\Config.cfg", "Sound", "MP3", 1)
        CMDEffectSound.Picture = General_Load_Picture_From_Resource("101.gif")
    Else
        Call WriteVar(App.Path & "\INIT\Config.cfg", "Sound", "MP3", 0)
        CMDEffectSound.Picture = LoadPicture("")
    End If
End Sub

Private Sub CMDResolucion_Click()
    If GetVar(App.Path & "\INIT\Config.cfg", "Video", "Res") = 0 Then
        Call WriteVar(App.Path & "\INIT\Config.cfg", "Video", "Res", 1)
        CMDResolucion.Picture = General_Load_Picture_From_Resource("101.gif")
    Else
        Call WriteVar(App.Path & "\INIT\Config.cfg", "Video", "Res", 0)
        CMDResolucion.Picture = LoadPicture("")
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Call addConsole("Esperando...", 0, 200, 0, False, False)
    Me.Picture = General_Load_Picture_From_Resource("95.gif")
    Progres.Picture = General_Load_Picture_From_Resource("96.gif")

    If GetVar(App.Path & "\INIT\Config.cfg", "Video", "Res") = 1 Then
        CMDResolucion.Picture = General_Load_Picture_From_Resource("101.gif")
    Else
        CMDResolucion.Picture = LoadPicture("")
    End If

    If GetVar(App.Path & "\INIT\Config.cfg", "Sound", "MP3") = 1 Then
        CMDEffectSound.Picture = General_Load_Picture_From_Resource("101.gif")
    Else
        CMDEffectSound.Picture = LoadPicture("")
    End If
    
    If GetVar(App.Path & "\INIT\Config.cfg", "Sound", "Wav") = 1 Then
        CMDAmbiental.Picture = General_Load_Picture_From_Resource("101.gif")
    Else
        CMDAmbiental.Picture = LoadPicture("")
    End If

    Skin Me, vbRed
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then Call Auto_Drag(Me.hwnd)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CMDJugar.Picture = LoadPicture("")
    CMDWeb.Picture = LoadPicture("")
    CMDManual.Picture = LoadPicture("")
    CMDAutoUpdate.Picture = LoadPicture("")
End Sub

Private Sub Picture1_Click()
    Call WriteVar(App.Path & "\Init\config.cfg", "INIT", "Res", "1")
    CMDResolucion.Visible = False
    reso = False
End Sub

Private Sub Timer1_Timer()

    'Avance de la barra de descarga
    If Progres.Width = 182 Then
        Timer1.Enabled = False
    Else
        Progres.Width = Progres.Width + 5
    End If
    'Psb1.Text = CLng(Psb1.Percent) & "%"
End Sub

'>> Funciones/Subs
Function Search()
On Error Resume Next
           
    Dim iX As Integer
    Dim tX As Integer
    Dim DifX As Integer
    Dim strsX As String
           
'LINK1
    'Variable que contiene el numero de actualización correcto del servidor
    iX = Inet1.OpenURL("http://aowinter.com.ar/update/version.txt")
    'Variable que contiene el numero de actualización del cliente
    tX = GetVar(App.Path & "\INIT\config.cfg", "UPDATE", "X")
    'Variable con la diferencia de actualizaciones servidor-cliente
    DifX = iX - tX
           
    If Not (DifX = 0) Then 'Si la diferencia no es nula,
        Call addConsole("Iniciando, se descargarán " & DifX & " actualizaciones.", 200, 200, 200, True, False)   '>> Informacion
        For i = 1 To DifX 'Descargamos todas las versiones de diferencia
'LINK2
            strURL = "http://aowinter.com.ar/update/Parche" & CStr(i + tX) & ".zip" 'URL del parche .zip
            Darchivo = App.Path & "\INIT\Parche" & i + tX & ".zip" 'Directorio del parche
            Call addConsole("   Descargando parche nº " & i, 0, 0, 255, False, True)    '>> Informacion
            Call AutoDownload(i + tX) 'Descargamos todas las versiones faltantes a partir de la nuestra
            Call addConsole("   Parche nº " & i & " descargado satisfactoriamente.", 0, 0, 255, False, True)    '>> Informacion
               
            Call addConsole(" Actualizaciones: " & i & "/" & DifX, 100, 100, 100, True, False)   '>> Informacion
        Next i
    Else
        Call addConsole("No hay actualizaciones pendientes", 200, 200, 200, True, False)    '>> Informacion
    End If
           
    Call WriteVar(App.Path & "\INIT\config.cfg", "UPDATE", "X", CStr(iX)) 'Avisamos al cliente que está actualizado
           
    EnProceso = False
           
    Call addConsole("El cliente ya está listo para jugar", 200, 200, 200, True, False)  '>> Informacion
     

End Function

Function Analizar()
    Dim iX As Integer
    Dim tX As Integer
    Dim DifX As Integer
    Dim strsX As String
           
'LINK1
    'Variable que contiene el numero de actualización correcto del servidor
    iX = Inet1.OpenURL("http://aowinter.com.ar/update/version.txt")
    'Variable que contiene el numero de actualización del cliente
    tX = GetVar(App.Path & "\INIT\config.cfg", "UPDATE", "X")
    'Variable con la diferencia de actualizaciones servidor-cliente
    DifX = iX - tX
           
    If Not (DifX = 0) Then 'Si la diferencia no es nula,
        Call addConsole("Iniciando, se descargarán " & DifX & " actualizaciones.", 200, 200, 200, True, False)   '>> Informacion
        For i = 1 To DifX 'Descargamos todas las versiones de diferencia
'LINK2
            strURL = "http://aowinter.com.ar/update/Parche" & CStr(i + tX) & ".zip" 'URL del parche .zip
            Darchivo = App.Path & "\INIT\Parche" & i + tX & ".zip" 'Directorio del parche
            Call addConsole("   Descargando parche nº " & i, 0, 0, 255, False, True)    '>> Informacion
            Call AutoDownload(i + tX) 'Descargamos todas las versiones faltantes a partir de la nuestra
            Call addConsole("   Parche nº " & i & " descargado satisfactoriamente.", 0, 0, 255, False, True)    '>> Informacion
               
            Call addConsole(" Actualizaciones: " & i & "/" & DifX, 100, 100, 100, True, False)   '>> Informacion
        Next i
    Else
        Call addConsole("No hay actualizaciones pendientes", 200, 200, 200, True, False)    '>> Informacion
    End If
           
    Call WriteVar(App.Path & "\INIT\config.cfg", "UPDATE", "X", CStr(iX)) 'Avisamos al cliente que está actualizado
           
    EnProceso = False
           
    Call addConsole("El cliente ya está listo para jugar", 200, 200, 200, True, False)  '>> Informacion
           
    Me.Visible = False
    If General_File_Exists(App.Path & "\Winter AO Ultimate.EXE", vbNormal) Then
        Call WriteVar(App.Path & "\init\Config.CFG", "UPDATE", "Y", "1")
        DoEvents
        Call Shell(App.Path & "\Winter AO Ultimate.EXE", vbNormalFocus)
        End
    Else
        MsgBox "No se encontro el ejecutable del juego ""0Winter AO Ultimate.EXE""."
        End
    End If
    
End Function
       
Public Sub AutoDownload(Numero As Integer)
On Error Resume Next
           
    Inet1.AccessType = icUseDefault
    Dim B() As Byte
           
    'Informacion...
    Progres.Width = 0
    Progres.Visible = True
    Timer1.Enabled = True
           
    B() = Inet1.OpenURL(strURL, icByteArray)
           
    'Descargamos y guardamos el archivo
    Open Darchivo For Binary Access _
    Write As #1
    Put #1, , B()
    Close #1
    Timer1.Enabled = False
    Progres.Width = 100
    Timer1.Enabled = False
    'Informacion
    Call addConsole("   Instalando actualización.", 0, 100, 255, False, False)    '>> Informacion
           
    'Unzipeamos
    UnZip Darchivo, App.Path & "\"
    
    DoEvents
    
    If General_File_Exists(App.Path & "\Recursos\tmp.WAO", vbNormal) Then
    '    'Instalamos el Parche
        Extract_Patch App.Path & "\Recursos", App.Path & "\Recursos\tmp.WAO"
    '
    '    'Esperamos a que termine
        DoEvents
    '
    '    'Borramos el Parche
        Kill App.Path & "\Recursos\tmp.WAO"
     End If
           
    'Borramos el zip
    Kill Darchivo
End Sub
'<< End

'*********************CONTROLES*************************
Private Sub CMDWeb_click()
Call ShellExecute(0, "Open", "http://aowinter.com.ar/", "", App.Path, 0)
End Sub

Private Sub CMDManual_click()
Call ShellExecute(0, "Open", "http://aowinter.com.ar/manual/", "", App.Path, 0)
End Sub

Private Sub CMDAutoUpdate_Click()
    If EnProceso Then Exit Sub
    Call addConsole("Buscando actualizaciones. Espere porfavor...", 0, 200, 0, False, False) '>> Informacion
    EnProceso = True
    Search 'Iniciamos la función Analizar =).
End Sub

Private Sub CMDJugar_click()
    
    If EnProceso Then Exit Sub
    
    Call addConsole("Buscando actualizaciones. Espere porfavor...", 0, 200, 0, False, False) '>> Informacion
    EnProceso = True
    Analizar 'Iniciamos la función Analizar =).
End Sub

Private Sub CMDCerrar_Click()
    End
End Sub

Private Sub CMDJugar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CMDJugar.Picture = General_Load_Picture_From_Resource("97.gif")
End Sub
Private Sub CMDAutoUpdate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CMDAutoUpdate.Picture = General_Load_Picture_From_Resource("98.gif")
End Sub
Private Sub CMDWeb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CMDWeb.Picture = General_Load_Picture_From_Resource("99.gif")
End Sub
Private Sub CMDManual_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CMDManual.Picture = General_Load_Picture_From_Resource("100.gif")
End Sub
'******************FIN CONTROLES************************

Sub Skin(Frm As Form, Color As Long)
Frm.BackColor = Color
Dim Ret As Long
Ret = GetWindowLong(Frm.hwnd, G_E)
Ret = Ret Or W_E
SetWindowLong Frm.hwnd, G_E, Ret
SetLayeredWindowAttributes Frm.hwnd, Color, 0, LW_KEY
End Sub
