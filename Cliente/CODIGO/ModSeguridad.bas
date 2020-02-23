Attribute VB_Name = "ModSeguridad"
Option Explicit
'*******************************************************************************
'EN ESTE MODULO SE PUEDE ENCONTRAR VARIAS FUNCIONES Y SUBS QUE CONTROLAN LA _
SEGURIDAD DEL JUEGO, PARA ASI EVITAR TRAMPAS. ALGUNAS FUNCIONES Y CODIGOS FUERON _
TOMADAS DE ALGUNAS LIBERACIONES Y RECURSOS DE INTERNET. _
By LORWIK
'*******************************************************************************

'**************Detecta externos mediante nombre de ventanas*********************
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'*******************************************************************************
'**************Dectecta Externo mediante palabras claves************************
Declare Function EnumWindows Lib "user32" ( _
                 ByVal wndenmprc As Long, _
                 ByVal lParam As Long) As Long
 
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
                 ByVal hwnd As Long, _
                 ByVal lpString As String, _
                 ByVal cch As Long) As Long
 
 
Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                 ByVal hwnd As Long, _
                 ByVal wMsg As Long, _
                 ByVal wParam As Long, _
                 lParam As Any) As Long
 
 
 
Const WM_SYSCOMMAND = &H112
Const SC_CLOSE = &HF060&
 
 
Private Caption As String
'*******************************************************************************

'*******************Detecta si se cambio el nombre al exe***********************
Public OriginalClientName As String
Public ClientName As String
Public DetectName As String
'*******************************************************************************

'**********************************Anti Debugger********************************
Private Declare Function IsDebuggerPresent Lib "kernel32" () As Long
'*******************************************************************************

'********************************Anti Speed Hack********************************
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Time As Long
Private count As Integer
'*******************************************************************************

'**************Dectecta Externo mediante palabras claves************************
Public Sub Externos(Nombre As String)
    Caption = Nombre
    Call EnumWindows(AddressOf recorrerVentanas, 0)
End Sub
 
 
Public Function recorrerVentanas(ByVal hwnd As Long, _
                ByVal param As Long) As Long
 
Dim Buffer As String * 256
Dim tWindows As String
Dim Size_buffer As Long
 
 
    Size_buffer = GetWindowText(hwnd, Buffer, Len(Buffer))
   
    tWindows = Left$(Buffer, Size_buffer)
   
    'comparamos ventanas
 
    If InStr(UCase(tWindows), UCase(Caption)) <> 0 Then
       
       ' le mandamos el sendmesagge y cerramos
        SendMessage hwnd, WM_SYSCOMMAND, SC_CLOSE, ByVal 0&
    End If
 
   
    recorrerVentanas = 1
End Function
'*******************************************************************************

'**************Detecta externos mediante nombre de ventanas*********************
Public Sub SecuLwK()
'Lorwik> La verdad esque no me gusta mucho, pero mejor esto que nada xD
If FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.1.1")) Then
    Call HayExterno("CHEAT ENGINE 5.1.1")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.0")) Then
    Call HayExterno("CHEAT ENGINE 5.0")
ElseIf FindWindow(vbNullString, UCase$("Pts")) Then
    Call HayExterno("Auto Pots")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.2")) Then
    Call HayExterno("CHEAT ENGINE 5.2")
ElseIf FindWindow(vbNullString, UCase$("SOLOCOVO?")) Then
    Call HayExterno("SOLOCOVO?")
ElseIf FindWindow(vbNullString, UCase$("-=[ANUBYS RADAR]=-")) Then
    Call HayExterno("-=[ANUBYS RADAR]=-")
ElseIf FindWindow(vbNullString, UCase$("CRAZY SPEEDER 1.05")) Then
    Call HayExterno("CRAZY SPEEDER 1.05")
ElseIf FindWindow(vbNullString, UCase$("SET !XSPEED.NET")) Then
    Call HayExterno("SET !XSPEED.NET")
ElseIf FindWindow(vbNullString, UCase$("SPEEDERXP V1.80 - UNREGISTERED")) Then
    Call HayExterno("SPEEDERXP V1.80 - UNREGISTERED")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.3")) Then
    Call HayExterno("CHEAT ENGINE 5.3")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.1")) Then
    Call HayExterno("CHEAT ENGINE 5.1")
ElseIf FindWindow(vbNullString, UCase$("A SPEEDER")) Then
    Call HayExterno("A SPEEDER")
ElseIf FindWindow(vbNullString, UCase$("MEMO :P")) Then
    Call HayExterno("MEMO :P")
ElseIf FindWindow(vbNullString, UCase$("ORK4M VERSION 1.5")) Then
    Call HayExterno("ORK4M VERSION 1.5")
ElseIf FindWindow(vbNullString, UCase$("BY FEDEX")) Then
    Call HayExterno("By Fedex")
ElseIf FindWindow(vbNullString, UCase$("!XSPEED.NET +4.59")) Then
    Call HayExterno("!Xspeeder")
ElseIf FindWindow(vbNullString, UCase$("CAMBIA TITULOS DE CHEATS BY FEDEX")) Then
    Call HayExterno("Cambia titulos")
ElseIf FindWindow(vbNullString, UCase$("NEWENG OCULTO")) Then
    Call HayExterno("Cambia titulos")
ElseIf FindWindow(vbNullString, UCase$("SERBIO ENGINE")) Then
    Call HayExterno("Serbio Engine")
ElseIf FindWindow(vbNullString, UCase$("REYMIX ENGINE 5.3 PUBLIC")) Then
    Call HayExterno("ReyMix Engine")
ElseIf FindWindow(vbNullString, UCase$("REY ENGINE 5.2")) Then
    Call HayExterno("ReyMix Engine")
ElseIf FindWindow(vbNullString, UCase$("AUTOCLICK - BY NIO_SHOOTER")) Then
    Call HayExterno("AutoClick")
ElseIf FindWindow(vbNullString, UCase$("TONNER MINER! :D [REG][SKLOV] 2.0")) Then
    Call HayExterno("Tonner")
ElseIf FindWindow(vbNullString, UCase$("Buffy The vamp Slayer")) Then
    Call HayExterno("Buffy The vamp Slayer")
ElseIf FindWindow(vbNullString, UCase$("Blorb Slayer 1.12.552 (BETA)")) Then
    Call HayExterno("Blorb Slayer 1.12.552 (BETA)")
ElseIf FindWindow(vbNullString, UCase$("PumaEngine3.0")) Then
    Call HayExterno("PumaEngine3.0")
ElseIf FindWindow(vbNullString, UCase$("Vicious Engine 5.0")) Then
    Call HayExterno("Vicious Engine 5.0")
ElseIf FindWindow(vbNullString, UCase$("AkumaEngine33")) Then
    Call HayExterno("AkumaEngine33")
ElseIf FindWindow(vbNullString, UCase$("Spuc3ngine")) Then
    Call HayExterno("Spuc3ngine")
ElseIf FindWindow(vbNullString, UCase$("Ultra Engine")) Then
    Call HayExterno("Ultra Engine")
ElseIf FindWindow(vbNullString, UCase$("Engine")) Then
    Call HayExterno("Engine")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V5.4")) Then
    Call HayExterno("Cheat Engine V5.4")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.4")) Then
    Call HayExterno("Cheat Engine V4.4")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.4 German Add-On")) Then
    Call HayExterno("Cheat Engine V4.4 German Add-On")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.3")) Then
    Call HayExterno("Cheat Engine V4.3")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.2")) Then
    Call HayExterno("Cheat Engine V4.2")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.1.1")) Then
    Call HayExterno("Cheat Engine V4.1.1")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.3")) Then
    Call HayExterno("Cheat Engine V3.3")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.2")) Then
    Call HayExterno("Cheat Engine V3.2")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.1")) Then
    Call HayExterno("Cheat Engine V3.1")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine")) Then
    Call HayExterno("Cheat Engine")
ElseIf FindWindow(vbNullString, UCase$("danza engine 5.2.150")) Then
    Call HayExterno("danza engine 5.2.150")
ElseIf FindWindow(vbNullString, UCase$("zenx engine")) Then
    Call HayExterno("zenx engine")
ElseIf FindWindow(vbNullString, UCase$("MACROMAKER")) Then
    Call HayExterno("Macro Maker")
ElseIf FindWindow(vbNullString, UCase$("MACREOMAKER - EDIT MACRO")) Then
    Call HayExterno("Macro Maker")
ElseIf FindWindow(vbNullString, UCase$("By Fedex")) Then
    Call HayExterno("Macro Fedex")
ElseIf FindWindow(vbNullString, UCase$("Macro Mage 1.0")) Then
    Call HayExterno("Macro Mage")
ElseIf FindWindow(vbNullString, UCase$("Auto* v0.4 (c) 2001 [Agresión] Powa")) Then
    Call HayExterno("Macro Fisher")
ElseIf FindWindow(vbNullString, UCase$("Kizsada")) Then
    Call HayExterno("Macro K33")
ElseIf FindWindow(vbNullString, UCase$("Makro K33")) Then
    Call HayExterno("Macro K33")
ElseIf FindWindow(vbNullString, UCase$("Super Saiyan")) Then
    Call HayExterno("El Chit del Geri")
ElseIf FindWindow(vbNullString, UCase$("Makro-Piringulete")) Then
    Call HayExterno("Piringulete")
ElseIf FindWindow(vbNullString, UCase$("Makro-Piringulete 2003")) Then
    Call HayExterno("Piringulete 2003")
ElseIf FindWindow(vbNullString, UCase$("TUKY2005")) Then
    Call HayExterno("Makro Tuky")
ElseIf FindWindow(vbNullString, UCase$("Macro Configurable")) Then
Call HayExterno("Macro Configurable")
ElseIf FindWindow(vbNullString, UCase$("Macro")) Then
Call HayExterno("Macro")

End If

End Sub
Public Function HayExterno(ByVal Chit As String)
    Call MsgBox("Se ha detectado una aplicacion externa prohibida, seras expulsado del juego. Lee el reglamento si tienes dudas. (Te estaremos vigilando O.O)")
    End
End Function
'******************************************************************

'***********Detecta si se le cambio el nombre al exe***************
Public Function ChangeName() As Boolean
If OriginalClientName <> ClientName Then
ChangeName = True
Exit Function
End If
ChangeName = False
End Function
Public Sub ClientOn()
MsgBox "Se ha detectado cambio de nombre en el ejecutable. ¡No es posible ejecutar el cliente!.", vbCritical, "Winter-AO Ultimate"
End Sub
'*******************************************************************

'***************************Anti Debugger***************************
Public Function Debugger() As Boolean
If IsDebuggerPresent Then
Debugger = True
Exit Function
End If
Debugger = False
End Function
Public Sub AntiDebugger()
MsgBox "Se ha detectado un intento de Debuggear el cliente, su cliente será cerrado.!", vbCritical, "Winter-AO Ultimate"
End Sub
'*******************************************************************

'************************Anti Speed Hack****************************
Public Sub AntiShInitialize()
Time = GetTickCount()
End Sub
Public Function AntiSh(ByVal FramesPerSec) As Boolean
If GetTickCount - Time > 1050 Or GetTickCount - Time < 950 Then
        count = count + 1
    Else
        count = 0
    End If
    
    If FramesPerSec < 5 Then
    count = count + 1
    End If
    
    If count > 30 Then
       AntiSh = True
       Exit Function
    End If

Time = GetTickCount()
AntiSh = False
End Function
Public Sub AntiShOn()
MsgBox "Se ha detectado el uso de SpeedHack, el cliente será cerrado!.", vbCritical, "Winter-AO Ultimate"
End Sub

'*******************************************************************

