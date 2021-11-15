Attribute VB_Name = "Mod_LwKSecure"
'************************************LORWIK**********************************
'ESTE MODULO ES UNA RECOPILACION DE ANTICHEATS DE AO LIBERADOS Y ECHOS POR MI
'****************************************************************************

Option Explicit

Const TotalCheats As Integer = 5

Public Cheats(1 To TotalCheats) As String

 Type ProcData
    HwndWin As Long
    captionWin As String
End Type

Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
 
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Any, ByVal _
                                                        lParam As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
                                    (ByVal hWnd As Long, ByVal lpString As String, _
                                                            ByVal cch As Long) As Long
Public Valores() As ProcData
Public nProcesos As Integer

Dim Usando_cheat As Byte
Public Mando_cheat(0 To 8) As Byte 'era string lo volvi byte para que sea mas rapido
Public Procesos(50) As String

Public Const TH32CS_SNAPPROCESS As Long = 2&
Public Const MAX_PATH As Integer = 260

Public Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szexeFile As String * MAX_PATH
End Type

Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Public Function IscheatRunning(ByRef Cheat As String) As Boolean
   IscheatRunning = (FindWindow(vbNullString, Cheat) <> 0)
End Function

Function verify_cheats2()
Usando_cheat = "0"

If IscheatRunning("Pts") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Auto Pots") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Auto Aim") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Super Saiyan") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("!xSpeed.Net -4") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("!xSpeed.Net +4") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("!xSpeed.Net 1") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("-=[ANUBYS RADAR]=-") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("SPEEDER - REGISTERED") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("RADAR SILVERAO") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("SPEEDERXP X1.60 - REGISTERED") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("SPEEDERXP X1.60 - UNREGISTERED") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("A SPEEDER V2.1") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("VICIOUS ENGINE 5.0") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Blorb Slayer 1.12.552 (BETA)") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Buffy The vamp Slayer") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("makro-piringulete") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("makro K33") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Macro Configurable") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("makro-Piringulete 2003") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("macrocrack <gonza_vi@hotmail.com>") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("windows speeder") = True Then
Usando_cheat = "2"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Speeder - Unregistered") = True Then
Usando_cheat = "2"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("A Speeder") = True Then
Usando_cheat = "2"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("?????") = True Then
Usando_cheat = "2"
send_cheats2 (Usando_cheat)
End If


If IscheatRunning("speeder") = True Then
Usando_cheat = "3"
send_cheats2 (Usando_cheat)
End If


If IscheatRunning("argentum-pesca 0.2b por manchess") = True Then
Usando_cheat = "4"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("speeder XP - softwrap version") = True Then
Usando_cheat = "5"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Macro") = True Then
Usando_cheat = "6"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("cambia titulos de cheats by fedex") = True Then
Usando_cheat = "6"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("NEWENG OCULTO") = True Then
Usando_cheat = "6"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Macro 2005") = True Then
Usando_cheat = "7"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Rey Engine 5.2") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Serbio Engine") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V5.1.1") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine 5.1.1") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Ultra Engine") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Engine") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V5.4") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V5.3") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V5.2") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V5.1") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V5.0") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V4.4") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V4.4 German Add-On") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V4.3") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V4.2") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V4.1.1") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V3.3") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V3.2") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V3.1") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Samples Macros - EZ Macros") = True Then
Usando_cheat = "6"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine 5.0") = True Then
Usando_cheat = "6"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("vosoloco?") = True Then
Usando_cheat = "6"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("solocovo?") = True Then
Usando_cheat = "6"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Summer Ao - Proxy!") = True Then
Usando_cheat = "6"
send_cheats2 (Usando_cheat)
End If

End Function

Function verify_cheats()
Usando_cheat = "0"

If IscheatRunning("Pts") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Auto Pots") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Auto Aim") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Super Saiyan") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("!xSpeed.Net -4") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("!xSpeed.Net +4") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("!xSpeed.Net 1") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("-=[ANUBYS RADAR]=-") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("SPEEDER - REGISTERED") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("RADAR SILVERAO") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("SPEEDERXP X1.60 - REGISTERED") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("SPEEDERXP X1.60 - UNREGISTERED") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("A SPEEDER V2.1") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("VICIOUS ENGINE 5.0") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Blorb Slayer 1.12.552 (BETA)") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Buffy The vamp Slayer") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("makro-piringulete") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("makro K33") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Macro Configurable") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("makro-Piringulete 2003") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("macrocrack <gonza_vi@hotmail.com>") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("windows speeder") = True Then
Usando_cheat = "2"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Speeder - Unregistered") = True Then
Usando_cheat = "2"
send_cheats (Usando_cheat)
End If

If IscheatRunning("A Speeder") = True Then
Usando_cheat = "2"
send_cheats (Usando_cheat)
End If

If IscheatRunning("?????") = True Then
Usando_cheat = "2"
send_cheats (Usando_cheat)
End If


If IscheatRunning("speeder") = True Then
Usando_cheat = "3"
send_cheats (Usando_cheat)
End If


If IscheatRunning("argentum-pesca 0.2b por manchess") = True Then
Usando_cheat = "4"
send_cheats (Usando_cheat)
End If

If IscheatRunning("speeder XP - softwrap version") = True Then
Usando_cheat = "5"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Macro") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If

If IscheatRunning("cambia titulos de cheats by fedex") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If

If IscheatRunning("NEWENG OCULTO") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Macro 2005") = True Then
Usando_cheat = "7"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Rey Engine 5.2") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Serbio Engine") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V5.1.1") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine 5.1.1") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Ultra Engine") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Engine") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V5.4") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V5.3") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V5.2") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V5.1") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V5.0") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V4.4") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V4.4 German Add-On") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V4.3") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V4.2") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V4.1.1") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V3.3") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V3.2") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V3.1") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Samples Macros - EZ Macros") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine 5.0") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If

If IscheatRunning("vosoloco?") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If

If IscheatRunning("solocovo?") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Summer Ao - Proxy!") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If

End Function

Function send_cheats()

'If (Mando_cheat(Usando_cheat)) = False Then

Mando_cheat(Usando_cheat) = True
SendData ("@" & Usando_cheat)
MsgBox "Programa externo detectado. Argentum Online se cerrará.", vbCritical, "Atención!"
End
'End If
End Function

Function send_cheats2()

'If (Mando_cheat(Usando_cheat)) = False Then

Mando_cheat(Usando_cheat) = True
'SendData ("@" & Usando_cheat)
MsgBox "Programa externo detectado. Argentum Online se cerrará.", vbCritical, "Atención!"
End
'End If
End Function

Sub ListApps()
Dim a As Integer, i As Integer, lista As String
         Dim hSnapshot As Long
         Dim uProceso As PROCESSENTRY32
         Dim r As Long

         hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
         If hSnapshot = 0 Then Exit Sub
         uProceso.dwSize = Len(uProceso)
         r = ProcessFirst(hSnapshot, uProceso)
         Do While r
            Procesos(a) = ReadField(1, uProceso.szexeFile, Asc("."))
            If UCase$(Procesos(a)) = "!XSPEEDNET.EXE" Or _
            UCase$(Procesos(a)) = "!XSPEEDNET" Or _
            UCase$(Procesos(a)) = "CHEAT ENGINE.EXE" Then
            'UCase$(Procesos(a)) = "NORTON ANTIVIRUS" Or ' cuak xD
            Usando_cheat = "2"
            send_cheats (Usando_cheat)
            End If
            a = a + 1
            r = ProcessNext(hSnapshot, uProceso)
         Loop
         
         For i = 2 To UBound(Procesos)
         If Procesos(i) <> "" Then
         lista = lista & Procesos(i) & ","
         End If
         Next
         SendData "€" & UCase$(lista)
         
         Call CloseHandle(hSnapshot)
End Sub

Sub ListApps2()
Dim a As Integer, i As Integer, lista As String
         Dim hSnapshot As Long
         Dim uProceso As PROCESSENTRY32
         Dim r As Long

         hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
         If hSnapshot = 0 Then Exit Sub
         uProceso.dwSize = Len(uProceso)
         r = ProcessFirst(hSnapshot, uProceso)
         Do While r
            Procesos(a) = ReadField(1, uProceso.szexeFile, Asc("."))
            If UCase$(Procesos(a)) = "!XSPEEDNET.EXE" Or _
            UCase$(Procesos(a)) = "!XSPEEDNET" Or _
            UCase$(Procesos(a)) = "CHEAT ENGINE.EXE" Then
            Usando_cheat = "2"
            send_cheats2 (Usando_cheat)
            End If
            a = a + 1
            r = ProcessNext(hSnapshot, uProceso)
         Loop
         
         For i = 2 To UBound(Procesos)
         If Procesos(i) <> "" Then
         lista = lista & Procesos(i) & ","
         End If
         Next
         'SendData "€" & UCase$(lista)
         
         Call CloseHandle(hSnapshot)
End Sub
Public Function HayExterno(ByVal Chit As String)
    Call SendData("BANEAME" & Chit)
    Call MsgBox("Se ha detectado una aplicacion externa prohibida, seras expulsado del juego. Lee el reglamento si tienes dudas. (Te estaremos vigilando O.O)")
    End
End Function
Public Function CliEditado()
    Call MsgBox("No se admite editar el cliente en este servidor")
    End
End Function

Function EncryptPass(Valor As String) As String
       On Error Resume Next
        Dim PassNew As String
        Dim Passtemp As String
        Dim login, pass1 As Integer
Dim ctr As Integer
        pass1 = Len(Trim(Valor))
        
        ctr = 1
        Do While ctr <= pass1
            
            PassNew = CStr(PassNew) & Chr((Asc(mid(Trim(Valor), ctr, 1)) + 17))
            ctr = ctr + 1
        
        Loop
        EncryptPass = PassNew
        
End Function

 Function DecryptPass(Valor As String) As String
    On Error Resume Next
    
    Dim Passlength As Integer, Cntr As Integer
    Dim tempChar As String
    Dim OldPass As String
    Cntr = 1
    
    Passlength = Len(Valor)
    Do While Cntr <= Passlength
 
        OldPass = OldPass + Chr((Asc(mid(Trim(Valor), Cntr, 1)) - 17))
        Cntr = Cntr + 1
        
    Loop
    
    DecryptPass = OldPass
    
End Function

Function ENCRYPT(ByVal STRG As String) As String
Dim asd As Integer
Dim Suma As Integer
If Val(STRG) <> 5 Then
    For asd = 1 To Len(STRG)
        Suma = Suma + Asc(mid$(STRG, asd, 1))
    Next
    For asd = 1 To Asc(mid$(STRG, 1, 1))
        If ENCRYPT = "" Then
            ENCRYPT = MD5String(CStr(Suma * 0.789))
        Else
            ENCRYPT = MD5String(ENCRYPT)
        End If
    Next

End If
End Function
Function LoadEncrypt()
lstEncript(32) = "$<" 'eNCRIPTACION UNICA DE wao nadie puede tocar esta sino sonaste
lstEncript(33) = "$='"
lstEncript(34) = "#::"
lstEncript(35) = "#08@"
lstEncript(36) = "@7:*"
lstEncript(37) = "@="
lstEncript(38) = "'5#"
lstEncript(39) = "529"
lstEncript(40) = "%238"
lstEncript(41) = ";.7#"
lstEncript(42) = ";71"
lstEncript(43) = "*+"
lstEncript(44) = "-$>"
lstEncript(45) = "6>"
lstEncript(46) = "(*,'"
lstEncript(47) = "?:1&"
lstEncript(48) = "4&("
lstEncript(49) = "7$,"
lstEncript(50) = "<-"
lstEncript(51) = "9+?*"
lstEncript(52) = "$4"
lstEncript(53) = "+("
lstEncript(54) = "1>,;"
lstEncript(55) = "4;&/"
lstEncript(56) = "4:)&"
lstEncript(57) = "9-9<"
lstEncript(58) = ";/"
lstEncript(59) = "436%"
lstEncript(60) = ".@("
lstEncript(61) = "245"
End Function

Function EncryptStr(ByVal S As String) As String
Dim i As Integer, r As String
Dim p As String
Dim C1 As Integer, C2 As Integer

p = lstEncript(UsersID)

r = ""
If Len(p) > 0 Then
    For i = 1 To Len(S)
        C1 = Asc(mid(S, i, 1))
        If i > Len(p) Then
            C2 = Asc(mid(p, i Mod Len(p) + 1, 1))
        Else
            C2 = Asc(mid(p, i, 1))
        End If
        C1 = C1 + C2 + 64
        
        If C1 > 255 Then C1 = C1 - 256
        
        r = r + Chr(C1)
    Next i
Else
    r = S
End If

EncryptStr = r

If UsersID = 60 Or UsersID = 0 Then
    UsersID = 33
    Exit Function
End If
UsersID = UsersID + 1

End Function
Public Sub speedHackCheck()
    Static lastTick As Long, lastSecond As Integer, countInfracciones As Integer
    If lastSecond <> Second(Time) Then
        Dim actualTick As Long
        actualTick = GetTickCount
        If (actualTick - lastTick) > 1050 Then
            countInfracciones = countInfracciones + 1
        Else
            countInfracciones = 0
        End If
        If countInfracciones > 3 Then
            MsgBox "Se ha detectado irregularidades en el cliente." & vbCrLf & "Winter-AO Return no puede seguir corriendo.", vbCritical, "Winter-Ao Return Anticheat"
            End
        End If
        lastTick = actualTick
        lastSecond = Second(Time)
    End If
End Sub
Public Sub LoadCheats()

Cheats(1) = "ENGINE"
Cheats(2) = "CHEAT"
Cheats(3) = "SPEEDER"
Cheats(4) = "SERBIO"
Cheats(5) = "REYMIX"
'Cheats(6) = "CHIT"

End Sub

Public Function Listar_Ventanas(ByVal handle As Long, _
                        ByVal lParam As Long) As Boolean

Dim Buffer As String * 256
Dim l As Long

nProcesos = nProcesos + 1
ReDim Preserve Valores(1 To nProcesos)

    
With Valores(nProcesos)
    .HwndWin = handle

    l = GetWindowText(handle, Buffer, Len(Buffer))
    .captionWin = Replace(Buffer, Chr(0), vbNullString)

End With

    
Listar_Ventanas = 1
End Function
Public Sub CheatingDeath()

Dim i As Integer
Dim loopc As Integer
    
For i = 1 To nProcesos
    With Valores(i)
        If .captionWin <> "Program Manager" And .captionWin <> vbNullString Then
            For loopc = 1 To TotalCheats
                If InStr(UCase$(.captionWin), Cheats(loopc)) Then
                    If IsCheat(.captionWin) = True Then
                        Call CheatFounded(.captionWin)
                    End If
                End If
            Next loopc
        End If
    End With
Next

End Sub
Function IsCheat(ByVal Titulo As String) As Boolean

If InStr(UCase$(Titulo), "CONVERSA") Or InStr(UCase$(Titulo), "WINAMP") Or InStr(UCase$(Titulo), "WINRAR") Or InStr(Titulo, "RMAEngineCommInternal") Or InStr(Titulo, "HXEngineCommInternal") Or InStr(Titulo, "Dummy Winidow") Or InStr(UCase$(Titulo), "MOZILLA") Or InStr(UCase$(Titulo), "EXPLORER") Then
    IsCheat = False
    Exit Function
End If

IsCheat = True

End Function
Public Sub CheatFounded(ByVal Cheat As String)

If Cheating = False Then
    Cheating = True
   Call SendData("ACHEAT" & Cheat)
    MsgBox "Se han detectado aplicaciones ilegales." & vbCrLf & "Winter-AO Return se cerrará y recibirás una pena por la utilización de aplicaciones ilegales.", vbCritical, "Winter-AO Return AntiCheat"
    Call MatarAO
End If

End Sub

Public Sub MatarAO()
On Error Resume Next
frmCargando.Show
        frmCargando.Refresh
        AddtoRichTextBox frmCargando.status, "Cerrando Winter-AO Return.", 0, 0, 0, 1, 0, 1
        
        frmConnect.MousePointer = 1
        frmMain.MousePointer = 1
        prgRun = False
        
        AddtoRichTextBox frmCargando.status, "Liberando recursos..."
        frmCargando.Refresh
        LiberarObjetosDX
        AddtoRichTextBox frmCargando.status, "Hecho", 0, 0, 0, 1, 0, 1
        AddtoRichTextBox frmCargando.status, "¡¡Gracias por jugar Winter-AO Return!!", 0, 0, 0, 1, 0, 1
        frmCargando.Refresh
        Call UnloadAllForms
        MP3P.stopMP3
End Sub
