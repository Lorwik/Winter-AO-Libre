Attribute VB_Name = "TCP"
'Winter-AO 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Winter-AO is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

#If UsarQueSocket = 0 Then
' General constants used with most of the controls
Public Const INVALID_HANDLE As Integer = -1
Public Const CONTROL_ERRIGNORE As Integer = 0
Public Const CONTROL_ERRDISPLAY As Integer = 1


' SocietWrench Control Actions
Public Const SOCKET_OPEN As Integer = 1
Public Const SOCKET_CONNECT As Integer = 2
Public Const SOCKET_LISTEN As Integer = 3
Public Const SOCKET_ACCEPT As Integer = 4
Public Const SOCKET_CANCEL As Integer = 5
Public Const SOCKET_FLUSH As Integer = 6
Public Const SOCKET_CLOSE As Integer = 7
Public Const SOCKET_DISCONNECT As Integer = 7
Public Const SOCKET_ABORT As Integer = 8

' SocketWrench Control States
Public Const SOCKET_NONE As Integer = 0
Public Const SOCKET_IDLE As Integer = 1
Public Const SOCKET_LISTENING As Integer = 2
Public Const SOCKET_CONNECTING As Integer = 3
Public Const SOCKET_ACCEPTING As Integer = 4
Public Const SOCKET_RECEIVING As Integer = 5
Public Const SOCKET_SENDING As Integer = 6
Public Const SOCKET_CLOSING As Integer = 7

' Societ Address Families
Public Const AF_UNSPEC As Integer = 0
Public Const AF_UNIX As Integer = 1
Public Const AF_INET As Integer = 2

' Societ Types
Public Const SOCK_STREAM As Integer = 1
Public Const SOCK_DGRAM As Integer = 2
Public Const SOCK_RAW As Integer = 3
Public Const SOCK_RDM As Integer = 4
Public Const SOCK_SEQPACKET As Integer = 5

' Protocol Types
Public Const IPPROTO_IP As Integer = 0
Public Const IPPROTO_ICMP As Integer = 1
Public Const IPPROTO_GGP As Integer = 2
Public Const IPPROTO_TCP As Integer = 6
Public Const IPPROTO_PUP As Integer = 12
Public Const IPPROTO_UDP As Integer = 17
Public Const IPPROTO_IDP As Integer = 22
Public Const IPPROTO_ND As Integer = 77
Public Const IPPROTO_RAW As Integer = 255
Public Const IPPROTO_MAX As Integer = 256


' Network Addpesses
Public Const INADDR_ANY As String = "0.0.0.0"
Public Const INADDR_LOOPBACK As String = "127.0.0.1"
Public Const INADDR_NONE As String = "255.055.255.255"

' Shutdown Values
Public Const SOCKET_READ As Integer = 0
Public Const SOCKET_WRITE As Integer = 1
Public Const SOCKET_READWRITE As Integer = 2

' SocketWrench Error Pesponse
Public Const SOCKET_ERRIGNORE As Integer = 0
Public Const SOCKET_ERRDISPLAY As Integer = 1

' SocketWrench Error Codes
Public Const WSABASEERR As Integer = 24000
Public Const WSAEINTR As Integer = 24004
Public Const WSAEBADF As Integer = 24009
Public Const WSAEACCES As Integer = 24013
Public Const WSAEFAULT As Integer = 24014
Public Const WSAEINVAL As Integer = 24022
Public Const WSAEMFILE As Integer = 24024
Public Const WSAEWOULDBLOCK As Integer = 24035
Public Const WSAEINPROGRESS As Integer = 24036
Public Const WSAEALREADY As Integer = 24037
Public Const WSAENOTSOCK As Integer = 24038
Public Const WSAEDESTADDRREQ As Integer = 24039
Public Const WSAEMSGSIZE As Integer = 24040
Public Const WSAEPROTOTYPE As Integer = 24041
Public Const WSAENOPROTOOPT As Integer = 24042
Public Const WSAEPROTONOSUPPORT As Integer = 24043
Public Const WSAESOCKTNOSUPPORT As Integer = 24044
Public Const WSAEOPNOTSUPP As Integer = 24045
Public Const WSAEPFNOSUPPORT As Integer = 24046
Public Const WSAEAFNOSUPPORT As Integer = 24047
Public Const WSAEADDRINUSE As Integer = 24048
Public Const WSAEADDRNOTAVAIL As Integer = 24049
Public Const WSAENETDOWN As Integer = 24050
Public Const WSAENETUNREACH As Integer = 24051
Public Const WSAENETRESET As Integer = 24052
Public Const WSAECONNABORTED As Integer = 24053
Public Const WSAECONNRESET As Integer = 24054
Public Const WSAENOBUFS As Integer = 24055
Public Const WSAEISCONN As Integer = 24056
Public Const WSAENOTCONN As Integer = 24057
Public Const WSAESHUTDOWN As Integer = 24058
Public Const WSAETOOMANYREFS As Integer = 24059
Public Const WSAETIMEDOUT As Integer = 24060
Public Const WSAECONNREFUSED As Integer = 24061
Public Const WSAELOOP As Integer = 24062
Public Const WSAENAMETOOLONG As Integer = 24063
Public Const WSAEHOSTDOWN As Integer = 24064
Public Const WSAEHOSTUNREACH As Integer = 24065
Public Const WSAENOTEMPTY As Integer = 24066
Public Const WSAEPROCLIM As Integer = 24067
Public Const WSAEUSERS As Integer = 24068
Public Const WSAEDQUOT As Integer = 24069
Public Const WSAESTALE As Integer = 24070
Public Const WSAEREMOTE As Integer = 24071
Public Const WSASYSNOTREADY As Integer = 24091
Public Const WSAVERNOTSUPPORTED As Integer = 24092
Public Const WSANOTINITIALISED As Integer = 24093
Public Const WSAHOST_NOT_FOUND As Integer = 25001
Public Const WSATRY_AGAIN As Integer = 25002
Public Const WSANO_RECOVERY As Integer = 25003
Public Const WSANO_DATA As Integer = 25004
Public Const WSANO_ADDRESS As Integer = 2500
#End If

Sub DarCuerpoYCabeza(ByVal Userindex As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last modified: 14/03/2007
'Elije una cabeza para el usuario y le da un body
'*************************************************
Dim NewBody As Integer
Dim NewHead As Integer
Dim UserRaza As Byte
Dim UserGenero As Byte
UserGenero = UserList(Userindex).genero
UserRaza = UserList(Userindex).raza
Select Case UserGenero
   Case eGenero.Hombre
        Select Case UserRaza
            Case eRaza.Humano
                NewHead = RandomNumber(1, 40)
                NewBody = 1
            Case eRaza.Elfo
                NewHead = RandomNumber(101, 112)
                NewBody = 2
            Case eRaza.Drow
                NewHead = RandomNumber(200, 210)
                NewBody = 3
            Case eRaza.Enano
                NewHead = RandomNumber(300, 306)
                NewBody = 300
            Case eRaza.Gnomo
                NewHead = RandomNumber(401, 406)
                NewBody = 300
            Case eRaza.orco
                NewHead = RandomNumber(516, 525)
                NewBody = 502
        End Select
   Case eGenero.Mujer
        Select Case UserRaza
            Case eRaza.Humano
                NewHead = RandomNumber(70, 79)
                NewBody = 1
            Case eRaza.Elfo
                NewHead = RandomNumber(170, 178)
                NewBody = 2
            Case eRaza.Drow
                NewHead = RandomNumber(270, 278)
                NewBody = 3
            Case eRaza.Gnomo
                NewHead = RandomNumber(370, 372)
                NewBody = 300
            Case eRaza.Enano
                NewHead = RandomNumber(470, 476)
                NewBody = 300
            Case eRaza.orco
                NewHead = RandomNumber(526, 531)
                NewBody = 503
        End Select
End Select
UserList(Userindex).Char.Head = NewHead
UserList(Userindex).Char.Body = NewBody
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))
    
    If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
        AsciiValidos = False
        Exit Function
    End If
    
Next i

AsciiValidos = True

End Function

Function Numeric(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))
    
    If (car < 48 Or car > 57) Then
        Numeric = False
        Exit Function
    End If
    
Next i

Numeric = True

End Function


Function NombrePermitido(ByVal Nombre As String) As Boolean
Dim i As Integer

For i = 1 To UBound(ForbidenNames)
    If InStr(Nombre, ForbidenNames(i)) Then
            NombrePermitido = False
            Exit Function
    End If
Next i

NombrePermitido = True

End Function

Function ValidateSkills(ByVal Userindex As Integer) As Boolean

Dim LoopC As Integer

For LoopC = 1 To NUMSKILLS
    If UserList(Userindex).Stats.UserSkills(LoopC) < 0 Then
        Exit Function
        If UserList(Userindex).Stats.UserSkills(LoopC) > 100 Then UserList(Userindex).Stats.UserSkills(LoopC) = 100
    End If
Next LoopC

ValidateSkills = True
    
End Function

Sub ConnectNewUser(ByVal Userindex As Integer, ByRef Name As String, ByVal UserRaza As eRaza, ByVal UserSexo As eGenero, ByVal UserClase As eClass, ByVal Cabeza As Integer, ByVal CuentaName As String)
    '*************************************************
    'Author: Unknown
    'Last modified: 20/4/2007
    'Conecta un nuevo Usuario
    '23/01/2007 Pablo (ToxicWaste) - Agregué ResetFaccion al crear usuario
    '24/01/2007 Pablo (ToxicWaste) - Agregué el nuevo mana inicial de los magos.
    '12/02/2007 Pablo (ToxicWaste) - Puse + 1 de const al Elfo normal.
    '20/04/2007 Pablo (ToxicWaste) - Puse -1 de fuerza al Elfo.
    '09/01/2008 Pablo (ToxicWaste) - Ahora los modificadores de Raza se controlan desde Balance.dat
    '*************************************************
    Dim i As Long
    
    With UserList(Userindex)
    
    If Not AsciiValidos(Name) Or LenB(Name) = 0 Then
        Call WriteErrorMsg(Userindex, "Nombre invalido.")
        Exit Sub
    End If
    
    If .flags.UserLogged Then
        Call LogCheating("El usuario " & .Name & " ha intentado crear a " & Name & " desde la IP " & .ip)
        
        'Kick player ( and leave character inside :D )!
        Call CloseSocketSL(Userindex)
        Call Cerrar_Usuario(Userindex)
        
        Exit Sub
    End If
    
    Dim LoopC As Long
    Dim totalskpts As Long
    
    '¿Existe el personaje?
    If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) = True Then
        Call WriteErrorMsg(Userindex, "Ya existe el personaje.")
        Exit Sub
    End If
    
    'Tiró los dados antes de llegar acá??
    If .Stats.UserAtributos(eAtributos.Fuerza) = 0 Then
        Call WriteErrorMsg(Userindex, "Debes de asignar los atributos.")
        Exit Sub
    End If
    
    .flags.Muerto = 0
    .flags.Escondido = 0
    
    .Reputacion.AsesinoRep = 0
    .Reputacion.BandidoRep = 0
    .Reputacion.BurguesRep = 0
    .Reputacion.LadronesRep = 0
    .Reputacion.NobleRep = 1000
    .Reputacion.PlebeRep = 30
    
    .Reputacion.Promedio = 30 / 6
    
    
    .Name = Name
    .clase = UserClase
    .raza = UserRaza
    .genero = UserSexo
    .email = (GetVar(App.Path & "\Cuentas\" & CuentaName & ".acc", "CUENTA", "MAIL"))
    .Hogar = eCiudad.cRamx
    
            
        If .Stats.UserAtributos(eAtributos.Fuerza) + .Stats.UserAtributos(eAtributos.Agilidad) + .Stats.UserAtributos(eAtributos.Inteligencia) + .Stats.UserAtributos(eAtributos.Carisma) + .Stats.UserAtributos(eAtributos.Constitucion) > 74 Then
            Call WriteErrorMsg(Userindex, "Error en la asignacion de atributos, vuelva a asignarlos.")
            Call FlushBuffer(Userindex)
            Call CloseSocket(Userindex)
            Exit Sub
        End If
        
        If .Stats.UserAtributos(eAtributos.Fuerza) > 18 Then
            Call WriteErrorMsg(Userindex, "Error en la asignacion de atributos, vuelva a asignarlos.")
            Call FlushBuffer(Userindex)
            Call CloseSocket(Userindex)
            Exit Sub
        End If
        
        If .Stats.UserAtributos(eAtributos.Inteligencia) > 18 Then
            Call WriteErrorMsg(Userindex, "Error en la asignacion de atributos, vuelva a asignarlos.")
            Call FlushBuffer(Userindex)
            Call CloseSocket(Userindex)
            Exit Sub
        End If
        
        If .Stats.UserAtributos(eAtributos.Carisma) > 18 Then
            Call WriteErrorMsg(Userindex, "Error en la asignacion de atributos, vuelva a asignarlos.")
            Call FlushBuffer(Userindex)
            Call CloseSocket(Userindex)
            Exit Sub
        End If
        
        If .Stats.UserAtributos(eAtributos.Constitucion) > 18 Then
            Call WriteErrorMsg(Userindex, "Error en la asignacion de atributos, vuelva a asignarlos.")
            Call FlushBuffer(Userindex)
            Call CloseSocket(Userindex)
            Exit Sub
        End If
    
    '[Pablo (Toxic Waste) 9/01/08]
    .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + ModRaza(UserRaza).Fuerza
    .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + ModRaza(UserRaza).Agilidad
    .Stats.UserAtributos(eAtributos.Inteligencia) = .Stats.UserAtributos(eAtributos.Inteligencia) + ModRaza(UserRaza).Inteligencia
    .Stats.UserAtributos(eAtributos.Carisma) = .Stats.UserAtributos(eAtributos.Carisma) + ModRaza(UserRaza).Carisma
    .Stats.UserAtributos(eAtributos.Constitucion) = .Stats.UserAtributos(eAtributos.Constitucion) + ModRaza(UserRaza).Constitucion
    '[/Pablo (Toxic Waste)]
    
    For i = 1 To NUMSKILLS
        .Stats.UserSkills(i) = 0
    Next i
        
    .Stats.SkillPts = 10
    
    .Char.heading = eHeading.SOUTH
    
    Call DarCuerpoYCabeza(Userindex)
    .Char.Head = Cabeza
    .OrigChar = .Char
     
    .Char.WeaponAnim = NingunArma
    .Char.ShieldAnim = NingunEscudo
    .Char.CascoAnim = NingunCasco
    .Char.Aura = 0
    
    Dim MiInt As Long
    MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.Constitucion) \ 3)
    
    .Stats.MaxHP = 15 + MiInt
    .Stats.MinHP = 15 + MiInt
    
    MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.Agilidad) \ 6)
    If MiInt = 1 Then MiInt = 2
    
    .Stats.MaxSta = 20 * MiInt
    .Stats.MinSta = 20 * MiInt
    
    
    .Stats.MaxAGU = 100
    .Stats.MinAGU = 100
    
    .Stats.MaxHam = 100
    .Stats.MinHam = 100
    
    
    '<-----------------MANA----------------------->
    If UserClase = eClass.Mage Then 'Cambio en mana inicial (ToxicWaste)
        MiInt = .Stats.UserAtributos(eAtributos.Inteligencia) * 3
        .Stats.MaxMAN = MiInt
        .Stats.MinMAN = MiInt
    ElseIf UserClase = eClass.Cleric Or UserClase = eClass.Druid _
        Or UserClase = eClass.Bard Or UserClase = eClass.Assasin Then
            .Stats.MaxMAN = 50
            .Stats.MinMAN = 50
    ElseIf UserClase = eClass.Bandit Then 'Mana Inicial del Bandido (ToxicWaste)
            .Stats.MaxMAN = 150
            .Stats.MinMAN = 150
    Else
        .Stats.MaxMAN = 0
        .Stats.MinMAN = 0
    End If
    
    If UserClase = eClass.Mage Or UserClase = eClass.Cleric Or _
       UserClase = eClass.Druid Or UserClase = eClass.Bard Or _
       UserClase = eClass.Assasin Then
            .Stats.UserHechizos(1) = 2
    End If
    
    .Stats.MaxHIT = 2
    .Stats.MinHIT = 1
    
    .Stats.GLD = 0
    .Stats.Puntos = 0
    
    .Stats.Exp = 0
    .Stats.ELU = 300
    .Stats.ELV = 1
    
       '???????????????? INVENTARIO ¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿
    Dim Slot As Byte
    Dim IsPaladin As Boolean
    
    IsPaladin = UserClase = eClass.Paladin
    
    'Pociones Rojas (Newbie)
    Slot = 1
    .Invent.Object(Slot).ObjIndex = 461
    .Invent.Object(Slot).amount = 100
    
    'Pociones azules (Newbie)
    If .Stats.MaxMAN > 0 Or IsPaladin Then
        Slot = Slot + 1
        .Invent.Object(Slot).ObjIndex = 1097
        .Invent.Object(Slot).amount = 200
    
    Else
        'Pociones amarillas (Newbie)
        Slot = Slot + 1
        .Invent.Object(Slot).ObjIndex = 1098
        .Invent.Object(Slot).amount = 50
    
        'Pociones verdes (Newbie)
        Slot = Slot + 1
        .Invent.Object(Slot).ObjIndex = 462
        .Invent.Object(Slot).amount = 50
    
    End If
    
    ' Ropa (Newbie)
    Slot = Slot + 1
    Select Case UserRaza
        Case eRaza.Humano
            .Invent.Object(Slot).ObjIndex = 463
        Case eRaza.Elfo
            .Invent.Object(Slot).ObjIndex = 464
        Case eRaza.Drow
            .Invent.Object(Slot).ObjIndex = 465
        Case eRaza.Enano
            .Invent.Object(Slot).ObjIndex = 466
        Case eRaza.Gnomo
            .Invent.Object(Slot).ObjIndex = 466
        Case eRaza.orco
            If eGenero.Hombre Then
                .Invent.Object(Slot).ObjIndex = 1150
            Else
                .Invent.Object(Slot).ObjIndex = 1151
            End If
    End Select
    
    ' Equipo ropa
    .Invent.Object(Slot).amount = 1
    .Invent.Object(Slot).Equipped = 1
    
    .Invent.ArmourEqpSlot = Slot
    .Invent.ArmourEqpObjIndex = .Invent.Object(Slot).ObjIndex

    'Arma (Newbie)
    Slot = Slot + 1
    Select Case UserClase
        Case eClass.Hunter
            ' Arco (Newbie)
            .Invent.Object(Slot).ObjIndex = 859
        Case eClass.Paladin
            ' Espada Larga (Newbie)
            .Invent.Object(Slot).ObjIndex = 860
        Case Else
            ' Daga (Newbie)
            .Invent.Object(Slot).ObjIndex = 460
    End Select
    
    ' Equipo arma
    .Invent.Object(Slot).amount = 1
    .Invent.Object(Slot).Equipped = 1
    
    .Invent.WeaponEqpObjIndex = .Invent.Object(Slot).ObjIndex
    .Invent.WeaponEqpSlot = Slot
    
    .Char.WeaponAnim = GetWeaponAnim(Userindex, .Invent.WeaponEqpObjIndex)

    ' Municiones (Newbie)
    If UserClase = eClass.Hunter Then
        Slot = Slot + 1
        .Invent.Object(Slot).ObjIndex = 861
        .Invent.Object(Slot).amount = 150
        
        ' Equipo flechas
        .Invent.Object(Slot).Equipped = 1
        .Invent.MunicionEqpSlot = Slot
        .Invent.MunicionEqpObjIndex = 861
    End If

    ' Manzanas (Newbie)
    Slot = Slot + 1
    .Invent.Object(Slot).ObjIndex = 467
    .Invent.Object(Slot).amount = 100
    
    ' Jugos (Nwbie)
    Slot = Slot + 1
    .Invent.Object(Slot).ObjIndex = 468
    .Invent.Object(Slot).amount = 100
    
    ' Sin casco y escudo
    .Char.ShieldAnim = NingunEscudo
    .Char.CascoAnim = NingunCasco
    
    ' Total Items
    .Invent.NroItems = Slot
     
    #If ConUpTime Then
        .LogOnTime = Now
        .UpTime = 0
    #End If
    
    End With
    
    'Valores Default de facciones al Activar nuevo usuario
    Call ResetFacciones(Userindex)
    
    Call WriteVar(CharPath & UCase$(Name) & ".chr", "INIT", "Password", (GetVar(App.Path & "\Cuentas\" & CuentaName & ".acc", "CUENTA", "MAIL"))) 'grabamos el password aqui afuera, para no mantenerlo cargado en memoria
    
    'Lo pongo aqui ya que en caso de errores se podria hackear cuentas.
    Call AgregarPersonaje(Userindex, CuentaName, Name)
    
    Call SaveUser(Userindex, CharPath & UCase$(Name) & ".chr")
      
    'Open User
    Call ConnectUser(Userindex, Name)
  
End Sub

#If UsarQueSocket = 1 Or UsarQueSocket = 2 Then

Sub CloseSocket(ByVal Userindex As Integer)

On Error GoTo Errhandler
    
    If Userindex = LastUser Then
        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1
            If LastUser < 1 Then Exit Do
        Loop
    End If
    
    If UserList(Userindex).flags.AutoTorneo = True Then
        Call Rondas_UsuarioDesconecta(Userindex)
    End If
    
    'Call SecurityIp.IpRestarConexion(GetLongIp(UserList(UserIndex).ip))
    
    If UserList(Userindex).ConnID <> -1 Then
        Call CloseSocketSL(Userindex)
    End If
    
    'Es el mismo user al que está revisando el centinela??
    'IMPORTANTE!!! hacerlo antes de resetear así todavía sabemos el nombre del user
    ' y lo podemos loguear
    If Centinela.RevisandoUserIndex = Userindex Then _
        Call modCentinela.CentinelaUserLogout
    
    'mato los comercios seguros
    If UserList(Userindex).ComUsu.DestUsu > 0 Then
        If UserList(UserList(Userindex).ComUsu.DestUsu).flags.UserLogged Then
            If UserList(UserList(Userindex).ComUsu.DestUsu).ComUsu.DestUsu = Userindex Then
                Call WriteConsoleMsg(UserList(Userindex).ComUsu.DestUsu, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(UserList(Userindex).ComUsu.DestUsu)
                Call FlushBuffer(UserList(Userindex).ComUsu.DestUsu)
            End If
        End If
    End If
    
    'Empty buffer for reuse
    Call UserList(Userindex).incomingData.ReadASCIIStringFixed(UserList(Userindex).incomingData.length)
    
    If UserList(Userindex).flags.UserLogged Then
        If NumUsers > 0 Then NumUsers = NumUsers - 1
        Call CloseUser(Userindex)
    Else
        Call ResetUserSlot(Userindex)
    End If
    
    UserList(Userindex).ConnID = -1
    UserList(Userindex).ConnIDValida = False
    
Exit Sub

Errhandler:
    UserList(Userindex).ConnID = -1
    UserList(Userindex).ConnIDValida = False
    Call ResetUserSlot(Userindex)

    Call LogError("CloseSocket - Error = " & Err.Number & " - Descripción = " & Err.description & " - UserIndex = " & Userindex)
End Sub

#ElseIf UsarQueSocket = 0 Then

Sub CloseSocket(ByVal Userindex As Integer)
On Error GoTo Errhandler
    
    
    
    UserList(Userindex).ConnID = -1

    If Userindex = LastUser And LastUser > 1 Then
        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1
            If LastUser <= 1 Then Exit Do
        Loop
    End If

    If UserList(Userindex).flags.UserLogged Then
            If NumUsers <> 0 Then NumUsers = NumUsers - 1
            Call CloseUser(Userindex)
    End If

    frmMain.Socket2(Userindex).Cleanup
    Unload frmMain.Socket2(Userindex)
    Call ResetUserSlot(Userindex)

Exit Sub

Errhandler:
    UserList(Userindex).ConnID = -1
    Call ResetUserSlot(Userindex)
End Sub







#ElseIf UsarQueSocket = 3 Then

Sub CloseSocket(ByVal Userindex As Integer, Optional ByVal cerrarlo As Boolean = True)

On Error GoTo Errhandler

Dim NURestados As Boolean
Dim CoNnEcTiOnId As Long


    NURestados = False
    CoNnEcTiOnId = UserList(Userindex).ConnID
    
    'call logindex(UserIndex, "******> Sub CloseSocket. ConnId: " & CoNnEcTiOnId & " Cerrarlo: " & Cerrarlo)
    
    
  
    UserList(Userindex).ConnID = -1 'inabilitamos operaciones en socket

    If Userindex = LastUser And LastUser > 1 Then
        Do
            LastUser = LastUser - 1
            If LastUser <= 1 Then Exit Do
        Loop While UserList(LastUser).flags.UserLogged = True
    End If

    If UserList(Userindex).flags.UserLogged Then
            If NumUsers <> 0 Then NumUsers = NumUsers - 1
            NURestados = True
            Call CloseUser(Userindex)
    End If
    
    Call ResetUserSlot(Userindex)
    
    'limpiada la userlist... reseteo el socket, si me lo piden
    'Me lo piden desde: cerrada intecional del servidor (casi todas
    'las llamadas a CloseSocket del codigo)
    'No me lo piden desde: disconnect remoto (el on_close del control
    'de alejo realiza la desconexion automaticamente). Esto puede pasar
    'por ejemplo, si el cliente cierra el AO.
    If cerrarlo Then Call frmMain.TCPServ.CerrarSocket(CoNnEcTiOnId)

Exit Sub

Errhandler:
    Call LogError("CLOSESOCKETERR: " & Err.description & " UI:" & Userindex)
    
    If Not NURestados Then
        If UserList(Userindex).flags.UserLogged Then
            If NumUsers > 0 Then
                NumUsers = NumUsers - 1
            End If
            Call LogError("Cerre sin grabar a: " & UserList(Userindex).Name)
        End If
    End If
    
    Call LogError("El usuario no guardado tenia connid " & CoNnEcTiOnId & ". Socket no liberado.")
    Call ResetUserSlot(Userindex)

End Sub


#End If

'[Alejo-21-5]: Cierra un socket sin limpiar el slot
Sub CloseSocketSL(ByVal Userindex As Integer)

#If UsarQueSocket = 1 Then

If UserList(Userindex).ConnID <> -1 And UserList(Userindex).ConnIDValida Then
    Call BorraSlotSock(UserList(Userindex).ConnID)
    Call WSApiCloseSocket(UserList(Userindex).ConnID)
    UserList(Userindex).ConnIDValida = False
End If

#ElseIf UsarQueSocket = 0 Then

If UserList(Userindex).ConnID <> -1 And UserList(Userindex).ConnIDValida Then
    frmMain.Socket2(Userindex).Cleanup
    Unload frmMain.Socket2(Userindex)
    UserList(Userindex).ConnIDValida = False
End If

#ElseIf UsarQueSocket = 2 Then

If UserList(Userindex).ConnID <> -1 And UserList(Userindex).ConnIDValida Then
    Call frmMain.Serv.CerrarSocket(UserList(Userindex).ConnID)
    UserList(Userindex).ConnIDValida = False
End If

#End If
End Sub

''
' Send an string to a Slot
'
' @param userIndex The index of the User
' @param Datos The string that will be send
' @remarks If UsarQueSocket is 3 it won`t use the clsByteQueue

Public Function EnviarDatosASlot(ByVal Userindex As Integer, ByRef Datos As String) As Long
'***************************************************
'Author: Unknown
'Last Modification: 01/10/07
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
'Now it uses the clsByteQueue class and don`t make a FIFO Queue of String
'***************************************************

#If UsarQueSocket = 1 Then '**********************************************
    On Error GoTo Err
    
    Dim Ret As Long
    
    Ret = WsApiEnviar(Userindex, Datos)
    
    If Ret <> 0 And Ret <> WSAEWOULDBLOCK Then
        ' Close the socket avoiding any critical error
        Call CloseSocketSL(Userindex)
        Call Cerrar_Usuario(Userindex)
    End If
Exit Function
    
Err:

#ElseIf UsarQueSocket = 0 Then '**********************************************
    
    If frmMain.Socket2(Userindex).Write(Datos, Len(Datos)) < 0 Then
        If frmMain.Socket2(Userindex).LastError = WSAEWOULDBLOCK Then
            ' WSAEWOULDBLOCK, put the data again in the outgoingData Buffer
            Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(Datos)
        Else
            'Close the socket avoiding any critical error
            Call Cerrar_Usuario(Userindex)
        End If
    End If
#ElseIf UsarQueSocket = 2 Then '**********************************************

    'Return value for this Socket:
    '--0) OK
    '--1) WSAEWOULDBLOCK
    '--2) ERROR
    
    Dim Ret As Long

    Ret = frmMain.Serv.Enviar(.ConnID, Datos, Len(Datos))
            
    If Ret = 1 Then
        ' WSAEWOULDBLOCK, put the data again in the outgoingData Buffer
        Call .outgoingData.WriteASCIIStringFixed(Datos)
    ElseIf Ret = 2 Then
        'Close socket avoiding any critical error
        Call CloseSocketSL(Userindex)
        Call Cerrar_Usuario(Userindex)
    End If
    

#ElseIf UsarQueSocket = 3 Then
    'THIS SOCKET DOESN`T USE THE BYTE QUEUE CLASS
    Dim rv As Long
    'al carajo, esto encola solo!!! che, me aprobará los
    'parciales también?, este control hace todo solo!!!!
    On Error GoTo ErrorHandler
        
        If UserList(Userindex).ConnID = -1 Then
            Call LogError("TCP::EnviardatosASlot, se intento enviar datos a un userIndex con ConnId=-1")
            Exit Function
        End If
        
        If frmMain.TCPServ.Enviar(UserList(Userindex).ConnID, Datos, Len(Datos)) = 2 Then Call CloseSocket(Userindex)

Exit Function
ErrorHandler:
    Call LogError("TCP::EnviarDatosASlot. UI/ConnId/Datos: " & Userindex & "/" & UserList(Userindex).ConnID & "/" & Datos)
#End If '**********************************************

End Function
Function EstaPCarea(index As Integer, Index2 As Integer) As Boolean


Dim X As Integer, Y As Integer
For Y = UserList(index).Pos.Y - MinYBorder + 1 To UserList(index).Pos.Y + MinYBorder - 1
        For X = UserList(index).Pos.X - MinXBorder + 1 To UserList(index).Pos.X + MinXBorder - 1

            If MapData(UserList(index).Pos.map, X, Y).Userindex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
        Next X
Next Y
EstaPCarea = False
End Function

Function HayPCarea(Pos As WorldPos) As Boolean


Dim X As Integer, Y As Integer
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                If MapData(Pos.map, X, Y).Userindex > 0 Then
                    HayPCarea = True
                    Exit Function
                End If
            End If
        Next X
Next Y
HayPCarea = False
End Function

Function HayOBJarea(Pos As WorldPos, ObjIndex As Integer) As Boolean


Dim X As Integer, Y As Integer
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If MapData(Pos.map, X, Y).ObjInfo.ObjIndex = ObjIndex Then
                HayOBJarea = True
                Exit Function
            End If
        
        Next X
Next Y
HayOBJarea = False
End Function
Function ValidateChr(ByVal Userindex As Integer) As Boolean

ValidateChr = UserList(Userindex).Char.Head <> 0 _
                And UserList(Userindex).Char.Body <> 0 _
                And ValidateSkills(Userindex)

End Function

Sub ConnectUser(ByVal Userindex As Integer, ByRef Name As String)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 12/06/2009
'26/03/2009: ZaMa - Agrego por default que el color de dialogo de los dioses, sea como el de su nick.
'12/06/2009: ZaMa - Agrego chequeo de nivel al loguear
'***************************************************
    Dim N As Integer
    Dim tStr As String
    
    With UserList(Userindex)
    
    '¿Ya esta conectado el personaje?
    If CheckForSameName(Name) Then
        Call CloseSocketSL(Userindex)
        Call WriteErrorMsg(Userindex, "El personaje esta conectado.")
        Exit Sub
    End If
    
    If .flags.UserLogged Then
        Call LogCheating("El usuario " & .Name & " ha intentado loguear a " & Name & " desde la IP " & .ip)
        'Kick player ( and leave character inside :D )!
        Call CloseSocketSL(Userindex)
        Call Cerrar_Usuario(Userindex)
        Call WriteErrorMsg(Userindex, "El personaje esta conectado.")
        Exit Sub
    End If
    
    'Reseteamos los FLAGS
    .flags.Escondido = 0
    .flags.TargetNPC = 0
    .flags.TargetNpcTipo = eNPCType.Comun
    .flags.TargetObj = 0
    .flags.TargetUser = 0
    .Char.FX = 0
    
    'Controlamos no pasar el maximo de usuarios
    If NumUsers >= MaxUsers Then
        Call WriteErrorMsg(Userindex, "El servidor ha alcanzado el máximo de usuarios soportado, por favor vuelva a intertarlo más tarde.")
        Call FlushBuffer(Userindex)
        Call CloseSocket(Userindex)
        Exit Sub
    End If
    
    '¿Este IP ya esta conectado?
    If AllowMultiLogins = 0 Then
        If CheckForSameIP(Userindex, .ip) = True Then
            Call WriteErrorMsg(Userindex, "No es posible usar más de un personaje al mismo tiempo.")
            Call FlushBuffer(Userindex)
            Call CloseSocket(Userindex)
            Exit Sub
        End If
    End If
    
    '¿Existe el personaje?
    If Not FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) Then
        Call WriteErrorMsg(Userindex, "El personaje no existe.")
        Call FlushBuffer(Userindex)
        Call CloseSocket(Userindex)
        Exit Sub
    End If
    
    'Reseteamos los privilegios
    .flags.Privilegios = 0
    
    'Vemos que clase de user es (se lo usa para setear los privilegios al loguear el PJ)
    If EsAdmin(Name) Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.Admin
        Call LogGM(Name, "Se conecto con ip:" & .ip)
    ElseIf EsDios(Name) Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.Dios
        Call LogGM(Name, "Se conecto con ip:" & .ip)
    ElseIf EsSemiDios(Name) Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.SemiDios
        Call LogGM(Name, "Se conecto con ip:" & .ip)
    ElseIf EsConsejero(Name) Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.Consejero
        Call LogGM(Name, "Se conecto con ip:" & .ip)
    Else
        .flags.Privilegios = .flags.Privilegios Or PlayerType.User
        .flags.AdminPerseguible = True
    End If
    
    'Add RM flag if needed
    If EsRolesMaster(Name) Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.RoleMaster
    End If
    
    If ServerSoloGMs > 0 Then
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) = 0 Then
            Call WriteErrorMsg(Userindex, "Servidor restringido a administradores. Por favor reintente en unos momentos.")
            Call FlushBuffer(Userindex)
            Call CloseSocket(Userindex)
            Exit Sub
        End If
    End If
    
    'Cargamos el personaje
    Dim Leer As New clsIniReader
    
    Call Leer.Initialize(CharPath & UCase$(Name) & ".chr")
    
    'Cargamos los datos del personaje
    Call LoadUserInit(Userindex, Leer)
    
    Call LoadUserStats(Userindex, Leer)
    
    Call LoadQuestStats(Userindex, Leer)
    
    If Not ValidateChr(Userindex) Then
        Call WriteErrorMsg(Userindex, "Error en el personaje.")
        Call CloseSocket(Userindex)
        Exit Sub
    End If
    
    Call LoadUserReputacion(Userindex, Leer)
    
    Set Leer = Nothing
    
    If .Invent.EscudoEqpSlot = 0 Then .Char.ShieldAnim = NingunEscudo
    If .Invent.CascoEqpSlot = 0 Then .Char.CascoAnim = NingunCasco
    If .Invent.WeaponEqpSlot = 0 Then .Char.WeaponAnim = NingunArma
    
    If (.flags.Muerto = 0) Then
        .flags.SeguroResu = False
        Call WriteResuscitationSafeOff(Userindex)
    Else
        .flags.SeguroResu = True
        Call WriteResuscitationSafeOn(Userindex)
    End If
    
    Call UpdateUserInv(True, Userindex, 0)
    Call UpdateUserHechizos(True, Userindex, 0)
    
    If UserList(Userindex).flags.Equitando = 1 Then
         UserList(Userindex).Char.Body = ObjData(UserList(Userindex).Invent.MonturaObjIndex).Ropaje
    End If
    
    If .flags.Paralizado Then
        Call WriteParalizeOK(Userindex)
    End If
    
    ''
    'TODO : Feo, esto tiene que ser parche cliente
    If .flags.Estupidez = 0 Then
        Call WriteDumbNoMore(Userindex)
    End If
    
    'Posicion de comienzo
    If .Pos.map = 0 Then
        Select Case .Hogar
            Case eCiudad.cRamx
                .Pos = Ramx
        End Select
    Else
        If Not MapaValido(.Pos.map) Then
            Call WriteErrorMsg(Userindex, "EL PJ se encuenta en un mapa inválido.")
            Call FlushBuffer(Userindex)
            Call CloseSocket(Userindex)
            Exit Sub
        End If
    End If
    
    'Tratamos de evitar en lo posible el "Telefrag". Solo 1 intento de loguear en pos adjacentes.
    'Codigo por Pablo (ToxicWaste) y revisado por Nacho (Integer), corregido para que realmetne ande y no tire el server por Juan Martín Sotuyo Dodero (Maraxus)
    If MapData(.Pos.map, .Pos.X, .Pos.Y).Userindex <> 0 Or MapData(.Pos.map, .Pos.X, .Pos.Y).NpcIndex <> 0 Then
        Dim FoundPlace As Boolean
        Dim esAgua As Boolean
        Dim tX As Long
        Dim tY As Long
        
        FoundPlace = False
        esAgua = HayAgua(.Pos.map, .Pos.X, .Pos.Y)
        
        For tY = .Pos.Y - 1 To .Pos.Y + 1
            For tX = .Pos.X - 1 To .Pos.X + 1
                If esAgua Then
                    'reviso que sea pos legal en agua, que no haya User ni NPC para poder loguear.
                    If LegalPos(.Pos.map, tX, tY, True, False) Then
                        FoundPlace = True
                        Exit For
                    End If
                Else
                    'reviso que sea pos legal en tierra, que no haya User ni NPC para poder loguear.
                    If LegalPos(.Pos.map, tX, tY, False, True) Then
                        FoundPlace = True
                        Exit For
                    End If
                End If
            Next tX
            
            If FoundPlace Then _
                Exit For
        Next tY
        
        If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
            .Pos.X = tX
            .Pos.Y = tY
        Else
            'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
            If MapData(.Pos.map, .Pos.X, .Pos.Y).Userindex <> 0 Then
                'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
                If UserList(MapData(.Pos.map, .Pos.X, .Pos.Y).Userindex).ComUsu.DestUsu > 0 Then
                    'Le avisamos al que estaba comerciando que se tuvo que ir.
                    If UserList(UserList(MapData(.Pos.map, .Pos.X, .Pos.Y).Userindex).ComUsu.DestUsu).flags.UserLogged Then
                        Call FinComerciarUsu(UserList(MapData(.Pos.map, .Pos.X, .Pos.Y).Userindex).ComUsu.DestUsu)
                        Call WriteConsoleMsg(UserList(MapData(.Pos.map, .Pos.X, .Pos.Y).Userindex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                        Call FlushBuffer(UserList(MapData(.Pos.map, .Pos.X, .Pos.Y).Userindex).ComUsu.DestUsu)
                    End If
                    'Lo sacamos.
                    If UserList(MapData(.Pos.map, .Pos.X, .Pos.Y).Userindex).flags.UserLogged Then
                        Call FinComerciarUsu(MapData(.Pos.map, .Pos.X, .Pos.Y).Userindex)
                        Call WriteErrorMsg(MapData(.Pos.map, .Pos.X, .Pos.Y).Userindex, "Alguien se ha conectado donde te encontrabas, por favor reconéctate...")
                        Call FlushBuffer(MapData(.Pos.map, .Pos.X, .Pos.Y).Userindex)
                    End If
                End If
                
                Call CloseSocket(MapData(.Pos.map, .Pos.X, .Pos.Y).Userindex)
            End If
        End If
    End If
    
    'Nombre de sistema
    .Name = Name
    
    .showName = True 'Por default los nombres son visibles
    
    'If in the water, and has a boat, equip it!
    If .Invent.BarcoObjIndex > 0 And _
            (HayAgua(.Pos.map, .Pos.X, .Pos.Y) Or BodyIsBoat(.Char.Body)) Then
        Dim Barco As ObjData
        Barco = ObjData(.Invent.BarcoObjIndex)
        .Char.Head = 0
        If .flags.Muerto = 0 Then
    
            If .Faccion.ArmadaReal = 1 Then
                .Char.Body = iFragataReal
            ElseIf .Faccion.FuerzasCaos = 1 Then
                .Char.Body = iFragataCaos
            Else
                If criminal(Userindex) Then
                    If Barco.Ropaje = iBarca Then .Char.Body = iBarcaPk
                    If Barco.Ropaje = iGalera Then .Char.Body = iGaleraPk
                    If Barco.Ropaje = iGaleon Then .Char.Body = iGaleonPk
                Else
                    If Barco.Ropaje = iBarca Then .Char.Body = iBarcaCiuda
                    If Barco.Ropaje = iGalera Then .Char.Body = iGaleraCiuda
                    If Barco.Ropaje = iGaleon Then .Char.Body = iGaleonCiuda
                End If
            End If
        Else
            .Char.Body = iFragataFantasmal
        End If
        
        .Char.ShieldAnim = NingunEscudo
        .Char.WeaponAnim = NingunArma
        .Char.CascoAnim = NingunCasco
        .flags.Navegando = 1
    End If
    
    
    'Info
    Call WriteUserIndexInServer(Userindex) 'Enviamos el User index
    Call WriteChangeMap(Userindex, .Pos.map, MapInfo(.Pos.map).Zona)  'Carga el mapa
    Call WritePlayMidi(Userindex, val(ReadField(1, MapInfo(.Pos.map).Music, 45)))
    
    If .flags.Privilegios = PlayerType.Admin Then
        .flags.ChatColor = RGB(0, 255, 0)
    ElseIf .flags.Privilegios = PlayerType.Dios Then
        .flags.ChatColor = RGB(0, 255, 0)
    ElseIf .flags.Privilegios <> PlayerType.User And .flags.Privilegios <> (PlayerType.User Or PlayerType.ChaosCouncil) And .flags.Privilegios <> (PlayerType.User Or PlayerType.RoyalCouncil) Then
        .flags.ChatColor = RGB(0, 255, 0)
    ElseIf .flags.Privilegios = (PlayerType.User Or PlayerType.RoyalCouncil) Then
        .flags.ChatColor = RGB(0, 255, 255)
    ElseIf .flags.Privilegios = (PlayerType.User Or PlayerType.ChaosCouncil) Then
        .flags.ChatColor = RGB(255, 128, 64)
    Else
        .flags.ChatColor = vbWhite
    End If
    
    
    ''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
    #If ConUpTime Then
        .LogOnTime = Now
    #End If
    
    'Crea  el personaje del usuario
    Call MakeUserChar(True, .Pos.map, Userindex, .Pos.map, .Pos.X, .Pos.Y)
    
    Call WriteUserCharIndexInServer(Userindex)
    ''[/el oso]
        
    Call DoTileEvents(Userindex, .Pos.map, .Pos.X, .Pos.Y)
    
    Call CheckUserLevel(Userindex)
    Call WriteUpdateUserStats(Userindex)
    
    Call WriteUpdateHit(Userindex)
    Call WriteUpdateArmor(Userindex)
    Call WriteUpdateEscu(Userindex)
    Call WriteUpdateCasco(Userindex)
    
    Call WriteUpdateHungerAndThirst(Userindex)
    Call WriteUpdateStrenghtAndDexterity(Userindex)
    
    Call SendMOTD(Userindex)
    
    If haciendoBK Then
        Call WritePauseToggle(Userindex)
        Call WriteConsoleMsg(Userindex, "Servidor> Por favor espera algunos segundos, el WorldSave está ejecutándose.", FontTypeNames.FONTTYPE_SERVER)
    End If
    
    If EnPausa Then
        Call WritePauseToggle(Userindex)
        Call WriteConsoleMsg(Userindex, "Servidor> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde.", FontTypeNames.FONTTYPE_SERVER)
    End If
    
    If EnTesting And .Stats.ELV >= 18 Then
        Call WriteErrorMsg(Userindex, "Servidor en Testing por unos minutos, conectese con PJs de nivel menor a 18. No se conecte con Pjs que puedan resultar importantes por ahora pues pueden arruinarse.")
        Call FlushBuffer(Userindex)
        Call CloseSocket(Userindex)
        Exit Sub
    End If
    
    'Actualiza el Num de usuarios
    'DE ACA EN ADELANTE GRABA EL CHARFILE, OJO!
    NumUsers = NumUsers + 1
    .flags.UserLogged = True
    
    'usado para borrar Pjs
    Call WriteVar(CharPath & .Name & ".chr", "INIT", "Logged", "1")
    
    MapInfo(.Pos.map).NumUsers = MapInfo(.Pos.map).NumUsers + 1
    
    If .Stats.SkillPts > 0 Then
        Call WriteSendSkills(Userindex)
        Call WriteLevelUp(Userindex, .Stats.SkillPts)
    End If
    
    If NumUsers > recordusuarios Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Record de usuarios conectados simultaneamente." & "Hay " & NumUsers & " usuarios.", FontTypeNames.FONTTYPE_INFO))
        recordusuarios = NumUsers
        Call WriteVar(IniPath & "Server.ini", "INIT", "Record", str(recordusuarios))
    End If
    
    If .NroMascotas > 0 And MapInfo(.Pos.map).Pk Then
        Dim i As Integer
        For i = 1 To MAXMASCOTAS
            If .MascotasType(i) > 0 Then
                .MascotasIndex(i) = SpawnNpc(.MascotasType(i), .Pos, True, True)
                
                If .MascotasIndex(i) > 0 Then
                    Npclist(.MascotasIndex(i)).MaestroUser = Userindex
                    Call FollowAmo(.MascotasIndex(i))
                Else
                    .MascotasIndex(i) = 0
                End If
            End If
        Next i
    End If
    
    If .flags.Navegando = 1 Then
        Call WriteNavigateToggle(Userindex)
    End If
    
    If .flags.Equitando = 1 Then
        Call WriteEquitandoToggle(Userindex)
    End If
    
    If .GuildIndex > 0 Then
        'welcome to the show baby...
        If Not modGuilds.m_ConectarMiembroAClan(Userindex, .GuildIndex) Then
            Call WriteConsoleMsg(Userindex, "Tu estado no te permite entrar al clan.", FontTypeNames.FONTTYPE_GUILD)
        End If
    End If
    
    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))
    
    Call WriteLoggedMessage(Userindex)
    
    'Inicia en Modo Combate [MaxTus]
    .flags.ModoCombate = True
    Call WriteConsoleMsg(Userindex, "MODO COMBATE activado.", FontTypeNames.FONTTYPE_FIGHT)
    
    Call modGuilds.SendGuildNews(Userindex)
    
    '**************Lorwik/Noche**************
    '***********www.lwk-foros.net***********
    If Anocheceria = 0 Then
        Call writeNoche(Userindex, 0)
    ElseIf Anocheceria = 1 Then
        Call writeNoche(Userindex, 1)
    ElseIf Anocheceria = 2 Then
        Call writeNoche(Userindex, 2)
    ElseIf Anocheceria = 3 Then
        Call writeNoche(Userindex, 3)
    End If
    '******************************************
    
    tStr = modGuilds.a_ObtenerRechazoDeChar(.Name)
    
    If LenB(tStr) <> 0 Then
        Call WriteShowMessageBox(Userindex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr)
    End If
    
    'Load the user statistics
    Call Statistics.UserConnected(Userindex)
    
    Call MostrarNumUsers
    
    N = FreeFile
    Open App.Path & "\logs\numusers.log" For Output As N
    Print #N, NumUsers
    Close #N
    
    N = FreeFile
    'Log
    Open App.Path & "\logs\Connect.log" For Append Shared As #N
    Print #N, .Name & " ha entrado al juego. UserIndex:" & Userindex & " " & Time & " " & Date
    Close #N
    
    End With
End Sub

Sub SendMOTD(ByVal Userindex As Integer)
    Dim j As Long
    
    Call WriteConsoleMsg(Userindex, "Mensajes de entrada:", FontTypeNames.FONTTYPE_INFO)
    For j = 1 To MaxLines
        Call WriteConsoleMsg(Userindex, MOTD(j).texto, FontTypeNames.FONTTYPE_INFO)
    Next j
    
    Call WriteConsoleMsg(Userindex, "El Castillo Winter esta en manos del clan: " & GetVar(App.Path & "\Castillos.ini", "CLANES", "WINTER"), FontTypeNames.FONTTYPE_INFO)
End Sub

Sub ResetFacciones(ByVal Userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
'*************************************************
    With UserList(Userindex).Faccion
        .ArmadaReal = 0
        .CiudadanosMatados = 0
        .CriminalesMatados = 0
        .FuerzasCaos = 0
        .FechaIngreso = "No ingresó a ninguna Facción"
        .RecibioArmaduraCaos = 0
        .RecibioArmaduraReal = 0
        .RecibioExpInicialCaos = 0
        .RecibioExpInicialReal = 0
        .RecompensasCaos = 0
        .RecompensasReal = 0
        .Reenlistadas = 0
        .NivelIngreso = 0
        .MatadosIngreso = 0
        .NextRecompensa = 0
    End With
End Sub

Sub ResetContadores(ByVal Userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'05/20/2007 Integer - Agregue todas las variables que faltaban.
'*************************************************
    With UserList(Userindex).Counters
        .AGUACounter = 0
        .AttackCounter = 0
        .Ceguera = 0
        .COMCounter = 0
        .Estupidez = 0
        .Frio = 0
        .HPCounter = 0
        .IdleCount = 0
        .Invisibilidad = 0
        .Paralisis = 0
        .Pena = 0
        .Piedra = 0
        .PiqueteC = 0
        .STACounter = 0
        .Veneno = 0
        .Trabajando = 0
        .Ocultando = 0
        .bPuedeMeditar = False
        .Lava = 0
        .Mimetismo = 0
        .Saliendo = False
        .Salir = 0
        .TiempoOculto = 0
        .TimerMagiaGolpe = 0
        .TimerGolpeMagia = 0
        .TimerLanzarSpell = 0
        .TimerPuedeAtacar = 0
        .TimerPuedeUsarArco = 0
        .TimerPuedeTrabajar = 0
        .TimerUsar = 0
        .Makro = 0
        .Resucitar = 0
    End With
End Sub

Sub ResetCharInfo(ByVal Userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(Userindex).Char
        .Body = 0
        .CascoAnim = 0
        .CharIndex = 0
        .FX = 0
        .Head = 0
        .loops = 0
        .heading = 0
        .loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
        .Aura = 0
    End With
End Sub

Sub ResetBasicUserInfo(ByVal Userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(Userindex)
        .Name = vbNullString
        .desc = vbNullString
        .DescRM = vbNullString
        .Pos.map = 0
        .Pos.X = 0
        .Pos.Y = 0
        .ip = vbNullString
        .clase = 0
        .email = vbNullString
        .genero = 0
        .Hogar = 0
        .raza = 0
        
        .PartyIndex = 0
        .PartySolicitud = 0
        
        With .Stats
            .Banco = 0
            .ELV = 0
            .ELU = 0
            .Exp = 0
            .def = 0
            '.CriminalesMatados = 0
            .NPCsMuertos = 0
            .UsuariosMatados = 0
            .SkillPts = 0
            .GLD = 0
            .UserAtributos(1) = 0
            .UserAtributos(2) = 0
            .UserAtributos(3) = 0
            .UserAtributos(4) = 0
            .UserAtributos(5) = 0
            .UserAtributosBackUP(1) = 0
            .UserAtributosBackUP(2) = 0
            .UserAtributosBackUP(3) = 0
            .UserAtributosBackUP(4) = 0
            .UserAtributosBackUP(5) = 0
        End With
        
    End With
End Sub

Sub ResetReputacion(ByVal Userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(Userindex).Reputacion
        .AsesinoRep = 0
        .BandidoRep = 0
        .BurguesRep = 0
        .LadronesRep = 0
        .NobleRep = 0
        .PlebeRep = 0
        .NobleRep = 0
        .Promedio = 0
    End With
End Sub

Sub ResetGuildInfo(ByVal Userindex As Integer)
    If UserList(Userindex).EscucheClan > 0 Then
        Call modGuilds.GMDejaDeEscucharClan(Userindex, UserList(Userindex).EscucheClan)
        UserList(Userindex).EscucheClan = 0
    End If
    If UserList(Userindex).GuildIndex > 0 Then
        Call modGuilds.m_DesconectarMiembroDelClan(Userindex, UserList(Userindex).GuildIndex)
    End If
    UserList(Userindex).GuildIndex = 0
End Sub

Sub ResetUserFlags(ByVal Userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 06/28/2008
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'03/29/2006 Maraxus - Reseteo el CentinelaOK también.
'06/28/2008 NicoNZ - Agrego el flag Inmovilizado
'*************************************************
    With UserList(Userindex).flags
        .Comerciando = False
        .Ban = 0
        .Escondido = 0
        .DuracionEfecto = 0
        .NpcInv = 0
        .StatsChanged = 0
        .TargetNPC = 0
        .TargetNpcTipo = eNPCType.Comun
        .TargetObj = 0
        .TargetObjMap = 0
        .TargetObjX = 0
        .TargetObjY = 0
        .TargetUser = 0
        .TipoPocion = 0
        .TomoPocion = False
        .Descuento = vbNullString
        .Hambre = 0
        .Sed = 0
        .Descansar = False
        .ModoCombate = False
        .Vuela = 0
        .Navegando = 0
        .Equitando = 0
        .Oculto = 0
        .Envenenado = 0
        .invisible = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Metamorfosis = 0
        .Maldicion = 0
        .Bendicion = 0
        .Meditando = 0
        .Privilegios = 0
        .PuedeMoverse = 0
        .OldBody = 0
        .OldHead = 0
        .AdminInvisible = 0
        .ValCoDe = 0
        .Hechizo = 0
        .TimesWalk = 0
        .StartWalk = 0
        .CountSH = 0
        .Silenciado = 0
        .CentinelaOK = False
        .AdminPerseguible = False
        .Anomalia = 0
        .Resucitando = 0
    End With
End Sub

Sub ResetUserSpells(ByVal Userindex As Integer)
    Dim LoopC As Long
    For LoopC = 1 To MAXUSERHECHIZOS
        UserList(Userindex).Stats.UserHechizos(LoopC) = 0
    Next LoopC
End Sub

Sub ResetUserPets(ByVal Userindex As Integer)
    Dim LoopC As Long
    
    UserList(Userindex).NroMascotas = 0
        
    For LoopC = 1 To MAXMASCOTAS
        UserList(Userindex).MascotasIndex(LoopC) = 0
        UserList(Userindex).MascotasType(LoopC) = 0
    Next LoopC
End Sub

Sub ResetUserBanco(ByVal Userindex As Integer)
    Dim LoopC As Long
    
    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
          UserList(Userindex).BancoInvent.Object(LoopC).amount = 0
          UserList(Userindex).BancoInvent.Object(LoopC).Equipped = 0
          UserList(Userindex).BancoInvent.Object(LoopC).ObjIndex = 0
    Next LoopC
    
    UserList(Userindex).BancoInvent.NroItems = 0
End Sub

Public Sub LimpiarComercioSeguro(ByVal Userindex As Integer)
    With UserList(Userindex).ComUsu
        If .DestUsu > 0 Then
            Call FinComerciarUsu(.DestUsu)
            Call FinComerciarUsu(Userindex)
        End If
    End With
End Sub

Sub ResetUserSlot(ByVal Userindex As Integer)

UserList(Userindex).ConnIDValida = False
UserList(Userindex).ConnID = -1

Call LimpiarComercioSeguro(Userindex)
Call ResetFacciones(Userindex)
Call ResetContadores(Userindex)
Call ResetCharInfo(Userindex)
Call ResetBasicUserInfo(Userindex)
Call ResetReputacion(Userindex)
Call ResetGuildInfo(Userindex)
Call ResetUserFlags(Userindex)
Call LimpiarInventario(Userindex)
Call ResetUserSpells(Userindex)
Call ResetUserPets(Userindex)
Call ResetUserBanco(Userindex)
Call ResetQuestStats(Userindex)
With UserList(Userindex).ComUsu
    .Acepto = False
    .cant = 0
    .DestNick = vbNullString
    .DestUsu = 0
    .Objeto = 0
End With

End Sub

Sub CloseUser(ByVal Userindex As Integer)
'Call LogTarea("CloseUser " & UserIndex)
On Error GoTo Errhandler

Dim N As Integer
Dim LoopC As Integer
Dim map As Integer
Dim Name As String
Dim i As Integer

Dim aN As Integer

aN = UserList(Userindex).flags.AtacadoPorNpc
If aN > 0 Then
      Npclist(aN).Movement = Npclist(aN).flags.OldMovement
      Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
      Npclist(aN).flags.AttackedBy = vbNullString
End If
aN = UserList(Userindex).flags.NPCAtacado
If aN > 0 Then
    If Npclist(aN).flags.AttackedFirstBy = UserList(Userindex).Name Then
        Npclist(aN).flags.AttackedFirstBy = vbNullString
    End If
End If
UserList(Userindex).flags.AtacadoPorNpc = 0
UserList(Userindex).flags.NPCAtacado = 0

map = UserList(Userindex).Pos.map
Name = UCase$(UserList(Userindex).Name)

Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateParticle(UserList(Userindex).Char.CharIndex, UserList(Userindex).Char.Particle, 1))
UserList(Userindex).Char.Particle = 0
UserList(Userindex).Char.loops = 0

UserList(Userindex).flags.UserLogged = False
UserList(Userindex).Counters.Saliendo = False

'Le devolvemos el body y head originales
If UserList(Userindex).flags.AdminInvisible = 1 Then Call DoAdminInvisible(Userindex)

'si esta en party le devolvemos la experiencia
If UserList(Userindex).PartyIndex > 0 Then Call mdParty.SalirDeParty(Userindex)

'Save statistics
Call Statistics.UserDisconnected(Userindex)

' Grabamos el personaje del usuario
Call SaveUser(Userindex, CharPath & Name & ".chr")

'Actualizamos el Ranking
Call ActualizarRanking(Userindex)

'usado para borrar Pjs
Call WriteVar(CharPath & UserList(Userindex).Name & ".chr", "INIT", "Logged", "0")


'Quitar el dialogo
'If MapInfo(Map).NumUsers > 0 Then
'    Call SendToUserArea(UserIndex, "QDL" & UserList(UserIndex).Char.charindex)
'End If

If MapInfo(map).NumUsers > 0 Then
    Call SendData(SendTarget.ToPCAreaButIndex, Userindex, PrepareMessageRemoveCharDialog(UserList(Userindex).Char.CharIndex))
End If

'Borrar el personaje
If UserList(Userindex).Char.CharIndex > 0 Then
    Call EraseUserChar(Userindex, UserList(Userindex).flags.AdminInvisible = 1)
End If

'Borrar mascotas
For i = 1 To MAXMASCOTAS
    If UserList(Userindex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(Userindex).MascotasIndex(i)).flags.NPCActive Then _
            Call QuitarNPC(UserList(Userindex).MascotasIndex(i))
    End If
Next i

'Update Map Users
MapInfo(map).NumUsers = MapInfo(map).NumUsers - 1

If MapInfo(map).NumUsers < 0 Then
    MapInfo(map).NumUsers = 0
End If

' Si el usuario habia dejado un msg en la gm's queue lo borramos
If Ayuda.Existe(UserList(Userindex).Name) Then Call Ayuda.Quitar(UserList(Userindex).Name)

Call ResetUserSlot(Userindex)

Call MostrarNumUsers

N = FreeFile(1)
Open App.Path & "\logs\Connect.log" For Append Shared As #N
Print #N, Name & " ha dejado el juego. " & "User Index:" & Userindex & " " & Time & " " & Date
Close #N

Exit Sub

Errhandler:
Call LogError("Error en CloseUser. Número " & Err.Number & " Descripción: " & Err.description)

End Sub

Sub ReloadSokcet()
On Error GoTo Errhandler
#If UsarQueSocket = 1 Then

    Call LogApiSock("ReloadSokcet() " & NumUsers & " " & LastUser & " " & MaxUsers)
    
    If NumUsers <= 0 Then
        Call WSApiReiniciarSockets
    Else
'       Call apiclosesocket(SockListen)
'       SockListen = ListenForConnect(Puerto, hWndMsg, "")
    End If

#ElseIf UsarQueSocket = 0 Then

    frmMain.Socket1.Cleanup
    Call ConfigListeningSocket(frmMain.Socket1, Puerto)
    
#ElseIf UsarQueSocket = 2 Then

    

#End If

Exit Sub
Errhandler:
    Call LogError("Error en CheckSocketState " & Err.Number & ": " & Err.description)

End Sub

Public Sub EcharPjsNoPrivilegiados()
Dim LoopC As Long

For LoopC = 1 To LastUser
    If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
        If UserList(LoopC).flags.Privilegios And PlayerType.User Then
            Call CloseSocket(LoopC)
        End If
    End If
Next LoopC

End Sub

Private Sub ActualizarRanking(ByVal Userindex As Integer)

    With UserList(Userindex)
        
        'Actualizamos el Rank del nivel
        If .Stats.ELV >= val(GetVar(App.Path & "\rank.dat", "MaxNivel", "Nivel")) Then
            Call WriteVar(App.Path & "\rank.dat", "MaxNivel", "Nivel", .Stats.ELV)
            Call WriteVar(App.Path & "\rank.dat", "MaxNivel", "NivelName", .Name)
        End If
        'Actualizamos el Rank del oro
        If .Stats.GLD + .Stats.Banco >= val(GetVar(App.Path & "\rank.dat", "MaxOro", "Oro")) Then
            Call WriteVar(App.Path & "\rank.dat", "MaxOro", "Oro", .Stats.GLD + .Stats.Banco)
            Call WriteVar(App.Path & "\rank.dat", "MaxOro", "NombreOro", .Name)
        End If
        'Actualizamos el Rank de las muertes
        If .Faccion.CriminalesMatados + .Faccion.CiudadanosMatados >= val(GetVar(App.Path & "\rank.dat", "MaxDead", "Muertes")) Then
            Call WriteVar(App.Path & "\rank.dat", "MaxDead", "Muertes", .Faccion.CriminalesMatados + .Faccion.CiudadanosMatados)
            Call WriteVar(App.Path & "\rank.dat", "MaxDead", "NombreDead", .Name)
        End If
        
        '*************TOP 5*****************
        If .Faccion.CriminalesMatados + .Faccion.CiudadanosMatados >= val(GetVar(App.Path & "\rank.dat", "Top5", "Muertes1")) Then
            Call WriteVar(App.Path & "\rank.dat", "Top5", "Nivel1", .Stats.ELV)
            Call WriteVar(App.Path & "\rank.dat", "Top5", "NivelName1", .Name)
            Call WriteVar(App.Path & "\rank.dat", "Top5", "Muertes1", .Faccion.CriminalesMatados + .Faccion.CiudadanosMatados)
            
        ElseIf .Faccion.CriminalesMatados + .Faccion.CiudadanosMatados >= val(GetVar(App.Path & "\rank.dat", "Top5", "Muertes2")) Then
            Call WriteVar(App.Path & "\rank.dat", "Top5", "Nivel2", .Stats.ELV)
            Call WriteVar(App.Path & "\rank.dat", "Top5", "NivelName2", .Name)
            Call WriteVar(App.Path & "\rank.dat", "Top5", "Muertes2", .Faccion.CriminalesMatados + .Faccion.CiudadanosMatados)
            
        ElseIf .Faccion.CriminalesMatados + .Faccion.CiudadanosMatados >= val(GetVar(App.Path & "\rank.dat", "Top5", "Muertes3")) Then
            Call WriteVar(App.Path & "\rank.dat", "Top5", "Nivel3", .Stats.ELV)
            Call WriteVar(App.Path & "\rank.dat", "Top5", "NivelName3", .Name)
            Call WriteVar(App.Path & "\rank.dat", "Top5", "Muertes3", .Faccion.CriminalesMatados + .Faccion.CiudadanosMatados)
            
        ElseIf .Faccion.CriminalesMatados + .Faccion.CiudadanosMatados >= val(GetVar(App.Path & "\rank.dat", "Top5", "Muertes4")) Then
            Call WriteVar(App.Path & "\rank.dat", "Top5", "Nivel4", .Stats.ELV)
            Call WriteVar(App.Path & "\rank.dat", "Top5", "NivelName4", .Name)
            Call WriteVar(App.Path & "\rank.dat", "Top5", "Muertes4", .Faccion.CriminalesMatados + .Faccion.CiudadanosMatados)
            
        ElseIf .Faccion.CriminalesMatados + .Faccion.CiudadanosMatados >= val(GetVar(App.Path & "\rank.dat", "Top5", "Muertes5")) Then
            Call WriteVar(App.Path & "\rank.dat", "Top5", "Nivel5", .Stats.ELV)
            Call WriteVar(App.Path & "\rank.dat", "Top5", "NivelName5", .Name)
            Call WriteVar(App.Path & "\rank.dat", "Top5", "Muertes5", .Faccion.CriminalesMatados + .Faccion.CiudadanosMatados)
        
        End If
    End With
End Sub
