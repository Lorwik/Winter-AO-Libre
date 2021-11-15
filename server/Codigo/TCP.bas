Attribute VB_Name = "TCP"
Option Explicit

'RUTAS DE ENVIO DE DATOS
Public Enum SendTarget
    ToIndex = 0         'Envia a un solo User
    toall = 1           'A todos los Users
    ToMap = 2           'Todos los Usuarios en el mapa
    ToPCArea = 3        'Todos los Users en el area de un user determinado
    ToNone = 4          'Ninguno
    ToAllButIndex = 5   'Todos menos el index
    ToMapButIndex = 6   'Todos en el mapa menos el indice
    ToGM = 7
    ToNPCArea = 8       'Todos los Users en el area de un user determinado
    ToGuildMembers = 9
    ToAdmins = 10
    ToPCAreaButIndex = 11
    ToAdminsAreaButConsejeros = 12
    ToDiosesYclan = 13
    ToConsejo = 14
    ToClanArea = 15
    ToConsejoCaos = 16
    ToRolesMasters = 17
    ToDeadArea = 18
    ToCiudadanos = 19
    ToCriminales = 20
    ToPartyArea = 21
    ToReal = 22
    ToCaos = 23
    ToCiudadanosYRMs = 24
    ToCriminalesYRMs = 25
    ToRealYRMs = 26
    ToCaosYRMs = 27
    tosubasta = 28
End Enum


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

Sub DarCuerpoYCabeza(ByRef UserBody As Integer, ByRef UserHead As Integer, ByVal Raza As String, ByVal Gen As String)
'TODO: Poner las heads en arrays, así se acceden por índices
'y no hay problemas de discontinuidad de los índices.
'También se debe usar enums para raza y sexo
Select Case Gen
   Case "Hombre"
        Select Case Raza
            Case "Humano"
                UserHead = RandomNumber(1, 30)
                UserBody = 1
            Case "Elfo"
                UserHead = RandomNumber(1, 13) + 100
                If UserHead = 113 Then UserHead = 201       'Un índice no es continuo.... :S muy feo
                UserBody = 2
            Case "Elfo Oscuro"
                UserHead = RandomNumber(1, 8) + 201
                UserBody = 3
            Case "Enano"
                UserHead = RandomNumber(1, 5) + 300
                UserBody = 52
            Case "ORCO"
                UserHead = RandomNumber(411, 417) + 300
                UserBody = 434
            Case "Gnomo"
                UserHead = RandomNumber(1, 6) + 400
                UserBody = 52
            Case Else
                UserHead = 1
                UserBody = 1
        End Select
   Case "Mujer"
        Select Case Raza
            Case "Humano"
                UserHead = RandomNumber(1, 7) + 69
                UserBody = 1
            Case "Elfo"
                UserHead = RandomNumber(1, 7) + 169
                UserBody = 2
            Case "Elfo Oscuro"
                UserHead = RandomNumber(1, 11) + 269
                UserBody = 3
            Case "Gnomo"
                UserHead = RandomNumber(1, 5) + 469
                UserBody = 52
            Case "Enano"
                UserHead = RandomNumber(1, 3) + 369
                UserBody = 52
            Case "ORCO"
                UserHead = RandomNumber(418, 422) + 369
                UserBody = 432
            Case Else
                UserHead = 70
                UserBody = 1
        End Select
End Select

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

Function ValidateSkills(ByVal userindex As Integer) As Boolean

Dim LoopC As Integer

For LoopC = 1 To NUMSKILLS
    If UserList(userindex).Stats.UserSkills(LoopC) < 0 Then
        Exit Function
        If UserList(userindex).Stats.UserSkills(LoopC) > 100 Then UserList(userindex).Stats.UserSkills(LoopC) = 100
    End If
Next LoopC

ValidateSkills = True
    
End Function
Public Function IsYourChr(ByVal Account As String, ByVal PJ As String)

Dim i As Integer
Dim NumPjs As Integer
Dim ChrToView As String



NumPjs = GetVar(App.Path & "\Accounts\" & Account & ".act", "PJS", "NumPjs")

IsYourChr = False

For i = 0 To NumPjs
    ChrToView = GetVar(App.Path & "\Accounts\" & Account & ".act", "PJS", "PJ" & i)
    If ChrToView = PJ Then IsYourChr = True
Next i

End Function

Sub ConnectAccount(ByVal userindex As Integer, name As String, Password As String)

Dim i As Integer
Dim Pjjj As String
Dim NumPjs As Integer
Dim ArchivodeUser As String
Dim pos() As String
Dim Oro() As Long
Dim Nivel() As String
Dim PuntosdeCanje() As Integer
Dim OroBanco() As Byte
Dim cosa As Integer



If UCase$(Password) <> UCase$(GetVar(App.Path & "\Accounts\" & name & ".act", name, "password")) Then
    Call SendData(SendTarget.ToIndex, userindex, 0, "ERRPassword incorrecto.")
    Call CloseSocket(userindex)
    Exit Sub
End If

UserList(userindex).Accounted = name
UserList(userindex).AccountedPass = UCase$(Password)

NumPjs = GetVar(App.Path & "\Accounts\" & name & ".act", "PJS", "NumPjs")

If TienePjs(name) = True Then
    Call SendData(SendTarget.ToIndex, userindex, 0, "INIAC" & name & "," & NumPjs + 1)
Else
    Call SendData(SendTarget.ToIndex, userindex, 0, "INIAC0")
End If
ArchivodeUser = App.Path & "\charfile\"
For i = 0 To NumPjs
    Pjjj = GetVar(App.Path & "\Accounts\" & name & ".act", "PJS", "PJ" & i)
    If Pjjj = "" Then Exit Sub
    Call LoadUserAccount(Pjjj & ".chr")
    Call SendData(SendTarget.ToIndex, userindex, 0, "ADDPJ" & Pjjj & "," & i & "," & PJEnCuenta & PJEnCuentaB)
Next i
End Sub
Sub ChrToAccount(ByVal Accounted As String, tName As String)

Dim NumPjs As Integer
Dim N As Integer

NumPjs = GetVar(App.Path & "\Accounts\" & Accounted & ".act", "PJS", "NumPjs")

If NumPjs = 0 And GetVar(App.Path & "\Accounts\" & Accounted & ".act", "PJS", "PJ" & NumPjs) = "" Then
    Call WriteVar(App.Path & "\Accounts\" & Accounted & ".act", "PJS", "NumPjs", NumPjs)
    Call WriteVar(App.Path & "\Accounts\" & Accounted & ".act", "PJS", "PJ" & NumPjs, tName)
    Exit Sub
End If

NumPjs = NumPjs + 1

Call WriteVar(App.Path & "\Accounts\" & Accounted & ".act", "PJS", "NumPjs", NumPjs)
Call WriteVar(App.Path & "\Accounts\" & Accounted & ".act", "PJS", "PJ" & NumPjs, tName)


End Sub
Sub CreateAccount(ByVal Account As String, Password As String, Mail As String, pregunta As String, Respuesta As String, userindex As Integer)

On Error GoTo errhandler

If FileExist(App.Path & "\Accounts\" & Account & ".act", vbNormal) = True Then
Call SendData(SendTarget.ToIndex, userindex, 0, "ERREl nombre de la cuenta ya está siendo utilizado por otro usuario.")
    Exit Sub
End If

Dim N As Integer
Dim i As Integer


N = FreeFile()

Open App.Path & "\Accounts\" & Account & ".act" For Output As N
    Print #N, "[" & Account & "]"
    Print #N, "password=" & Password
    Print #N, "mail=" & Mail
    Print #N, "Pregunta=" & pregunta
    Print #N, "Respuesta=" & Respuesta
    Print #N, "ban=0"
    Print #N, "[PJS]"
    Print #N, "NumPjs=0"
    Print #N, "PJ0="
    Print #N, "PJ1="
    Print #N, "PJ2="
    Print #N, "PJ3="
    Print #N, "PJ4="
    Print #N, "PJ5="
    Print #N, "PJ6="
    Print #N, "PJ7="
Close N

DoEvents

Call CloseSocket(userindex)

Call SendData(SendTarget.ToIndex, userindex, 0, "HLQ")

Exit Sub

errhandler:

Call LogError("NewAccount - Error = " & Err.Number & " - Descripción = " & Err.Description)

End Sub

Public Function TienePjs(ByVal Account As String) As Boolean

Dim frstPj As String

frstPj = GetVar(App.Path & "\Accounts\" & Account & ".act", "PJS", "PJ0")

If frstPj <> "" Then
    TienePjs = True
Else
    TienePjs = False
End If

End Function

'Barrin 3/3/03
'Agregué PadrinoName y Padrino password como opcionales, que se les da un valor siempre y cuando el servidor esté usando el sistema
Sub ConnectNewUser(userindex As Integer, name As String, UserRaza As String, UserSexo As String, UserClase As String, Hogar As String, _
                    US1 As String, US2 As String, US3 As String, US4 As String, US5 As String, _
                    US6 As String, US7 As String, US8 As String, US9 As String, US10 As String, _
                    US11 As String, US12 As String, US13 As String, US14 As String, US15 As String, _
                    US16 As String, US17 As String, US18 As String, US19 As String, US20 As String, _
                    US21 As String, US22 As String, Head As Integer)

If Not AsciiValidos(name) Then
    Call SendData(SendTarget.ToIndex, userindex, 0, "ERRNombre invalido.")
    Exit Sub
End If

Dim LoopC As Integer
Dim totalskpts As Long

'¿Existe el personaje?
If FileExist(CharPath & UCase$(name) & ".chr", vbNormal) = True Then
    Call SendData(SendTarget.ToIndex, userindex, 0, "ERRYa existe el personaje.")
    Exit Sub
End If

'Tiró los dados antes de llegar acá??
If UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = 0 Then
    Call SendData(SendTarget.ToIndex, userindex, 0, "ERRDebe tirar los dados antes de poder crear un personaje.")
    Exit Sub
End If

UserList(userindex).flags.Muerto = 0
UserList(userindex).flags.Escondido = 0



UserList(userindex).Reputacion.AsesinoRep = 0
UserList(userindex).Reputacion.BandidoRep = 0
UserList(userindex).Reputacion.BurguesRep = 0
UserList(userindex).Reputacion.LadronesRep = 0
UserList(userindex).Reputacion.NobleRep = 1000
UserList(userindex).Reputacion.PlebeRep = 30

UserList(userindex).Reputacion.Promedio = 30 / 6


UserList(userindex).name = name
UserList(userindex).Clase = UserClase
UserList(userindex).Raza = UserRaza
UserList(userindex).Genero = UserSexo
UserList(userindex).Hogar = Hogar

Select Case UCase$(UserRaza)
    Case "HUMANO"
        UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) + 2
        UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) + 1
        UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) = UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) + 2
    Case "ELFO"
        UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) - 1
        UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) + 2
        UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) + 3
        UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) = UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) + 2
    Case "ELFO OSCURO"
        UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) + 1
        UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) + 1
        UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) + 2
        UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) = UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) + 1
        UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) = UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) + 1
    Case "ENANO"
        UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) + 3
        UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) = UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) + 4
        UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) - 7
        UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) - 1
        UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) = UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) - 1
    Case "ORCO"
        UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) + 5
        UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) = UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) + 3
        UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) - 6
        UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) - 2
        UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) = UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) - 2
    Case "GNOMO"
        UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) - 4
        UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) + 4
        UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) + 3
        UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) = UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) + 1
        UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) = UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) - 1
End Select



UserList(userindex).Stats.UserSkills(1) = val(US1)
UserList(userindex).Stats.UserSkills(2) = val(US2)
UserList(userindex).Stats.UserSkills(3) = val(US3)
UserList(userindex).Stats.UserSkills(4) = val(US4)
UserList(userindex).Stats.UserSkills(5) = val(US5)
UserList(userindex).Stats.UserSkills(6) = val(US6)
UserList(userindex).Stats.UserSkills(7) = val(US7)
UserList(userindex).Stats.UserSkills(8) = val(US8)
UserList(userindex).Stats.UserSkills(9) = val(US9)
UserList(userindex).Stats.UserSkills(10) = val(US10)
UserList(userindex).Stats.UserSkills(11) = val(US11)
UserList(userindex).Stats.UserSkills(12) = val(US12)
UserList(userindex).Stats.UserSkills(13) = val(US13)
UserList(userindex).Stats.UserSkills(14) = val(US14)
UserList(userindex).Stats.UserSkills(15) = val(US15)
UserList(userindex).Stats.UserSkills(16) = val(US16)
UserList(userindex).Stats.UserSkills(17) = val(US17)
UserList(userindex).Stats.UserSkills(18) = val(US18)
UserList(userindex).Stats.UserSkills(19) = val(US19)
UserList(userindex).Stats.UserSkills(20) = val(US20)
UserList(userindex).Stats.UserSkills(21) = val(US21)
UserList(userindex).Stats.UserSkills(22) = val(US22)

totalskpts = 0

'Abs PREVINENE EL HACKEO DE LOS SKILLS %%%%%%%%%%%%%
For LoopC = 1 To NUMSKILLS
    totalskpts = totalskpts + Abs(UserList(userindex).Stats.UserSkills(LoopC))
Next LoopC



If totalskpts > 10 Then
    Call LogHackAttemp(UserList(userindex).name & " intento hackear los skills.")
    Call BorrarUsuario(UserList(userindex).name)
    Call CloseSocket(userindex)
    Exit Sub
End If
'%%%%%%%%%%%%% PREVENIR HACKEO DE LOS SKILLS %%%%%%%%%%%%%

UserList(userindex).Char.Heading = eHeading.SOUTH

Call DarCuerpoYCabeza(UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Raza, UserList(userindex).Genero)
UserList(userindex).Char.Head = Head
UserList(userindex).OrigChar = UserList(userindex).Char
   
 
UserList(userindex).Char.WeaponAnim = NingunArma
UserList(userindex).Char.ShieldAnim = NingunEscudo
UserList(userindex).Char.CascoAnim = NingunCasco

UserList(userindex).Stats.MET = 1
Dim MiInt As Long
MiInt = RandomNumber(1, UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) \ 3)

UserList(userindex).Stats.MaxHP = 15 + MiInt
UserList(userindex).Stats.MinHP = 15 + MiInt

UserList(userindex).Stats.FIT = 1


MiInt = RandomNumber(1, UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) \ 6)
If MiInt = 1 Then MiInt = 2

UserList(userindex).Stats.MaxSta = 20 * MiInt
UserList(userindex).Stats.MinSta = 20 * MiInt


UserList(userindex).Stats.MaxAGU = 100
UserList(userindex).Stats.MinAGU = 100

UserList(userindex).Stats.MaxHam = 100
UserList(userindex).Stats.MinHam = 100


'<-----------------MANA----------------------->
If UCase$(UserClase) = "MAGO" Or UCase$(UserClase) = "NIGROMANTE" Then
    MiInt = RandomNumber(1, UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)) / 3
    UserList(userindex).Stats.MaxMAN = 100 + MiInt
    UserList(userindex).Stats.MinMAN = 100 + MiInt
ElseIf UCase$(UserClase) = "CLERIGO" Or UCase$(UserClase) = "DRUIDA" _
    Or UCase$(UserClase) = "BARDO" Or UCase$(UserClase) = "ASESINO" Then
        MiInt = RandomNumber(1, UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)) / 4
        UserList(userindex).Stats.MaxMAN = 50
        UserList(userindex).Stats.MinMAN = 50
Else
    UserList(userindex).Stats.MaxMAN = 0
    UserList(userindex).Stats.MinMAN = 0
End If

If UCase$(UserClase) = "MAGO" Or UCase$(UserClase) = "NIGROMANTE" Or UCase$(UserClase) = "CLERIGO" Or _
   UCase$(UserClase) = "DRUIDA" Or UCase$(UserClase) = "BARDO" Or _
   UCase$(UserClase) = "ASESINO" Then
        UserList(userindex).Stats.UserHechizos(1) = 2
End If

UserList(userindex).Stats.MaxHIT = 2
UserList(userindex).Stats.MinHIT = 1

UserList(userindex).Stats.GLD = 0




UserList(userindex).Stats.Exp = 0
UserList(userindex).Stats.ELU = 300
UserList(userindex).Stats.ELV = 1

'???????????????? INVENTARIO ¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿
UserList(userindex).Invent.NroItems = 5

UserList(userindex).Invent.Object(1).ObjIndex = 467
UserList(userindex).Invent.Object(1).Amount = 100

UserList(userindex).Invent.Object(2).ObjIndex = 468
UserList(userindex).Invent.Object(2).Amount = 100

UserList(userindex).Invent.Object(3).ObjIndex = 460
UserList(userindex).Invent.Object(3).Amount = 1
UserList(userindex).Invent.Object(3).Equipped = 1

Select Case UserRaza
    Case "Humano"
        UserList(userindex).Invent.Object(4).ObjIndex = 463
    Case "Elfo"
        UserList(userindex).Invent.Object(4).ObjIndex = 464
    Case "Elfo Oscuro"
        UserList(userindex).Invent.Object(4).ObjIndex = 465
    Case "Enano"
        UserList(userindex).Invent.Object(4).ObjIndex = 466
    Case "Gnomo"
        UserList(userindex).Invent.Object(4).ObjIndex = 466
    Case "Orco"
        UserList(userindex).Invent.Object(4).ObjIndex = 463
End Select

UserList(userindex).Invent.Object(4).Amount = 1
UserList(userindex).Invent.Object(4).Equipped = 1

UserList(userindex).Invent.Object(5).ObjIndex = 461
UserList(userindex).Invent.Object(5).Amount = 100

UserList(userindex).Invent.ArmourEqpSlot = 4
UserList(userindex).Invent.ArmourEqpObjIndex = UserList(userindex).Invent.Object(4).ObjIndex

UserList(userindex).Invent.WeaponEqpObjIndex = UserList(userindex).Invent.Object(3).ObjIndex
UserList(userindex).Invent.WeaponEqpSlot = 3

Call SaveUser(userindex, CharPath & UCase$(name) & ".chr")

Call ChrToAccount(UserList(userindex).Accounted, name)

totalPjs = val(GetVar(IniPath & "Server.ini", "INIT", "PJS"))
totalPjs = totalPjs + 1

Call WriteVar(IniPath & "Server.ini", "INIT", "PJS", str(totalPjs))
  
'Open User
Call ConnectUser(userindex, name)
  
End Sub

#If UsarQueSocket = 1 Or UsarQueSocket = 2 Then

Sub CloseSocket(ByVal userindex As Integer, Optional ByVal cerrarlo As Boolean = True)
Dim LoopC As Integer

On Error GoTo errhandler
    
    If userindex = LastUser Then
        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1
            If LastUser < 1 Then Exit Do
        Loop
    End If
    
    If UserList(userindex).flags.automatico = True Then
Call Rondas_UsuarioDesconecta(userindex)
End If
    
    'Call SecurityIp.IpRestarConexion(GetLongIp(UserList(UserIndex).ip))
    
    If UserList(userindex).ConnID <> -1 Then
    Call ControlarPortalLum(userindex)
UserList(userindex).flags.TiroPortalL = 0
UserList(userindex).Counters.TimeTeleport = 0
UserList(userindex).Counters.CreoTeleport = False
        Call CloseSocketSL(userindex)
    End If
    
    'Es el mismo user al que está revisando el centinela??
    'IMPORTANTE!!! hacerlo antes de resetear así todavía sabemos el nombre del user
    ' y lo podemos loguear
    If Centinela.RevisandoUserIndex = userindex Then _
        Call modCentinela.CentinelaUserLogout
    
    'mato los comercios seguros
    If UserList(userindex).ComUsu.DestUsu > 0 Then
        If UserList(UserList(userindex).ComUsu.DestUsu).flags.UserLogged Then
            If UserList(UserList(userindex).ComUsu.DestUsu).ComUsu.DestUsu = userindex Then
                Call SendData(SendTarget.ToIndex, UserList(userindex).ComUsu.DestUsu, 0, "||Comercio cancelado por el otro usuario" & FONTTYPE_TALK)
                Call FinComerciarUsu(UserList(userindex).ComUsu.DestUsu)
            End If
        End If
    End If
    
    If UserList(userindex).flags.UserLogged Then
        If NumUsers > 0 Then NumUsers = NumUsers - 1
        Call CloseUser(userindex)
        
        Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
    Else
        Call ResetUserSlot(userindex)
    End If
    
    UserList(userindex).ConnID = -1
    UserList(userindex).ConnIDValida = False
    UserList(userindex).NumeroPaquetesPorMiliSec = 0
    
Exit Sub

errhandler:
    UserList(userindex).ConnID = -1
    UserList(userindex).ConnIDValida = False
    UserList(userindex).NumeroPaquetesPorMiliSec = 0
    Call ResetUserSlot(userindex)
    
#If UsarQueSocket = 1 Then
    If UserList(userindex).ConnID <> -1 Then
    Call ControlarPortalLum(userindex)
UserList(userindex).flags.TiroPortalL = 0
UserList(userindex).Counters.TimeTeleport = 0
UserList(userindex).Counters.CreoTeleport = False
        Call CloseSocketSL(userindex)
    End If
#End If

If UserList(userindex).flags.EnClanes = True Then
    Call ClanDesconecta(userindex)
End If


    Call LogError("CloseSocket - Error = " & Err.Number & " - Descripción = " & Err.Description & " - UserIndex = " & userindex)
End Sub

#ElseIf UsarQueSocket = 0 Then

Sub CloseSocket(ByVal userindex As Integer)
On Error GoTo errhandler
    
    
    
    UserList(userindex).ConnID = -1
    UserList(userindex).NumeroPaquetesPorMiliSec = 0

    If userindex = LastUser And LastUser > 1 Then
        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1
            If LastUser <= 1 Then Exit Do
        Loop
    End If

If UserList(userindex).flags.UserLogged Then
    If NumUsers <> 0 Then NumUsers = NumUsers - 1
    Call CloseUser(userindex)
    Call SendData(toall, 0, 0, "³" & NumUsers)
    End If

    frmMain.Socket2(userindex).Cleanup
    Unload frmMain.Socket2(userindex)
    Call ResetUserSlot(userindex)

Exit Sub

errhandler:
    UserList(userindex).ConnID = -1
    UserList(userindex).NumeroPaquetesPorMiliSec = 0
    Call ResetUserSlot(userindex)
End Sub







#ElseIf UsarQueSocket = 3 Then

Sub CloseSocket(ByVal userindex As Integer, Optional ByVal cerrarlo As Boolean = True)

On Error GoTo errhandler

Dim NURestados As Boolean
Dim CoNnEcTiOnId As Long


    NURestados = False
    CoNnEcTiOnId = UserList(userindex).ConnID
    
    'call logindex(UserIndex, "******> Sub CloseSocket. ConnId: " & CoNnEcTiOnId & " Cerrarlo: " & Cerrarlo)
    
    
  
    UserList(userindex).ConnID = -1 'inabilitamos operaciones en socket
    UserList(userindex).NumeroPaquetesPorMiliSec = 0

    If userindex = LastUser And LastUser > 1 Then
        Do
            LastUser = LastUser - 1
            If LastUser <= 1 Then Exit Do
        Loop While UserList(LastUser).ConnID = -1
    End If

 
If UserList(userindex).flags.UserLogged Then
   If NumUsers <> 0 Then NumUsers = NumUsers - 1
   NURestados = True
   Call CloseUser(userindex)
    Call SendData(toall, 0, 0, "³" & NumUsers)
   End If
    
    Call ResetUserSlot(userindex)
    
    'limpiada la userlist... reseteo el socket, si me lo piden
    'Me lo piden desde: cerrada intecional del servidor (casi todas
    'las llamadas a CloseSocket del codigo)
    'No me lo piden desde: disconnect remoto (el on_close del control
    'de alejo realiza la desconexion automaticamente). Esto puede pasar
    'por ejemplo, si el cliente cierra el AO.
    If cerrarlo Then Call frmMain.TCPServ.CerrarSocket(CoNnEcTiOnId)

Exit Sub

errhandler:
    UserList(userindex).NumeroPaquetesPorMiliSec = 0
    Call LogError("CLOSESOCKETERR: " & Err.Description & " UI:" & userindex)
    
    If Not NURestados Then
        If UserList(userindex).flags.UserLogged Then
            If NumUsers > 0 Then
                NumUsers = NumUsers - 1
            End If
            Call LogError("Cerre sin grabar a: " & UserList(userindex).name)
        End If
    End If
    
    Call LogError("El usuario no guardado tenia connid " & CoNnEcTiOnId & ". Socket no liberado.")
    Call ResetUserSlot(userindex)

End Sub


#End If

'[Alejo-21-5]: Cierra un socket sin limpiar el slot
Sub CloseSocketSL(ByVal userindex As Integer)

#If UsarQueSocket = 1 Then

If UserList(userindex).ConnID <> -1 And UserList(userindex).ConnIDValida Then
    Call BorraSlotSock(UserList(userindex).ConnID)
    Call WSApiCloseSocket(UserList(userindex).ConnID)
    UserList(userindex).ConnIDValida = False
End If

#ElseIf UsarQueSocket = 0 Then

If UserList(userindex).ConnID <> -1 And UserList(userindex).ConnIDValida Then
    frmMain.Socket2(userindex).Cleanup
    Unload frmMain.Socket2(userindex)
    UserList(userindex).ConnIDValida = False
End If

#ElseIf UsarQueSocket = 2 Then

If UserList(userindex).ConnID <> -1 And UserList(userindex).ConnIDValida Then
    Call frmMain.Serv.CerrarSocket(UserList(userindex).ConnID)
    UserList(userindex).ConnIDValida = False
End If

#End If
End Sub

Public Function EnviarDatosASlot(ByVal userindex As Integer, Datos As String) As Long

#If UsarQueSocket = 1 Then '**********************************************
    On Error GoTo Err
    
    Dim Ret As Long
    
    
    
    Ret = WsApiEnviar(userindex, Datos)
    
    If Ret <> 0 And Ret <> WSAEWOULDBLOCK Then
        Call CloseSocketSL(userindex)
        Call Cerrar_Usuario(userindex)
    End If
    EnviarDatosASlot = Ret
    Exit Function
    
Err:
        'If frmMain.SUPERLOG.Value = 1 Then LogCustom ("EnviarDatosASlot:: ERR Handler. userindex=" & UserIndex & " datos=" & Datos & " UL?/CId/CIdV?=" & UserList(UserIndex).flags.UserLogged & "/" & UserList(UserIndex).ConnID & "/" & UserList(UserIndex).ConnIDValida & " ERR: " & Err.Description)

#ElseIf UsarQueSocket = 0 Then '**********************************************

    Dim Encolar As Boolean
    Encolar = False
    
    EnviarDatosASlot = 0
    
    If UserList(userindex).ColaSalida.Count <= 0 Then
        If frmMain.Socket2(userindex).Write(Datos, Len(Datos)) < 0 Then
            If frmMain.Socket2(userindex).LastError = WSAEWOULDBLOCK Then
                UserList(userindex).SockPuedoEnviar = False
                Encolar = True
            Else
                Call Cerrar_Usuario(userindex)
            End If
        End If
    Else
        Encolar = True
    End If
    
    If Encolar Then
        Debug.Print "Encolando..."
        UserList(userindex).ColaSalida.Add Datos
    End If

#ElseIf UsarQueSocket = 2 Then '**********************************************

Dim Encolar As Boolean
Dim Ret As Long
    
    Encolar = False
    
    '//
    '// Valores de retorno:
    '//                     0: Todo OK
    '//                     1: WSAEWOULDBLOCK
    '//                     2: Error critico
    '//
    If UserList(userindex).ColaSalida.Count <= 0 Then
        Ret = frmMain.Serv.Enviar(UserList(userindex).ConnID, Datos, Len(Datos))
        If Ret = 1 Then
            Encolar = True
        ElseIf Ret = 2 Then
            Call CloseSocketSL(userindex)
            Call Cerrar_Usuario(userindex)
        End If
    Else
        Encolar = True
    End If
    
    If Encolar Then
        Debug.Print "Encolando..."
        UserList(userindex).ColaSalida.Add Datos
    End If

#ElseIf UsarQueSocket = 3 Then
    Dim rv As Long
    'al carajo, esto encola solo!!! che, me aprobará los
    'parciales también?, este control hace todo solo!!!!
    On Error GoTo ErrorHandler
        
        If UserList(userindex).ConnID = -1 Then
            Call LogError("TCP::EnviardatosASlot, se intento enviar datos a un userIndex con ConnId=-1")
            Exit Function
        End If
        
        If frmMain.TCPServ.Enviar(UserList(userindex).ConnID, Datos, Len(Datos)) = 2 Then Call CloseSocket(userindex, True)

Exit Function
ErrorHandler:
    Call LogError("TCP::EnviarDatosASlot. UI/ConnId/Datos: " & userindex & "/" & UserList(userindex).ConnID & "/" & Datos)
#End If '**********************************************

End Function
Sub PostMensaje(ByVal Mensaje As String)
frmMain.ListadoM.AddItem Mensaje
If frmMain.ListadoM.ListCount > 9 Then frmMain.ListadoM.RemoveItem 0
frmMain.ListadoM.ListIndex = frmMain.ListadoM.ListCount - 1
End Sub
Sub SendData(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal sndData As String)

On Error Resume Next

Dim LoopC As Integer
Dim x As Integer
Dim Y As Integer

sndData = sndData & ENDC

Select Case sndRoute

    Case SendTarget.ToPCArea
        For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
            For x = UserList(sndIndex).pos.x - MinXBorder + 1 To UserList(sndIndex).pos.x + MinXBorder - 1
               If InMapBounds(sndMap, x, Y) Then
                    If MapData(sndMap, x, Y).userindex > 0 Then
                       If UserList(MapData(sndMap, x, Y).userindex).ConnID <> -1 Then
                            Call EnviarDatosASlot(MapData(sndMap, x, Y).userindex, sndData)
                       End If
                    End If
               End If
            Next x
        Next Y
        Exit Sub
    
    Case SendTarget.ToIndex
        If UserList(sndIndex).ConnID <> -1 Then
            Call EnviarDatosASlot(sndIndex, sndData)
            Exit Sub
        End If


    Case SendTarget.ToNone
        Exit Sub
        
            Case SendTarget.ToAdmins
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID <> -1 Then
                If UserList(LoopC).flags.Privilegios > 0 Then
                    Call EnviarDatosASlot(LoopC, sndData)
               End If
            End If
        Next LoopC
                Call PostMensaje("ADMIN: " & Replace(ReadField(1, Right(sndData, Len(sndData) - 2), Asc("~")), ENDC, ""))
        Exit Sub
    Case SendTarget.ToAdmins
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID <> -1 Then
                If UserList(LoopC).flags.Privilegios > 0 Then
                    Call EnviarDatosASlot(LoopC, sndData)
               End If
            End If
        Next LoopC
        Exit Sub
        
    Case SendTarget.toall
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID <> -1 Then
                If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
                Call PostMensaje(Replace(ReadField(1, Right(sndData, Len(sndData) - 2), Asc("~")), ENDC, ""))
        Exit Sub
        
         Case SendTarget.tosubasta
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID <> -1 Then
                If UserList(LoopC).flags.ParticipaSubasta Then 'Esta logeado como usuario?
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToAllButIndex
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) And (LoopC <> sndIndex) Then
                If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToMap
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).flags.UserLogged Then
                    If UserList(LoopC).pos.Map = sndMap Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            End If
        Next LoopC
        Exit Sub
      
    Case SendTarget.ToMapButIndex
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) And LoopC <> sndIndex Then
                If UserList(LoopC).pos.Map = sndMap Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
            
    Case SendTarget.ToGuildMembers
        
        LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        While LoopC > 0
            If (UserList(LoopC).ConnID <> -1) Then
                Call EnviarDatosASlot(LoopC, sndData)
            End If
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        Wend
        
        Exit Sub


    Case SendTarget.ToDeadArea
        For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
            For x = UserList(sndIndex).pos.x - MinXBorder + 1 To UserList(sndIndex).pos.x + MinXBorder - 1
               If InMapBounds(sndMap, x, Y) Then
                    If MapData(sndMap, x, Y).userindex > 0 Then
                        If UserList(MapData(sndMap, x, Y).userindex).flags.Muerto = 1 Or UserList(MapData(sndMap, x, Y).userindex).flags.Privilegios >= 1 Then
                           If UserList(MapData(sndMap, x, Y).userindex).ConnID <> -1 Then
                                Call EnviarDatosASlot(MapData(sndMap, x, Y).userindex, sndData)
                           End If
                        End If
                    End If
               End If
            Next x
        Next Y
        Exit Sub

    '[Alejo-18-5]
    Case SendTarget.ToPCAreaButIndex
        For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
            For x = UserList(sndIndex).pos.x - MinXBorder + 1 To UserList(sndIndex).pos.x + MinXBorder - 1
               If InMapBounds(sndMap, x, Y) Then
                    If (MapData(sndMap, x, Y).userindex > 0) And (MapData(sndMap, x, Y).userindex <> sndIndex) Then
                       If UserList(MapData(sndMap, x, Y).userindex).ConnID <> -1 Then
                            Call EnviarDatosASlot(MapData(sndMap, x, Y).userindex, sndData)
                       End If
                    End If
               End If
            Next x
        Next Y
        Exit Sub
       
    Case SendTarget.ToClanArea
        For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
            For x = UserList(sndIndex).pos.x - MinXBorder + 1 To UserList(sndIndex).pos.x + MinXBorder - 1
               If InMapBounds(sndMap, x, Y) Then
                    If (MapData(sndMap, x, Y).userindex > 0) Then
                        If UserList(MapData(sndMap, x, Y).userindex).ConnID <> -1 Then
                            If UserList(sndIndex).GuildIndex > 0 And UserList(MapData(sndMap, x, Y).userindex).GuildIndex = UserList(sndIndex).GuildIndex Then
                                Call EnviarDatosASlot(MapData(sndMap, x, Y).userindex, sndData)
                            End If
                        End If
                    End If
               End If
            Next x
        Next Y
        Exit Sub



    Case SendTarget.ToPartyArea
        For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
            For x = UserList(sndIndex).pos.x - MinXBorder + 1 To UserList(sndIndex).pos.x + MinXBorder - 1
               If InMapBounds(sndMap, x, Y) Then
                    If (MapData(sndMap, x, Y).userindex > 0) Then
                        If UserList(MapData(sndMap, x, Y).userindex).ConnID <> -1 Then
                            If UserList(sndIndex).PartyIndex > 0 And UserList(MapData(sndMap, x, Y).userindex).PartyIndex = UserList(sndIndex).PartyIndex Then
                                Call EnviarDatosASlot(MapData(sndMap, x, Y).userindex, sndData)
                            End If
                        End If
                    End If
               End If
            Next x
        Next Y
        Exit Sub
        
    '[CDT 17-02-2004]
    Case SendTarget.ToAdminsAreaButConsejeros
        For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
            For x = UserList(sndIndex).pos.x - MinXBorder + 1 To UserList(sndIndex).pos.x + MinXBorder - 1
               If InMapBounds(sndMap, x, Y) Then
                    If (MapData(sndMap, x, Y).userindex > 0) And (MapData(sndMap, x, Y).userindex <> sndIndex) Then
                       If UserList(MapData(sndMap, x, Y).userindex).ConnID <> -1 Then
                            If UserList(MapData(sndMap, x, Y).userindex).flags.Privilegios > 1 Then
                                Call EnviarDatosASlot(MapData(sndMap, x, Y).userindex, sndData)
                            End If
                       End If
                    End If
               End If
            Next x
        Next Y
        Exit Sub
    '[/CDT]

    Case SendTarget.ToNPCArea
        For Y = Npclist(sndIndex).pos.Y - MinYBorder + 1 To Npclist(sndIndex).pos.Y + MinYBorder - 1
            For x = Npclist(sndIndex).pos.x - MinXBorder + 1 To Npclist(sndIndex).pos.x + MinXBorder - 1
               If InMapBounds(sndMap, x, Y) Then
                    If MapData(sndMap, x, Y).userindex > 0 Then
                       If UserList(MapData(sndMap, x, Y).userindex).ConnID <> -1 Then
                            Call EnviarDatosASlot(MapData(sndMap, x, Y).userindex, sndData)
                       End If
                    End If
               End If
            Next x
        Next Y
        Exit Sub

    Case SendTarget.ToDiosesYclan
        LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        While LoopC > 0
            If (UserList(LoopC).ConnID <> -1) Then
                Call EnviarDatosASlot(LoopC, sndData)
            End If
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        Wend

        LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
        While LoopC > 0
            If (UserList(LoopC).ConnID <> -1) Then
                Call EnviarDatosASlot(LoopC, sndData)
            End If
            LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
        Wend

        Exit Sub

    Case SendTarget.ToConsejo
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).flags.PertAlCons > 0 Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    Case SendTarget.ToConsejoCaos
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).flags.PertAlConsCaos > 0 Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    Case SendTarget.ToRolesMasters
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).flags.EsRolesMaster Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToCiudadanos
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If Not Criminal(LoopC) Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToCriminales
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If Criminal(LoopC) Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToReal
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).Faccion.ArmadaReal = 1 Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToCaos
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).Faccion.FuerzasCaos = 1 Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
        
    Case ToCiudadanosYRMs
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If Not Criminal(LoopC) Or UserList(LoopC).flags.EsRolesMaster Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case ToCriminalesYRMs
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If Criminal(LoopC) Or UserList(LoopC).flags.EsRolesMaster Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case ToRealYRMs
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).Faccion.ArmadaReal = 1 Or UserList(LoopC).flags.EsRolesMaster Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case ToCaosYRMs
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).Faccion.FuerzasCaos = 1 Or UserList(LoopC).flags.EsRolesMaster Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
End Select

End Sub

#If SeguridadAlkon Then

Sub SendCryptedMoveChar(ByVal Map As Integer, ByVal userindex As Integer, ByVal x As Integer, ByVal Y As Integer)
Dim LoopC As Integer

    For LoopC = 1 To LastUser
        If UserList(LoopC).pos.Map = Map Then
            If LoopC <> userindex Then
                If (UserList(LoopC).ConnID <> -1) Then
                    Call EnviarDatosASlot(LoopC, "+" & Encriptacion.MoveCharCrypt(LoopC, UserList(userindex).Char.CharIndex, x, Y) & ENDC)
                End If
            End If
        End If
    Next LoopC
    Exit Sub
    

End Sub

Sub SendCryptedData(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal sndData As String)
'No puse un optional parameter en senddata porque no estoy seguro
'como afecta la performance un parametro opcional
'Prefiero 1K mas de exe que arriesgar performance
On Error Resume Next

Dim LoopC As Integer
Dim x As Integer
Dim Y As Integer


Select Case sndRoute


    Case SendTarget.ToNone
        Exit Sub
        
    Case SendTarget.ToAdmins
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID <> -1 Then
'               If EsDios(UserList(LoopC).Name) Or EsSemiDios(UserList(LoopC).Name) Then
                If UserList(LoopC).flags.Privilegios > 0 Then
                    Call EnviarDatosASlot(LoopC, ProtoCrypt(sndData, LoopC) & ENDC)
               End If
            End If
        Next LoopC
        Exit Sub
        
    Case SendTarget.toall
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID <> -1 Then
                If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                    Call EnviarDatosASlot(LoopC, ProtoCrypt(sndData, LoopC) & ENDC)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToAllButIndex
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) And (LoopC <> sndIndex) Then
                If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                    Call EnviarDatosASlot(LoopC, ProtoCrypt(sndData, LoopC) & ENDC)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToMap
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).flags.UserLogged Then
                    If UserList(LoopC).pos.Map = sndMap Then
                        Call EnviarDatosASlot(LoopC, ProtoCrypt(sndData, LoopC) & ENDC)
                    End If
                End If
            End If
        Next LoopC
        Exit Sub
      
    Case SendTarget.ToMapButIndex
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) And LoopC <> sndIndex Then
                If UserList(LoopC).pos.Map = sndMap Then
                    Call EnviarDatosASlot(LoopC, ProtoCrypt(sndData, LoopC) & ENDC)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToGuildMembers
    
        LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        While LoopC > 0
            If (UserList(LoopC).ConnID <> -1) Then
                Call EnviarDatosASlot(LoopC, ProtoCrypt(sndData, LoopC) & ENDC)
            End If
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        Wend
        
        Exit Sub
    
    Case SendTarget.ToPCArea
        For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
            For x = UserList(sndIndex).pos.x - MinXBorder + 1 To UserList(sndIndex).pos.x + MinXBorder - 1
               If InMapBounds(sndMap, x, Y) Then
                    If MapData(sndMap, x, Y).userindex > 0 Then
                       If UserList(MapData(sndMap, x, Y).userindex).ConnID <> -1 Then
                            Call EnviarDatosASlot(MapData(sndMap, x, Y).userindex, ProtoCrypt(sndData, MapData(sndMap, x, Y).userindex) & ENDC)
                       End If
                    End If
               End If
            Next x
        Next Y
        Exit Sub

    '[Alejo-18-5]
    Case SendTarget.ToPCAreaButIndex
        For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
            For x = UserList(sndIndex).pos.x - MinXBorder + 1 To UserList(sndIndex).pos.x + MinXBorder - 1
               If InMapBounds(sndMap, x, Y) Then
                    If (MapData(sndMap, x, Y).userindex > 0) And (MapData(sndMap, x, Y).userindex <> sndIndex) Then
                       If UserList(MapData(sndMap, x, Y).userindex).ConnID <> -1 Then
                            Call EnviarDatosASlot(MapData(sndMap, x, Y).userindex, ProtoCrypt(sndData, MapData(sndMap, x, Y).userindex) & ENDC)
                       End If
                    End If
               End If
            Next x
        Next Y
        Exit Sub

    '[CDT 17-02-2004]
    Case SendTarget.ToAdminsAreaButConsejeros
        For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
            For x = UserList(sndIndex).pos.x - MinXBorder + 1 To UserList(sndIndex).pos.x + MinXBorder - 1
               If InMapBounds(sndMap, x, Y) Then
                    If (MapData(sndMap, x, Y).userindex > 0) And (MapData(sndMap, x, Y).userindex <> sndIndex) Then
                       If UserList(MapData(sndMap, x, Y).userindex).ConnID <> -1 Then
                            If UserList(MapData(sndMap, x, Y).userindex).flags.Privilegios > 1 Then
                                Call EnviarDatosASlot(MapData(sndMap, x, Y).userindex, ProtoCrypt(sndData, MapData(sndMap, x, Y).userindex) & ENDC)
                            End If
                       End If
                    End If
               End If
            Next x
        Next Y
        Exit Sub
    '[/CDT]

    Case SendTarget.ToNPCArea
        For Y = Npclist(sndIndex).pos.Y - MinYBorder + 1 To Npclist(sndIndex).pos.Y + MinYBorder - 1
            For x = Npclist(sndIndex).pos.x - MinXBorder + 1 To Npclist(sndIndex).pos.x + MinXBorder - 1
               If InMapBounds(sndMap, x, Y) Then
                    If MapData(sndMap, x, Y).userindex > 0 Then
                       If UserList(MapData(sndMap, x, Y).userindex).ConnID <> -1 Then
                            Call EnviarDatosASlot(MapData(sndMap, x, Y).userindex, ProtoCrypt(sndData, MapData(sndMap, x, Y).userindex) & ENDC)
                       End If
                    End If
               End If
            Next x
        Next Y
        Exit Sub

    Case SendTarget.ToIndex
        If UserList(sndIndex).ConnID <> -1 Then
             Call EnviarDatosASlot(sndIndex, ProtoCrypt(sndData, sndIndex) & ENDC)
             Exit Sub
        End If
    Case SendTarget.ToDiosesYclan
        
        LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        While LoopC > 0
            If (UserList(LoopC).ConnID <> -1) Then
                Call EnviarDatosASlot(LoopC, ProtoCrypt(sndData, LoopC) & ENDC)
            End If
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        Wend

        LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
        While LoopC > 0
            If (UserList(LoopC).ConnID <> -1) Then
                Call EnviarDatosASlot(LoopC, ProtoCrypt(sndData, LoopC) & ENDC)
            End If
            LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
        Wend

        Exit Sub
        

End Select

End Sub

#End If

Function EstaPCarea(Index As Integer, Index2 As Integer) As Boolean


Dim x As Integer, Y As Integer
For Y = UserList(Index).pos.Y - MinYBorder + 1 To UserList(Index).pos.Y + MinYBorder - 1
        For x = UserList(Index).pos.x - MinXBorder + 1 To UserList(Index).pos.x + MinXBorder - 1

            If MapData(UserList(Index).pos.Map, x, Y).userindex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
        Next x
Next Y
EstaPCarea = False
End Function

Function HayPCarea(pos As WorldPos) As Boolean


Dim x As Integer, Y As Integer
For Y = pos.Y - MinYBorder + 1 To pos.Y + MinYBorder - 1
        For x = pos.x - MinXBorder + 1 To pos.x + MinXBorder - 1
            If x > 0 And Y > 0 And x < 101 And Y < 101 Then
                If MapData(pos.Map, x, Y).userindex > 0 Then
                    HayPCarea = True
                    Exit Function
                End If
            End If
        Next x
Next Y
HayPCarea = False
End Function

Function HayOBJarea(pos As WorldPos, ObjIndex As Integer) As Boolean


Dim x As Integer, Y As Integer
For Y = pos.Y - MinYBorder + 1 To pos.Y + MinYBorder - 1
        For x = pos.x - MinXBorder + 1 To pos.x + MinXBorder - 1
            If MapData(pos.Map, x, Y).OBJInfo.ObjIndex = ObjIndex Then
                HayOBJarea = True
                Exit Function
            End If
        
        Next x
Next Y
HayOBJarea = False
End Function

Function ValidateChr(ByVal userindex As Integer) As Boolean

ValidateChr = UserList(userindex).Char.Head <> 0 _
                And UserList(userindex).Char.Body <> 0 _
                And ValidateSkills(userindex)

End Function

Sub ConnectUser(ByVal userindex As Integer, name As String)
Dim N As Integer
Dim tStr As String

'Reseteamos los FLAGS
UserList(userindex).flags.Guerra = False
Call SendData(SendTarget.ToIndex, userindex, 0, "|G0")
UserList(userindex).flags.Escondido = 0
UserList(userindex).flags.TargetNPC = 0
UserList(userindex).flags.TargetNpcTipo = eNPCType.Comun
UserList(userindex).flags.TargetObj = 0
UserList(userindex).flags.TargetUser = 0
UserList(userindex).Char.FX = 0

If UCase$(name) = "LORWIK" And UserList(userindex).ip <> GetVar(IniPath & "Server.ini", "IPGM", "LorwikIp") Then
    Call SendData(ToIndex, userindex, 0, "ERRTu no eres un GM. Dejate de joder y juega con tu usuario.")
    Call CloseSocket(userindex)
    Exit Sub
End If

If UCase$(name) = "STICK" And UserList(userindex).ip <> GetVar(IniPath & "Server.ini", "IPGM", "StickIp") Then
    Call SendData(ToIndex, userindex, 0, "ERRTu no eres un GM. Dejate de joder y juega con tu usuario.")
    Call CloseSocket(userindex)
    Exit Sub
End If

If UCase$(name) = "SANTI" And UserList(userindex).ip <> GetVar(IniPath & "Server.ini", "IPGM", "SantiIp") Then
    Call SendData(ToIndex, userindex, 0, "ERRTu no eres un GM. Dejate de joder y juega con tu usuario.")
    Call CloseSocket(userindex)
    Exit Sub
End If

'Controlamos no pasar el maximo de usuarios
If NumUsers >= MaxUsers Then
    Call SendData(SendTarget.ToIndex, userindex, 0, "ERREl servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo mas tarde.")
    Call CloseSocket(userindex)
    Exit Sub
End If

'¿Este IP ya esta conectado?
If AllowMultiLogins = 0 Then
    If CheckForSameIP(userindex, UserList(userindex).ip) = True Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "ERRNo es posible usar mas de un personaje al mismo tiempo.")
        Call CloseSocket(userindex)
        Exit Sub
    End If
End If

'¿Existe el personaje?
If Not FileExist(CharPath & UCase$(name) & ".chr", vbNormal) Then
    Call SendData(SendTarget.ToIndex, userindex, 0, "ERREl personaje no existe.")
    Call CloseSocket(userindex)
    Exit Sub
End If

'¿Ya esta conectado el personaje?
If CheckForSameName(userindex, name) Then
    If UserList(NameIndex(name)).Counters.Saliendo Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "ERREl usuario está saliendo.")
    Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "ERRPerdon, un usuario con el mismo nombre se há logoeado.")
    End If
    Call CloseSocket(userindex)
    Exit Sub
End If
Call WriteVar(CharPath & UCase$(name) & ".chr", "FLAGS", "Equitando", "0")
'Cargamos el personaje
Dim Leer As New clsIniReader

Call Leer.Initialize(CharPath & UCase$(name) & ".chr")

'Cargamos los datos del personaje
Call LoadUserInit(userindex, Leer)

Call LoadUserStats(userindex, Leer)

If Not ValidateChr(userindex) Then
    Call SendData(SendTarget.ToIndex, userindex, 0, "ERRError en el personaje.")
    Call CloseSocket(userindex)
    Exit Sub
End If

Call LoadUserReputacion(userindex, Leer)

Set Leer = Nothing

If UserList(userindex).Invent.EscudoEqpSlot = 0 Then UserList(userindex).Char.ShieldAnim = NingunEscudo
If UserList(userindex).Invent.CascoEqpSlot = 0 Then UserList(userindex).Char.CascoAnim = NingunCasco
If UserList(userindex).Invent.WeaponEqpSlot = 0 Then UserList(userindex).Char.WeaponAnim = NingunArma


Call UpdateUserInv(True, userindex, 0)
Call UpdateUserHechizos(True, userindex, 0)

If UserList(userindex).flags.Navegando = 1 Then
     UserList(userindex).Char.Body = ObjData(UserList(userindex).Invent.BarcoObjIndex).Ropaje
     UserList(userindex).Char.Head = 0
     UserList(userindex).Char.WeaponAnim = NingunArma
     UserList(userindex).Char.ShieldAnim = NingunEscudo
     UserList(userindex).Char.CascoAnim = NingunCasco
End If

If UserList(userindex).flags.Muerto = 1 Then
Call SendData(SendTarget.ToIndex, userindex, 0, "ZRE")
End If

If UserList(userindex).flags.Equitando = 1 Then
     UserList(userindex).Char.Body = ObjData(UserList(userindex).Invent.MonturaObjIndex).Ropaje
End If


If UserList(userindex).flags.Paralizado Then
#If SeguridadAlkon Then
    If EncriptarProtocolosCriticos Then
        Call SendCryptedData(SendTarget.ToIndex, userindex, 0, "PARADOK")
    Else
#End If
        Call SendData(SendTarget.ToIndex, userindex, 0, "PARADOK")
#If SeguridadAlkon Then
    End If
#End If
End If

'Feo, esto tiene que ser parche cliente
If UserList(userindex).flags.Estupidez = 0 Then Call SendData(SendTarget.ToIndex, userindex, 0, "NESTUP")
'

'Posicion de comienzo
If UserList(userindex).pos.Map = 0 Then
    If UCase$(UserList(userindex).Hogar) = "RAMX" Then
             UserList(userindex).pos = Ramx
             Else
        UserList(userindex).Hogar = "RAMX"
        UserList(userindex).pos = Ramx
    End If
Else

   ''TELEFRAG
    If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex <> 0 Then
        ''si estaba en comercio seguro...
        If UserList(MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex).ComUsu.DestUsu > 0 Then
            If UserList(UserList(MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex).ComUsu.DestUsu).flags.UserLogged Then
                Call FinComerciarUsu(UserList(MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex).ComUsu.DestUsu)
                Call SendData(SendTarget.ToIndex, UserList(MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex).ComUsu.DestUsu, 0, "||Comercio cancelado. El otro usuario se ha desconectado." & FONTTYPE_TALK)
            End If
            If UserList(MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex).flags.UserLogged Then
                Call FinComerciarUsu(MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex)
                Call SendData(SendTarget.ToIndex, MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex, 0, "ERRAlguien se ha conectado donde te encontrabas, por favor reconéctate...")
            End If
        End If
        Call CloseSocket(MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex)
    End If
   
   
    If UserList(userindex).flags.Muerto = 1 Then
        Call Empollando(userindex)
    End If
End If

If Not MapaValido(UserList(userindex).pos.Map) Then
    Call SendData(SendTarget.ToIndex, userindex, 0, "ERREL PJ se encuenta en un mapa invalido.")
    Call CloseSocket(userindex)
    Exit Sub
End If

'Nombre de sistema
UserList(userindex).name = name

UserList(userindex).showName = True 'Por default los nombres son visibles

'Info
Call SendData(SendTarget.ToIndex, userindex, 0, "IU" & userindex) 'Enviamos el User index
Call SendData(SendTarget.ToIndex, userindex, 0, "CM" & UserList(userindex).pos.Map & "," & MapInfo(UserList(userindex).pos.Map).MapVersion) 'Carga el mapa
Call SendData(SendTarget.ToIndex, userindex, 0, "TM" & MapInfo(UserList(userindex).pos.Map).Music)
Call SendData(SendTarget.ToIndex, userindex, 0, "N~" & MapInfo(UserList(userindex).pos.Map).name)


'Vemos que clase de user es (se lo usa para setear los privilegios alcrear el PJ)
UserList(userindex).flags.EsRolesMaster = EsRolesMaster(name)
If EsAdmin(name) Then
    UserList(userindex).flags.Privilegios = PlayerType.Admin
    Call LogGM(UserList(userindex).name, "Se conecto con ip:" & UserList(userindex).ip, False)
ElseIf EsDios(name) Then
    UserList(userindex).flags.Privilegios = PlayerType.Dios
    Call LogGM(UserList(userindex).name, "Se conecto con ip:" & UserList(userindex).ip, False)
ElseIf EsSemiDios(name) Then
    UserList(userindex).flags.Privilegios = PlayerType.SemiDios
    Call LogGM(UserList(userindex).name, "Se conecto con ip:" & UserList(userindex).ip, False)
ElseIf EsConsejero(name) Then
    UserList(userindex).flags.Privilegios = PlayerType.Consejero
    Call LogGM(UserList(userindex).name, "Se conecto con ip:" & UserList(userindex).ip, True)
Else
    UserList(userindex).flags.Privilegios = PlayerType.User
End If

''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
UserList(userindex).Counters.IdleCount = 0
'Crea  el personaje del usuario
Call MakeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y)

Call SendData(SendTarget.ToIndex, userindex, 0, "IP" & UserList(userindex).Char.CharIndex)
''[/el oso]

Call SendUserStatsBox(userindex)

Call EnviarHambreYsed(userindex)

Call SendMOTD(userindex)

If haciendoBK Then
    Call SendData(SendTarget.ToIndex, userindex, 0, "BKW")
    Call SendData(SendTarget.ToIndex, userindex, 0, "||Servidor> Por favor espera algunos segundos, WorldSave esta ejecutandose." & FONTTYPE_SERVER)
End If

If EnPausa Then
    Call SendData(SendTarget.ToIndex, userindex, 0, "BKW")
    Call SendData(SendTarget.ToIndex, userindex, 0, "||Servidor> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde." & FONTTYPE_SERVER)
End If

If EnTesting And UserList(userindex).Stats.ELV >= 18 Then
    Call SendData(SendTarget.ToIndex, userindex, 0, "ERRServidor en Testing por unos minutos, conectese con PJs de nivel menor a 18. No se conecte con Pjs que puedan resultar importantes por ahora pues pueden arruinarse.")
    Call CloseSocket(userindex)
    Exit Sub
End If

'Actualiza el Num de usuarios
'DE ACA EN ADELANTE GRABA EL CHARFILE, OJO!
NumUsers = NumUsers + 1
UserList(userindex).flags.UserLogged = True
Call SendData(toall, 0, 0, "³" & NumUsers)

'usado para borrar Pjs
Call WriteVar(CharPath & UserList(userindex).name & ".chr", "INIT", "Logged", "1")

Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
Call SendData(toall, 0, 0, "³" & NumUsers)


MapInfo(UserList(userindex).pos.Map).NumUsers = MapInfo(UserList(userindex).pos.Map).NumUsers + 1

If UserList(userindex).Stats.SkillPts > 0 Then
    Call EnviarSkills(userindex)
    Call EnviarSubirNivel(userindex, UserList(userindex).Stats.SkillPts)
End If

If NumUsers > DayStats.MaxUsuarios Then DayStats.MaxUsuarios = NumUsers

If NumUsers > recordusuarios Then
   Call SendData(SendTarget.toall, 0, 0, "PRB19," & NumUsers)
    recordusuarios = NumUsers
    Call WriteVar(IniPath & "Server.ini", "INIT", "Record", str(recordusuarios))
    
    Call EstadisticasWeb.Informar(RECORD_USUARIOS, recordusuarios)
End If

If UserList(userindex).NroMacotas > 0 Then
    Dim i As Integer
    For i = 1 To MAXMASCOTAS
        If UserList(userindex).MascotasType(i) > 0 Then
            UserList(userindex).MascotasIndex(i) = SpawnNpc(UserList(userindex).MascotasType(i), UserList(userindex).pos, True, True)
            
            If UserList(userindex).MascotasIndex(i) > 0 Then
                Npclist(UserList(userindex).MascotasIndex(i)).MaestroUser = userindex
                Call FollowAmo(UserList(userindex).MascotasIndex(i))
            Else
                UserList(userindex).MascotasIndex(i) = 0
            End If
        End If
    Next i
End If

If UserList(userindex).flags.Equitando = 1 Then Call SendData(SendTarget.ToIndex, userindex, 0, "ET")
If UserList(userindex).flags.Navegando = 1 Then Call SendData(SendTarget.ToIndex, userindex, 0, "NAVEG")

If Criminal(userindex) Then
    'Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Miembro de las fuerzas del caos > Seguro desactivado <" & FONTTYPE_FIGHT)
    Call SendData(SendTarget.ToIndex, userindex, 0, "SEGOFF")
    UserList(userindex).flags.Seguro = False
Else
    UserList(userindex).flags.Seguro = True
    Call SendData(SendTarget.ToIndex, userindex, 0, "SEGON")
End If

If ServerSoloGMs > 0 Then
    If UserList(userindex).flags.Privilegios < ServerSoloGMs Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "ERRServidor restringido a administradores de jerarquia mayor o igual a: " & ServerSoloGMs & ". Por favor intente en unos momentos.")
        Call CloseSocket(userindex)
        Exit Sub
    End If
End If

If UserList(userindex).GuildIndex > 0 Then
    'welcome to the show baby...
    If Not modGuilds.m_ConectarMiembroAClan(userindex, UserList(userindex).GuildIndex) Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Tu estado no te permite entrar al clan." & FONTTYPE_GUILD)
    End If
End If

Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXWARP & "," & 0)

Call SendData(SendTarget.ToIndex, userindex, 0, "LOGGED")

Call modGuilds.SendGuildNews(userindex)

If UserList(userindex).flags.NoActualizado Then
    Call SendData(SendTarget.ToIndex, userindex, 0, "REAU")
End If

If Lloviendo Then Call SendData(SendTarget.ToIndex, userindex, 0, "LLU")

If Anochecer = 1 Then
Call SendData(ToIndex, userindex, 0, "NUB" & 1)
ElseIf MedioDia = 1 Then
Call SendData(ToIndex, userindex, 0, "MDI" & 1)
ElseIf Atardecer = 1 Then
Call SendData(ToIndex, userindex, 0, "TAR" & 1)
ElseIf Amanecer = 1 Then
Call SendData(ToIndex, userindex, 0, "MAÑ" & 1)
End If

tStr = modGuilds.a_ObtenerRechazoDeChar(UserList(userindex).name)

If tStr <> vbNullString Then
    Call SendData(SendTarget.ToIndex, userindex, 0, "!!Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr & ENDC)
End If

Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "PF" & UserList(userindex).Stats.UserAtributos(Fuerza))
Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "PG" & UserList(userindex).Stats.UserAtributos(Agilidad))

Call MostrarNumUsers


N = FreeFile
Open App.Path & "\logs\numusers.log" For Output As N
Print #N, NumUsers
Close #N

N = FreeFile
'Log
Open App.Path & "\logs\Connect.log" For Append Shared As #N
Print #N, UserList(userindex).name & " ha entrado al juego. UserIndex:" & userindex & " " & Time & " " & Date
Close #N

End Sub
Sub SendMOTD(ByVal userindex As Integer)
    Dim j As Long
    
    Call SendData(SendTarget.ToIndex, userindex, 0, "||" & "Mensajes de entrada:" & FONTTYPE_INFO)
    
    For j = 1 To MaxLines
        Call SendData(SendTarget.ToIndex, userindex, 0, "||" & Chr$(3) & MOTD(j).texto)
    Next j
    Call SendData(ToIndex, userindex, 0, "||El Castillo Oeste esta en manos del clan: " & GetVar(App.Path & "\Castillos.ini", "CLANES", "OESTE") & FONTTYPE_FENIX)
    Call SendData(ToIndex, userindex, 0, "||El Castillo Norte esta en manos del clan: " & GetVar(App.Path & "\Castillos.ini", "CLANES", "NORTE") & FONTTYPE_FENIX)
    Call SendData(ToIndex, userindex, 0, "||El Castillo Este esta en manos del clan: " & GetVar(App.Path & "\Castillos.ini", "CLANES", "ESTE") & FONTTYPE_FENIX)
End Sub

Sub ResetFacciones(ByVal userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(userindex).Faccion
        .ArmadaReal = 0
        .FuerzasCaos = 0
        .CiudadanosMatados = 0
        .CriminalesMatados = 0
        .RecibioArmaduraCaos = 0
        .RecibioArmaduraReal = 0
        .RecibioExpInicialCaos = 0
        .RecibioExpInicialReal = 0
        .RecompensasCaos = 0
        .RecompensasReal = 0
        .Reenlistadas = 0
    End With
End Sub

Sub ResetContadores(ByVal userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(userindex).Counters
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
        .Pasos = 0
        .Pena = 0
        .PiqueteC = 0
        .STACounter = 0
        .Veneno = 0
        .Trabajando = 0
        .Ocultando = 0

        .TimerLanzarSpell = 0
        .TimerPuedeAtacar = 0
        .TimerPuedeTrabajar = 0
        .TimerUsar = 0
    End With
End Sub

Sub ResetCharInfo(ByVal userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(userindex).Char
        .Body = 0
        .CascoAnim = 0
        .CharIndex = 0
        .FX = 0
        .Head = 0
        .loops = 0
        .Heading = 0
        .loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
    End With
End Sub

Sub ResetBasicUserInfo(ByVal userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(userindex)
        .name = ""
        .modName = ""
        .Password = ""
        .Desc = ""
        .DescRM = ""
        .pos.Map = 0
        .pos.x = 0
        .pos.Y = 0
        .ip = ""
        .RDBuffer = ""
        .Clase = ""
        .email = ""
        .Genero = ""
        .Hogar = ""
        .Raza = ""

        .RandKey = 0
        .PrevCheckSum = 0
        .PacketNumber = 0

        .EmpoCont = 0
        .PartyIndex = 0
        .PartySolicitud = 0
        
        With .Stats
            .Banco = 0
            .ELV = 0
            .ELU = 0
            .Exp = 0
            .def = 0
            .CriminalesMatados = 0
            .NPCsMuertos = 0
            .UsuariosMatados = 0
            .FIT = 0
            .SkillPts = 0
        End With
    End With
End Sub

Sub ResetReputacion(ByVal userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(userindex).Reputacion
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

Sub ResetGuildInfo(ByVal userindex As Integer)
    If UserList(userindex).EscucheClan > 0 Then
        Call modGuilds.GMDejaDeEscucharClan(userindex, UserList(userindex).EscucheClan)
        UserList(userindex).EscucheClan = 0
    End If
    If UserList(userindex).GuildIndex > 0 Then
        Call modGuilds.m_DesconectarMiembroDelClan(userindex, UserList(userindex).GuildIndex)
    End If
    UserList(userindex).GuildIndex = 0
End Sub

Sub ResetUserFlags(ByVal userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/29/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'03/29/2006 Maraxus - Reseteo el CentinelaOK también.
'*************************************************
    With UserList(userindex).flags
        .Comerciando = False
        .Ban = 0
        .ParticipaSubasta = False
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
        .Descuento = ""
        .Hambre = 0
        .Sed = 0
        .Descansar = False
        .ModoCombate = False
        .Vuela = 0
        .Navegando = 0
        .Equitando = 0
        .Oculto = 0
        .Envenenado = 0
        .Invisible = 0
        .Paralizado = 0
        .Maldicion = 0
        .Bendicion = 0
        .Meditando = 0
        .Privilegios = PlayerType.User
        .PuedeMoverse = 0
        .OldBody = 0
        .OldHead = 0
        .AdminInvisible = 0
        .ValCoDe = 0
        .Hechizo = 0
        .TimesWalk = 0
        .StartWalk = 0
        .CountSH = 0
        .EstaEmpo = 0
        .PertAlCons = 0
        .PertAlConsCaos = 0
        .Silenciado = 0
        .CentinelaOK = False
    End With
End Sub

Sub ResetUserSpells(ByVal userindex As Integer)
    Dim LoopC As Long
    For LoopC = 1 To MAXUSERHECHIZOS
        UserList(userindex).Stats.UserHechizos(LoopC) = 0
    Next LoopC
End Sub

Sub ResetUserPets(ByVal userindex As Integer)
    Dim LoopC As Long
    
    UserList(userindex).NroMacotas = 0
        
    For LoopC = 1 To MAXMASCOTAS
        UserList(userindex).MascotasIndex(LoopC) = 0
        UserList(userindex).MascotasType(LoopC) = 0
    Next LoopC
End Sub

Sub ResetUserBanco(ByVal userindex As Integer)
    Dim LoopC As Long
    
    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
          UserList(userindex).BancoInvent.Object(LoopC).Amount = 0
          UserList(userindex).BancoInvent.Object(LoopC).Equipped = 0
          UserList(userindex).BancoInvent.Object(LoopC).ObjIndex = 0
    Next LoopC
    
    UserList(userindex).BancoInvent.NroItems = 0
End Sub

Public Sub LimpiarComercioSeguro(ByVal userindex As Integer)
    With UserList(userindex).ComUsu
        If .DestUsu > 0 Then
            Call FinComerciarUsu(.DestUsu)
            Call FinComerciarUsu(userindex)
        End If
    End With
End Sub

Sub ResetUserSlot(ByVal userindex As Integer)

Dim UsrTMP As User

Set UserList(userindex).CommandsBuffer = Nothing


Set UserList(userindex).ColaSalida = Nothing
UserList(userindex).SockPuedoEnviar = False
UserList(userindex).ConnIDValida = False
UserList(userindex).ConnID = -1

Call LimpiarComercioSeguro(userindex)
Call ResetFacciones(userindex)
Call ResetContadores(userindex)
Call ResetCharInfo(userindex)
Call ResetBasicUserInfo(userindex)
Call ResetReputacion(userindex)
Call ResetGuildInfo(userindex)
Call ResetUserFlags(userindex)
Call LimpiarInventario(userindex)
Call ResetUserSpells(userindex)
Call ResetUserPets(userindex)
Call ResetUserBanco(userindex)

With UserList(userindex).ComUsu
    .Acepto = False
    .Cant = 0
    .DestNick = ""
    .DestUsu = 0
    .Objeto = 0
End With

UserList(userindex) = UsrTMP

End Sub


Sub CloseUser(ByVal userindex As Integer)
'Call LogTarea("CloseUser " & UserIndex)

On Error GoTo errhandler

Dim N As Integer
Dim x As Integer
Dim Y As Integer
Dim LoopC As Integer
Dim Map As Integer
Dim name As String
Dim Raza As String
Dim Clase As String
Dim i As Integer

Dim aN As Integer

aN = UserList(userindex).flags.AtacadoPorNpc
If aN > 0 Then
      Npclist(aN).Movement = Npclist(aN).flags.OldMovement
      Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
      Npclist(aN).flags.AttackedBy = ""
End If
UserList(userindex).flags.AtacadoPorNpc = 0

Map = UserList(userindex).pos.Map
x = UserList(userindex).pos.x
Y = UserList(userindex).pos.Y
name = UCase$(UserList(userindex).name)
Raza = UserList(userindex).Raza
Clase = UserList(userindex).Clase

UserList(userindex).Char.FX = 0
UserList(userindex).Char.loops = 0
Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & 0 & "," & 0)

UserList(userindex).flags.UserLogged = False
UserList(userindex).Counters.Saliendo = False

If userindex = GranPoder Then
    Call SendData(SendTarget.toall, userindex, 0, "||" & UserList(userindex).name & " ha abandonado el reino." & FONTTYPE_GUILD)
    Call OtorgarGranPoder(0)
End If

'Le devolvemos el body y head originales
If UserList(userindex).flags.AdminInvisible = 1 Then Call DoAdminInvisible(userindex)

'si esta en party le devolvemos la experiencia
If UserList(userindex).PartyIndex > 0 Then Call mdParty.SalirDeParty(userindex)

' Grabamos el personaje del usuario
Call SaveUser(userindex, CharPath & name & ".chr")

'usado para borrar Pjs
Call WriteVar(CharPath & UserList(userindex).name & ".chr", "INIT", "Logged", "0")


'Quitar el dialogo
'If MapInfo(Map).NumUsers > 0 Then
'    Call SendToUserArea(UserIndex, "QDL" & UserList(UserIndex).Char.charindex)
'End If

If MapInfo(Map).NumUsers > 0 Then
    Call SendData(SendTarget.ToMapButIndex, userindex, Map, "QDL" & UserList(userindex).Char.CharIndex)
End If



'Borrar el personaje
If UserList(userindex).Char.CharIndex > 0 Then
    Call EraseUserChar(SendTarget.ToMap, userindex, Map, userindex)
End If

'Borrar mascotas
For i = 1 To MAXMASCOTAS
    If UserList(userindex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(userindex).MascotasIndex(i)).flags.NPCActive Then _
            Call QuitarNPC(UserList(userindex).MascotasIndex(i))
    End If
Next i

'Update Map Users
MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1

If MapInfo(Map).NumUsers < 0 Then
    MapInfo(Map).NumUsers = 0
End If

' Si el usuario habia dejado un msg en la gm's queue lo borramos
If Ayuda.Existe(UserList(userindex).name) Then Call Ayuda.Quitar(UserList(userindex).name)
If Torneox.Existe(UserList(userindex).name) Then Call Torneox.Quitar(UserList(userindex).name)

Call ResetUserSlot(userindex)

Call MostrarNumUsers

 Call SendData(toall, 0, 0, "³" & NumUsers)

N = FreeFile(1)
Open App.Path & "\logs\Connect.log" For Append Shared As #N
Print #N, name & " há dejado el juego. " & "User Index:" & userindex & " " & Time & " " & Date
Close #N

Exit Sub

'Si el usuario sale del juego en duelo, lo llevamos a ramx
If UserList(userindex).flags.EnDuelo = 1 Then
Call WarpUserChar(userindex, 1, 50, 50)
End If

errhandler:
Call LogError("Error en CloseUser. Número " & Err.Number & " Descripción: " & Err.Description)


End Sub

Sub SetUserPacketID(ByVal userindex As Integer)
    If UserList(userindex).sID = 60 Or UserList(userindex).sID = 0 Then
        UserList(userindex).sID = 33
        Exit Sub
    End If
    UserList(userindex).sID = UserList(userindex).sID + 1
End Sub
Sub HandleData(ByVal userindex As Integer, ByVal rData As String)

'
' ATENCION: Cambios importantes en HandleData.
' =========
'
'           La funcion se encuentra dividida en 2,
'           una parte controla los comandos que
'           empiezan con "/" y la otra los comanos
'           que no. (Basado en la idea de Barrin)
'


Call LogTarea("Sub HandleData :" & rData & " " & UserList(userindex).name)

'Nunca jamas remover o comentar esta linea !!!
'Nunca jamas remover o comentar esta linea !!!
'Nunca jamas remover o comentar esta linea !!!
On Error GoTo ErrorHandler:
'Nunca jamas remover o comentar esta linea !!!
'Nunca jamas remover o comentar esta linea !!!
'Nunca jamas remover o comentar esta linea !!!
'
'Ah, no me queres hacer caso ? Entonces
'atenete a las consecuencias!!
'

    Dim CadenaOriginal As String
    
    Dim LoopC As Integer
    Dim nPos As WorldPos
    Dim tStr As String
    Dim tInt As Integer
    Dim tLong As Long
    Dim tIndex As Integer
    Dim tName As String
    Dim tMessage As String
    Dim AuxInd As Integer
    Dim Arg1 As String
    Dim Arg2 As String
    Dim Arg3 As String
    Dim Arg4 As String
    Dim Ver As String
    Dim encpass As String
    Dim Pass As String
    Dim mapa As Integer
    Dim name As String
    Dim ind
    Dim N As Integer
    Dim wpaux As WorldPos
    Dim mifile As Integer
    Dim x As Integer
    Dim Y As Integer
    Dim DummyInt As Integer
    Dim T() As String
    Dim i As Integer
    
    Dim sndData As String
    Dim cliMD5 As String
    Dim ClientChecksum As String
    Dim ServerSideChecksum As Long
    Dim IdleCountBackup As Long
    
    Call SetUserPacketID(userindex)
    
    rData = UnEncryptStr(rData, lstEncript(UserList(userindex).sID))

    CadenaOriginal = rData
    
    '¿Tiene un indece valido?
    If userindex <= 0 Then
        Call CloseSocket(userindex)
        Exit Sub
    End If
    
      If Left$(rData, 3) = "QLR" Then 'QuestList request
        Call SendQuestList(userindex)
        Exit Sub
    ElseIf Left$(rData, 3) = "QIR" Then
        tInt = UserList(userindex).Stats.UserQuests(val(Right(rData, Len(rData) - 3))).QuestIndex
        
        If tInt < 1 Then Exit Sub
        
        If QuestList(tInt).NpcKillIndex And QuestList(tInt).CantNPCs Then
            tStr = UserList(userindex).Stats.UserQuests(val(Right(rData, Len(rData) - 3))).NPCsKilled & " / " & QuestList(tInt).CantNPCs
        Else
            tStr = "Ninguna"
        End If

        Call SendData(SendTarget.ToIndex, userindex, 0, "QI" & QuestList(tInt).Nombre & "-" & QuestList(tInt).Descripcion & "-" & tStr)
        Exit Sub
    ElseIf Left$(rData, 2) = "QA" Then
        tInt = val(Right(rData, Len(rData) - 2))
        i = UserList(userindex).Stats.UserQuests(tInt).QuestIndex
        
        If tInt < 1 Or i < 1 Then Exit Sub
        
        UserList(userindex).Stats.UserQuests(tInt).QuestIndex = 0
        UserList(userindex).Stats.UserQuests(tInt).NPCsKilled = 0
        
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Has abandonado la quest " & Chr(34) & QuestList(i).Nombre & Chr(34) & FONTTYPE_INFO)
        Call SendQuestList(userindex)
        Exit Sub
    End If
    
    If Left$(rData, 13) = "gIvEmEvAlcOde" Then
#If SeguridadAlkon Then
        '<<<<<<<<<<< MODULO PRIVADO DE CADA IMPLEMENTACION >>>>>>
        UserList(userindex).flags.ValCoDe = RandomNumber(20000, 32000)
        UserList(userindex).RandKey = RandomNumber(0, 99999)
        UserList(userindex).PrevCheckSum = UserList(userindex).RandKey
        UserList(userindex).PacketNumber = 100
        UserList(userindex).KeyCrypt = (UserList(userindex).RandKey And 17320) Xor (UserList(userindex).flags.ValCoDe Xor 4232)
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        Call SendData(SendTarget.ToIndex, userindex, 0, "VAL" & UserList(userindex).RandKey & "," & UserList(userindex).flags.ValCoDe & "," & Encriptacion.StringValidacion)
        Exit Sub
    Else
        '<<<<<<<<<<< MODULO PRIVADO DE CADA IMPLEMENTACION >>>>>>
        ClientChecksum = Right$(rData, Len(rData) - InStrRev(rData, Chr$(126)))
        tStr = Left$(rData, Len(rData) - Len(ClientChecksum) - 1)
        ServerSideChecksum = CheckSum(UserList(userindex).PrevCheckSum, tStr)
        UserList(userindex).PrevCheckSum = ServerSideChecksum
        
        If CLng(ClientChecksum) <> ServerSideChecksum Then
            Call LogError("Checksum error userindex: " & userindex & " rdata: " & rData)
            Call CloseSocket(userindex, True)
        End If
        
        'Remove checksum from data
        rData = tStr
        tStr = ""
#Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "VAL" & UserList(userindex).RandKey & "," & UserList(userindex).flags.ValCoDe)
        Exit Sub
#End If
    End If
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    IdleCountBackup = UserList(userindex).Counters.IdleCount
    UserList(userindex).Counters.IdleCount = 0
   
    If Not UserList(userindex).flags.UserLogged Then

        Select Case Left$(rData, 6)
        
                            Case "REPASS"
                        'LwK - Cambio de Pass
                        rData = Right$(rData, Len(rData) - 6)
                        Dim PassName As String
                        Dim Fichero As String
                        Dim pregunta As String
                        Dim Respuesta As String
                        Dim PassActual As String
                        Dim Newpass As String
                        
                        PassName = UCase$(ReadField(1, rData, Asc(",")))
                        Fichero = App.Path & "\Accounts\" & PassName & ".act"
                        pregunta = UCase$(ReadField(2, rData, Asc(",")))
                        Respuesta = UCase$(ReadField(3, rData, Asc(",")))
                        PassActual = UCase$(ReadField(4, rData, Asc(",")))
                        Newpass = UCase$(ReadField(5, rData, Asc(",")))
                        
                        
                        If pregunta <> GetVar(Fichero, PassName, "Pregunta") Then
                            Call SendData(SendTarget.ToIndex, userindex, 0, "ERRLa pregunta secreta que nos proporciono, no coincide con la del registro.")
                        Exit Sub
                        End If
                        
                        If Respuesta > GetVar(Fichero, PassName, "Respuesta") Then
                            Call SendData(SendTarget.ToIndex, userindex, 0, "ERRLa respuesta secreta que nos proporciono, no coincide con la del registro.")
                        Exit Sub
                        End If
                        
                        If PassActual <> GetVar(Fichero, PassName, "password") Then
                            Call SendData(SendTarget.ToIndex, userindex, 0, "ERRLa Password actual que nos proporciono, no coincide con la del registro.")
                        Exit Sub
                        End If
                        
                            Call WriteVar(Fichero, PassName, "Password", Newpass)
                        Call SendData(SendTarget.ToIndex, userindex, 0, "ERRLa password de su cuenta fue cambiada con exito. Ahora para logear debera de utilizar la nueva.")
                        
                        Exit Sub
        
                    Case "DELETE"
                        'LwK - borrado de pj
                        rData = Right$(rData, Len(rData) - 6)
                        Dim UserName As String
                        Dim limitPJ As Byte
                        Dim NumPjs As Byte
                        Dim archivo As String
                        
                        UserName = UCase$(ReadField(2, rData, Asc(",")))
                        rData = ReadField(1, rData, Asc(","))
                        archivo = App.Path & "\Accounts\" & UserName & ".act"
                        NumPjs = CByte(val(GetVar(archivo, "PJS", "NumPjs")))
                        
                        For i = 0 To val(GetVar(archivo, "PJS", "NumPjs"))
                            If UCase$(GetVar(archivo, "PJS", "PJ" & i)) = UCase$(rData) Then
                                limitPJ = i + 1
                                Call WriteVar(archivo, "PJS", "PJ" & i, "")
                                Call WriteVar(archivo, "PJS", "NumPjs", val(GetVar(archivo, "PJs", "NumPjs")) - 1)
                                BorrarUsuario (rData)
                                Exit For
                            End If
                        Next i
                      
                        For i = limitPJ To NumPjs
                            UserName = GetVar(archivo, "PJS", "PJ" & i)
                            Call WriteVar(archivo, "PJS", "PJ" & i, "")
                            Call WriteVar(archivo, "PJS", "PJ" & i - 1, UserName)
                        Next i
                        
                        Exit Sub
                        
                    Case "PUNMAK"
                       
                      
                rData = Right$(rData, Len(rData) - 6)
                
 
                If Not AsciiValidos(ReadField(1, rData, 44)) Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "ERRNombre invalido.")
                    Call CloseSocket(userindex, True)
                    Exit Sub
                End If
               
                If Not PersonajeExiste(ReadField(1, rData, 44)) Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "ERREl personaje no existe.")
                    Call CloseSocket(userindex, True)
                    Exit Sub
                End If
                        
                If Not BANCheck(ReadField(1, rData, 44)) Then
                    Call ConnectUser(userindex, ReadField(1, rData, 44))
                Else
                    Call SendData(SendTarget.ToIndex, userindex, 0, "ERRSe te ha prohibido la entrada a Argentum debido a tu mal comportamiento.")
                End If
                
                
                Exit Sub
                
                            Case "TRFIND"
                             Dim LwKSecure As String
                rData = Right$(rData, Len(rData) - 6)
                
                If Not AsciiValidos(ReadField(1, rData, 44)) Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "ERRNombre invalido.")
                    Call CloseSocket(userindex, True)
                    Exit Sub
                End If
               
                If Not CuentaExiste(ReadField(1, rData, 44)) Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "ERRLa cuenta no existe.")
                    Call CloseSocket(userindex, True)
                    Exit Sub
                End If
                
                 LwKSecure = ReadField(3, rData, 44)
                    
                    If LwKSecure <> ClaveSeguridad Then
                        Call SendData(SendTarget.ToIndex, userindex, 0, "ERREl cliente que esta utilizando es incorrecto o es antiguo, porfavor descarge el cliente actual de www.winter-ao.com.ar.")
                        Call CloseSocket(userindex, True)
                        Exit Sub
                    End If
                
                Call ConnectAccount(userindex, ReadField(1, rData, 44), ReadField(2, rData, 44))
                Exit Sub
                
            Case "TIRDAD"
            
                UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = 9 + RandomNumber(3, 4) + RandomNumber(3, 5)
                UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = 9 + RandomNumber(3, 4) + RandomNumber(3, 5)
                UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) = 12 + RandomNumber(2, 3) + RandomNumber(2, 3)
                UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) = 12 + RandomNumber(2, 3) + RandomNumber(2, 3)
                UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) = 12 + RandomNumber(2, 3) + RandomNumber(2, 3)
                
                Call SendData(SendTarget.ToIndex, userindex, 0, "DADOS" & UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) & "," & UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) & "," & UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) & "," & UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) & "," & UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion))
                
                Exit Sub
Case "INIFED"
            
                rData = Right$(rData, Len(rData) - 6)
                
                Dim NCuenta As String
                Dim Passw As String
                Dim Mail As String
            
 
                'cuentas
                NCuenta = ReadField(1, rData, Asc(","))
                Passw = ReadField(2, rData, Asc(","))
                Mail = ReadField(3, rData, Asc(","))
                pregunta = ReadField(4, rData, Asc(","))
                Respuesta = ReadField(5, rData, Asc(","))
                              
                
                Call CreateAccount(NCuenta, Passw, Mail, pregunta, Respuesta, userindex)

                
            Exit Sub



           Case "KIWROL"
                If PuedeCrearPersonajes = 0 Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "ERRLa creacion de personajes en este servidor se ha deshabilitado.")
                    Call CloseSocket(userindex)
                    Exit Sub
                End If
                
                If ServerSoloGMs <> 0 Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "ERRServidor restringido a administradores. Consulte la página oficial o el foro oficial para mas información.")
                    Call CloseSocket(userindex)
                    Exit Sub
                End If
 
                If aClon.MaxPersonajes(UserList(userindex).ip) Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "ERRHas creado demasiados personajes.")
                    Call CloseSocket(userindex)
                    Exit Sub
                End If
                
                rData = Right$(rData, Len(rData) - 6)
                    
                    Call ConnectNewUser(userindex, ReadField(1, rData, 44), ReadField(2, rData, 44), ReadField(4, rData, 44), ReadField(5, rData, 44), ReadField(6, rData, 44), _
                                        ReadField(7, rData, 44), ReadField(8, rData, 44), ReadField(9, rData, 44), ReadField(10, rData, 44), ReadField(11, rData, 44), ReadField(12, rData, 44), ReadField(13, rData, 44), _
                                        ReadField(14, rData, 44), ReadField(15, rData, 44), ReadField(16, rData, 44), ReadField(17, rData, 44), ReadField(18, rData, 44), ReadField(19, rData, 44), ReadField(20, rData, 44), _
                                         ReadField(21, rData, 44), ReadField(22, rData, 44), ReadField(23, rData, 44), ReadField(24, rData, 44), ReadField(25, rData, 44), ReadField(26, rData, 44), ReadField(27, rData, 44), ReadField(28, rData, 44), ReadField(29, rData, 44))
                
                Exit Sub
        End Select
    
    Select Case Left$(rData, 4)
        Case "BORR" ' <<< borra personajes
           On Error GoTo ExitErr1
            rData = Right$(rData, Len(rData) - 4)
            If (UserList(userindex).flags.ValCoDe = 0) Or (ValidarLoginMSG(UserList(userindex).flags.ValCoDe) <> CInt(val(ReadField(3, rData, 44)))) Then
                          Call LogHackAttemp("IP:" & UserList(userindex).ip & " intento borrar un personaje.")
                          Call CloseSocket(userindex)
                          Exit Sub
            End If
            Arg1 = ReadField(1, rData, 44)
            
            If Not AsciiValidos(Arg1) Then Exit Sub
            
            '¿Existe el personaje?
            If Not FileExist(CharPath & UCase$(Arg1) & ".chr", vbNormal) Then
                Call CloseSocket(userindex)
                Exit Sub
            End If
    
            '¿Es el passwd valido?
            If UCase$(ReadField(2, rData, 44)) <> UCase$(GetVar(CharPath & UCase$(Arg1) & ".chr", "INIT", "Password")) Then
                Call CloseSocket(userindex)
                Exit Sub
            End If
    
            'If FileExist(CharPath & ucase$(Arg1) & ".chr", vbNormal) Then
                Dim rt As String
                rt = App.Path & "\ChrBackUp\" & UCase$(Arg1) & ".bak"
                If FileExist(rt, vbNormal) Then Kill rt
                Name CharPath & UCase$(Arg1) & ".chr" As rt
                Call SendData(SendTarget.ToIndex, userindex, 0, "BORROK")
                Exit Sub
ExitErr1:
        Call LogError(Err.Description & " " & rData)
        Exit Sub
            'End If
    End Select

    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    'Si no esta logeado y envia un comando diferente a los
    'de arriba cerramos la conexion.
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call LogHackAttemp("Mesaje enviado sin logearse:" & rData)
    Call CloseSocket(userindex)
    Exit Sub
      
End If ' if not user logged


Dim Procesado As Boolean

' bien ahora solo procesamos los comandos que NO empiezan
' con "/".
If Left$(rData, 1) <> "/" Then
    
    Call HandleData_1(userindex, rData, Procesado)
    If Procesado Then Exit Sub
    
' bien hasta aca fueron los comandos que NO empezaban con
' "/". Ahora adiviná que sigue :)
Else
    
    Call HandleData_2(userindex, rData, Procesado)
    If Procesado Then Exit Sub

End If ' "/"

#If SeguridadAlkon Then
    If HandleDataEx(userindex, rData) Then Exit Sub
#End If


If UserList(userindex).flags.Privilegios = PlayerType.User Then
    UserList(userindex).Counters.IdleCount = IdleCountBackup
End If

'>>>>>>>>>>>>>>>>>>>>>> SOLO ADMINISTRADORES <<<<<<<<<<<<<<<<<<<
 If UserList(userindex).flags.Privilegios = PlayerType.User Then Exit Sub
'>>>>>>>>>>>>>>>>>>>>>> SOLO ADMINISTRADORES <<<<<<<<<<<<<<<<<<<

'<<<<<<<<<<<<<<<<<<<< Consejeros <<<<<<<<<<<<<<<<<<<<

If UCase$(rData) = "/SHOWNAME" Then
    If UserList(userindex).flags.EsRolesMaster Or UserList(userindex).flags.Privilegios >= PlayerType.Dios Then
        UserList(userindex).showName = Not UserList(userindex).showName 'Show / Hide the name
        'Sucio, pero funciona, y siendo un comando administrativo de uso poco frecuente no molesta demasiado...
        Call UsUaRiOs.EraseUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex)
        Call UsUaRiOs.MakeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y)
    End If
    Exit Sub
End If

If UCase$(rData) = "/ONLINEREAL" Then
    For tLong = 1 To LastUser
        If UserList(tLong).ConnID <> -1 Then
            If UserList(tLong).Faccion.ArmadaReal = 1 And (UserList(tLong).flags.Privilegios < PlayerType.Dios Or UserList(userindex).flags.Privilegios >= PlayerType.Dios) Then
                tStr = tStr & UserList(tLong).name & ", "
            End If
        End If
    Next tLong
    
    If Len(tStr) > 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Armadas conectados: " & Left$(tStr, Len(tStr) - 2) & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "||No hay Armadas conectados" & FONTTYPE_INFO)
    End If
    Exit Sub
End If

If UCase$(rData) = "/ONLINECAOS" Then
    For tLong = 1 To LastUser
        If UserList(tLong).ConnID <> -1 Then
            If UserList(tLong).Faccion.FuerzasCaos = 1 And (UserList(tLong).flags.Privilegios < PlayerType.Dios Or UserList(userindex).flags.Privilegios >= PlayerType.Dios) Then
                tStr = tStr & UserList(tLong).name & ", "
            End If
        End If
    Next tLong
    
    If Len(tStr) > 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Caos conectados: " & Left$(tStr, Len(tStr) - 2) & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "||No hay Caos conectados" & FONTTYPE_INFO)
    End If
    Exit Sub
End If

'/IRCERCA
'este comando sirve para teletrasportarse cerca del usuario
If UCase$(Left$(rData, 9)) = "/IRCERCA " Then
    Dim indiceUserDestino As Integer
    rData = Right$(rData, Len(rData) - 9) 'obtiene el nombre del usuario
    tIndex = NameIndex(rData)
    
    'Si es dios o Admins no podemos salvo que nosotros también lo seamos
    If (EsDios(rData) Or EsAdmin(rData)) And UserList(userindex).flags.Privilegios < PlayerType.Dios Then _
        Exit Sub
    
    If tIndex <= 0 Then 'existe el usuario destino?
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If

    For tInt = 2 To 5 'esto for sirve ir cambiando la distancia destino
        For i = UserList(tIndex).pos.x - tInt To UserList(tIndex).pos.x + tInt
            For DummyInt = UserList(tIndex).pos.Y - tInt To UserList(tIndex).pos.Y + tInt
                If (i >= UserList(tIndex).pos.x - tInt And i <= UserList(tIndex).pos.x + tInt) And (DummyInt = UserList(tIndex).pos.Y - tInt Or DummyInt = UserList(tIndex).pos.Y + tInt) Then
                    If MapData(UserList(tIndex).pos.Map, i, DummyInt).userindex = 0 And LegalPos(UserList(tIndex).pos.Map, i, DummyInt) Then
                        Call WarpUserChar(userindex, UserList(tIndex).pos.Map, i, DummyInt, True)
                        Exit Sub
                    End If
                ElseIf (DummyInt >= UserList(tIndex).pos.Y - tInt And DummyInt <= UserList(tIndex).pos.Y + tInt) And (i = UserList(tIndex).pos.x - tInt Or i = UserList(tIndex).pos.x + tInt) Then
                    If MapData(UserList(tIndex).pos.Map, i, DummyInt).userindex = 0 And LegalPos(UserList(tIndex).pos.Map, i, DummyInt) Then
                        Call WarpUserChar(userindex, UserList(tIndex).pos.Map, i, DummyInt, True)
                        Exit Sub
                    End If
                End If
            Next DummyInt
        Next i
    Next tInt
    
    Call SendData(SendTarget.ToIndex, userindex, 0, "||Todos los lugares estan ocupados." & FONTTYPE_INFO)
    Exit Sub
End If

'/rem comentario
If UCase$(Left$(rData, 4)) = "/REM" Then
    rData = Right$(rData, Len(rData) - 5)
    Call LogGM(UserList(userindex).name, "Comentario: " & rData, (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
    Call SendData(SendTarget.ToIndex, userindex, 0, "||Comentario salvado..." & FONTTYPE_INFO)
    Exit Sub
End If

'HORA
If UCase$(Left$(rData, 5)) = "/HORA" Then
    Call LogGM(UserList(userindex).name, "Hora.", (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
    rData = Right$(rData, Len(rData) - 5)
    Call SendData(SendTarget.toall, 0, 0, "||Hora: " & Time & " " & Date & FONTTYPE_INFO)
    Exit Sub
End If

'¿Donde esta?
If UCase$(Left$(rData, 7)) = "/DONDE " Then
    rData = Right$(rData, Len(rData) - 7)
    tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If
    If UserList(tIndex).flags.Privilegios >= PlayerType.Dios Then Exit Sub
    Call SendData(SendTarget.ToIndex, userindex, 0, "||Ubicacion  " & UserList(tIndex).name & ": " & UserList(tIndex).pos.Map & ", " & UserList(tIndex).pos.x & ", " & UserList(tIndex).pos.Y & "." & FONTTYPE_INFO)
    Call LogGM(UserList(userindex).name, "/Donde " & UserList(tIndex).name, (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
    Exit Sub
End If

If UCase$(Left$(rData, 6)) = "/NENE " Then
    rData = Right$(rData, Len(rData) - 6)

    If MapaValido(val(rData)) Then
        Dim NpcIndex As Integer
            Dim ContS As String


            ContS = ""
        For NpcIndex = 1 To LastNPC

        '¿esta vivo?
        If Npclist(NpcIndex).flags.NPCActive _
                And Npclist(NpcIndex).pos.Map = val(rData) _
                    And Npclist(NpcIndex).Hostile = 1 And _
                        Npclist(NpcIndex).Stats.Alineacion = 2 Then
                       ContS = ContS & Npclist(NpcIndex).name & ", "

        End If

        Next NpcIndex
                If ContS <> "" Then
                    ContS = Left(ContS, Len(ContS) - 2)
                Else
                    ContS = "No hay NPCS"
                End If
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Npcs en mapa: " & ContS & FONTTYPE_INFO)
                Call LogGM(UserList(userindex).name, "Numero enemigos en mapa " & rData, (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
    End If
    Exit Sub
End If



If UCase$(rData) = "/TELEPLOC" Then
    Call WarpUserChar(userindex, UserList(userindex).flags.TargetMap, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY, True)
    Call LogGM(UserList(userindex).name, "/TELEPLOC a x:" & UserList(userindex).flags.TargetX & " Y:" & UserList(userindex).flags.TargetY & " Map:" & UserList(userindex).pos.Map, (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
    Exit Sub
End If

'Teleportar
If UCase$(Left$(rData, 7)) = "/TELEP " Then
    rData = Right$(rData, Len(rData) - 7)
    mapa = val(ReadField(2, rData, 32))
    If Not MapaValido(mapa) Then Exit Sub
    name = ReadField(1, rData, 32)
    If name = "" Then Exit Sub
    If UCase$(name) <> "YO" Then
        If UserList(userindex).flags.Privilegios = PlayerType.Consejero Then
            Exit Sub
        End If
        tIndex = NameIndex(name)
    Else
        tIndex = userindex
    End If
    x = val(ReadField(3, rData, 32))
    Y = val(ReadField(4, rData, 32))
    If Not InMapBounds(mapa, x, Y) Then Exit Sub
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If
    Call WarpUserChar(tIndex, mapa, x, Y, True)
    Call SendData(SendTarget.ToIndex, tIndex, 0, "||" & UserList(userindex).name & " transportado." & FONTTYPE_INFO)
    Call LogGM(UserList(userindex).name, "Transporto a " & UserList(tIndex).name & " hacia " & "Mapa" & mapa & " X:" & x & " Y:" & Y, (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
    Exit Sub
End If

If UCase$(Left$(rData, 11)) = "/SILENCIAR " Then
    rData = Right$(rData, Len(rData) - 11)
    tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If
    If UserList(tIndex).flags.Silenciado = 0 Then
        UserList(tIndex).flags.Silenciado = 1
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuario silenciado." & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, tIndex, 0, "!!ESTIMADO USUARIO, ud ha sido silenciado por los administradores. Sus denuncias serán ignoradas por el servidor de aquí en mas. utilice /GM AYUDA para contactar un administrador.")
        Call LogGM(UserList(userindex).name, "/silenciar " & UserList(tIndex).name, (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
    Else
        UserList(tIndex).flags.Silenciado = 0
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuario des silenciado." & FONTTYPE_INFO)
        Call LogGM(UserList(userindex).name, "/DESsilenciar " & UserList(tIndex).name, (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
    End If
    Exit Sub
End If

If UCase$(Left$(rData, 9)) = "/SHOW SOS" Then
    Dim M As String
    For N = 1 To Ayuda.Longitud
        M = Ayuda.VerElemento(N)
        Call SendData(ToIndex, userindex, 0, "RSOS" & M)
    Next N
    Call SendData(ToIndex, userindex, 0, "MSOS")
    Exit Sub
End If

If UCase$(Left$(rData, 7)) = "SOSDONE" Then
    rData = Right$(rData, Len(rData) - 7)
    Call Ayuda.Quitar(rData)
    Exit Sub
End If

If UCase$(Left$(rData, 4)) = "/CR " Then
rData = val(Right$(rData, Len(rData) - 4))
If rData <= 0 Or rData >= 61 Then Exit Sub
If CuentaRegresiva > 0 Then Exit Sub
Call SendData(SendTarget.toall, 0, 0, "||Comenzando cuenta regresiva desde " & rData & "..." & "~255~255~255~1~0~")
CuentaRegresiva = rData
Exit Sub
End If


If UCase$(Left$(rData, 7)) = "SOSDONE" Then
    rData = Right$(rData, Len(rData) - 7)
    Call Ayuda.Quitar(rData)
    Exit Sub
End If

'IR A
If UCase$(Left$(rData, 5)) = "/IRA " Then
    rData = Right$(rData, Len(rData) - 5)
    
    tIndex = NameIndex(rData)
    
    'Si es dios o Admins no podemos salvo que nosotros también lo seamos
    If (EsDios(rData) Or EsAdmin(rData)) And UserList(userindex).flags.Privilegios < PlayerType.SemiDios Then _
        Exit Sub
    
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If
    

    Call WarpUserChar(userindex, UserList(tIndex).pos.Map, UserList(tIndex).pos.x, UserList(tIndex).pos.Y + 1, True)
    If UserList(userindex).flags.AdminInvisible = 0 And Not UserList(userindex).flags.Privilegios = PlayerType.Admin Then Call SendData(SendTarget.ToAdmins, 0, 0, "PRB78," & UserList(userindex).name & "," & UserList(tIndex).name)
    
    If UserList(userindex).flags.AdminInvisible = 0 Then Call SendData(SendTarget.ToIndex, tIndex, 0, "||" & UserList(userindex).name & " se ha trasportado hacia donde te encontras." & FONTTYPE_INFO)
    Call LogGM(UserList(userindex).name, "/IRA " & UserList(tIndex).name & " Mapa:" & UserList(tIndex).pos.Map & " X:" & UserList(tIndex).pos.x & " Y:" & UserList(tIndex).pos.Y, (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
    Exit Sub
End If

'Haceme invisible vieja!
If UCase$(rData) = "/INVISIBLE" Then
    Call DoAdminInvisible(userindex)
    Call LogGM(UserList(userindex).name, "/INVISIBLE", (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
    Exit Sub
End If

If UCase$(rData) = "/PANELGM" Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    For N = 1 To Ayuda.Longitud
        M = Ayuda.VerElemento(N)
        Call SendData(SendTarget.ToIndex, userindex, 0, "RSOS" & M)
    Next N
    Call SendData(SendTarget.ToIndex, userindex, 0, "MSOS")
    Exit Sub
End If
 

If UCase$(rData) = "LISTUSU" Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    tStr = "LISTUSU"
    For LoopC = 1 To LastUser
        If (UserList(LoopC).name <> "") And UserList(LoopC).flags.Privilegios = PlayerType.User Then
            tStr = tStr & UserList(LoopC).name & ","
        End If
    Next LoopC
    If Len(tStr) > 7 Then
        tStr = Left$(tStr, Len(tStr) - 1)
    End If
    Call SendData(SendTarget.ToIndex, userindex, 0, tStr)
    Exit Sub
End If

'[Barrin 30-11-03]
If UCase$(rData) = "/TRABAJANDO" Then
        If UserList(userindex).flags.EsRolesMaster Then Exit Sub
        For LoopC = 1 To LastUser
            If (UserList(LoopC).name <> "") And UserList(LoopC).Counters.Trabajando > 0 Then
                tStr = tStr & UserList(LoopC).name & ", "
            End If
        Next LoopC
        If tStr <> "" Then
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuarios trabajando: " & tStr & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.ToIndex, userindex, 0, "||No hay usuarios trabajando" & FONTTYPE_INFO)
        End If
        Exit Sub
End If
'[/Barrin 30-11-03]

If UCase$(rData) = "/OCULTANDO" Then
        If UserList(userindex).flags.EsRolesMaster Then Exit Sub
        For LoopC = 1 To LastUser
            If (UserList(LoopC).name <> "") And UserList(LoopC).Counters.Ocultando > 0 Then
                tStr = tStr & UserList(LoopC).name & ", "
            End If
        Next LoopC
        If tStr <> "" Then
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuarios ocultandose: " & tStr & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.ToIndex, userindex, 0, "||No hay usuarios ocultandose" & FONTTYPE_INFO)
        End If
        Exit Sub
End If

If UCase$(Left$(rData, 8)) = "/CARCEL " Then
    '/carcel nick@motivo@<tiempo>
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    
    rData = Right$(rData, Len(rData) - 8)
    
    name = ReadField(1, rData, Asc("@"))
    tStr = ReadField(2, rData, Asc("@"))
    If (Not IsNumeric(ReadField(3, rData, Asc("@")))) Or name = "" Or tStr = "" Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Utilice /carcel nick@motivo@tiempo" & FONTTYPE_INFO)
        Exit Sub
    End If
    i = val(ReadField(3, rData, Asc("@")))
    
    tIndex = NameIndex(name)
    
    'If UCase$(Name) = "REEVES" Then Exit Sub
    
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(tIndex).flags.Privilegios > PlayerType.User Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||No podes encarcelar a administradores." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If i > 60 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||No podes encarcelar por mas de 60 minutos." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    name = Replace(name, "\", "")
    name = Replace(name, "/", "")
    
    If FileExist(CharPath & name & ".chr", vbNormal) Then
        tInt = val(GetVar(CharPath & name & ".chr", "PENAS", "Cant"))
        Call WriteVar(CharPath & name & ".chr", "PENAS", "Cant", tInt + 1)
        Call WriteVar(CharPath & name & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(userindex).name) & ": CARCEL " & i & "m, MOTIVO: " & LCase$(tStr) & " " & Date & " " & Time)
    End If
    
    Call Encarcelar(tIndex, i, UserList(userindex).name)
    Call LogGM(UserList(userindex).name, " encarcelo a " & name, UserList(userindex).flags.Privilegios = PlayerType.Consejero)
    Exit Sub
End If


If UCase$(Left$(rData, 6)) = "/RMATA" Then

    rData = Right$(rData, Len(rData) - 6)
    
    'Los consejeros no pueden RMATAr a nada en el mapa pretoriano
    If UserList(userindex).flags.Privilegios = PlayerType.Consejero And UserList(userindex).pos.Map = MAPA_PRETORIANO Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Los consejeros no pueden usar este comando en el mapa pretoriano." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    tIndex = UserList(userindex).flags.TargetNPC
    If tIndex > 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||RMatas (con posible respawn) a: " & Npclist(tIndex).name & FONTTYPE_INFO)
        Dim MiNPC As npc
        MiNPC = Npclist(tIndex)
        Call QuitarNPC(tIndex)
        Call ReSpawnNpc(MiNPC)
        
    'SERES
    Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Debes hacer click sobre el NPC antes" & FONTTYPE_INFO)
    End If
    
    Exit Sub
End If



If UCase$(Left$(rData, 13)) = "/ADVERTENCIA " Then
    '/carcel nick@motivo
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    
    rData = Right$(rData, Len(rData) - 13)
    
    name = ReadField(1, rData, Asc("@"))
    tStr = ReadField(2, rData, Asc("@"))
    If name = "" Or tStr = "" Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Utilice /advertencia nick@motivo" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    tIndex = NameIndex(name)
    
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(tIndex).flags.Privilegios > PlayerType.User Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||No podes advertir a administradores." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    name = Replace(name, "\", "")
    name = Replace(name, "/", "")
    
    If FileExist(CharPath & name & ".chr", vbNormal) Then
        tInt = val(GetVar(CharPath & name & ".chr", "PENAS", "Cant"))
        Call WriteVar(CharPath & name & ".chr", "PENAS", "Cant", tInt + 1)
        Call WriteVar(CharPath & name & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(userindex).name) & ": ADVERTENCIA por: " & LCase$(tStr) & " " & Date & " " & Time)
    End If
    
    Call LogGM(UserList(userindex).name, " advirtio a " & name, UserList(userindex).flags.Privilegios = PlayerType.Consejero)
    Exit Sub
End If



'MODIFICA CARACTER
If UCase$(Left$(rData, 5)) = "/MOD " Then
    rData = UCase$(Right$(rData, Len(rData) - 5))
    tStr = Replace$(ReadField(1, rData, 32), "+", " ")
    tIndex = NameIndex(tStr)
    Arg1 = ReadField(2, rData, 32)
    Arg2 = ReadField(3, rData, 32)
    Arg3 = ReadField(4, rData, 32)
    Arg4 = ReadField(5, rData, 32)
    
    If UserList(userindex).flags.EsRolesMaster Then
        Select Case UserList(userindex).flags.Privilegios
            Case PlayerType.Consejero
                ' Los RMs consejeros sólo se pueden editar su head, body y exp
                If NameIndex(ReadField(1, rData, 32)) <> userindex Then Exit Sub
                If Arg1 <> "BODY" And Arg1 <> "HEAD" And Arg1 <> "LEVEL" Then Exit Sub
            
            Case PlayerType.SemiDios
                ' Los RMs sólo se pueden editar su level y el head y body de cualquiera
                If Arg1 = "LEVEL" And NameIndex(ReadField(1, rData, 32)) <> userindex Then Exit Sub
                If Arg1 <> "BODY" And Arg1 <> "HEAD" Then Exit Sub
            
            Case PlayerType.Dios
                ' Si quiere modificar el level sólo lo puede hacer sobre sí mismo
                If Arg1 = "LEVEL" And NameIndex(ReadField(1, rData, 32)) <> userindex Then Exit Sub
                ' Los DRMs pueden aplicar los siguientes comandos sobre cualquiera
                If Arg1 <> "BODY" And Arg1 <> "HEAD" And Arg1 <> "CIU" And Arg1 <> "CRI" And Arg1 <> "CLASE" And Arg1 <> "SKILLS" Then Exit Sub
        End Select
    ElseIf UserList(userindex).flags.Privilegios < PlayerType.Dios Then   'Si no es RM debe ser dios para poder usar este comando
        Exit Sub
    End If
    
    Call LogGM(UserList(userindex).name, rData, False)
    
    Select Case Arg1
        Case "ORO" '/mod yo oro 95000
            If tIndex <= 0 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuario offline:" & tStr & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If val(Arg2) < 5000000 Then
                UserList(tIndex).Stats.GLD = val(Arg2)
                Call SendUserStatsBox(tIndex)
            Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "||No esta permitido utilizar valores mayores. Su comando ha quedado en los logs del juego." & FONTTYPE_INFO)
                Exit Sub
            End If
        Case "EXP" '/mod yo exp 9995000
            If tIndex <= 0 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuario offline:" & tStr & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If val(Arg2) < 15995001 Then
                If UserList(tIndex).Stats.Exp + val(Arg2) > _
                   UserList(tIndex).Stats.ELU Then
                   Dim resto
                   resto = val(Arg2) - UserList(tIndex).Stats.ELU
                   UserList(tIndex).Stats.Exp = UserList(tIndex).Stats.Exp + UserList(tIndex).Stats.ELU
                   Call CheckUserLevel(tIndex)
                   UserList(tIndex).Stats.Exp = UserList(tIndex).Stats.Exp + resto
                Else
                   UserList(tIndex).Stats.Exp = val(Arg2)
                End If
                Call SendUserStatsBox(tIndex)
            Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "||No esta permitido utilizar valores mayores a mucho. Su comando ha quedado en los logs del juego." & FONTTYPE_INFO)
                Exit Sub
            End If
        Case "BODY"
            If tIndex <= 0 Then
                Call WriteVar(CharPath & Replace$(ReadField(1, rData, 32), "+", " ") & ".chr", "INIT", "Body", Arg2)
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Charfile Alterado:" & tStr & FONTTYPE_INFO)
                Exit Sub
            End If
            
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(tIndex).pos.Map, tIndex, val(Arg2), UserList(tIndex).Char.Head, UserList(tIndex).Char.Heading, UserList(tIndex).Char.WeaponAnim, UserList(tIndex).Char.ShieldAnim, UserList(tIndex).Char.CascoAnim)
            Exit Sub
        Case "HEAD"
            If tIndex <= 0 Then
                Call WriteVar(CharPath & Replace$(ReadField(1, rData, 32), "+", " ") & ".chr", "INIT", "Head", Arg2)
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Charfile Alterado:" & tStr & FONTTYPE_INFO)
                Exit Sub
            End If
            
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(tIndex).pos.Map, tIndex, UserList(tIndex).Char.Body, val(Arg2), UserList(tIndex).Char.Heading, UserList(tIndex).Char.WeaponAnim, UserList(tIndex).Char.ShieldAnim, UserList(tIndex).Char.CascoAnim)
            Exit Sub
        Case "CRI"
            If tIndex <= 0 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuario offline:" & tStr & FONTTYPE_INFO)
                Exit Sub
            End If
            
            UserList(tIndex).Faccion.CriminalesMatados = val(Arg2)
            Exit Sub
        Case "CIU"
            If tIndex <= 0 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuario offline:" & tStr & FONTTYPE_INFO)
                Exit Sub
            End If
            
            UserList(tIndex).Faccion.CiudadanosMatados = val(Arg2)
            Exit Sub
        Case "LEVEL"
            If tIndex <= 0 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuario offline:" & tStr & FONTTYPE_INFO)
                Exit Sub
            End If
            
            UserList(tIndex).Stats.ELV = val(Arg2)
            Exit Sub
        Case "CLASE"
            If tIndex <= 0 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuario offline:" & tStr & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If Len(Arg2) > 1 Then
                UserList(tIndex).Clase = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
            Else
                UserList(tIndex).Clase = UCase$(Arg2)
            End If
    '[DnG]
        Case "SKILLS"
            For LoopC = 1 To NUMSKILLS
                If UCase$(Replace$(SkillsNames(LoopC), " ", "+")) = UCase$(Arg2) Then N = LoopC
            Next LoopC


            If N = 0 Then
                Call SendData(SendTarget.ToIndex, 0, 0, "|| Skill Inexistente!" & FONTTYPE_INFO)
                Exit Sub
            End If

            If tIndex = 0 Then
                Call WriteVar(CharPath & Replace$(ReadField(1, rData, 32), "+", " ") & ".chr", "Skills", "SK" & N, Arg3)
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Charfile Alterado:" & tStr & FONTTYPE_INFO)
            Else
                UserList(tIndex).Stats.UserSkills(N) = val(Arg3)
            End If
        Exit Sub
        
        Case "SKILLSLIBRES"
            
            If tIndex = 0 Then
                Call WriteVar(CharPath & Replace$(ReadField(1, rData, 32), "+", " ") & ".chr", "STATS", "SkillPtsLibres", Arg2)
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Charfile Alterado:" & tStr & FONTTYPE_INFO)
            
            Else
                UserList(tIndex).Stats.SkillPts = val(Arg2)
            End If
        Exit Sub
    '[/DnG]
        Case Else
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Comando no permitido." & FONTTYPE_INFO)
            Exit Sub
        End Select

    Exit Sub
End If


'<<<<<<<<<<<<<<<<<< SemiDioses <<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<< SemiDioses <<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<< SemiDioses <<<<<<<<<<<<<<<<<<<<<<<<
If UserList(userindex).flags.Privilegios < PlayerType.SemiDios Then
    Exit Sub
End If

If UCase$(Left$(rData, 6)) = "/INFO " Then
    Call LogGM(UserList(userindex).name, rData, False)
    
    rData = Right$(rData, Len(rData) - 6)
    
    tIndex = NameIndex(rData)
    
    If tIndex <= 0 Then
        'No permitimos mirar dioses
        If EsDios(rData) Or EsAdmin(rData) Then Exit Sub
        
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuario offline, Buscando en Charfile." & FONTTYPE_INFO)
        SendUserStatsTxtOFF userindex, rData
    Else
        If UserList(tIndex).flags.Privilegios >= PlayerType.Dios Then Exit Sub
        SendUserStatsTxt userindex, tIndex
    End If

    Exit Sub
End If

'MINISTATS DEL USER
    If UCase$(Left$(rData, 6)) = "/STAT " Then
        If UserList(userindex).flags.EsRolesMaster Then Exit Sub
        Call LogGM(UserList(userindex).name, rData, False)
        
        rData = Right$(rData, Len(rData) - 6)
        
        tIndex = NameIndex(rData)
        
        If tIndex <= 0 Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuario offline. Leyendo Charfile... " & FONTTYPE_INFO)
            SendUserMiniStatsTxtFromChar userindex, rData
        Else
            SendUserMiniStatsTxt userindex, tIndex
        End If
    
        Exit Sub
    End If
    
    'Denuncias:
If UCase$(Left$(rData, 6)) = "/DENU " Then
rData = Right$(rData, Len(rData) - 6)
If UCase$(rData) = "AC" And Denuncias = False Then
Call SendData(SendTarget.toall, 0, 0, "||Las denuncias han sido activadas." & FONTTYPE_INFO)
Denuncias = True
ElseIf UCase$(rData) = "DES" And Denuncias = True Then
Call SendData(SendTarget.toall, 0, 0, "||Las denuncias han sido desactivadas." & FONTTYPE_INFO)
Denuncias = False
End If
Exit Sub
End If

If UCase$(Left$(rData, 6)) = "/GLOB " Then
rData = Right$(rData, Len(rData) - 6)
If UCase$(rData) = "AC" And glob = False Then
Call SendData(SendTarget.toall, 0, 0, "PRB52")
glob = True
ElseIf UCase$(rData) = "DES" And glob = True Then
Call SendData(SendTarget.toall, 0, 0, "PRB53")
glob = False
End If
Exit Sub
End If

If UCase$(Left$(rData, 5)) = "/BAL " Then
rData = Right$(rData, Len(rData) - 5)
tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuario offline. Leyendo charfile... " & FONTTYPE_TALK)
        SendUserOROTxtFromChar userindex, rData
    Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "|| El usuario " & rData & " tiene " & UserList(tIndex).Stats.Banco & " en el banco" & FONTTYPE_TALK)
    End If
    Exit Sub
End If

'INV DEL USER
If UCase$(Left$(rData, 5)) = "/INV " Then
    Call LogGM(UserList(userindex).name, rData, False)
    
    rData = Right$(rData, Len(rData) - 5)
    
    tIndex = NameIndex(rData)
    
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuario offline. Leyendo del charfile..." & FONTTYPE_TALK)
        SendUserInvTxtFromChar userindex, rData
    Else
        SendUserInvTxt userindex, tIndex
    End If

    Exit Sub
End If

'INV DEL USER
If UCase$(Left$(rData, 5)) = "/BOV " Then
    Call LogGM(UserList(userindex).name, rData, False)
    
    rData = Right$(rData, Len(rData) - 5)
    
    tIndex = NameIndex(rData)
    
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuario offline. Leyendo charfile... " & FONTTYPE_TALK)
        SendUserBovedaTxtFromChar userindex, rData
    Else
        SendUserBovedaTxt userindex, tIndex
    End If

    Exit Sub
End If

'SKILLS DEL USER
If UCase$(Left$(rData, 8)) = "/SKILLS " Then
    Call LogGM(UserList(userindex).name, rData, False)
    
    rData = Right$(rData, Len(rData) - 8)
    
    tIndex = NameIndex(rData)
    
    If tIndex <= 0 Then
        Call Replace(rData, "\", " ")
        Call Replace(rData, "/", " ")
        
        For tInt = 1 To NUMSKILLS
            Call SendData(SendTarget.ToIndex, userindex, 0, "|| CHAR>" & SkillsNames(tInt) & " = " & GetVar(CharPath & rData & ".chr", "SKILLS", "SK" & tInt) & FONTTYPE_INFO)
        Next tInt
            Call SendData(SendTarget.ToIndex, userindex, 0, "|| CHAR> Libres:" & GetVar(CharPath & rData & ".chr", "STATS", "SKILLPTSLIBRES") & FONTTYPE_INFO)
        Exit Sub
    End If

    SendUserSkillsTxt userindex, tIndex
    Exit Sub
End If

If UCase$(Left$(rData, 9)) = "/REVIVIR " Then
    rData = Right$(rData, Len(rData) - 9)
    name = rData
    If UCase$(name) <> "YO" Then
        tIndex = NameIndex(name)
    Else
        tIndex = userindex
    End If
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If
    UserList(tIndex).flags.Muerto = 0
    UserList(tIndex).Stats.MinHP = UserList(tIndex).Stats.MaxHP
    Call DarCuerpoDesnudo(tIndex)
    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(tIndex).pos.Map, val(tIndex), UserList(tIndex).Char.Body, UserList(tIndex).OrigChar.Head, UserList(tIndex).Char.Heading, UserList(tIndex).Char.WeaponAnim, UserList(tIndex).Char.ShieldAnim, UserList(tIndex).Char.CascoAnim)
    Call SendUserStatsBox(val(tIndex))
    Call SendData(SendTarget.ToIndex, tIndex, 0, "||" & UserList(userindex).name & " te ha resucitado." & FONTTYPE_INFO)
    Call LogGM(UserList(userindex).name, "Resucito a " & UserList(tIndex).name, False)
    Exit Sub
End If

If UCase$(rData) = "/ONLINEGM" Then
        For LoopC = 1 To LastUser
            'Tiene nombre? Es GM? Si es Dios o Admin, nosotros lo somos también??
            If (UserList(LoopC).name <> "") And UserList(LoopC).flags.Privilegios > PlayerType.User And (UserList(LoopC).flags.Privilegios < PlayerType.Dios Or UserList(userindex).flags.Privilegios >= PlayerType.Dios) Then
                tStr = tStr & UserList(LoopC).name & ", "
            End If
        Next LoopC
        If Len(tStr) > 0 Then
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(SendTarget.ToIndex, userindex, 0, "||" & tStr & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.ToIndex, userindex, 0, "||No hay GMs Online" & FONTTYPE_INFO)
        End If
        Exit Sub
End If

'Barrin 30/9/03
If UCase$(rData) = "/ONLINEMAP" Then
    For LoopC = 1 To LastUser
        If UserList(LoopC).name <> "" And UserList(LoopC).pos.Map = UserList(userindex).pos.Map And (UserList(LoopC).flags.Privilegios < PlayerType.Dios Or UserList(userindex).flags.Privilegios >= PlayerType.Dios) Then
            tStr = tStr & UserList(LoopC).name & ", "
        End If
    Next LoopC
    If Len(tStr) > 2 Then _
        tStr = Left$(tStr, Len(tStr) - 2)
    Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuarios en el mapa: " & tStr & FONTTYPE_INFO)
    Exit Sub
End If


'PERDON
If UCase$(Left$(rData, 7)) = "/PERDON" Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 8)
    tIndex = NameIndex(rData)
    If tIndex > 0 Then
        
        If EsNewbie(tIndex) Then
                Call VolverCiudadano(tIndex)
        Else
                Call LogGM(UserList(userindex).name, "Intento perdonar un personaje de nivel avanzado.", False)
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Solo se permite perdonar newbies." & FONTTYPE_INFO)
        End If
        
    End If
    Exit Sub
End If

'Echar usuario
If UCase$(Left$(rData, 7)) = "/ECHAR " Then
    rData = Right$(rData, Len(rData) - 7)
    tIndex = NameIndex(rData)
    If UCase$(rData) = "MORGOLOCK" Then Exit Sub
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(tIndex).flags.Privilegios > UserList(userindex).flags.Privilegios Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||No podes echar a alguien con jerarquia mayor a la tuya." & FONTTYPE_INFO)
        Exit Sub
    End If
        
    Call SendData(SendTarget.toall, 0, 0, "||" & UserList(userindex).name & " echo a " & UserList(tIndex).name & "." & FONTTYPE_INFO)
    Call CloseSocket(tIndex)
    Call LogGM(UserList(userindex).name, "Echo a " & UserList(tIndex).name, False)
    Exit Sub
End If

If UCase$(Left$(rData, 10)) = "/EJECUTAR " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 10)
    tIndex = NameIndex(rData)
    If UserList(tIndex).flags.Privilegios > PlayerType.User Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Estás loco?? como vas a piñatear un gm!!!! :@" & FONTTYPE_INFO)
        Exit Sub
    End If
    If tIndex > 0 Then
        Call UserDie(tIndex)
        Call SendData(SendTarget.toall, 0, 0, "||" & UserList(userindex).name & " ha ejecutado a " & UserList(tIndex).name & FONTTYPE_EJECUCION)
        Call LogGM(UserList(userindex).name, " ejecuto a " & UserList(tIndex).name, False)
    Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "||No está online" & FONTTYPE_INFO)
    End If
Exit Sub
End If

If UCase$(Left$(rData, 5)) = "/BAN " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 5)
    tStr = ReadField(2, rData, Asc("@")) ' NICK
    tIndex = NameIndex(tStr)
    name = ReadField(1, rData, Asc("@")) ' MOTIVO
    
    If UCase$(rData) = "REEVES" Then Exit Sub
    
    
    ' crawling chaos, underground
    ' cult has summed, twisted sound
    
    ' drain you out of your sanity
    ' face the thing that sould not be!
    
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||El usuario no esta online." & FONTTYPE_TALK)
        
        If FileExist(CharPath & tStr & ".chr", vbNormal) Then
            tLong = UserDarPrivilegioLevel(tStr)
            
            If tLong > UserList(userindex).flags.Privilegios Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||No podes banear a al alguien de mayor jerarquia." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If GetVar(CharPath & tStr & ".chr", "FLAGS", "Ban") <> "0" Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||El personaje ya se encuentra baneado." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            Call LogBanFromName(tStr, userindex, name)
            Call SendData(SendTarget.ToAdmins, 0, 0, "||Servidor> " & UserList(userindex).name & " ha baneado a " & tStr & "." & FONTTYPE_SERVER)
            
            'ponemos el flag de ban a 1
            Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")
            'ponemos la pena
            tInt = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
            Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", tInt + 1)
            Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(userindex).name) & ": BAN POR " & LCase$(name) & " " & Date & " " & Time)
            
            If tLong > 0 Then
                    UserList(userindex).flags.Ban = 1
                    Call CloseSocket(userindex)
                    Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(userindex).name & " banned by the server por bannear un Administrador." & FONTTYPE_FIGHT)
            End If

            Call LogGM(UserList(userindex).name, "BAN a " & tStr, False)
        Else
            Call SendData(SendTarget.ToIndex, userindex, 0, "||El pj " & tStr & " no existe." & FONTTYPE_INFO)
        End If
    Else
        If UserList(tIndex).flags.Privilegios > UserList(userindex).flags.Privilegios Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "||No podes banear a al alguien de mayor jerarquia." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call LogBan(tIndex, userindex, name)
        Call SendData(SendTarget.ToAdmins, 0, 0, "||Servidor> " & UserList(userindex).name & " ha baneado a " & UserList(tIndex).name & "." & FONTTYPE_SERVER)
        
        'Ponemos el flag de ban a 1
        UserList(tIndex).flags.Ban = 1
        
        If UserList(tIndex).flags.Privilegios > PlayerType.User Then
            UserList(userindex).flags.Ban = 1
            Call CloseSocket(userindex)
            Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(userindex).name & " banned by the server por bannear un Administrador." & FONTTYPE_FIGHT)
        End If
        
        Call LogGM(UserList(userindex).name, "BAN a " & UserList(tIndex).name, False)
        
        'ponemos el flag de ban a 1
        Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")
        'ponemos la pena
        tInt = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
        Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", tInt + 1)
        Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(userindex).name) & ": BAN POR " & LCase$(name) & " " & Date & " " & Time)
        
        Call CloseSocket(tIndex)
    End If

    Exit Sub
End If

If UCase$(Left$(rData, 7)) = "/UNBAN " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 7)
    
    rData = Replace(rData, "\", "")
    rData = Replace(rData, "/", "")
    
    If Not FileExist(CharPath & rData & ".chr", vbNormal) Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Charfile inexistente (no use +)" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call UnBan(rData)
    
    'penas
    i = val(GetVar(CharPath & rData & ".chr", "PENAS", "Cant"))
    Call WriteVar(CharPath & rData & ".chr", "PENAS", "Cant", i + 1)
    Call WriteVar(CharPath & rData & ".chr", "PENAS", "P" & i + 1, LCase$(UserList(userindex).name) & ": UNBAN. " & Date & " " & Time)
    
    Call LogGM(UserList(userindex).name, "/UNBAN a " & rData, False)
    Call SendData(SendTarget.ToIndex, userindex, 0, "||" & rData & " unbanned." & FONTTYPE_INFO)
    

    Exit Sub
End If


'SEGUIR
If UCase$(rData) = "/SEGUIR" Then
    If UserList(userindex).flags.TargetNPC > 0 Then
        Call DoFollow(UserList(userindex).flags.TargetNPC, UserList(userindex).name)
    End If
    Exit Sub
End If

If UCase$(Left$(rData, 5)) = "/PUN " Then
rData = Right(rData, Len(rData) - 5)
tIndex = NameIndex(rData)

If tIndex <= 0 Then
Call SendData(SendTarget.ToIndex, userindex, 0, "||El Usuario no esta online." & FONTTYPE_INFO)
Exit Sub
End If

Call SendData(SendTarget.ToIndex, tIndex, 0, "||Has recibido 5 puntos." & FONTTYPE_INFO)
UserList(tIndex).Stats.PuntosTorneo = UserList(tIndex).Stats.PuntosTorneo + 5
Exit Sub
End If

If UCase$(Left$(rData, 5)) = "/TOR " Then
    rData = Right$(rData, Len(rData) - 5)
   
    Dim Data As String
   
    If Hay_Torneo = False Then
         Hay_Torneo = True
         Torneo_Nivel_Minimo = val(ReadField(1, rData, 32))
         Torneo_Nivel_Maximo = val(ReadField(2, rData, 32))
         Torneo_Cantidad = val(ReadField(3, rData, 32))
         Torneo_Clases_Validas2(1) = val(ReadField(4, rData, 32))
         Torneo_Clases_Validas2(2) = val(ReadField(5, rData, 32))
         Torneo_Clases_Validas2(3) = val(ReadField(6, rData, 32))
         Torneo_Clases_Validas2(4) = val(ReadField(7, rData, 32))
         Torneo_Clases_Validas2(5) = val(ReadField(8, rData, 32))
         Torneo_Clases_Validas2(6) = val(ReadField(9, rData, 32))
         Torneo_Clases_Validas2(7) = val(ReadField(10, rData, 32))
         Torneo_Clases_Validas2(8) = val(ReadField(11, rData, 32))
         Torneo_SumAuto = val(ReadField(12, rData, 32))
         Torneo_Map = val(ReadField(13, rData, 32))
         Torneo_X = val(ReadField(14, rData, 32))
         Torneo_Y = val(ReadField(15, rData, 32))
         Torneo_Alineacion_Validas2(1) = val(ReadField(16, rData, 32))
         Torneo_Alineacion_Validas2(1) = val(ReadField(17, rData, 32))
         Torneo_Alineacion_Validas2(1) = val(ReadField(18, rData, 32))
         Torneo_Alineacion_Validas2(1) = val(ReadField(19, rData, 32))
         
         Call SendData(SendTarget.toall, 0, 0, "||Organizado por : " & UserList(userindex).name & FONTTYPE_CELESTE_NEGRITA)
         Call SendData(SendTarget.toall, 0, 0, "||Nivel máximo: " & Torneo_Nivel_Maximo & FONTTYPE_CELESTE_NEGRITA)
         Call SendData(SendTarget.toall, 0, 0, "||Nivel minimo: " & Torneo_Nivel_Minimo & FONTTYPE_CELESTE_NEGRITA)
         Call SendData(SendTarget.toall, 0, 0, "||Límite de usuarios: " & Torneo_Cantidad & FONTTYPE_CELESTE_NEGRITA)
         
         For i = 1 To 8
             If Torneo_Clases_Validas2(i) = 1 Then
                 Data = Data & Torneo_Clases_Validas(i) & ","
             End If
         Next
         Data = Left$(Data, Len(Data) - 1) & "."
         
         Call SendData(SendTarget.toall, 0, 0, "||Clases válidas: " & Data & FONTTYPE_CELESTE_NEGRITA)
         Data = ""
         
         For i = 1 To 4
             If Torneo_Alineacion_Validas2(i) = 1 Then
                 Data = Data & Torneo_Alineacion_Validas(i) & ","
             End If
         Next
         Data = Left$(Data, Len(Data) - 1) & "."
         
         Call SendData(SendTarget.toall, 0, 0, "||Alineaciónes válidas: " & Data & FONTTYPE_CELESTE_NEGRITA)
         Call SendData(SendTarget.toall, 0, 0, "||Escriba en consola '/TORNEO' para participar." & FONTTYPE_CELESTE_NEGRITA)
     Else
         Call SendData(SendTarget.ToIndex, userindex, 0, "||Ya hay un torneo." & FONTTYPE_INFO)
    End If
    Exit Sub
End If
 
If UCase$(rData) = "/CERRARTORNEO" Then
    If Hay_Torneo = True Then
        Call SendData(SendTarget.toall, 0, 0, "||Inscripciones Cerradas." & FONTTYPE_CELESTE_NEGRITA)
        Hay_Torneo = False
        Torneox.Reset
    End If
    Exit Sub
End If

'Summon
If UCase$(Left$(rData, 5)) = "/SUM " Then
    rData = Right$(rData, Len(rData) - 5)
    
    tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||El jugador no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call SendData(SendTarget.ToIndex, tIndex, 0, "||" & UserList(userindex).name & " há sido trasportado." & FONTTYPE_INFO)
    Call WarpUserChar(tIndex, UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y + 1, True)
    
    Call LogGM(UserList(userindex).name, "/SUM " & UserList(tIndex).name & " Map:" & UserList(userindex).pos.Map & " X:" & UserList(userindex).pos.x & " Y:" & UserList(userindex).pos.Y, False)
    Exit Sub
End If

'Crear criatura
If UCase$(Left$(rData, 3)) = "/CC" Then
   Call EnviarSpawnList(userindex)
   Exit Sub
End If

'Spawn!!!!!
If UCase$(Left$(rData, 3)) = "SPA" Then
    rData = Right$(rData, Len(rData) - 3)
    
    If val(rData) > 0 And val(rData) < UBound(SpawnList) + 1 Then _
          Call SpawnNpc(SpawnList(val(rData)).NpcIndex, UserList(userindex).pos, True, False)
          
          Call LogGM(UserList(userindex).name, "Sumoneo " & SpawnList(val(rData)).NpcName, False)
          
    Exit Sub
End If

'Resetea el inventario
If UCase$(rData) = "/RESETINV" Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 9)
    If UserList(userindex).flags.TargetNPC = 0 Then Exit Sub
    Call ResetNpcInv(UserList(userindex).flags.TargetNPC)
    Call LogGM(UserList(userindex).name, "/RESETINV " & Npclist(UserList(userindex).flags.TargetNPC).name, False)
    Exit Sub
End If

'/Clean
If UCase$(rData) = "/LIMPIAR" Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call LimpiarMundo
    Exit Sub
End If

'Mensaje del servidor
If UCase$(Left$(rData, 6)) = "/RMSG " Then
    rData = Right$(rData, Len(rData) - 6)
    Call LogGM(UserList(userindex).name, "Mensaje Broadcast:" & rData, False)
    If rData <> "" Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(userindex).name & "> " & rData & FONTTYPE_GM & ENDC)
    End If
    Exit Sub
End If

'Mensaje del servidor
If UCase$(Left$(rData, 6)) = "/SVSG " Then
    rData = Right$(rData, Len(rData) - 6)
    Call LogGM(UserList(userindex).name, "Mensaje Broadcast:" & rData, False)
    If rData <> "" Then
Call SendData(SendTarget.toall, 0, 0, "||" & "Servidor> " & rData & FONTTYPE_SERVER & ENDC)
    End If
    Exit Sub
End If

'Ip del nick
If UCase$(Left$(rData, 9)) = "/NICK2IP " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 9)
    tIndex = NameIndex(UCase$(rData))
    Call LogGM(UserList(userindex).name, "NICK2IP Solicito la IP de " & rData, UserList(userindex).flags.Privilegios = PlayerType.Consejero)
    If tIndex > 0 Then
        If (UserList(userindex).flags.Privilegios > PlayerType.User And UserList(tIndex).flags.Privilegios = PlayerType.User) Or (UserList(userindex).flags.Privilegios >= PlayerType.Dios) Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "||El ip de " & rData & " es " & UserList(tIndex).ip & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.ToIndex, userindex, 0, "||No tienes los privilegios necesarios" & FONTTYPE_INFO)
        End If
    Else
       Call SendData(SendTarget.ToIndex, userindex, 0, "||No hay ningun personaje con ese nick" & FONTTYPE_INFO)
    End If
    Exit Sub
End If
 
'Ip del nick
If UCase$(Left$(rData, 9)) = "/IP2NICK " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 9)

    If InStr(rData, ".") < 1 Then
        tInt = NameIndex(rData)
        If tInt < 1 Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Pj Offline" & FONTTYPE_INFO)
            Exit Sub
        End If
        rData = UserList(tInt).ip
    End If
    tStr = vbNullString
    Call LogGM(UserList(userindex).name, "IP2NICK Solicito los Nicks de IP " & rData, UserList(userindex).flags.Privilegios = PlayerType.Consejero)
    For LoopC = 1 To LastUser
        If UserList(LoopC).ip = rData And UserList(LoopC).name <> "" And UserList(LoopC).flags.UserLogged Then
            If (UserList(userindex).flags.Privilegios > PlayerType.User And UserList(LoopC).flags.Privilegios = PlayerType.User) Or (UserList(userindex).flags.Privilegios >= PlayerType.Dios) Then
                tStr = tStr & UserList(LoopC).name & ", "
            End If
        End If
    Next LoopC
    
    Call SendData(SendTarget.ToIndex, userindex, 0, "||Los personajes con ip " & rData & " son: " & tStr & FONTTYPE_INFO)
    Exit Sub
End If


If UCase$(Left$(rData, 8)) = "/ONCLAN " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 8)
    tInt = GuildIndex(rData)
    
    If tInt > 0 Then
        tStr = modGuilds.m_ListaDeMiembrosOnline(userindex, tInt)
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Clan " & UCase(rData) & ": " & tStr & FONTTYPE_GUILDMSG)
    End If
End If


'Crear Teleport
If UCase(Left(rData, 4)) = "/CT " Then
    If Not UserList(userindex).flags.EsRolesMaster And UserList(userindex).flags.Privilegios < PlayerType.Dios Then Exit Sub
    '/ct mapa_dest x_dest y_dest
    rData = Right(rData, Len(rData) - 4)
    Call LogGM(UserList(userindex).name, "/CT: " & rData, False)
    mapa = ReadField(1, rData, 32)
    x = ReadField(2, rData, 32)
    Y = ReadField(3, rData, 32)
    
    If MapaValido(mapa) = False Or InMapBounds(mapa, x, Y) = False Then
        Exit Sub
    End If
    If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y - 1).OBJInfo.ObjIndex > 0 Then
        Exit Sub
    End If
    If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y - 1).TileExit.Map > 0 Then
        Exit Sub
    End If
    
    If MapData(mapa, x, Y).OBJInfo.ObjIndex > 0 Then
        Call SendData(SendTarget.ToIndex, userindex, mapa, "||Hay un objeto en el piso en ese lugar" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Dim ET As Obj
    ET.Amount = 1
    ET.ObjIndex = 378
    
    Call MakeObj(SendTarget.ToMap, 0, UserList(userindex).pos.Map, ET, UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y - 1)
    
    ET.Amount = 1
    ET.ObjIndex = 651
    
    Call MakeObj(SendTarget.ToMap, 0, mapa, ET, mapa, x, Y)
    
    MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y - 1).TileExit.Map = mapa
    MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y - 1).TileExit.x = x
    MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y - 1).TileExit.Y = Y
    
    Exit Sub
End If

'Destruir Teleport
'toma el ultimo click
If UCase(Left(rData, 3)) = "/DT" Then
    '/dt
    If Not UserList(userindex).flags.EsRolesMaster And UserList(userindex).flags.Privilegios < PlayerType.Dios Then Exit Sub
    Call LogGM(UserList(userindex).name, "/DT", False)
    
    mapa = UserList(userindex).flags.TargetMap
    x = UserList(userindex).flags.TargetX
    Y = UserList(userindex).flags.TargetY
    
    If ObjData(MapData(mapa, x, Y).OBJInfo.ObjIndex).OBJType = eOBJType.otTeleport And _
        MapData(mapa, x, Y).TileExit.Map > 0 Then
        Call EraseObj(SendTarget.ToMap, 0, mapa, MapData(mapa, x, Y).OBJInfo.Amount, mapa, x, Y)
        Call EraseObj(SendTarget.ToMap, 0, MapData(mapa, x, Y).TileExit.Map, 1, MapData(mapa, x, Y).TileExit.Map, MapData(mapa, x, Y).TileExit.x, MapData(mapa, x, Y).TileExit.Y)
        MapData(mapa, x, Y).TileExit.Map = 0
        MapData(mapa, x, Y).TileExit.x = 0
        MapData(mapa, x, Y).TileExit.Y = 0
    End If
    
    Exit Sub
End If

If UCase$(Left$(rData, 9)) = "/SETDESC " Then
    If Not UserList(userindex).flags.EsRolesMaster And UserList(userindex).flags.Privilegios < PlayerType.Dios Then Exit Sub
    rData = Right$(rData, Len(rData) - 9)
    DummyInt = UserList(userindex).flags.TargetUser
    If DummyInt > 0 Then
        UserList(DummyInt).DescRM = rData
    Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Haz click sobre un personaje antes!" & FONTTYPE_INFO)
    End If
    Exit Sub
    
End If




Select Case UCase$(Left$(rData, 13))
    Case "/FORCEMIDIMAP"
        If Len(rData) > 13 Then
            rData = Right$(rData, Len(rData) - 14)
        Else
            Call SendData(SendTarget.ToIndex, userindex, 0, "||El formato correcto de este comando es /FORCEMIDMAP MIDI MAPA, siendo el MAPA opcional" & FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Solo dioses, admins y RMS
        If UserList(userindex).flags.Privilegios < PlayerType.Dios And Not UserList(userindex).flags.EsRolesMaster Then Exit Sub
        
        'Obtenemos el número de midi
        Arg1 = ReadField(1, rData, vbKeySpace)
        ' y el de mapa
        Arg2 = ReadField(2, rData, vbKeySpace)
        
        'Si el mapa no fue enviado tomo el actual
        If IsNumeric(Arg2) Then
            tInt = CInt(Arg2)
        Else
            tInt = UserList(userindex).pos.Map
        End If
        
        If IsNumeric(Arg1) Then
            If Arg1 = "0" Then
                'Ponemos el default del mapa
                Call SendData(SendTarget.ToMap, 0, tInt, "TM" & CStr(MapInfo(UserList(userindex).pos.Map).Music))
            Else
                'Ponemos el pedido por el GM
                Call SendData(SendTarget.ToMap, 0, tInt, "TM" & Arg1)
            End If
        Else
            Call SendData(SendTarget.ToIndex, userindex, 0, "||El formato correcto de este comando es /FORCEMIDMAP MIDI MAPA, siendo el MAPA opcional" & FONTTYPE_INFO)
        End If
        Exit Sub
    
    Case "/FORCEWAVMAP "
        rData = Right$(rData, Len(rData) - 13)
        'Solo dioses, admins y RMS
        If UserList(userindex).flags.Privilegios < PlayerType.Dios And Not UserList(userindex).flags.EsRolesMaster Then Exit Sub
        
        'Obtenemos el número de wav
        Arg1 = ReadField(1, rData, vbKeySpace)
        ' el de mapa
        Arg2 = ReadField(2, rData, vbKeySpace)
        ' el de X
        Arg3 = ReadField(3, rData, vbKeySpace)
        ' y el de Y (las coords X-Y sólo tendrán sentido al implementarse el panning en la 11.6)
        Arg4 = ReadField(4, rData, vbKeySpace)
        
        'Si el mapa no fue enviado tomo el actual
        If IsNumeric(Arg2) And IsNumeric(Arg3) And IsNumeric(Arg4) Then
            tInt = CInt(Arg2)
        Else
            tInt = UserList(userindex).pos.Map
            Arg3 = CStr(UserList(userindex).pos.x)
            Arg4 = CStr(UserList(userindex).pos.Y)
        End If
        
        If IsNumeric(Arg1) Then
            'Ponemos el pedido por el GM
            Call SendData(SendTarget.ToMap, 0, tInt, "TW" & Arg1)
        Else
            Call SendData(SendTarget.ToIndex, userindex, 0, "||El formato correcto de este comando es /FORCEWAVMAP WAV MAPA X Y, siendo la posición opcional" & FONTTYPE_INFO)
        End If
        Exit Sub
End Select

Select Case UCase$(Left$(rData, 8))
    Case "/REALMSG"
        'Solo dioses, admins y RMS
        If UserList(userindex).flags.Privilegios > PlayerType.SemiDios Or UserList(userindex).flags.EsRolesMaster Then
            tStr = Right$(rData, Len(rData) - 9)
            
            If InStr(1, tStr, "~") = 0 Then
                Call SendData(SendTarget.ToRealYRMs, 0, 0, "||ARMADA REAL> " & tStr & FONTTYPE_TALK)
            Else
                Call SendData(SendTarget.ToRealYRMs, 0, 0, "||ARMADA REAL> " & tStr)
            End If
        End If
        Exit Sub
    
    Case "/CAOSMSG"
        'Solo dioses, admins y RMS
        If UserList(userindex).flags.Privilegios > PlayerType.SemiDios Or UserList(userindex).flags.EsRolesMaster Then
            tStr = Right$(rData, Len(rData) - 9)
            
            If InStr(1, tStr, "~") = 0 Then
                Call SendData(SendTarget.ToCaosYRMs, 0, 0, "||FUERZAS DEL CAOS> " & tStr & FONTTYPE_TALK)
            Else
                Call SendData(SendTarget.ToCaosYRMs, 0, 0, "||FUERZAS DEL CAOS> " & tStr)
            End If
        End If
        Exit Sub
    
    Case "/CIUMSG "
        'Solo dioses, admins y RMS
        If UserList(userindex).flags.Privilegios > PlayerType.SemiDios Or UserList(userindex).flags.EsRolesMaster Then
            tStr = Right$(rData, Len(rData) - 8)
            
            If InStr(1, tStr, "~") = 0 Then
                Call SendData(SendTarget.ToCiudadanosYRMs, 0, 0, "||CIUDADANOS> " & tStr & FONTTYPE_TALK)
            Else
                Call SendData(SendTarget.ToCiudadanosYRMs, 0, 0, "||CIUDADANOS> " & tStr)
            End If
        End If
        Exit Sub
    
    Case "/CRIMSG "
        'Solo dioses, admins y RMS
        If UserList(userindex).flags.Privilegios > PlayerType.SemiDios Or UserList(userindex).flags.EsRolesMaster Then
            tStr = Right$(rData, Len(rData) - 8)
            
            If InStr(1, tStr, "~") = 0 Then
                Call SendData(SendTarget.ToCriminalesYRMs, 0, 0, "||CRIMINALES> " & tStr & FONTTYPE_TALK)
            Else
                Call SendData(SendTarget.ToCriminalesYRMs, 0, 0, "||CRIMINALES> " & tStr)
            End If
        End If
        Exit Sub
    
    Case "/TALKAS "
        'Solo dioses, admins y RMS
        If UserList(userindex).flags.Privilegios > PlayerType.SemiDios Or UserList(userindex).flags.EsRolesMaster Then
            'Asegurarse haya un NPC seleccionado
            If UserList(userindex).flags.TargetNPC > 0 Then
                tStr = Right$(rData, Len(rData) - 8)
                
                Call SendData(SendTarget.ToNPCArea, UserList(userindex).flags.TargetNPC, Npclist(UserList(userindex).flags.TargetNPC).pos.Map, "||" & vbWhite & "°" & tStr & "°" & CStr(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
            Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Debes seleccionar el NPC por el que quieres hablar antes de usar este comando" & FONTTYPE_INFO)
            End If
        End If
        Exit Sub
End Select




'<<<<<<<<<<<<<<<<<<<<< Dioses >>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<< Dioses >>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<< Dioses >>>>>>>>>>>>>>>>>>>>>>>>
If UserList(userindex).flags.Privilegios < PlayerType.Dios Then
    Exit Sub
End If


'[Barrin 30-11-03]
'Quita todos los objetos del area
If UCase$(rData) = "/MASSDEST" Then
    For Y = UserList(userindex).pos.Y - MinYBorder + 1 To UserList(userindex).pos.Y + MinYBorder - 1
            For x = UserList(userindex).pos.x - MinXBorder + 1 To UserList(userindex).pos.x + MinXBorder - 1
                If x > 0 And Y > 0 And x < 101 And Y < 101 Then _
                    If MapData(UserList(userindex).pos.Map, x, Y).OBJInfo.ObjIndex > 0 Then _
                    If ItemNoEsDeMapa(MapData(UserList(userindex).pos.Map, x, Y).OBJInfo.ObjIndex) Then Call EraseObj(SendTarget.ToMap, userindex, UserList(userindex).pos.Map, 10000, UserList(userindex).pos.Map, x, Y)
            Next x
    Next Y
    Call LogGM(UserList(userindex).name, "/MASSDEST", (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
    Exit Sub
End If
'[/Barrin 30-11-03]


'[yb]
If UCase$(Left$(rData, 12)) = "/ACEPTCONSE " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 12)
    tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuario offline" & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.toall, 0, 0, "||" & rData & " fue aceptado en el honorable Consejo Real de Banderbill." & FONTTYPE_CONSEJO)
        UserList(tIndex).flags.PertAlCons = 1
        Call WarpUserChar(tIndex, UserList(tIndex).pos.Map, UserList(tIndex).pos.x, UserList(tIndex).pos.Y, False)
    End If
    Exit Sub
End If

If UCase$(Left$(rData, 16)) = "/ACEPTCONSECAOS " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 16)
    tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuario offline" & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.toall, 0, 0, "||" & rData & " fue aceptado en el Consejo de la Legión Oscura." & FONTTYPE_CONSEJOCAOS)
        UserList(tIndex).flags.PertAlConsCaos = 1
        Call WarpUserChar(tIndex, UserList(tIndex).pos.Map, UserList(tIndex).pos.x, UserList(tIndex).pos.Y, False)
    End If
    Exit Sub
End If

If Left$(UCase$(rData), 5) = "/PISO" Then
    For x = 5 To 95
        For Y = 5 To 95
            tIndex = MapData(UserList(userindex).pos.Map, x, Y).OBJInfo.ObjIndex
            If tIndex > 0 Then
                If ObjData(tIndex).OBJType <> 4 Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||(" & x & "," & Y & ") " & ObjData(tIndex).name & FONTTYPE_INFO)
                End If
            End If
        Next Y
    Next x
    Exit Sub
End If

If UCase$(Left$(rData, 10)) = "/ESTUPIDO " Then
    If UserList(userindex).flags.EsRolesMaster = 1 Then Exit Sub
    'para deteccion de aoice
    rData = UCase$(Right$(rData, Len(rData) - 10))
    i = NameIndex(rData)
    If i <= 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Offline" & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, i, 0, "DUMB")
    End If
    Exit Sub
End If
If UCase$(Left$(rData, 12)) = "/GUERRAAUTO " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    rData = UCase$(Right$(rData, Len(rData) - 12))
        If UCase$(rData) = "ON" Then
            GuerrasAuto userindex, 1
        ElseIf UCase$(rData) = "OFF" Then
            GuerrasAuto userindex, 0
        End If
    Exit Sub
End If
If UCase$(Left$(rData, 12)) = "/NOESTUPIDO " Then
    If UserList(userindex).flags.EsRolesMaster = 1 Then Exit Sub
    'para deteccion de aoice
    rData = UCase$(Right$(rData, Len(rData) - 12))
    i = NameIndex(rData)
    If i <= 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Offline" & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, i, 0, "NESTUP")
    End If
    Exit Sub
End If

If Left$(UCase$(rData), 13) = "/DUMPSECURITY" Then
    Call SecurityIp.DumpTables
    Exit Sub
End If

If UCase$(Left$(rData, 11)) = "/KICKCONSE " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 11)
    tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        If FileExist(CharPath & rData & ".chr") Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuario offline, Echando de los consejos" & FONTTYPE_INFO)
            Call WriteVar(CharPath & UCase(rData) & ".chr", "CONSEJO", "PERTENECE", 0)
            Call WriteVar(CharPath & UCase(rData) & ".chr", "CONSEJO", "PERTENECECAOS", 0)
        Else
            Call SendData(SendTarget.ToIndex, userindex, 0, "||No se encuentra el charfile " & CharPath & rData & ".chr" & FONTTYPE_INFO)
            Exit Sub
        End If
    Else
        If UserList(tIndex).flags.PertAlCons > 0 Then
            Call SendData(SendTarget.ToIndex, tIndex, 0, "||Has sido echado en el consejo de banderbill" & FONTTYPE_TALK & ENDC)
            UserList(tIndex).flags.PertAlCons = 0
            Call WarpUserChar(tIndex, UserList(tIndex).pos.Map, UserList(tIndex).pos.x, UserList(tIndex).pos.Y)
            Call SendData(SendTarget.toall, 0, 0, "||" & rData & " fue expulsado del consejo de Banderbill" & FONTTYPE_CONSEJO)
        End If
        If UserList(tIndex).flags.PertAlConsCaos > 0 Then
            Call SendData(SendTarget.ToIndex, tIndex, 0, "||Has sido echado en el consejo de la legión oscura" & FONTTYPE_TALK & ENDC)
            UserList(tIndex).flags.PertAlConsCaos = 0
            Call WarpUserChar(tIndex, UserList(tIndex).pos.Map, UserList(tIndex).pos.x, UserList(tIndex).pos.Y)
            Call SendData(SendTarget.toall, 0, 0, "||" & rData & " fue expulsado del consejo de la Legión Oscura" & FONTTYPE_CONSEJOCAOS)
        End If
    End If
    Exit Sub
End If
'[/yb]



If UCase$(Left$(rData, 8)) = "/TRIGGER" Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)
    
    rData = Trim(Right(rData, Len(rData) - 8))
    mapa = UserList(userindex).pos.Map
    x = UserList(userindex).pos.x
    Y = UserList(userindex).pos.Y
    If rData <> "" Then
        tInt = MapData(mapa, x, Y).trigger
        MapData(mapa, x, Y).trigger = val(rData)
    End If
    Call SendData(SendTarget.ToIndex, userindex, 0, "||Trigger " & MapData(mapa, x, Y).trigger & " en mapa " & mapa & " " & x & ", " & Y & FONTTYPE_INFO)
    Exit Sub
End If



If UCase(rData) = "/BANIPLIST" Then
   
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)
    tStr = "||"
    For LoopC = 1 To BanIps.Count
        tStr = tStr & BanIps.Item(LoopC) & ", "
    Next LoopC
    tStr = tStr & FONTTYPE_INFO
    Call SendData(SendTarget.ToIndex, userindex, 0, tStr)
    Exit Sub
End If

If UCase(rData) = "/BANIPRELOAD" Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call BanIpGuardar
    Call BanIpCargar
    Exit Sub
End If

If UCase(Left(rData, 14)) = "/MIEMBROSCLAN " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    rData = Trim(Right(rData, Len(rData) - 9))
    If Not FileExist(App.Path & "\guilds\" & rData & "-members.mem") Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "|| No existe el clan: " & rData & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call LogGM(UserList(userindex).name, "MIEMBROSCLAN a " & rData, False)

    tInt = val(GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "INIT", "NroMembers"))
    
    For i = 1 To tInt
        tStr = GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "Members", "Member" & i)
        'tstr es la victima
        Call SendData(SendTarget.ToIndex, userindex, 0, "||" & tStr & "<" & rData & ">." & FONTTYPE_INFO)
    Next i

    Exit Sub
End If



If UCase(Left(rData, 9)) = "/BANCLAN " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    rData = Trim(Right(rData, Len(rData) - 9))
    If Not FileExist(App.Path & "\guilds\" & rData & "-members.mem") Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "|| No existe el clan: " & rData & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call SendData(SendTarget.toall, 0, 0, "|| " & UserList(userindex).name & " banned al clan " & UCase$(rData) & FONTTYPE_FIGHT)
    
    'baneamos a los miembros
    Call LogGM(UserList(userindex).name, "BANCLAN a " & rData, False)

    tInt = val(GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "INIT", "NroMembers"))
    
    For i = 1 To tInt
        tStr = GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "Members", "Member" & i)
        'tstr es la victima
        Call Ban(tStr, "Administracion del servidor", "Clan Banned")
        tIndex = NameIndex(tStr)
        If tIndex > 0 Then
            'esta online
            UserList(tIndex).flags.Ban = 1
            Call CloseSocket(tIndex)
        End If
        
        Call SendData(SendTarget.toall, 0, 0, "||   " & tStr & "<" & rData & "> ha sido expulsado del servidor." & FONTTYPE_FIGHT)

        'ponemos el flag de ban a 1
        Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")

        'ponemos la pena
        N = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
        Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", N + 1)
        Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & N + 1, LCase$(UserList(userindex).name) & ": BAN AL CLAN: " & rData & " " & Date & " " & Time)

    Next i

    Exit Sub
End If


'Ban x IP
If UCase(Left(rData, 7)) = "/BANIP " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Dim BanIP As String, XNick As Boolean
    
    rData = Right$(rData, Len(rData) - 7)
    tStr = Replace(ReadField(1, rData, Asc(" ")), "+", " ")
    'busca primero la ip del nick
    tIndex = NameIndex(tStr)
    If tIndex <= 0 Then
        XNick = False
        Call LogGM(UserList(userindex).name, "/BanIP " & rData, False)
        BanIP = tStr
    Else
        XNick = True
        Call LogGM(UserList(userindex).name, "/BanIP " & UserList(tIndex).name & " - " & UserList(tIndex).ip, False)
        BanIP = UserList(tIndex).ip
    End If
    
    rData = Right$(rData, Len(rData) - Len(tStr))
    
    If BanIpBuscar(BanIP) > 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||La IP " & BanIP & " ya se encuentra en la lista de bans." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call BanIpAgrega(BanIP)
    Call SendData(SendTarget.ToAdmins, userindex, 0, "||" & UserList(userindex).name & " Baneo la IP " & BanIP & FONTTYPE_FIGHT)
    
    If XNick = True Then
        Call LogBan(tIndex, userindex, "Ban por IP desde Nick por " & rData)
        
        Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(userindex).name & " echo a " & UserList(tIndex).name & "." & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(userindex).name & " Banned a " & UserList(tIndex).name & "." & FONTTYPE_FIGHT)
        
        'Ponemos el flag de ban a 1
        UserList(tIndex).flags.Ban = 1
        
        Call LogGM(UserList(userindex).name, "Echo a " & UserList(tIndex).name, False)
        Call LogGM(UserList(userindex).name, "BAN a " & UserList(tIndex).name, False)
        Call CloseSocket(tIndex)
    End If
    
    Exit Sub
End If

'Desbanea una IP
If UCase(Left(rData, 9)) = "/UNBANIP " Then
    
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    
    rData = Right(rData, Len(rData) - 9)
    Call LogGM(UserList(userindex).name, "/UNBANIP " & rData, False)
    
'    For LoopC = 1 To BanIps.Count
'        If BanIps.Item(LoopC) = rdata Then
'            BanIps.Remove LoopC
'            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La IP " & BanIP & " se ha quitado de la lista de bans." & FONTTYPE_INFO)
'            Exit Sub
'        End If
'    Next LoopC
'
'    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La IP " & rdata & " NO se encuentra en la lista de bans." & FONTTYPE_INFO)
    
    If BanIpQuita(rData) Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||La IP """ & rData & """ se ha quitado de la lista de bans." & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "||La IP """ & rData & """ NO se encuentra en la lista de bans." & FONTTYPE_INFO)
    End If
    
    Exit Sub
End If



'Crear Item
If UCase(Left(rData, 4)) = "/CI " Then
    rData = Right$(rData, Len(rData) - 4)
    Call LogGM(UserList(userindex).name, "/CI: " & rData, False)
    
    If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y - 1).OBJInfo.ObjIndex > 0 Then
        Exit Sub
    End If
    If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y - 1).TileExit.Map > 0 Then
        Exit Sub
    End If
    If val(rData) < 1 Or val(rData) > NumObjDatas Then
        Exit Sub
    End If
    
    'Is the object not null?
    If ObjData(val(rData)).name = "" Then Exit Sub
    
    Dim Objeto As Obj
    
    Call SendData(SendTarget.ToIndex, userindex, 0, "||ATENCION: FUERON CREADOS ***100*** ITEMS!, TIRE Y /DEST LOS QUE NO NECESITE!!" & FONTTYPE_GUILD)
    
    Objeto.Amount = 100
    Objeto.ObjIndex = val(rData)
    
    Call MakeObj(SendTarget.ToMap, 0, UserList(userindex).pos.Map, Objeto, UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y - 1)
    
    Exit Sub
End If
If UCase$(Left$(rData, 13)) = "/VERPROCESOS " Then
    rData = Right$(rData, Len(rData) - 13)
    tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, tIndex, 0, "PRCS" & userindex)
    End If
    Exit Sub
End If

If UCase$(Left$(rData, 5)) = "/DEST" Then
    Call LogGM(UserList(userindex).name, "/DEST", False)
    rData = Right$(rData, Len(rData) - 5)
    Call EraseObj(SendTarget.ToMap, userindex, UserList(userindex).pos.Map, 10000, UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y)
    Exit Sub
End If

If UCase$(Left$(rData, 8)) = "/NOCAOS " Then
    rData = Right$(rData, Len(rData) - 8)
    Call LogGM(UserList(userindex).name, "ECHO DEL CAOS A: " & rData, False)

    tIndex = NameIndex(rData)
    
    If tIndex > 0 Then
        UserList(tIndex).Faccion.FuerzasCaos = 0
        UserList(tIndex).Faccion.Reenlistadas = 200
        Call SendData(SendTarget.ToIndex, userindex, 0, "|| " & rData & " expulsado de las fuerzas del caos y prohibida la reenlistada" & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, tIndex, 0, "|| " & UserList(userindex).name & " te ha expulsado en forma definitiva de las fuerzas del caos." & FONTTYPE_FIGHT)
    Else
        If FileExist(CharPath & rData & ".chr") Then
            Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "EjercitoCaos", 0)
            Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Reenlistadas", 200)
            Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Extra", "Expulsado por " & UserList(userindex).name)
            Call SendData(SendTarget.ToIndex, userindex, 0, "|| " & rData & " expulsado de las fuerzas del caos y prohibida la reenlistada" & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.ToIndex, userindex, 0, "|| " & rData & ".chr inexistente." & FONTTYPE_INFO)
        End If
    End If
    Exit Sub
End If

If UCase$(Left$(rData, 8)) = "/NOREAL " Then
    rData = Right$(rData, Len(rData) - 8)
    Call LogGM(UserList(userindex).name, "ECHO DE LA REAL A: " & rData, False)

    rData = Replace(rData, "\", "")
    rData = Replace(rData, "/", "")

    tIndex = NameIndex(rData)

    If tIndex > 0 Then
        UserList(tIndex).Faccion.ArmadaReal = 0
        UserList(tIndex).Faccion.Reenlistadas = 200
        Call SendData(SendTarget.ToIndex, userindex, 0, "|| " & rData & " expulsado de las fuerzas reales y prohibida la reenlistada" & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, tIndex, 0, "|| " & UserList(userindex).name & " te ha expulsado en forma definitiva de las fuerzas reales." & FONTTYPE_FIGHT)
    Else
        If FileExist(CharPath & rData & ".chr") Then
            Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "EjercitoReal", 0)
            Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Reenlistadas", 200)
            Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Extra", "Expulsado por " & UserList(userindex).name)
            Call SendData(SendTarget.ToIndex, userindex, 0, "|| " & rData & " expulsado de las fuerzas reales y prohibida la reenlistada" & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.ToIndex, userindex, 0, "|| " & rData & ".chr inexistente." & FONTTYPE_INFO)
        End If
    End If
    Exit Sub
End If

If UCase$(Left$(rData, 11)) = "/FORCEMIDI " Then
    rData = Right$(rData, Len(rData) - 11)
    If Not IsNumeric(rData) Then
        Exit Sub
    Else
        Call SendData(SendTarget.toall, 0, 0, "|| " & UserList(userindex).name & " broadcast musica: " & rData & FONTTYPE_SERVER)
        Call SendData(SendTarget.toall, 0, 0, "TM" & rData)
    End If
End If

If UCase$(Left$(rData, 10)) = "/FORCEWAV " Then
    rData = Right$(rData, Len(rData) - 10)
    If Not IsNumeric(rData) Then
        Exit Sub
    Else
        Call SendData(SendTarget.toall, 0, 0, "TW" & rData)
    End If
End If


If UCase$(Left$(rData, 12)) = "/BORRARPENA " Then
    '/borrarpena pj pena
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    
    rData = Right$(rData, Len(rData) - 12)
    
    name = ReadField(1, rData, Asc("@"))
    tStr = ReadField(2, rData, Asc("@"))
    
    If name = "" Or tStr = "" Or Not IsNumeric(tStr) Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Utilice /borrarpj Nick@NumeroDePena" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    name = Replace(name, "\", "")
    name = Replace(name, "/", "")
    
    If FileExist(CharPath & name & ".chr", vbNormal) Then
        rData = GetVar(CharPath & name & ".chr", "PENAS", "P" & val(tStr))
        Call WriteVar(CharPath & name & ".chr", "PENAS", "P" & val(tStr), LCase$(UserList(userindex).name) & ": <Pena borrada> " & Date & " " & Time)
    End If
    
    Call LogGM(UserList(userindex).name, " borro la pena: " & tStr & "-" & rData & " de " & name, UserList(userindex).flags.Privilegios = PlayerType.Consejero)
    Exit Sub
End If


' lo limite a torneode 5 rondas de 32 participantes, pero si quierren de mas participantes, cambien el < 6 por un numero mayor.
If UCase$(Left$(rData, 9)) = "/SATUROS " Then
rData = Right$(rData, Len(rData) - 9)
Dim torneos As Integer
torneos = CInt(rData)
If (torneos > 0 And torneos < 6) Then Call Torneos_Inicia(userindex, torneos)
End If

' con esto cancelamos el torneo
If UCase(rData) = "/CANCELAR" Then
Call Rondas_Cancela
Exit Sub
End If


'Bloquear
If UCase$(Left$(rData, 5)) = "/BLOQ" Then
    Call LogGM(UserList(userindex).name, "/BLOQ", False)
    rData = Right$(rData, Len(rData) - 5)
    If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).Blocked = 0 Then
        MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).Blocked = 1
        Call Bloquear(SendTarget.ToMap, userindex, UserList(userindex).pos.Map, UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y, 1)
    Else
        MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).Blocked = 0
        Call Bloquear(SendTarget.ToMap, userindex, UserList(userindex).pos.Map, UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y, 0)
    End If
    Exit Sub
End If



'Quitar NPC
If UCase$(rData) = "/MATA" Then
    rData = Right$(rData, Len(rData) - 5)
    If UserList(userindex).flags.TargetNPC = 0 Then Exit Sub
    Call QuitarNPC(UserList(userindex).flags.TargetNPC)
    Call LogGM(UserList(userindex).name, "/MATA " & Npclist(UserList(userindex).flags.TargetNPC).name, False)
    Exit Sub
End If

'Quita todos los NPCs del area
If UCase$(rData) = "/MASSKILL" Then
    For Y = UserList(userindex).pos.Y - MinYBorder + 1 To UserList(userindex).pos.Y + MinYBorder - 1
            For x = UserList(userindex).pos.x - MinXBorder + 1 To UserList(userindex).pos.x + MinXBorder - 1
                If x > 0 And Y > 0 And x < 101 And Y < 101 Then _
                    If MapData(UserList(userindex).pos.Map, x, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(UserList(userindex).pos.Map, x, Y).NpcIndex)
            Next x
    Next Y
    Call LogGM(UserList(userindex).name, "/MASSKILL", False)
    Exit Sub
End If

'Ultima ip de un char
If UCase(Left(rData, 8)) = "/LASTIP " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)
    rData = Right(rData, Len(rData) - 8)
    
    'No se si sea MUY necesario, pero por si las dudas... ;)
    rData = Replace(rData, "\", "")
    rData = Replace(rData, "/", "")
    
    If FileExist(CharPath & rData & ".chr", vbNormal) Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||La ultima IP de """ & rData & """ fue : " & GetVar(CharPath & rData & ".chr", "INIT", "LastIP") & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Charfile """ & rData & """ inexistente." & FONTTYPE_INFO)
    End If
    Exit Sub
End If





'cambia el MOTD
If UCase(rData) = "/MOTDCAMBIA" Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)
    tStr = "ZMOTD"
    For LoopC = 1 To MaxLines
        tStr = tStr & MOTD(LoopC).texto & vbCrLf
    Next LoopC
    If Right(tStr, 2) = vbCrLf Then tStr = Left(tStr, Len(tStr) - 2)
    Call SendData(SendTarget.ToIndex, userindex, 0, tStr)
    Exit Sub
End If

If UCase(Left(rData, 5)) = "ZMOTD" Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)
    rData = Right(rData, Len(rData) - 5)
    T = Split(rData, vbCrLf)
    
    MaxLines = UBound(T) - LBound(T) + 1
    ReDim MOTD(1 To MaxLines)
    Call WriteVar(App.Path & "\Dat\Motd.ini", "INIT", "NumLines", CStr(MaxLines))
    
    N = LBound(T)
    For LoopC = 1 To MaxLines
        Call WriteVar(App.Path & "\Dat\Motd.ini", "Motd", "Line" & LoopC, T(N))
        MOTD(LoopC).texto = T(N)
        N = N + 1
    Next LoopC
    
    Exit Sub
End If


'Quita todos los NPCs del area
If UCase$(rData) = "/LIMPIAR" Then
        If UserList(userindex).flags.EsRolesMaster Then Exit Sub
        Call LimpiarMundo
        Exit Sub
End If

'Mensaje del sistema
If UCase$(Left$(rData, 6)) = "/SMSG " Then
    rData = Right$(rData, Len(rData) - 6)
    Call LogGM(UserList(userindex).name, "Mensaje de sistema:" & rData, False)
    Call SendData(SendTarget.toall, 0, 0, "!!" & rData & ENDC)
    Exit Sub
End If

'Crear criatura, toma directamente el indice
If UCase$(Left$(rData, 5)) = "/ACC " Then
   rData = Right$(rData, Len(rData) - 5)
   Call LogGM(UserList(userindex).name, "Sumoneo a " & Npclist(val(rData)).name & " en mapa " & UserList(userindex).pos.Map, (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
   Call SpawnNpc(val(rData), UserList(userindex).pos, True, False)
   Exit Sub
End If

'Crear criatura con respawn, toma directamente el indice
If UCase$(Left$(rData, 6)) = "/RACC " Then
 
   rData = Right$(rData, Len(rData) - 6)
   Call LogGM(UserList(userindex).name, "Sumoneo con respawn " & Npclist(val(rData)).name & " en mapa " & UserList(userindex).pos.Map, (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
   Call SpawnNpc(val(rData), UserList(userindex).pos, True, True)
   Exit Sub
End If

If UCase$(Left$(rData, 5)) = "/AI1 " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
   rData = Right$(rData, Len(rData) - 5)
   ArmaduraImperial1 = val(rData)
   Exit Sub
End If

If UCase$(Left$(rData, 5)) = "/AI2 " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
   rData = Right$(rData, Len(rData) - 5)
   ArmaduraImperial2 = val(rData)
   Exit Sub
End If

If UCase$(Left$(rData, 5)) = "/AI3 " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
   rData = Right$(rData, Len(rData) - 5)
   ArmaduraImperial3 = val(rData)
   Exit Sub
End If

If UCase$(Left$(rData, 5)) = "/AI4 " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
   rData = Right$(rData, Len(rData) - 5)
   TunicaMagoImperial = val(rData)
   Exit Sub
End If

If UCase$(Left$(rData, 5)) = "/AC1 " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
   rData = Right$(rData, Len(rData) - 5)
   ArmaduraCaos1 = val(rData)
   Exit Sub
End If

If UCase$(Left$(rData, 5)) = "/AC2 " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
   rData = Right$(rData, Len(rData) - 5)
   ArmaduraCaos2 = val(rData)
   Exit Sub
End If

If UCase$(Left$(rData, 5)) = "/AC3 " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
   rData = Right$(rData, Len(rData) - 5)
   ArmaduraCaos3 = val(rData)
   Exit Sub
End If

If UCase$(Left$(rData, 5)) = "/AC4 " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
   rData = Right$(rData, Len(rData) - 5)
   TunicaMagoCaos = val(rData)
   Exit Sub
End If



'Comando para depurar la navegacion
If UCase$(rData) = "/NAVE" Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    If UserList(userindex).flags.Navegando = 1 Then
        UserList(userindex).flags.Navegando = 0
    Else
        UserList(userindex).flags.Navegando = 1
    End If
    Exit Sub
End If

If UCase$(rData) = "/HABILITAR" Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    If ServerSoloGMs > 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Servidor habilitado para todos" & FONTTYPE_INFO)
        ServerSoloGMs = 0
    Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Servidor restringido a administradores." & FONTTYPE_INFO)
        ServerSoloGMs = 1
    End If
    Exit Sub
End If
Select Case UCase$(rData)
Case "/DENUNCIASON"
If Denuncias = True Then
Call SendData(ToIndex, userindex, 0, "||Las denuncias ya estan activadas" & FONTTYPE_INFO)
Exit Sub
End If
If Denuncias = False Then
Denuncias = True
Call SendData(toall, 0, 0, "||Las denuncias han sido activadas" & FONTTYPE_VENENO)
Exit Sub
End If

Case "/DENUNCIASOFF"
If Denuncias = False Then
Call SendData(ToIndex, userindex, 0, "||Las denuncias estan desactivadas" & FONTTYPE_INFO)
Exit Sub
End If
If Denuncias = True Then
Denuncias = False
Call SendData(toall, 0, 0, "||Las denuncias han sido desactivadas" & FONTTYPE_VENENO)
Exit Sub
End If

End Select
'Apagamos
If UCase$(rData) = "/APAGAR" Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)
    Call SendData(SendTarget.toall, userindex, 0, "||" & UserList(userindex).name & " INTENTA APAGAR EL SERVIDOR!!!" & FONTTYPE_FIGHT)
'    If UCase$(UserList(UserIndex).Name) <> "ALEJOLP" Then
'        Call LogGM(UserList(UserIndex).Name, "¡¡¡Intento apagar el server!!!", False)
'        Exit Sub
'    End If
    'Log
    mifile = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #mifile
    Print #mifile, Date & " " & Time & " server apagado por " & UserList(userindex).name & ". "
    Close #mifile
    Unload frmMain
    Exit Sub
End If

'Reiniciamos
'If UCase$(rdata) = "/REINICIAR" Then
'    Call LogGM(UserList(UserIndex).Name, rdata, False)
'    If UCase$(UserList(UserIndex).Name) <> "ALEJOLP" Then
'        Call LogGM(UserList(UserIndex).Name, "¡¡¡Intento apagar el server!!!", False)
'        Exit Sub
'    End If
'    'Log
'    mifile = FreeFile
'    Open App.Path & "\logs\Main.log" For Append Shared As #mifile
'    Print #mifile, Date & " " & Time & " server reiniciado por " & UserList(UserIndex).Name & ". "
'    Close #mifile
'    ReiniciarServer = 666
'    Exit Sub
'End If

'CONDENA
If UCase$(Left$(rData, 7)) = "/CONDEN" Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)
    rData = Right$(rData, Len(rData) - 8)
    tIndex = NameIndex(rData)
    If tIndex > 0 Then Call VolverCriminal(tIndex)
    Exit Sub
End If

If UCase$(Left$(rData, 7)) = "/RAJAR " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)
    rData = Right$(rData, Len(rData) - 7)
    tIndex = NameIndex(UCase$(rData))
    If tIndex > 0 Then
        Call ResetFacciones(tIndex)
    End If
    Exit Sub
End If

If UCase$(Left$(rData, 11)) = "/RAJARCLAN " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)
    rData = Right$(rData, Len(rData) - 11)
    tInt = modGuilds.m_EcharMiembroDeClan(userindex, rData)  'me da el guildindex
    If tInt = 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "|| No pertenece a ningun clan o es fundador." & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "|| Expulsado." & FONTTYPE_INFO)
        Call SendData(SendTarget.ToGuildMembers, tInt, 0, "|| " & rData & " ha sido expulsado del clan por los administradores del servidor" & FONTTYPE_GUILD)
    End If
    Exit Sub
End If

'lst email
If UCase$(Left$(rData, 11)) = "/LASTEMAIL " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 11)
    If FileExist(CharPath & rData & ".chr") Then
        tStr = GetVar(CharPath & rData & ".chr", "CONTACTO", "email")
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Last email de " & rData & ":" & tStr & FONTTYPE_INFO)
    End If
Exit Sub
End If


'altera password
If UCase$(Left$(rData, 7)) = "/APASS " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)
    rData = Right$(rData, Len(rData) - 7)
    tStr = ReadField(1, rData, Asc("@"))
    If tStr = "" Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||usar /APASS <pjsinpass>@<pjconpass>" & FONTTYPE_INFO)
        Exit Sub
    End If
    tIndex = NameIndex(tStr)
    If tIndex > 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||El usuario a cambiarle el pass (" & tStr & ") esta online, no se puede si esta online" & FONTTYPE_INFO)
        Exit Sub
    End If
    Arg1 = ReadField(2, rData, Asc("@"))
    If Arg1 = "" Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||usar /APASS <pjsinpass> <pjconpass>" & FONTTYPE_INFO)
        Exit Sub
    End If
    If Not FileExist(CharPath & tStr & ".chr") Or Not FileExist(CharPath & Arg1 & ".chr") Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||alguno de los PJs no existe " & tStr & "@" & Arg1 & FONTTYPE_INFO)
    Else
        Arg2 = GetVar(CharPath & Arg1 & ".chr", "INIT", "Password")
        Call WriteVar(CharPath & tStr & ".chr", "INIT", "Password", Arg2)
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Password de " & tStr & " cambiado a: " & Arg2 & FONTTYPE_INFO)
    End If
Exit Sub
End If

'altera email
If UCase$(Left$(rData, 8)) = "/AEMAIL " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)
    rData = Right$(rData, Len(rData) - 8)
    tStr = ReadField(1, rData, Asc("-"))
    If tStr = "" Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||usar /AEMAIL <pj>-<nuevomail>" & FONTTYPE_INFO)
        Exit Sub
    End If
    tIndex = NameIndex(tStr)
    If tIndex > 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||El usuario esta online, no se puede si esta online" & FONTTYPE_INFO)
        Exit Sub
    End If
    Arg1 = ReadField(2, rData, Asc("-"))
    If Arg1 = "" Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||usar /AEMAIL <pj>-<nuevomail>" & FONTTYPE_INFO)
        Exit Sub
    End If
    If Not FileExist(CharPath & tStr & ".chr") Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||No existe el charfile " & CharPath & tStr & ".chr" & FONTTYPE_INFO)
    Else
        Call WriteVar(CharPath & tStr & ".chr", "CONTACTO", "Email", Arg1)
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Email de " & tStr & " cambiado a: " & Arg1 & FONTTYPE_INFO)
    End If
Exit Sub
End If


If UCase$(Left$(rData, 7)) = "/ANAME " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)
    rData = Right$(rData, Len(rData) - 7)
    tStr = ReadField(1, rData, Asc("@"))
    Arg1 = ReadField(2, rData, Asc("@"))
    
    
    If tStr = "" Or Arg1 = "" Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Usar: /ANAME origen@destino" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    tIndex = NameIndex(tStr)
    If tIndex > 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||El Pj esta online, debe salir para el cambio" & FONTTYPE_WARNING)
        Exit Sub
    End If
    
    If FileExist(CharPath & UCase(tStr) & ".chr") = False Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||El pj " & tStr & " es inexistente " & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Arg2 = GetVar(CharPath & UCase(tStr) & ".chr", "GUILD", "GUILDINDEX")
    If IsNumeric(Arg2) Then
        If CInt(Arg2) > 0 Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "||El pj " & tStr & " pertenece a un clan, debe salir del mismo con /salirclan para ser transferido. " & FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    If FileExist(CharPath & UCase(Arg1) & ".chr") = False Then
        FileCopy CharPath & UCase(tStr) & ".chr", CharPath & UCase(Arg1) & ".chr"
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Transferencia exitosa" & FONTTYPE_INFO)
        Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")
        'ponemos la pena
        tInt = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
        Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", tInt + 1)
        Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(userindex).name) & ": BAN POR Cambio de nick a " & UCase$(Arg1) & " " & Date & " " & Time)
    Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "||El nick solicitado ya existe" & FONTTYPE_INFO)
        Exit Sub
    End If
    Exit Sub
End If

If UCase$(rData) = "/CENTINELAACTIVADO" Then

    Centinela.RevisandoUserIndex = 0
    Centinela.clave = 0
    Centinela.TiempoRestante = 0
    
    If CentinelaNPCIndex Then
        Call QuitarNPC(CentinelaNPCIndex)
        CentinelaNPCIndex = 0
    End If
    
    If centinelaActivado Then
        Call SendData(SendTarget.ToAdmins, 0, 0, "El centinela ha sido activado.")
    Else
        Call SendData(SendTarget.ToAdmins, 0, 0, "El centinela ha sido desactivado.")
    End If
End If

If UCase$(Left$(rData, 9)) = "/DOBACKUP" Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)
    Call DoBackUp
    Exit Sub
End If
If UCase$(Left$(rData, 10)) = "/SHOWCMSG " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 10)
    Call modGuilds.GMEscuchaClan(userindex, rData)
    Exit Sub
End If
If UCase$(Left$(rData, 11)) = "/GUARDAMAPA" Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)
    Call GrabarMapa(UserList(userindex).pos.Map, App.Path & "\WorldBackUp\Mapa" & UserList(userindex).pos.Map)
    Exit Sub
End If

If UCase$(Left$(rData, 12)) = "/MODMAPINFO " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)
    rData = Right(rData, Len(rData) - 12)
    Select Case UCase(ReadField(1, rData, 32))
    Case "PK"
        tStr = ReadField(2, rData, 32)
        If tStr <> "" Then
            MapInfo(UserList(userindex).pos.Map).Pk = IIf(tStr = "0", True, False)
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(userindex).pos.Map & ".dat", "Mapa" & UserList(userindex).pos.Map, "Pk", tStr)
        End If
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Mapa " & UserList(userindex).pos.Map & " PK: " & MapInfo(UserList(userindex).pos.Map).Pk & FONTTYPE_INFO)
    Case "BACKUP"
        tStr = ReadField(2, rData, 32)
        If tStr <> "" Then
            MapInfo(UserList(userindex).pos.Map).BackUp = CByte(tStr)
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(userindex).pos.Map & ".dat", "Mapa" & UserList(userindex).pos.Map, "backup", tStr)
        End If
        
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Mapa " & UserList(userindex).pos.Map & " Backup: " & MapInfo(UserList(userindex).pos.Map).BackUp & FONTTYPE_INFO)
    End Select
    Exit Sub
End If

If UCase$(Left$(rData, 7)) = "/GRABAR" Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)
    Call mdParty.ActualizaExperiencias
    Call GuardarUsuarios
    Exit Sub
End If

If UCase$(Left$(rData, 11)) = "/BORRAR SOS" Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)
    Call Ayuda.Reset
    Exit Sub
End If

If UCase$(Left$(rData, 9)) = "/SHOW INT" Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)
    Call frmMain.mnuMostrar_Click
    Exit Sub
End If



If UCase(rData) = "/NOCHE" Then
Call Noche
TiempoClima = TiempoNoche
Clima = "Noche"
Exit Sub
End If
 
If UCase(rData) = "/TARDE" Then
Call Tarde
TiempoClima = TiempoTarde
Clima = "Tarde"
Exit Sub
End If
 
If UCase(rData) = "/MAÑANA" Then
Call Mañana
TiempoClima = TiempoMañana
Clima = "Mañana"
End If
 
If UCase(rData) = "/DIA" Then
Call Dia
TiempoClima = TiempoDia
Clima = "Dia"
Exit Sub
End If

If UCase$(rData) = "/ECHARTODOSPJS" Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)
    Call EcharPjsNoPrivilegiados
    Exit Sub
End If



If UCase$(rData) = "/TCPESSTATS" Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)
    Call SendData(SendTarget.ToIndex, userindex, 0, "||Los datos estan en BYTES." & FONTTYPE_INFO)
    With TCPESStats
        Call SendData(SendTarget.ToIndex, userindex, 0, "||IN/s: " & .BytesRecibidosXSEG & " OUT/s: " & .BytesEnviadosXSEG & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, userindex, 0, "||IN/s MAX: " & .BytesRecibidosXSEGMax & " -> " & .BytesRecibidosXSEGCuando & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, userindex, 0, "||OUT/s MAX: " & .BytesEnviadosXSEGMax & " -> " & .BytesEnviadosXSEGCuando & FONTTYPE_INFO)
    End With
    tStr = ""
    tLong = 0
    For LoopC = 1 To LastUser
        If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
            If UserList(LoopC).ColaSalida.Count > 0 Then
                tStr = tStr & UserList(LoopC).name & " (" & UserList(LoopC).ColaSalida.Count & "), "
                tLong = tLong + 1
            End If
        End If
    Next LoopC
    Call SendData(SendTarget.ToIndex, userindex, 0, "||Posibles pjs trabados: " & tLong & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, userindex, 0, "||" & tStr & FONTTYPE_INFO)
    Exit Sub
End If

If UCase$(rData) = "/RELOADNPCS" Then

    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)

    Call DescargaNpcsDat
    Call CargaNpcsDat
    Call SendData(SendTarget.ToIndex, userindex, 0, "|| Npcs.dat y npcsHostiles.dat recargados." & FONTTYPE_INFO)
    Exit Sub
End If

If UCase$(rData) = "/RELOADSINI" Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)
    Call LoadSini
    Call SendData(SendTarget.ToIndex, userindex, 0, "|| Server.ini Recargado" & FONTTYPE_INFO)
    Exit Sub
End If

If UCase$(rData) = "/RELOADMOTD" Then
   If UserList(userindex).flags.EsRolesMaster Then Exit Sub
   Call LogGM(UserList(userindex).name, rData, False)
   Call LoadMotd
   Call SendData(SendTarget.ToIndex, userindex, 0, "|| MOTD Recargado." & FONTTYPE_INFO)
   Exit Sub
End If

If UCase$(rData) = "/RELOADHECHIZOS" Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)
    Call CargarHechizos
    Call SendData(SendTarget.ToIndex, userindex, 0, "|| Hechizos Recargados." & FONTTYPE_INFO)
    Exit Sub
End If

If UCase$(rData) = "/RELOADOBJ" Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)
    Call LoadOBJData
    Call SendData(SendTarget.ToIndex, userindex, 0, "|| Objetos Recargados." & FONTTYPE_INFO)
    Exit Sub
End If

If UCase$(rData) = "/RELOADGUILD" Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)
    Call LoadGuildsDB
    Call SendData(SendTarget.ToIndex, userindex, 0, "|| Clanes Recargados." & FONTTYPE_INFO)
    Exit Sub
End If

If UCase$(rData) = "/REINICIAR" Then
    If UserList(userindex).name <> "LORWIK" Or UCase$(UserList(userindex).name) <> "MARAXUS" Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)
    Call ReiniciarServidor(True)
    Exit Sub
End If

If UCase$(rData) = "/AUTOUPDATE" Then
    If UserList(userindex).name <> "LORWIK" Or UCase$(UserList(userindex).name) <> "MARAXUS" Then Exit Sub
    Call LogGM(UserList(userindex).name, rData, False)
    Call SendData(SendTarget.ToIndex, userindex, 0, "|| TID: " & CStr(ReiniciarAutoUpdate()) & FONTTYPE_INFO)
    Exit Sub
End If

#If SeguridadAlkon Then
    HandleDataDiosEx userindex, rData
#End If

Exit Sub

ErrorHandler:
 Call LogError("HandleData. CadOri:" & CadenaOriginal & " Nom:" & UserList(userindex).name & "UI:" & userindex & " N: " & Err.Number & " D: " & Err.Description)
 'Resume
 'Call CloseSocket(UserIndex)
 Call Cerrar_Usuario(userindex)
 
 

End Sub

Sub ReloadSokcet()
On Error GoTo errhandler
#If UsarQueSocket = 1 Then

    Call LogApiSock("ReloadSokcet() " & NumUsers & " " & LastUser & " " & MaxUsers)
    
    If NumUsers <= 0 Then
        Call WSApiReiniciarSockets
    Else
'       Call apiclosesocket(SockListen)
'       SockListen = ListenFORCOnnect(Puerto, hWndMsg, "")
    End If

#ElseIf UsarQueSocket = 0 Then

    frmMain.Socket1.Cleanup
    Call ConfigListeningSocket(frmMain.Socket1, Puerto)
    
#ElseIf UsarQueSocket = 2 Then

    

#End If

Exit Sub
errhandler:
    Call LogError("Error en CheckSocketState " & Err.Number & ": " & Err.Description)

End Sub

Public Sub EnviarNoche(ByVal userindex As Integer)

Call SendData(SendTarget.ToIndex, userindex, 0, "NOC" & IIf(DeNoche And (MapInfo(UserList(userindex).pos.Map).Zona = Campo Or MapInfo(UserList(userindex).pos.Map).Zona = Ciudad), "1", "0"))
Call SendData(SendTarget.ToIndex, userindex, 0, "NOC" & IIf(DeNoche, "1", "0"))

End Sub

Public Sub EcharPjsNoPrivilegiados()
Dim LoopC As Long

For LoopC = 1 To LastUser
    If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
        If UserList(LoopC).flags.Privilegios < PlayerType.Consejero Then
            Call CloseSocket(LoopC)
        End If
    End If
Next LoopC

End Sub

