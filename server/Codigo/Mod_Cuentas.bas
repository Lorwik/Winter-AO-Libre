Attribute VB_Name = "Mod_Cuentas"
Option Explicit
Public Type pjs
    NamePJ As String
    LvlPJ As Byte
    ClasePJ As eClass
End Type
Public Type Acc
    name As String
    Pass As String
    
    CantPjs As Byte
    PJ(1 To 8) As pjs
End Type
Public Cuenta As Acc

Public Sub CrearCuenta(ByVal UserIndex As Integer, ByVal name As String, ByVal Pass As String, ByVal email As String, ByVal preg As String, ByVal resp As String)
Dim ciclo As Byte
'¿Posee caracteres invalidos?
If Not AsciiValidos(name) Or LenB(name) = 0 Then
    Call WriteErrorMsg(UserIndex, "Nombre invalido.")
    Exit Sub
End If

'Si ya existe la cuenta
If FileExist(App.Path & "\Cuentas\" & name & ".acc", vbNormal) Then
    Call WriteErrorMsg(UserIndex, "El nombre de la cuenta ya existe, por favor ingresa otro.")
    Exit Sub
End If

Call WriteVar(App.Path & "\Cuentas\" & name & ".acc", "CUENTA", "NOMBRE", name)
Call WriteVar(App.Path & "\Cuentas\" & name & ".acc", "CUENTA", "PASSWORD", Pass)
Call WriteVar(App.Path & "\Cuentas\" & name & ".acc", "CUENTA", "MAIL", email)
Call WriteVar(App.Path & "\Cuentas\" & name & ".acc", "CUENTA", "FECHA_CREACION", Now)
Call WriteVar(App.Path & "\Cuentas\" & name & ".acc", "CUENTA", "FECHA_ULTIMA_VISITA", Now)
Call WriteVar(App.Path & "\Cuentas\" & name & ".acc", "CUENTA", "BAN", "0")
Call WriteVar(App.Path & "\Cuentas\" & name & ".acc", "CUENTA", "PREGUNTA", preg)
Call WriteVar(App.Path & "\Cuentas\" & name & ".acc", "CUENTA", "RESPUESTA", resp)
'************************RELLENO LOS PJs************************'
Call WriteVar(App.Path & "\Cuentas\" & name & ".acc", "PERSONAJES", "CANTIDAD_PJS", "0")
For ciclo = 1 To 8
    Call WriteVar(App.Path & "\Cuentas\" & name & ".acc", "PERSONAJES", "PJ" & ciclo, "")
Next ciclo
'************************************************************'

Call EnviarCuenta(UserIndex, "", "0", "0", "1")
End Sub

Public Sub ConectarCuenta(ByVal UserIndex As Integer, ByVal name As String, ByVal Pass As String)
'Si NO existe la cuenta
If Not FileExist(App.Path & "\Cuentas\" & name & ".acc", vbNormal) Then
    Call WriteErrorMsg(UserIndex, "El nombre de la cuenta es inexistente.")
    Exit Sub
End If

With Cuenta
'Si la contraseña es correcta
If Pass = GetVar(App.Path & "\Cuentas\" & name & ".acc", "CUENTA", "PASSWORD") Then
    If GetVar(App.Path & "\Cuentas\" & name & ".acc", "CUENTA", "BAN") <> "0" Then
        Call WriteErrorMsg(UserIndex, "Se ha denegado el acceso a tu cuenta por mal comportamiento en el servidor. Por favor comunicate con los administradores del juego para más información.")
        Exit Sub
    Else
        .CantPjs = GetVar(App.Path & "\Cuentas\" & name & ".acc", "PERSONAJES", "CANTIDAD_PJS")
        
    Dim i As Integer
    
    If Not .CantPjs = 0 Then
        For i = 1 To .CantPjs
            .PJ(i).NamePJ = GetVar(App.Path & "\Cuentas\" & name & ".acc", "PERSONAJES", "PJ" & i)
            Call EnviarCuenta(UserIndex, .PJ(i).NamePJ, i, .CantPjs, "1")
        Next i
   Else
        Call EnviarCuenta(UserIndex, "", 0, 0, "1")
   End If
   
        Call WriteVar(App.Path & "\Cuentas\" & name & ".acc", "CUENTA", "FECHA_ULTIMA_VISITA", Now)
    End If
Else
    Call WriteErrorMsg(UserIndex, "La contraseña es incorrecta. Por favor intentalo nuevamente.")
    Exit Sub
End If
End With
End Sub

Public Sub AgregarPersonaje(ByVal UserIndex As Integer, ByVal CuentaName As String, ByVal UserName As String)
Dim CantidadPJs As Byte
CantidadPJs = GetVar(App.Path & "\Cuentas\" & CuentaName & ".acc", "PERSONAJES", "CANTIDAD_PJS")

WriteVar App.Path & "\Cuentas\" & CuentaName & ".acc", "PERSONAJES", "CANTIDAD_PJS", CantidadPJs + 1
WriteVar App.Path & "\Cuentas\" & CuentaName & ".acc", "PERSONAJES", "PJ" & (CantidadPJs + 1), UserName

WriteVar App.Path & "\Charfile\" & UserName & ".CHR", "INIT", "CUENTA", UCase(CuentaName)
End Sub

Public Sub BorrarPersonaje(ByVal UserIndex As Integer, ByVal CuentaName As String, ByVal PJClickeado As String, ByVal CuentaPassword As String)
Dim limitPJ As Byte
Dim NumPjs As Byte
Dim Archivo As String
Dim i As Byte

    'Si NO existe la cuenta
    If Not FileExist(App.Path & "\Cuentas\" & CuentaName & ".acc", vbNormal) Then
        Call WriteErrorMsg(UserIndex, "El nombre de la cuenta es inexistente.")
        Exit Sub
    End If
    
    If Not CuentaPassword = GetVar(App.Path & "\Cuentas\" & CuentaName & ".acc", "CUENTA", "PASSWORD") Then
        'ESTA TRATANDO DE HACKEAR.
        Call WriteErrorMsg(UserIndex, "HACKER PUTO.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
    End If

    
Archivo = App.Path & "\Cuentas\" & CuentaName & ".acc"
NumPjs = CByte(val(GetVar(Archivo, "PERSONAJES", "CANTIDAD_PJS")))

Dim PJABorrar As String
PJABorrar = GetVar(Archivo, "PERSONAJES", "PJ" & PJClickeado)
    
For i = 1 To val(GetVar(Archivo, "PERSONAJES", "CANTIDAD_PJS"))
    If UCase$(GetVar(Archivo, "PERSONAJES", "PJ" & i)) = UCase$(PJABorrar) Then
        limitPJ = i + 1
        Call WriteVar(Archivo, "PERSONAJES", "PJ" & i, "")
        Call WriteVar(Archivo, "PERSONAJES", "CANTIDAD_PJS", val(GetVar(Archivo, "PERSONAJES", "CANTIDAD_PJS")) - 1)
        BorrarUsuario (PJABorrar)
Exit For
    End If
Next i
                      
For i = limitPJ To NumPjs
    PJABorrar = GetVar(Archivo, "PERSONAJES", "PJ" & i)
    Call WriteVar(Archivo, "PERSONAJES", "PJ" & i, "")
    Call WriteVar(Archivo, "PERSONAJES", "PJ" & i - 1, PJABorrar)
Next i
End Sub

Public Sub CambiarPassword(ByVal UserIndex As Integer, ByVal CuentaName As String, ByVal PsswdAnte As String, ByVal PasswdNew As String)
If Not PasswdNew <> GetVar(App.Path & "\Cuentas\" & CuentaName & ".acc", "CUENTA", "PASSWORD") Then
    Call WriteErrorMsg(UserIndex, "La contraseña antigua que escribio no es correcta.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
End If

Call WriteVar(App.Path & "\Cuentas\" & CuentaName & ".acc", "CUENTA", "PASSWORD", PasswdNew)

Call WriteErrorMsg(UserIndex, "La contraseña fue cambiada.")
End Sub

Public Sub RecuperarCuenta(ByVal UserIndex As Integer, ByVal CuentaName As String, ByVal NewPsswd As String, ByVal Pregunta As String, ByVal Respuesta As String)

    If Not FileExist(App.Path & "\Cuentas\" & CuentaName & ".acc", vbNormal) Then
        Call WriteErrorMsg(UserIndex, "El nombre de la cuenta es inexistente.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    If Pregunta <> GetVar(App.Path & "\Cuentas\" & CuentaName & ".acc", "CUENTA", "PREGUNTA") Then
        Call WriteErrorMsg(UserIndex, "La pregunta secreta no es correcta.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    If Respuesta <> GetVar(App.Path & "\Cuentas\" & CuentaName & ".acc", "CUENTA", "RESPUESTA") Then
        Call WriteErrorMsg(UserIndex, "La respuesta secreta no es correcta.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
Call WriteVar(App.Path & "\Cuentas\" & CuentaName & ".acc", "CUENTA", "PASSWORD", NewPsswd)

Call WriteErrorMsg(UserIndex, "La contraseña fue cambiada.")
End Sub
