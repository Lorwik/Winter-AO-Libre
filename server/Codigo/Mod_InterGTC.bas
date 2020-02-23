Attribute VB_Name = "Mod_InterGTC"
Option Explicit
 
Public Lac_Camina As Long
Public Lac_Pociones As Long
Public Lac_Pegar As Long
Public Lac_Lanzar As Long
Public Lac_Usar As Long
Public Lac_Tirar As Long
 
Public Type TLac
    LCaminar As New clsInterGTC
    LPociones As New clsInterGTC
    LPegar As New clsInterGTC
    LUsar As New clsInterGTC
    LTirar As New clsInterGTC
    LLanzar As New clsInterGTC
End Type
 
Public Sub LoadAntiCheat()
    Dim i As Integer
 
    Lac_Camina = CLng(val(GetVar$(App.Path & "\AntiCheats.ini", "INTERVALOS", "Caminar")))
    Lac_Lanzar = CLng(val(GetVar$(App.Path & "\AntiCheats.ini", "INTERVALOS", "Lanzar")))
    Lac_Usar = CLng(val(GetVar$(App.Path & "\AntiCheats.ini", "INTERVALOS", "Usar")))
    Lac_Tirar = CLng(val(GetVar$(App.Path & "\AntiCheats.ini", "INTERVALOS", "Tirar")))
    Lac_Pociones = CLng(val(GetVar$(App.Path & "\AntiCheats.ini", "INTERVALOS", "Pociones")))
    Lac_Pegar = CLng(val(GetVar$(App.Path & "\AntiCheats.ini", "INTERVALOS", "Pegar")))
 
    For i = 1 To MaxUsers
        ResetearLac i
    Next
   
End Sub
 
Public Sub ResetearLac(UserIndex As Integer)
 
With UserList(UserIndex).Lac
 
    .LCaminar.Init Lac_Camina
    .LPociones.Init Lac_Pociones
    .LUsar.Init Lac_Usar
    .LPegar.Init Lac_Pegar
    .LLanzar.Init Lac_Lanzar
    .LTirar.Init Lac_Tirar
   
End With
 
End Sub
 
Public Sub CargaLac(UserIndex As Integer)
 
With UserList(UserIndex).Lac
 
    Set .LCaminar = New clsInterGTC
    Set .LLanzar = New clsInterGTC
    Set .LPegar = New clsInterGTC
    Set .LPociones = New clsInterGTC
    Set .LTirar = New clsInterGTC
    Set .LUsar = New clsInterGTC
 
    .LCaminar.Init Lac_Camina
    .LPociones.Init Lac_Pociones
    .LUsar.Init Lac_Usar
    .LPegar.Init Lac_Pegar
    .LLanzar.Init Lac_Lanzar
    .LTirar.Init Lac_Tirar
   
End With
 
End Sub
 
Public Sub DescargaLac(UserIndex As Integer)
 
With UserList(UserIndex).Lac
 
    Set .LCaminar = Nothing
    Set .LLanzar = Nothing
    Set .LPegar = Nothing
    Set .LPociones = Nothing
    Set .LTirar = Nothing
    Set .LUsar = Nothing
   
End With
 
End Sub
