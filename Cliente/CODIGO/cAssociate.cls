VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAssociate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Title As String
Public Class As String
Public Extension As String
Public AppCommand As String
Public DefaultIcon As String

Private Declare Function RegCreateKey& Lib "advapi32.DLL" Alias "RegCreateKeyA" (ByVal hKey&, ByVal lpszSubKey$, hKey&)
Private Declare Function RegSetValue& Lib "advapi32.DLL" Alias "RegSetValueA" (ByVal hKey&, ByVal lpszSubKey$, ByVal fdwType&, ByVal lpszValue$, ByVal dwLength&)

Private Const ERROR_SUCCESS = 0&
Private Const ERROR_BADDB = 1&
Private Const ERROR_BADKEY = 2&
Private Const ERROR_CANTOPEN = 3&
Private Const ERROR_CANTREAD = 4&
Private Const ERROR_CANTWRITE = 5&
Private Const ERROR_OUTOFMEMORY = 6&
Private Const ERROR_INVALID_PARAMETER = 7&
Private Const ERROR_ACCESS_DENIED = 8&
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const MAX_PATH = 256&
Private Const REG_SZ = 1

Public Sub Associate()
Dim sKey As String
Dim KeyVal As String
Dim rc As Long
Dim hKey As Long
Dim AppTitle As String, AppClass As String, _
    AppCmdLine As String, DefIcon As String
    
    AppTitle = Title
    AppClass = Class
    AppCmdLine = AppCommand
    DefIcon = DefaultIcon
    
    If AppTitle = "" Then
        Err.Raise 898 + 24, "Associate Class", "No se especificó el titulo"
        Exit Sub
    ElseIf AppCmdLine = "" Then
        Err.Raise 898 + 32, "Associate Class", "Linea de comandos no especificada"
        Exit Sub
    End If
  
    sKey = AppClass
    KeyVal = AppTitle
    rc& = RegCreateKey&(HKEY_CLASSES_ROOT, sKey, hKey&)
    rc& = RegSetValue&(hKey&, "", REG_SZ, KeyVal, 0&)
    
  
    KeyVal = AppClass
    rc& = RegCreateKey&(HKEY_CLASSES_ROOT, Extension, hKey&)
    rc& = RegSetValue&(hKey&, "", REG_SZ, KeyVal, 0&)
       
    KeyVal = AppCmdLine & " %1"
    rc& = RegCreateKey&(HKEY_CLASSES_ROOT, sKey, hKey&)
    rc& = RegSetValue&(hKey&, "shell\open\command", REG_SZ, KeyVal, MAX_PATH)
    

    rc& = RegSetValue&(hKey&, "DefaultIcon", REG_SZ, DefIcon, Len(DefIcon))
End Sub


