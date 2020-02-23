Attribute VB_Name = "General"
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Sub ActualizarDlls()

    'objeto para el manejo de ficheros
    'y directorios en visual basic
    Dim dlls As New Scripting.FileSystemObject
    
    'ruta donde se encuentran las dlls originales
    Set directorio = dlls.GetFolder(App.Path & "\librerias")
    
    On Error Resume Next
    ' por cada fichero en el directorio
    'dlls copiar fichero y registrar dll
    For Each fichero In directorio.Files
    'copiar la dll al directorio
    'de sistema de windows
    dlls.CopyFile App.Path & "\librerias\" & fichero.Name, RutaDelSistema & "\" & fichero.Name
    
    'Mostramos que libreria se esta registrando
    Call AddtoRichTextBox(frmMain.Consola, fichero.Name, 255, 255, 255, True, False, False)
    'registar la dll copia
    Shell RutaDelSistema & "\regsvr32 /s " & RutaDelSistema & "\" & fichero.Name, vbHide
    
    Call AddtoRichTextBox(frmMain.Consola, fichero.Name & " Registrado.", 255, 249, 125, True, False, False)
    Next
    
    Call AddtoRichTextBox(frmMain.Consola, "Todas las librerias fueron registradas.", 255, 0, 0, True, False, False)
    Call AddtoRichTextBox(frmMain.Consola, "By Lorwik - www.RincondelAO.com.ar", 255, 0, 0, True, False, False)
End Sub

'funcion que devuelve
'la ruta del directorio del sistema de windows
Public Function RutaDelSistema()

    Dim Car As String * 128
    Dim Longitud, Es As Integer
    Dim Camino As String
    
    Longitud = 128
    
    Es = GetSystemDirectory(Car, Longitud)
    Camino = RTrim$(LCase$(Left$(Car, Es)))
    RutaDelSistema = Camino
End Function

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
'Juan Martín Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
'******************************************r
    With RichTextBox
        If Len(.Text) > 1000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
        
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
        
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
        
        'RichTextBox.Refresh
    End With
End Sub

