Attribute VB_Name = "LoadGrh"
Option Explicit

Private Type tCabecera 'Cabecera de los con
    desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public Numheads As Integer
Public NumCascos As Integer
Public NumCuerpos As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer
Public NumEscudosAnims As Integer
Public NumFxs As Integer
Public grhCount As Long
Public fileVersion As Long
Public MiCabecera As tCabecera

Public Function LoadGrhData() As Boolean
On Error GoTo ErrorHandler
    Dim Grh As Long
    Dim frame As Long
    Dim handle As Integer
    Dim fileVersion As Long
   
    'Open files
    handle = FreeFile()
    

    Open App.path & "\INIT\Graficos.ind" For Binary Access Read As handle
    Seek #handle, 1
   
    'Get file version
    Get handle, , fileVersion
   
    'Get number of grhs
    Get handle, , grhCount
   
    'Resize arrays
    ReDim GrhData(1 To grhCount) As GrhData
    
    Get handle, , Grh
    
    While Not Grh <= 0
        With GrhData(Grh)
        
            'Get number of frames
            Get handle, , .NumFrames
            If .NumFrames <= 0 Then GoTo ErrorHandler
            
            
           GrhData(Grh).active = True

            ReDim .Frames(1 To GrhData(Grh).NumFrames)
           
            If .NumFrames > 1 Then
                'Read a animation GRH set
                For frame = 1 To .NumFrames
                    Get handle, , .Frames(frame)
                    
                    If .Frames(frame) <= 0 Or .Frames(frame) > grhCount Then
                        GoTo ErrorHandler
                    End If
                Next frame
               
                Get handle, , .Speed
               
                If .Speed <= 0 Then GoTo ErrorHandler
               
                'Compute width and height
                .pixelHeight = GrhData(.Frames(1)).pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
               
                .pixelWidth = GrhData(.Frames(1)).pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
               
                .TileWidth = GrhData(.Frames(1)).TileWidth
                If .TileWidth <= 0 Then GoTo ErrorHandler
               
                .TileHeight = GrhData(.Frames(1)).TileHeight
                If .TileHeight <= 0 Then GoTo ErrorHandler
            Else
                'Read in normal GRH data
                Get handle, , .FileNum
                If .FileNum <= 0 Then GoTo ErrorHandler
               
                Get handle, , GrhData(Grh).SX
                If .SX < 0 Then GoTo ErrorHandler
               
                Get handle, , .SY
                If .SY < 0 Then GoTo ErrorHandler
               
                Get handle, , .pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
               
                Get handle, , .pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
               
                'Compute width and height
                .TileWidth = .pixelWidth / 32
                .TileHeight = .pixelHeight / 32
               
                .Frames(1) = Grh
            End If
            frmmain.cargados.Caption = "Se cargaron " & Grh & " Grh's."
        End With
    Get handle, , Grh
    Wend
   
    Close handle
   
Dim Count As Long
 
    LoadGrhData = True
Exit Function
 
ErrorHandler:
    LoadGrhData = False
End Function

Public Sub Makeini()
Dim i As Long
Dim file_id As Byte
Dim wrote As Long
Dim Animation As String
Dim frame As Integer

Open "Graficos.ini" For Output As #1
Print #1, "[Init]"
Print #1, "NumGrh=" & grhCount
Print #1,

Print #1, "[Graphics]"
file_id = 1
For i = 1 To grhCount
    If wrote = 2000 Then
        wrote = wrote
    End If

    If GrhData(i).NumFrames = 0 And GrhData(i).FileNum = 0 And GrhData(i).SX = 0 Then
    Else
        wrote = wrote + 1
        If GrhData(i).NumFrames > 1 Then
            Animation = ""
            For frame = 1 To GrhData(i).NumFrames
                Animation = Animation & GrhData(i).Frames(frame) & "-"
            Next frame

            Print #1, "Grh" & i & "=" & GrhData(i).NumFrames & "-" & Animation & GrhData(i).Speed
        Else
            Print #1, "Grh" & i & "=" & GrhData(i).NumFrames & "-" & GrhData(i).FileNum & "-" & GrhData(i).SX & "-" & GrhData(i).SY & "-" & GrhData(GrhData(i).Frames(1)).pixelWidth & "-" & GrhData(GrhData(i).Frames(1)).pixelHeight
        End If
    End If
Next i
Close #1

End Sub
Public Sub MakeiniHead()
Dim i As Long
Dim file_id As Byte
Dim wrote As Long
Dim Animation As String
Dim frame As Integer

Open "Cabezas.ini" For Output As #1
Print #1, "[Init]"
Print #1, "NumHeads=" & grhCount
Print #1,

file_id = 1
For i = 1 To NumCascos
    If wrote = 2000 Then
        wrote = wrote
    End If

Print #1, "[HEAD" & i & "]"

        wrote = wrote + 1
        Print #1, "Head1" & "=" & HeadData(i).Head(1).GrhIndex
        Print #1, "Head2" & "=" & HeadData(i).Head(2).GrhIndex
        Print #1, "Head3" & "=" & HeadData(i).Head(3).GrhIndex
        Print #1, "Head4" & "=" & HeadData(i).Head(4).GrhIndex
        Print #1, ""
Next i
Close #1

frmmain.cargados.Caption = "Cabezas desindexadas"

End Sub
Public Function Indexar() As Boolean
    Dim handle
    Dim frame As Long
    Dim i As Long
    Dim tempint As Integer
    Dim MiCabecera As tCabecera
    Dim path As String
    
    path = App.path & "\init\Graficos.ind"
    
    handle = FreeFile()
    
    If FileExists(path) Then
        Call Kill(path)
    End If
    
    Open path For Binary Access Write As handle
    
    'Increment file version
    fileVersion = fileVersion + 1
    
    Put handle, , fileVersion
    
    Put handle, , CLng(UBound(GrhData()))
    
    'Store Grh List
    For i = 1 To UBound(GrhData())
        If GrhData(i).NumFrames > 0 Then
            Put handle, , i
            
            With GrhData(i)
                'Set number of frames
                Put handle, , .NumFrames
                
                If .NumFrames > 1 Then
                    'Read a animation GRH set
                    For frame = 1 To .NumFrames
                        Put handle, , .Frames(frame)
                    Next frame
                    
                    Put handle, , .Speed
                Else
                    'Write in normal GRH data
                    Put handle, , .FileNum
                    
                    Put handle, , .SX
                    
                    Put handle, , .SY
                        
                    Put handle, , .pixelWidth
                    
                    Put handle, , .pixelHeight
                End If
            End With
        End If
    Next i
    
    Close handle
    
    Indexar = True
End Function
Public Function IndexarHead() As Boolean
    Dim handle
    Dim frame As Long
    Dim i As Long
    Dim tempint As Integer
    Dim MiCabecera As tCabecera
    Dim path As String
    
    path = App.path & "\init\Cabezas.ind"
    
    handle = FreeFile()
    
    If FileExists(path) Then
        Call Kill(path)
    End If
    
    Open path For Binary Access Write As handle
    
    'Increment file version
    fileVersion = fileVersion + 1
    
    Put handle, , fileVersion
    
    Put handle, , CLng(UBound(HeadData()))
    
        For i = 1 To UBound(HeadData())
            With HeadData(i)
                'Write in normal GRH data
                Put handle, , .Head(1)
                    
                Put handle, , .Head(2)
                    
                Put handle, , .Head(3)
                        
                Put handle, , .Head(4)
            End With
        Next i
    Close handle
    
    IndexarHead = True
End Function
Sub CargarCabezas()
    Dim N As Integer
    Dim i As Long
    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open App.path & "\init\Cabezas.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
End Sub
Sub CargarCascos()
    Dim N As Integer
    Dim i As Long

    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open App.path & "\init\Cascos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
End Sub
Sub CargarCuerpos()
    Dim N As Integer
    Dim i As Long
    Dim MisCuerpos() As tIndiceCuerpo
    
    N = FreeFile()
    Open App.path & "\init\Personajes.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
            InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
            InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
            InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
            
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
        End If
    Next i
    
    Close #N
End Sub
Sub CargarAnimArmas()
On Error Resume Next

    Dim loopc As Long
    Dim arch As String
    
    arch = App.path & "\init\" & "armas.dat"
    
    NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For loopc = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(loopc).WeaponWalk(1), Val(GetVar(arch, "ARMA" & loopc, "Dir1")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(2), Val(GetVar(arch, "ARMA" & loopc, "Dir2")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(3), Val(GetVar(arch, "ARMA" & loopc, "Dir3")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(4), Val(GetVar(arch, "ARMA" & loopc, "Dir4")), 0
    Next loopc
End Sub
Sub CargarAnimEscudos()
On Error Resume Next

    Dim loopc As Long
    Dim arch As String
    
    arch = App.path & "\init\" & "escudos.dat"
    
    NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopc = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(GetVar(arch, "ESC" & loopc, "Dir1")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(GetVar(arch, "ESC" & loopc, "Dir2")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(GetVar(arch, "ESC" & loopc, "Dir3")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(GetVar(arch, "ESC" & loopc, "Dir4")), 0
    Next loopc
End Sub
Sub CargarFxs()
    Dim N As Integer
    Dim i As Long
    
    N = FreeFile()
    Open App.path & "\init\Fxs.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #N, , FxData(i)
    Next i
    
    Close #N
End Sub
