VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTexManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'MODULO CLSTEXMANAGER REPARADO POR LORWIK
'WWWW.RINCONDELAO.COM.AR


Const HASH_TABLE_SIZE As Long = 337
Private Const BYTES_PER_MB As Long = 1048576                        '1Mb = 1024 Kb = 1024 * 1024 bytes = 1048576 bytes
Private Const MIN_MEMORY_TO_USE As Long = 16 * BYTES_PER_MB          '4 Mb

Private Type SURFACE_ENTRY_DYN
    FileName As Integer
    UltimoAcceso As Long
    Texture As Direct3DTexture8
    size As Long
    texture_width As Integer
    texture_height As Integer
End Type

Private Type HashNode
    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN
End Type

Private TexList(HASH_TABLE_SIZE - 1) As HashNode

Private mD3D As D3DX8
Private device As Direct3DDevice8

Private mMaxEntries As Integer
Private mCantidadGraficos As Integer
Private maxBytesToUse As Long
Private mFreeMemoryBytes As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long
'MODULO CLSTEXMANAGER REPARADO POR LORWIK
'WWWW.RINCONDELAO.COM.AR

Private Sub Class_Terminate()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'WWWW.RINCONDELAO.COM.AR
'**************************************************************
    Dim i As Long
    Dim j As Long
    
    'Destroy every surface in memory
    For i = 0 To HASH_TABLE_SIZE - 1
        With TexList(i)
            For j = 1 To .surfaceCount
                Set .SurfaceEntry(j).Texture = Nothing
            Next j
            
            'Destroy the arrays
            Erase .SurfaceEntry
        End With
    Next i
End Sub

Public Function GetTexture(ByVal FileName As Integer, ByRef textwidth As Long, ByRef textheight As Long) As Direct3DTexture8
'WWWW.RINCONDELAO.COM.AR
If FileName = 0 Then
Debug.Print "0 GRH ATMPT TO BE LOADED"
Exit Function
End If

    Dim i As Long
    ' Search the index on the list
    With TexList(FileName Mod HASH_TABLE_SIZE)
        For i = 1 To .surfaceCount
            If .SurfaceEntry(i).FileName = FileName Then
                .SurfaceEntry(i).UltimoAcceso = GetTickCount
                textwidth = .SurfaceEntry(i).texture_width
                textheight = .SurfaceEntry(i).texture_height
                Set GetTexture = .SurfaceEntry(i).Texture
                Exit Function
            End If
        Next i
    End With

    'Not in memory, load it!
    Set GetTexture = CrearGrafico(FileName, textwidth, textheight)
End Function
'MODULO CLSTEXMANAGER REPARADO POR LORWIK
'WWWW.RINCONDELAO.COM.AR

Public Function Init(ByRef D3D8 As D3DX8, ByRef d3d_device As Direct3DDevice8, ByVal MaxMemory As Long) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'
'**************************************************************
    mCantidadGraficos = 0
    
    'Seteamos el objeto
    Set mD3D = D3D8
    Set device = d3d_device
    mFreeMemoryBytes = 0
    maxBytesToUse = MIN_MEMORY_TO_USE
    
    Init = True
End Function

Private Function CrearGrafico(ByVal Archivo As Integer, ByRef texwidth As Long, ByRef textheight As Long) As Direct3DTexture8
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'menduz was here
'
'**************************************************************
On Error GoTo ErrHandler
    Dim surface_desc As D3DSURFACE_DESC
    Dim texture_info As D3DXIMAGE_INFO
    Dim index As Integer
    Dim InfoHead As INFOHEADER
    Dim Data() As Byte
    
    index = Archivo Mod HASH_TABLE_SIZE
    With TexList(index)
        .surfaceCount = .surfaceCount + 1
        ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
        With .SurfaceEntry(.surfaceCount)
            'Nombre
            .FileName = Archivo
            
            'Ultimo acceso
            .UltimoAcceso = GetTickCount
            
            InfoHead = File_Find(App.Path & "\RECURSOS\Graphics.WAO", CStr(Archivo) & ".bmp")
                        'Si existe usamos BMP
        If InfoHead.lngFileSize <> 0 Then
            Extract_File_Memory graphics, App.Path & "\Recursos\", LCase$(CStr(Archivo) & ".bmp"), Data()
            
            Set .Texture = mD3D.CreateTextureFromFileInMemoryEx(device, Data(0), UBound(Data) + 1, _
                D3DX_DEFAULT, D3DX_DEFAULT, 3, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                D3DX_FILTER_NONE, &HFF000000, texture_info, ByVal 0)
                
            Erase Data
        End If
            .Texture.GetLevelDesc 0, surface_desc
            .texture_width = texture_info.Width
            .texture_height = texture_info.Height
            .size = surface_desc.size
            texwidth = .texture_width
            textheight = .texture_height
            Set CrearGrafico = .Texture
            mFreeMemoryBytes = mFreeMemoryBytes + surface_desc.size
        End With
    End With
    Debug.Print mFreeMemoryBytes / 1024 / 1024; " MB LIBRES"
    Do While mFreeMemoryBytes < 0
        If Not RemoveLRU() Then
            Exit Do
        End If
    Loop
Exit Function
ErrHandler:
Debug.Print "ERROR EN GRHLOAD>" & Archivo & ".bmp"
End Function

'MODULO CLSTEXMANAGER REPARADO POR LORWIK
'WWWW.RINCONDELAO.COM.AR

Private Function RemoveLRU() As Boolean
'**************************************************************
'Author: Juan Mart?n Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Removes the Least Recently Used surface to make some room for new ones
'WWWW.RINCONDELAO.COM.AR
'**************************************************************
    Dim LRUi As Long
    Dim LRUj As Long
    Dim LRUtime As Long
    Dim i As Long
    Dim j As Long
    Dim surface_desc As D3DSURFACE_DESC
    
    LRUtime = GetTickCount
    
    'Check out through the whole list for the least recently used
    For i = 0 To HASH_TABLE_SIZE - 1
        With TexList(i)
            For j = 1 To .surfaceCount
                If LRUtime > .SurfaceEntry(j).UltimoAcceso Then
                    LRUi = i
                    LRUj = j
                    LRUtime = .SurfaceEntry(j).UltimoAcceso
                End If
            Next j
        End With
    Next i
    
    'Retrieve the surface desc
    Call TexList(LRUi).SurfaceEntry(LRUj).Texture.GetLevelDesc(0, surface_desc)
    
    'Remove it
    Set TexList(LRUi).SurfaceEntry(LRUj).Texture = Nothing
    TexList(LRUi).SurfaceEntry(LRUj).FileName = 0
    
    'Move back the list (if necessary)
    If LRUj Then
        RemoveLRU = True
        
        With TexList(LRUi)
            For j = LRUj To .surfaceCount - 1
                .SurfaceEntry(j) = .SurfaceEntry(j + 1)
            Next j
            
            .surfaceCount = .surfaceCount - 1
            If .surfaceCount Then
                ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
            Else
                Erase .SurfaceEntry
            End If
        End With
    End If
    
    'Update the used bytes
    mFreeMemoryBytes = mFreeMemoryBytes + surface_desc.size
End Function
'MODULO CLSTEXMANAGER REPARADO POR LORWIK
'WWWW.RINCONDELAO.COM.AR


