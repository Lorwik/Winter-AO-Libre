Attribute VB_Name = "modCompression"
Option Explicit
'To get free bytes in drive
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, FreeBytesToCaller As Currency, bytesTotal As Currency, FreeBytesTotal As Currency) As Long

'This structure will describe our binary file's
'size and number of contained files
Public Type FILEHEADER
    lngFileSize As Long                 'How big is this file? (Used to check integrity)
    intNumFiles As Integer              'How many files are inside?
End Type

'This structure will describe each file contained
'in our binary file
Public Type INFOHEADER
    lngFileStart As Long            'Where does the chunk start?
    lngFileSize As Long             'How big is this chunk of stored data?
    strFileName As String * 16      'What's the name of the file this data came from?
    lngFileSizeUncompressed As Long 'How big is the file compressed
End Type

Public Enum resource_file_type
    graphics
    MP3
    Wav
    Scripts
    Patch
    Interface
    Map
End Enum

Private Const PASSHEADFILE As Long = 749204

Private Const GRAPHIC_PATH As String = "\Graficos\"
Private Const MP3_PATH As String = "\Mp3\"
Private Const WAV_PATH As String = "\Wav\"
Private Const MAP_PATH As String = "\Mapas\"
Private Const INTERFACE_PATH As String = "\Interface\"
Private Const SCRIPT_PATH As String = "\Init\"
Private Const PATCH_PATH As String = "\Patches\"
Private Const OUTPUT_PATH As String = "\Output\"

Private Declare Function Compress Lib "zlib.dll" Alias "compress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function UnCompress Lib "zlib.dll" Alias "uncompress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const MAX_LENGTH = 512

Public Sub Compress_Data(ByRef Data() As Byte)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Compresses binary data avoiding data loses
'*****************************************************************
    Dim Dimensions As Long
    Dim DimBuffer As Long
    Dim BufTemp() As Byte
    Dim BufTemp2() As Byte
    Dim loopc As Long
    
    'Get size of the uncompressed data
    Dimensions = UBound(Data)
    
    'Prepare a buffer 1.06 times as big as the original size
    DimBuffer = Dimensions * 1.06
    ReDim BufTemp(DimBuffer)
    
    'Compress data using zlib
    Compress BufTemp(0), DimBuffer, Data(0), Dimensions
    
    'Deallocate memory used by uncompressed data
    Erase Data
    
    'Get rid of unused bytes in the compressed data buffer
    ReDim Preserve BufTemp(DimBuffer - 1)
    
    'Copy the compressed data buffer to the original data array so it will return to caller
    Data = BufTemp
    
    'Deallocate memory used by the temp buffer
    Erase BufTemp
    
    'Encrypt the first byte of the compressed data for extra security
    Data(0) = Data(0) Xor 166
End Sub

Public Sub Decompress_Data(ByRef Data() As Byte, ByVal OrigSize As Long)
    Dim BufTemp() As Byte
   
    ReDim BufTemp(OrigSize - 1)
    Data(0) = Data(0) Xor 166
    UnCompress BufTemp(0), OrigSize, Data(0), UBound(Data) + 1
   
    ReDim Data(OrigSize - 1)
    Data = BufTemp
    Erase BufTemp
End Sub

Public Function Extract_All_Files(ByVal file_type As resource_file_type, ByVal resource_path As String, Optional ByVal UseOutputFolder As Boolean = False) As Boolean
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Extracts all files from a resource file
'*****************************************************************
    Dim loopc As Long
    Dim SourceFilePath As String
    Dim OutputFilePath As String
    Dim SourceFile As Integer
    Dim SourceData() As Byte
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim handle As Integer
    
'Set up the error handler
On Local Error GoTo ErrHandler
    
    Select Case file_type
        Case graphics
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Graphics.WAO"
            Else
                SourceFilePath = resource_path & "\Graphics.WAO"
            End If
            OutputFilePath = resource_path & GRAPHIC_PATH
        
        Case MP3
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Musics.WAO"
            Else
                SourceFilePath = resource_path & "\Music_HD.WAO"
            End If
            OutputFilePath = resource_path & MP3_PATH
        
        Case Wav
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Sounds.WAO"
            Else
                SourceFilePath = resource_path & "\Sound_HD.WAO"
            End If
            OutputFilePath = resource_path & WAV_PATH
        
        Case Scripts
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Scripts.WAO"
            Else
                SourceFilePath = resource_path & "\Init.WAO"
            End If
            OutputFilePath = resource_path & SCRIPT_PATH
        
        Case Interface
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Interface.WAO"
            Else
                SourceFilePath = resource_path & "\Interface.WAO"
            End If
            OutputFilePath = resource_path & INTERFACE_PATH
            
        Case Map
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Maps.WAO"
            Else
                SourceFilePath = resource_path & "\Maps.WAO"
            End If
            OutputFilePath = resource_path & MAP_PATH
                
        Case Else
            Exit Function
    End Select
    
        If Not Len(Dir(OutputFilePath, vbDirectory)) > 0 Then MkDir OutputFilePath
        
    'Open the binary file
    SourceFile = FreeFile
    Open SourceFilePath For Binary Access Read Lock Write As SourceFile
    
    'Extract the FILEHEADER
    Get SourceFile, 1, FileHead
        
    FileHead.lngFileSize = FileHead.lngFileSize - PASSHEADFILE
        
    'Check the file for validity
    If LOF(SourceFile) <> FileHead.lngFileSize Then
        MsgBox "Resource file " & SourceFilePath & " seems to be corrupted.", , "Error"
        Close SourceFile
        Erase InfoHead
        Exit Function
    End If
    
    'Size the InfoHead array
    ReDim InfoHead(FileHead.intNumFiles - 1)
    
    'Extract the INFOHEADER
    Get SourceFile, , InfoHead
        
    'Extract all of the files from the binary file
    For loopc = 0 To UBound(InfoHead)
        
        'Check if there is enough memory
        If InfoHead(loopc).lngFileSizeUncompressed > General_Drive_Get_Free_Bytes(Left(App.Path, 3)) Then
            MsgBox "There is not enough free memory to continue extracting files."
            Exit Function
        End If
        
        'Resize the byte data array
        ReDim SourceData(InfoHead(loopc).lngFileSize - 1)
        
        'Get the data
        Get SourceFile, InfoHead(loopc).lngFileStart, SourceData
        
        'Decompress all data
        Decompress_Data SourceData, InfoHead(loopc).lngFileSizeUncompressed
        
        'Get a free handler
        handle = FreeFile
        
        'Create a new file and put in the data
        Open OutputFilePath & InfoHead(loopc).strFileName For Binary As handle
        
        Put handle, , SourceData
        
        Close handle
        
        Erase SourceData
        
        DoEvents
    Next loopc
    
    'Close the binary file
    Close SourceFile
    
    Erase InfoHead
    
    MsgBox "Finish"
    Extract_All_Files = True
Exit Function

ErrHandler:
    Close SourceFile
    Erase SourceData
    Erase InfoHead
    'Display an error message if it didn't work
    MsgBox "Unable to decode binary file. Reason: " & Err.Number & " : " & Err.Description, vbOKOnly, "Error"
End Function

Public Function Extract_Patch(ByVal resource_path As String, ByVal File_Name As String) As Boolean
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Comrpesses all files to a resource file
'*****************************************************************
    Dim loopc As Long
    Dim LoopC2 As Long
    Dim LoopC3 As Long
    Dim OutputFile As Integer
    Dim UpdatedFile As Integer
    Dim SourceFilePath As String
    Dim SourceFile As Integer
    Dim SourceData() As Byte
    Dim ResFileHead As FILEHEADER
    Dim ResInfoHead() As INFOHEADER
    Dim UpdatedInfoHead As INFOHEADER
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim RequiredSpace As Currency
    Dim FileExtension As String
    Dim DataOffset As Long
    Dim OutputFilePath As String
    
    'Done flags
    Dim bmp_done As Boolean
    Dim wav_done As Boolean
    Dim mp3_done As Boolean

    Dim ind_done As Boolean
    Dim dat_done As Boolean
    Dim ini_done As Boolean
    
    Dim map_done As Boolean
    Dim gif_done As Boolean
    Dim wao_done As Boolean
    
    '************************************************************************************************
    'This is similar to Extract, but has some small differences to make sure what is being updated
    '************************************************************************************************
'Set up the error handler
On Local Error GoTo ErrHandler
    
    'Open the binary file
    SourceFile = FreeFile
    SourceFilePath = File_Name

    Open SourceFilePath For Binary Access Read Lock Write As SourceFile
    
    'Extract the FILEHEADER
    Get SourceFile, 1, FileHead
        
    'Check the file for validity
    'If LOF(SourceFile) <> FileHead.lngFileSize Then
    '    MsgBox "Resource file " & SourceFilePath & " seems to be corrupted.", , "Error"
    '    Exit Function
    'End If
    
    'Size the InfoHead array
    ReDim InfoHead(FileHead.intNumFiles - 1)
    
    'Extract the INFOHEADER
    Get SourceFile, , InfoHead
    
    'Check if there is enough hard drive space to extract all files
    For loopc = 0 To UBound(InfoHead)
        RequiredSpace = RequiredSpace + InfoHead(loopc).lngFileSizeUncompressed
    Next loopc
    
    If RequiredSpace >= General_Drive_Get_Free_Bytes(Left(App.Path, 3)) Then
        Erase InfoHead
        MsgBox "¡No hay espacio suficiente para extraer el archivo!", , "Error"
        Exit Function
    End If
    
    'Extract all of the files from the binary file
    For loopc = 0 To UBound(InfoHead())
        'Check the extension of the file
        Select Case LCase(Right(Trim(InfoHead(loopc).strFileName), 3))
            '****************************************
            Case Is = "bmp"
                If bmp_done Then GoTo EndMainLoop
                FileExtension = "bmp"
                OutputFilePath = resource_path & "\Graphics.WAO"
                bmp_done = True
            '****************************************
            Case Is = "mp3"
                If mp3_done Then GoTo EndMainLoop
                FileExtension = "mp3"
                OutputFilePath = resource_path & "\Musics.WAO"
                mp3_done = True
            '****************************************
            Case Is = "wav"
                If wav_done Then GoTo EndMainLoop
                FileExtension = "wav"
                OutputFilePath = resource_path & "\Sounds.WAO"
                wav_done = True
            '****************************************
            'Scripts
            Case Is = "ind"
                If ind_done Then GoTo EndMainLoop
                FileExtension = "ind"
                OutputFilePath = resource_path & "\Scripts.WAO"
                ind_done = True
            Case Is = "dat"
                If dat_done Then GoTo EndMainLoop
                FileExtension = "dat"
                OutputFilePath = resource_path & "\Scripts.WAO"
                dat_done = True
            Case Is = "ini"
                If ini_done Then GoTo EndMainLoop
                FileExtension = "ini"
                OutputFilePath = resource_path & "\Scripts.WAO"
                ini_done = True
            Case Is = "wao"
                If ini_done Then GoTo EndMainLoop
                FileExtension = "ini"
                OutputFilePath = resource_path & "\Scripts.WAO"
                ini_done = True
            '****************************************
            'Graphics_INT
            Case Is = "gif"
                If ini_done Then GoTo EndMainLoop
                FileExtension = "gif"
                OutputFilePath = resource_path & "\Interface.WAO"
                gif_done = True
            '****************************************
            'Maps
            Case Is = "map"
                If map_done Then GoTo EndMainLoop
                FileExtension = "map"
                OutputFilePath = resource_path & "\Maps.WAO"
                map_done = True
                
        End Select
        
        OutputFile = FreeFile
        Open OutputFilePath For Binary Access Read Write As OutputFile
        
        'Get file header
        Get OutputFile, 1, ResFileHead
                
        'Resize the Info Header array
        ReDim ResInfoHead(ResFileHead.intNumFiles - 1)
        
        'Load the info header
        Get OutputFile, , ResInfoHead
                
        'Check how many of the files are new, and how many are replacements
        For LoopC2 = loopc To UBound(InfoHead())
            If LCase$(Right$(Trim$(InfoHead(LoopC2).strFileName), 3)) = FileExtension Then
                'Look for same name in the resource file
                For LoopC3 = 0 To UBound(ResInfoHead())
                    If ResInfoHead(LoopC3).strFileName = InfoHead(LoopC2).strFileName Then
                        Exit For
                    End If
                Next LoopC3
                
                'Update the File Head
                If LoopC3 > UBound(ResInfoHead()) Then
                    'Update number of files and size
                    ResFileHead.intNumFiles = ResFileHead.intNumFiles + 1
                    ResFileHead.lngFileSize = ResFileHead.lngFileSize + Len(InfoHead(0)) + InfoHead(LoopC2).lngFileSize
                Else
                    'We substract the size of the old file and add the one of the new one
                    ResFileHead.lngFileSize = ResFileHead.lngFileSize - ResInfoHead(LoopC3).lngFileSize + InfoHead(LoopC2).lngFileSize
                End If
            End If
        Next LoopC2
        
        'Get the offset of the compressed data
        DataOffset = CLng(ResFileHead.intNumFiles) * Len(ResInfoHead(0)) + Len(FileHead) + 1
                
        'Now we start saving the updated file
        UpdatedFile = FreeFile
        Open OutputFilePath & "2" For Binary Access Write Lock Read As UpdatedFile
        
        'Store the filehead
        Put UpdatedFile, 1, ResFileHead
        
        FileHead.lngFileSize = FileHead.lngFileSize + PASSHEADFILE
        
        'Start storing the Info Heads
        LoopC2 = loopc
        For LoopC3 = 0 To UBound(ResInfoHead())
            Do While LoopC2 <= UBound(InfoHead())
                If LCase$(ResInfoHead(LoopC3).strFileName) < LCase$(InfoHead(LoopC2).strFileName) Then Exit Do
                If LCase$(Right$(Trim$(InfoHead(LoopC2).strFileName), 3)) = FileExtension Then
                    'Copy the info head data
                    UpdatedInfoHead = InfoHead(LoopC2)
                    
                    'Set the file start pos and update the offset
                    UpdatedInfoHead.lngFileStart = DataOffset
                    DataOffset = DataOffset + UpdatedInfoHead.lngFileSize
                                        
                    Put UpdatedFile, , UpdatedInfoHead
                    
                    DoEvents
                    
                End If
                LoopC2 = LoopC2 + 1
            Loop
            
            'If the file was replaced in the patch, we skip it
            If LoopC2 Then
                If LCase$(ResInfoHead(LoopC3).strFileName) <= LCase$(InfoHead(LoopC2 - 1).strFileName) Then GoTo EndLoop
            End If
            
            'Copy the info head data
            UpdatedInfoHead = ResInfoHead(LoopC3)
            
            'Set the file start pos and update the offset
            UpdatedInfoHead.lngFileStart = DataOffset
            DataOffset = DataOffset + UpdatedInfoHead.lngFileSize
                        
            Put UpdatedFile, , UpdatedInfoHead
EndLoop:
        Next LoopC3
        
        'If there was any file in the patch that would go in the bottom of the list we put it now
        For LoopC2 = LoopC2 To UBound(InfoHead())
            If LCase$(Right$(Trim$(InfoHead(LoopC2).strFileName), 3)) = FileExtension Then
                'Copy the info head data
                UpdatedInfoHead = InfoHead(LoopC2)
                
                'Set the file start pos and update the offset
                UpdatedInfoHead.lngFileStart = DataOffset
                DataOffset = DataOffset + UpdatedInfoHead.lngFileSize
                                
                Put UpdatedFile, , UpdatedInfoHead
            End If
        Next LoopC2
        
        'Now we start adding the compressed data
        LoopC2 = loopc
        For LoopC3 = 0 To UBound(ResInfoHead())
            Do While LoopC2 <= UBound(InfoHead())
                If LCase$(ResInfoHead(LoopC3).strFileName) < LCase$(InfoHead(LoopC2).strFileName) Then Exit Do
                If LCase$(Right$(Trim$(InfoHead(LoopC2).strFileName), 3)) = FileExtension Then
                    'Get the compressed data
                    ReDim SourceData(InfoHead(LoopC2).lngFileSize - 1)
                    
                    Get SourceFile, InfoHead(LoopC2).lngFileStart, SourceData
                    
                    Put UpdatedFile, , SourceData
                End If
                LoopC2 = LoopC2 + 1
            Loop
            
            'If the file was replaced in the patch, we skip it
            If LoopC2 Then
                If LCase$(ResInfoHead(LoopC3).strFileName) <= LCase$(InfoHead(LoopC2 - 1).strFileName) Then GoTo EndLoop2
            End If
            
            'Get the compressed data
            ReDim SourceData(ResInfoHead(LoopC3).lngFileSize - 1)
            
            Get OutputFile, ResInfoHead(LoopC3).lngFileStart, SourceData
            
            Put UpdatedFile, , SourceData
EndLoop2:
        Next LoopC3
        
        'If there was any file in the patch that would go in the bottom of the lsit we put it now
        For LoopC2 = LoopC2 To UBound(InfoHead())
            If LCase$(Right$(Trim$(InfoHead(LoopC2).strFileName), 3)) = FileExtension Then
                'Get the compressed data
                ReDim SourceData(InfoHead(LoopC2).lngFileSize - 1)
                
                Get SourceFile, InfoHead(LoopC2).lngFileStart, SourceData
                
                Put UpdatedFile, , SourceData
            End If
        Next LoopC2
        
        'We are done updating the file
        Close UpdatedFile
        
        'Close and delete the old resource file
        Close OutputFile
        Kill OutputFilePath
        
        'Rename the new one
        Name OutputFilePath & "2" As OutputFilePath
        
        'Deallocate the memory used by the data array
        Erase SourceData
EndMainLoop:
    Next loopc
    
    'Close the binary file
    Close SourceFile
    
    Erase InfoHead
    Erase ResInfoHead
    
    Extract_Patch = True
Exit Function

ErrHandler:
    Erase SourceData
    Erase InfoHead

End Function

Public Function Compress_Files(ByVal file_type As resource_file_type, ByVal resource_path As String, ByVal dest_path As String) As Boolean
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Comrpesses all files to a resource file
'*****************************************************************
    Dim SourceFilePath As String
    Dim SourceFileExtension As String
    Dim OutputFilePath As String
    Dim SourceFile As Long
    Dim OutputFile As Long
    Dim SourceFileName As String
    Dim SourceData() As Byte
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim FileNames() As String
    Dim lngFileStart As Long
    Dim loopc As Long
    
'Set up the error handler
On Local Error GoTo ErrHandler
    
    Select Case file_type
        Case graphics
            SourceFilePath = resource_path & GRAPHIC_PATH
            SourceFileExtension = ".bmp"
            OutputFilePath = dest_path & "Graphics.WAO"

        Case MP3
            SourceFilePath = resource_path & MP3_PATH
            SourceFileExtension = ".mp3"
            OutputFilePath = dest_path & "Musics.WAO"
        
        Case Wav
            SourceFilePath = resource_path & WAV_PATH
            SourceFileExtension = ".wav"
            OutputFilePath = dest_path & "Sounds.WAO"
                
        Case Scripts
            SourceFilePath = resource_path & SCRIPT_PATH
            SourceFileExtension = ".*"
            OutputFilePath = dest_path & "Scripts.WAO"
        
        Case Patch
            SourceFilePath = resource_path & PATCH_PATH
            SourceFileExtension = ".*"
            OutputFilePath = dest_path & "tmp.WAO"
    
        Case Interface
            SourceFilePath = resource_path & INTERFACE_PATH
            SourceFileExtension = ".gif"
            OutputFilePath = dest_path & "Interface.WAO"
            
        Case Map
            SourceFilePath = resource_path & MAP_PATH
            SourceFileExtension = ".map"
            OutputFilePath = dest_path & "Maps.WAO"
    End Select
    
    'Get first file in the directoy
    SourceFileName = Dir$(SourceFilePath & "*" & SourceFileExtension, vbNormal)
    
    SourceFile = FreeFile
    
    'Get all other files i nthe directory
    While SourceFileName <> ""
        FileHead.intNumFiles = FileHead.intNumFiles + 1
        
        ReDim Preserve FileNames(FileHead.intNumFiles - 1)
        FileNames(FileHead.intNumFiles - 1) = LCase(SourceFileName)
        
        'Search new file
        SourceFileName = Dir$()
    Wend
    
    'If we found none, be can't compress a thing, so we exit
    If FileHead.intNumFiles = 0 Then
        MsgBox "There are no files of extension " & SourceFileExtension & " in " & SourceFilePath & ".", , "Error"
        Exit Function
    End If
    
    'Sort file names alphabetically (this will make patching much easier).
    General_Quick_Sort FileNames(), 0, UBound(FileNames)
    
    'Resize InfoHead array
    ReDim InfoHead(FileHead.intNumFiles - 1)
        
    'Destroy file if it previuosly existed
    If Dir(OutputFilePath, vbNormal) <> "" Then
        Kill OutputFilePath
    End If
    
    'Open a new file
    OutputFile = FreeFile
    Open OutputFilePath For Binary Access Read Write As OutputFile
    
    For loopc = 0 To FileHead.intNumFiles - 1
        'Find a free file number to use and open the file
        SourceFile = FreeFile
        Open SourceFilePath & FileNames(loopc) For Binary Access Read Lock Write As SourceFile
        
        'Store file name
        InfoHead(loopc).strFileName = FileNames(loopc)
        
        'Find out how large the file is and resize the data array appropriately
        ReDim SourceData(LOF(SourceFile) - 1)
        
        'Store the value so we can decompress it later on
        InfoHead(loopc).lngFileSizeUncompressed = LOF(SourceFile)
        
        'Get the data from the file
        Get SourceFile, , SourceData
        
        'Compress it
        Compress_Data SourceData
        
        'Save it to a temp file
        Put OutputFile, , SourceData
        
        'Set up the file header
        FileHead.lngFileSize = FileHead.lngFileSize + UBound(SourceData) + 1
        
        'Set up the info headers
        InfoHead(loopc).lngFileSize = UBound(SourceData) + 1
        
        Erase SourceData
        
        'Close temp file
        Close SourceFile
        
        DoEvents
    Next loopc
    
    'Finish setting the FileHeader data
    FileHead.lngFileSize = FileHead.lngFileSize + CLng(FileHead.intNumFiles) * Len(InfoHead(0)) + Len(FileHead)
    
    'Set InfoHead data
    lngFileStart = Len(FileHead) + CLng(FileHead.intNumFiles) * Len(InfoHead(0)) + 1
    For loopc = 0 To FileHead.intNumFiles - 1
        InfoHead(loopc).lngFileStart = lngFileStart
        lngFileStart = lngFileStart + InfoHead(loopc).lngFileSize
    Next loopc
        
    '************ Write Data
    FileHead.lngFileSize = FileHead.lngFileSize + PASSHEADFILE
    'Get all data stored so far
    ReDim SourceData(LOF(OutputFile) - 1)
    Seek OutputFile, 1
    Get OutputFile, , SourceData
    
    Seek OutputFile, 1
    
    'Store the data in the file
    Put OutputFile, , FileHead
    Put OutputFile, , InfoHead
    Put OutputFile, , SourceData
    
    'Close the file
    Close OutputFile
    
    Erase InfoHead
    Erase SourceData
    
    MsgBox "Finish"
Exit Function

ErrHandler:
    Erase SourceData
    Erase InfoHead
    'Display an error message if it didn't work
    MsgBox "Unable to create binary file. Reason: " & Err.Number & " : " & Err.Description, vbOKOnly, "Error"
End Function

Public Function Extract_File(ByVal file_type As resource_file_type, ByVal resource_path As String, ByVal File_Name As String, ByVal OutputFilePath As String, Optional ByVal UseOutputFolder As Boolean = False) As Boolean
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Extracts all files from a resource file
'*****************************************************************
    Dim loopc As Long
    Dim SourceFilePath As String
    Dim SourceData() As Byte
    Dim InfoHead As INFOHEADER
    Dim handle As Integer
    
'Set up the error handler
On Local Error GoTo ErrHandler
    
    Select Case file_type
        Case graphics
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Graphics.WAO"
            Else
                SourceFilePath = resource_path & "\Graphics.WAO"
            End If
            
        Case MP3
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Musics.WAO"
            Else
                SourceFilePath = resource_path & "\Musics.WAO"
            End If
        
        Case Wav
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Sounds.WAO"
            Else
                SourceFilePath = resource_path & "\Sounds.WAO"
            End If
        
        Case Scripts
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Scripts.WAO"
            Else
                SourceFilePath = resource_path & "\Scripts.WAO"
            End If
        
        Case Interface
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Interface.WAO"
            Else
                SourceFilePath = resource_path & "\Interface.WAO"
            End If
            
        Case Map
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Maps.WAO"
            Else
                SourceFilePath = resource_path & "\Maps.WAO"
            End If
            
        Case Else
            Exit Function
    End Select
    
    'Find the Info Head of the desired file
    InfoHead = File_Find(SourceFilePath, File_Name)
    
    If InfoHead.strFileName = "" Or InfoHead.lngFileSize = 0 Then Exit Function

    'Open the binary file
    handle = FreeFile
    Open SourceFilePath For Binary Access Read Lock Write As handle
    
    'Check the file for validity
    'If LOF(handle) <> InfoHead.lngFileSize Then
    '    Close handle
    '    MsgBox "Resource file " & SourceFilePath & " seems to be corrupted.", , "Error"
    '    Exit Function
    'End If
    
    'Make sure there is enough space in the HD
    If InfoHead.lngFileSizeUncompressed > General_Drive_Get_Free_Bytes(Left$(App.Path, 3)) Then
        Close handle
        MsgBox "There is not enough drive space to extract the compressed file.", , "Error"
        Exit Function
    End If
    
    'Extract file from the binary file
    
    'Resize the byte data array
    ReDim SourceData(InfoHead.lngFileSize - 1)
    
    'Get the data
    Get handle, InfoHead.lngFileStart, SourceData
    
    'Decompress all data
    Decompress_Data SourceData, InfoHead.lngFileSizeUncompressed
    
    'Close the binary file
    Close handle
    
    'Get a free handler
    handle = FreeFile
    
    Open OutputFilePath & InfoHead.strFileName For Binary As handle
    
    Put handle, 1, SourceData
    
    Close handle
    
    Erase SourceData
        
    Extract_File = True
Exit Function

ErrHandler:
    Close handle
    Erase SourceData
    'Display an error message if it didn't work
    'MsgBox "Unable to decode binary file. Reason: " & Err.number & " : " & Err.Description, vbOKOnly, "Error"
End Function

Public Sub Delete_File(ByVal file_path As String)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 3/03/2005
'Deletes a resource files
'*****************************************************************
    Dim handle As Integer
    Dim Data() As Byte
    
    On Error GoTo Error_Handler
    
    'We open the file to delete
    handle = FreeFile
    Open file_path For Binary Access Write Lock Read As handle
    'We replace all the bytes in it with 0s
    ReDim Data(LOF(handle) - 1)
    Put handle, 1, Data
    
    'We close the file
    Close handle
    
    'Now we delete it, knowing that if they retrieve it (some antivirus may create backup copies of deleted files), it will be useless
    Kill file_path
    
    Exit Sub
    
Error_Handler:
    Kill file_path
    Debug.Print Err.Description
End Sub
Public Function Extract_File_Memory(ByVal file_type As resource_file_type, ByVal resource_path As String, ByVal File_Name As String, ByRef SourceData() As Byte) As Boolean
 '******************************************************************************************
 'Lorwik> Con este sub podemos descomprimir los recursos en la memoria en vez de en el disco
 '******************************************************************************************
 
    Dim loopc As Long
    Dim SourceFilePath As String
    Dim InfoHead As INFOHEADER
    Dim handle As Integer
   
On Local Error GoTo ErrHandler
   
    Select Case file_type
        Case graphics
                SourceFilePath = resource_path & "\Graphics.WAO"
            
        Case MP3
                SourceFilePath = resource_path & "\Musics.WAO"
        
        Case Wav
                SourceFilePath = resource_path & "\Sounds.WAO"

        Case Scripts
                SourceFilePath = resource_path & "\Scripts.WAO"

        Case Interface
                SourceFilePath = resource_path & "\Interface.WAO"

        Case Map
                SourceFilePath = resource_path & "\Maps.WAO"

        Case Else
            Exit Function
    End Select
   
    InfoHead = File_Find(SourceFilePath, File_Name)
   
    If InfoHead.strFileName = "" Or InfoHead.lngFileSize = 0 Then Exit Function
 
    handle = FreeFile
    Open SourceFilePath For Binary Access Read Lock Write As handle
   
    If InfoHead.lngFileSizeUncompressed > General_Drive_Get_Free_Bytes(Left$(App.Path, 3)) Then
        Close handle
        MsgBox "There is not enough drive space to extract the compressed file.", , "Error"
        Exit Function
    End If
   
   
    ReDim SourceData(InfoHead.lngFileSize - 1)
   
    Get handle, InfoHead.lngFileStart, SourceData
        Decompress_Data SourceData, InfoHead.lngFileSizeUncompressed
    Close handle
       
    Extract_File_Memory = True
Exit Function
 
ErrHandler:
    Close handle
    Erase SourceData
End Function

Public Function File_Find(ByVal resource_file_path As String, ByVal File_Name As String) As INFOHEADER
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'Looks for a compressed file in a resource file. Uses binary search ;)
'**************************************************************
On Error GoTo ErrHandler
    Dim max As Integer  'Max index
    Dim min As Integer  'Min index
    Dim mid As Integer  'Middle index
    Dim file_handler As Integer
    Dim file_head As FILEHEADER
    Dim info_head As INFOHEADER
    
    'Fill file name with spaces for compatibility
    If Len(File_Name) < Len(info_head.strFileName) Then _
        File_Name = File_Name & Space$(Len(info_head.strFileName) - Len(File_Name))
    
    'Open resource file
    file_handler = FreeFile
    Open resource_file_path For Binary Access Read Lock Write As file_handler
    
    'Get file head
    Get file_handler, 1, file_head
    
    min = 1
    max = file_head.intNumFiles
    
    Do While min <= max
        mid = (min + max) / 2
        
        'Get the info header of the appropiate compressed file
        Get file_handler, CLng(Len(file_head) + CLng(Len(info_head)) * CLng((mid - 1)) + 1), info_head
                
        If File_Name < info_head.strFileName Then
            If max = mid Then
                max = max - 1
            Else
                max = mid
            End If
        ElseIf File_Name > info_head.strFileName Then
            If min = mid Then
                min = min + 1
            Else
                min = mid
            End If
        Else
            'Copy info head
            File_Find = info_head
            
            'Close file and exit
            Close file_handler
            Exit Function
        End If
    Loop
    
ErrHandler:
    'Close file
    Close file_handler
    File_Find.strFileName = ""
    File_Find.lngFileSize = 0
End Function

Public Function General_Get_Temp_Dir() As String
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/11/2005
'Gets windows temporary directory
'**************************************************************
   Dim s As String
   Dim c As Long
   s = Space$(MAX_LENGTH)
   c = GetTempPath(MAX_LENGTH, s)
   If c > 0 Then
       If c > Len(s) Then
           s = Space$(c + 1)
           c = GetTempPath(MAX_LENGTH, s)
       End If
   End If
   General_Get_Temp_Dir = IIf(c > 0, Left$(s, c), "")
End Function


Public Sub General_Quick_Sort(ByRef SortArray As Variant, ByVal first As Long, ByVal last As Long)
'**************************************************************
'Author: juan Martín Sotuyo Dodero
'Last Modify Date: 3/03/2005
'Good old QuickSort algorithm :)
'**************************************************************
    Dim Low As Long, High As Long
    Dim temp As Variant
    Dim List_Separator As Variant
    
    Low = first
    High = last
    List_Separator = SortArray((first + last) / 2)
    Do While (Low <= High)
        Do While SortArray(Low) < List_Separator
            Low = Low + 1
        Loop
        Do While SortArray(High) > List_Separator
            High = High - 1
        Loop
        If Low <= High Then
            temp = SortArray(Low)
            SortArray(Low) = SortArray(High)
            SortArray(High) = temp
            Low = Low + 1
            High = High - 1
        End If
    Loop
    If first < High Then General_Quick_Sort SortArray, first, High
    If Low < last Then General_Quick_Sort SortArray, Low, last
End Sub

Public Function General_Drive_Get_Free_Bytes(ByVal DriveName As String) As Currency
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 6/07/2004
'
'**************************************************************
    Dim RetVal As Long
    Dim FB As Currency
    Dim BT As Currency
    Dim FBT As Currency
    
    RetVal = GetDiskFreeSpace(Left(DriveName, 2), FB, BT, FBT)
    
    General_Drive_Get_Free_Bytes = FB * 10000 'convert result to actual size in bytes
End Function



