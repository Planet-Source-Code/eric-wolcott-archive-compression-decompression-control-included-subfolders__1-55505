VERSION 5.00
Begin VB.UserControl UserControl1 
   BackColor       =   &H00C0C0FF&
   ClientHeight    =   1275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2295
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1275
   ScaleWidth      =   2295
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   585
      TabIndex        =   0
      Top             =   0
      Width           =   615
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "UserControl1.ctx":0000
         Top             =   15
         Width           =   480
      End
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum StateStr
    i_Ready = 0
    i_Compressing_Archive = 1
    i_Done_Compressing = 2
    i_Building = 3
    i_Done_Building = 4
    i_Extracting = 5
    i_Done_Extracting = 6
    i_Enumerating_Files = 7
    i_Done_Enumerating_Files = 8
    i_Starting_To_Build = 9
    i_Starting_To_Extract = 10
    i_Reading_Archive = 11
    i_Decompressing_Archive = 12
    i_Done_Decompressing = 13
    i_Done = 14
    i_ReturnContents_Reading = 15
    i_ReturnContents_Done = 16
    i_ReturnContents_Decrypting = 17
    i_ReturnContents_Starting = 18
    i_ExtractSingle_Starting = 19
    i_ExtractSingle_Decompressing = 20
    i_ExtractSingle_Reading = 21
    i_ExtractSingle_Done = 22
End Enum

Public WithEvents Huffman1 As clsHuffman
Attribute Huffman1.VB_VarHelpID = -1
Public WithEvents Huffman2 As clsHuffman
Attribute Huffman2.VB_VarHelpID = -1
Public WithEvents Huffman3 As clsHuffman
Attribute Huffman3.VB_VarHelpID = -1
Public WithEvents Huffman4 As clsHuffman
Attribute Huffman4.VB_VarHelpID = -1
Public WithEvents Huffman5 As clsHuffman
Attribute Huffman5.VB_VarHelpID = -1

Public Event StatusChange(State As StateStr)
Public Event BuildArchiveProgress(Percent As Integer)
Public Event ExtractArchiveProgress(Percent As Integer)
Public Event CompressArchiveProgress(Percent As Integer)
Public Event DeCompressArchiveProgress(Percent As Integer)
Public Event DeCompressArchiveUserRequestProgress(Percent As Integer)
Public Event AddingFile(Filename As String, Filepath As String, FileSize As Long)
Public Event ExtractingFileFromArchive(Filename As String, Filepath As String, FileSize As Long)
Public Event CompressionDone(NormalArchiveSize As Long, CompressedArchiveSize As Long)
Public Event ReturnContents(Filename As String, FileSize As Long)
Public Event ReturnContentsDecryptingProcess(Percent As Integer)
Public Event ExtractSingleDecompressProcess(Percent As Integer)


Public Function SaveFilesToArchiv(ByVal sPath As String, ByVal sArchiv As String, CompressFile As Boolean) As Long
RaiseEvent StatusChange(i_Starting_To_Build)
Set Huffman1 = New clsHuffman
      Dim F As Integer
      Dim n As Integer
      Dim nLenFileName As Integer
      Dim nLenFileData As Long
      Dim DirName As String
      Dim FileData As String
      Dim File() As String
      Dim nFiles As Long
      Dim i As Long
      Dim lngUBound As Long
      Dim X, Y, z() As String
      Dim FileNameFull1 As String
      Dim FileNameFull2 As String
      Dim FileNameFull3 As String
RaiseEvent StatusChange(i_Enumerating_Files)
      z = AllFilesInFolders(sPath, True)
      If Right$(sPath, 1) <> "\" Then sPath = sPath + "\"
      nFiles = 0
      For X = 0 To UBound(z) - 1
      DirName = Replace(z(X), sPath, "")
        If DirName <> "." And DirName <> ".." Then
          nFiles = nFiles + 1
          If nFiles > lngUBound Then lngUBound = 2 * nFiles
          ReDim Preserve File(lngUBound)
          File(nFiles) = DirName
        End If
      Next X
RaiseEvent StatusChange(i_Done_Enumerating_Files)
      ReDim Preserve File(nFiles)
      If Dir(sArchiv) <> "" Then Kill sArchiv
      F = FreeFile
RaiseEvent StatusChange(i_Building)
    If CompressFile = True Then
      Open sArchiv & ".tmp" For Binary As #F
    Else
      Open sArchiv For Binary As #F
    End If
      Put #F, , nFiles
      For i = 1 To nFiles
        nLenFileName = Len(File(i))
        Put #F, , nLenFileName
        Put #F, , File(i)
        n = FreeFile
        FileNameFull1 = sPath + File(i)
        FileNameFull2 = Left(FileNameFull1, InStrR(FileNameFull1, "\"))
        FileNameFull3 = Right(FileNameFull1, Len(FileNameFull1) - InStrR(FileNameFull1, "\"))
        Open sPath + File(i) For Binary As #n
        FileData = Space$(LOF(n))
        Get #n, , FileData
        Close #n
        nLenFileData = Len(FileData)
        Put #F, , nLenFileData
        Put #F, , FileData
        DoEvents
RaiseEvent AddingFile(FileNameFull3, FileNameFull2, FileLen(sPath + File(i)))
RaiseEvent BuildArchiveProgress(Int((i / nFiles) * 100))
      Next i
      Close #F
      SaveFilesToArchiv = nFiles
RaiseEvent StatusChange(i_Done_Building)
      If CompressFile = True Then
RaiseEvent StatusChange(i_Compressing_Archive)
            DoEvents
            Call Huffman1.EncodeFile(sArchiv & ".tmp", sArchiv)
            DoEvents
RaiseEvent CompressionDone(FileLen(sArchiv & ".tmp"), FileLen(sArchiv))
            Kill sArchiv & ".tmp"
RaiseEvent StatusChange(i_Done_Compressing)
      Else
RaiseEvent StatusChange(i_Done)
      End If
      DoEvents
RaiseEvent StatusChange(i_Done)
    End Function
    
Public Function ExtractFilesFromArchiv(ByVal sArchiv As String, ByVal sDestDir As String, Decompress As Boolean) As Long
RaiseEvent StatusChange(i_Starting_To_Extract)
Set Huffman2 = New clsHuffman
      Dim F As Integer
      Dim n As Integer
      Dim nLenFileName As Integer
      Dim nLenFileData As Long
      Dim DirName As String
      Dim FileData As String
      Dim File As String
      Dim nFiles As Long
      Dim i As Long
      Dim FileNameFull1 As String
      Dim FileNameFull2 As String
      Dim FileNameFull3 As String
      If Dir(sArchiv) = "" Then
        MsgBox "The archive does not exist!", 16
        Exit Function
      End If
      If Decompress = True Then
RaiseEvent StatusChange(i_Decompressing_Archive)
            Call Huffman2.DecodeFile(sArchiv, sArchiv & ".tmp")
RaiseEvent StatusChange(i_Done_Decompressing)
      End If
      If Right$(sDestDir, 1) <> "\" Then _
        sDestDir = sDestDir + "\"
      F = FreeFile
RaiseEvent StatusChange(i_Reading_Archive)
      If Decompress = True Then
      Open sArchiv & ".tmp" For Binary As #F
      Else
      Open sArchiv For Binary As #F
      End If
      Get #F, , nFiles
RaiseEvent StatusChange(i_Extracting)
      For i = 1 To nFiles
        Get #F, , nLenFileName
        File = Space$(nLenFileName)
        Get #F, , File
        Get #F, , nLenFileData
        FileData = Space$(nLenFileData)
        Get #F, , FileData
        n = FreeFile
        FileNameFull1 = sDestDir + File
        FileNameFull2 = Left(FileNameFull1, InStrR(FileNameFull1, "\"))
        FileNameFull3 = Right(FileNameFull1, Len(FileNameFull1) - InStrR(FileNameFull1, "\"))
        MakeSureDirectoryPathExists FileNameFull1
RaiseEvent ExtractingFileFromArchive(FileNameFull3, FileNameFull2, nLenFileData)
        Open sDestDir + File For Output As #n
        Print #n, FileData;
        Close #n
        DoEvents
RaiseEvent ExtractArchiveProgress(Int((i / nFiles) * 100))
      Next i
      Close #F
      ExtractFilesFromArchiv = nFiles
RaiseEvent StatusChange(i_Done_Extracting)
      If Decompress = True Then
      Kill sArchiv & ".tmp"
      End If
    End Function
Public Function ReadFiles(ByVal sArchiv As String, Decompress As Boolean)
RaiseEvent StatusChange(i_ReturnContents_Starting)
      Set Huffman3 = New clsHuffman
      If Dir(sArchiv) = "" Then
        MsgBox "The archive does not exist!", 16
        Exit Function
      End If
      If Decompress = True Then
RaiseEvent StatusChange(i_ReturnContents_Decrypting)
      Call Huffman3.DecodeFile(sArchiv, sArchiv & ".tmp")
      End If
RaiseEvent StatusChange(i_ReturnContents_Reading)
      Dim F As Integer
      Dim n As Integer
      Dim nLenFileName As Integer
      Dim nLenFileData As Long
      Dim DirName As String
      Dim FileData As String
      Dim File As String
      Dim nFiles As Long
      
      F = FreeFile
      If Decompress = True Then
      Open sArchiv & ".tmp" For Binary As #F
      Else
      Open sArchiv For Binary As #F
      End If
      Get #F, , nFiles
      For i = 1 To nFiles
        Get #F, , nLenFileName
        File = Space$(nLenFileName)
        Get #F, , File
        Get #F, , nLenFileData
        FileData = Space$(nLenFileData)
        Get #F, , FileData
        DoEvents
RaiseEvent ReturnContents(File, nLenFileData)
      Next i
      Close #F
      If Decompress = True Then
      Kill sArchiv & ".tmp"
      End If
RaiseEvent StatusChange(i_ReturnContents_Done)
End Function

Public Function ExtractFile(ByVal sArchiv As String, Filename As String, sDestDir As String, Decompress As Boolean)
RaiseEvent StatusChange(i_ExtractSingle_Starting)
      Set Huffman4 = New clsHuffman
      If Dir(sArchiv) = "" Then
        MsgBox "The archive does not exist!", 16
        Exit Function
      End If
      If Decompress = True Then
RaiseEvent StatusChange(i_ExtractSingle_Decompressing)
      Call Huffman4.DecodeFile(sArchiv, sArchiv & ".tmp")
      End If
RaiseEvent StatusChange(i_ExtractSingle_Reading)
      Dim F As Integer
      Dim n As Integer
      Dim nLenFileName As Integer
      Dim nLenFileData As Long
      Dim DirName As String
      Dim FileData As String
      Dim File As String
      Dim nFiles As Long
      If Right$(sDestDir, 1) <> "\" Then _
        sDestDir = sDestDir + "\"
      F = FreeFile
      If Decompress = True Then
      Open sArchiv & ".tmp" For Binary As #F
      Else
      Open sArchiv For Binary As #F
      End If
      Get #F, , nFiles
      For i = 1 To nFiles
        Get #F, , nLenFileName
        File = Space$(nLenFileName)
        Get #F, , File
        Get #F, , nLenFileData
        FileData = Space$(nLenFileData)
        Get #F, , FileData
        
        If UCase(Filename) = UCase(File) Then
        n = FreeFile
        FileNameFull1 = sDestDir + File
        FileNameFull2 = Left(FileNameFull1, InStrR(FileNameFull1, "\"))
        MakeSureDirectoryPathExists FileNameFull2
        Open sDestDir + File For Output As #n
        Print #n, FileData;
        Close #n
        RaiseEvent StatusChange(i_ExtractSingle_Done)
        Exit Function
        End If
        
        DoEvents
      Next i
      Close #F
      If Decompress = True Then
      Kill sArchiv & ".tmp"
      End If
RaiseEvent StatusChange(i_ExtractSingle_Done)
End Function

Public Function DeCompressArchive(ByVal sArchiv As String, sDestDir As String)
Set Huffman5 = New clsHuffman
Huffman5.DecodeFile sArchiv, sDestDir
End Function

'---------------------------------------------------------------------------------'
Private Sub UserControl_Initialize()
RaiseEvent StatusChange(i_Ready)
End Sub

Private Sub UserControl_Resize()
UserControl.Width = Picture1.Width
UserControl.Height = Picture1.Height
End Sub

Private Sub Huffman1_Progress(Procent As Integer)
    RaiseEvent CompressArchiveProgress(Procent)
    DoEvents
End Sub
Private Sub Huffman2_Progress(Procent As Integer)
    RaiseEvent DeCompressArchiveProgress(Procent)
    DoEvents
End Sub

Private Sub Huffman3_Progress(Procent As Integer)
    RaiseEvent ReturnContentsDecryptingProcess(Procent)
    DoEvents
End Sub

Private Sub Huffman4_Progress(Procent As Integer)
    RaiseEvent ExtractSingleDecompressProcess(Procent)
    DoEvents
End Sub

Private Sub Huffman5_Progress(Procent As Integer)
    RaiseEvent DeCompressArchiveUserRequestProgress(Procent)
    DoEvents
End Sub
