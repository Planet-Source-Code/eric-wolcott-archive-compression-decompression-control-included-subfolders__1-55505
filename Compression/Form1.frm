VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Read File List From Archive"
      Height          =   2850
      Left            =   75
      TabIndex        =   20
      Top             =   6480
      Width           =   8805
      Begin VB.ListBox List1 
         Height          =   2205
         Left            =   1980
         TabIndex        =   25
         Top             =   585
         Width           =   4350
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Read File List From Archive"
         Height          =   2475
         Left            =   135
         TabIndex        =   24
         Top             =   255
         Width           =   1770
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1980
         TabIndex        =   23
         Text            =   "<Status>"
         Top             =   255
         Width           =   3030
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5040
         TabIndex        =   22
         Text            =   "%"
         Top             =   255
         Width           =   1290
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Extract Selected File From Archive"
         Height          =   2475
         Left            =   6435
         TabIndex        =   21
         Top             =   285
         Width           =   2265
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Extract Archive"
      Height          =   2625
      Left            =   75
      TabIndex        =   4
      Top             =   3840
      Width           =   8805
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         Height          =   1875
         Left            =   90
         ScaleHeight     =   1815
         ScaleWidth      =   8550
         TabIndex        =   13
         Top             =   225
         Width           =   8610
         Begin VB.TextBox Text6 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   45
            TabIndex        =   19
            Top             =   1500
            Width           =   8460
         End
         Begin Project1.CheckBox CheckBox6 
            Height          =   300
            Left            =   30
            TabIndex        =   14
            Top             =   1185
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   529
            Caption         =   ""
            Enabled         =   0   'False
         End
         Begin Project1.CheckBox CheckBox7 
            Height          =   270
            Left            =   30
            TabIndex        =   15
            Top             =   900
            Width           =   4080
            _ExtentX        =   7197
            _ExtentY        =   476
            Caption         =   ""
            Enabled         =   0   'False
         End
         Begin Project1.CheckBox CheckBox8 
            Height          =   285
            Left            =   30
            TabIndex        =   16
            Top             =   615
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   503
            Caption         =   ""
         End
         Begin Project1.CheckBox CheckBox9 
            Height          =   240
            Left            =   30
            TabIndex        =   17
            Top             =   345
            Width           =   3945
            _ExtentX        =   6959
            _ExtentY        =   423
            Caption         =   ""
            Enabled         =   0   'False
         End
         Begin Project1.CheckBox CheckBox10 
            Height          =   300
            Left            =   30
            TabIndex        =   18
            Top             =   60
            Width           =   3705
            _ExtentX        =   5689
            _ExtentY        =   529
            Caption         =   "Ready"
            Enabled         =   0   'False
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Decompress And Extract All Files"
         Height          =   390
         Left            =   60
         TabIndex        =   5
         Top             =   2130
         Width           =   8655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Make Archive"
      Height          =   3135
      Left            =   60
      TabIndex        =   0
      Top             =   630
      Width           =   8820
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   1890
         Left            =   135
         ScaleHeight     =   1830
         ScaleWidth      =   8475
         TabIndex        =   6
         Top             =   615
         Width           =   8535
         Begin Project1.CheckBox CheckBox5 
            Height          =   300
            Left            =   45
            TabIndex        =   12
            Top             =   1170
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   529
            Caption         =   ""
            Enabled         =   0   'False
         End
         Begin Project1.CheckBox CheckBox4 
            Height          =   270
            Left            =   45
            TabIndex        =   11
            Top             =   885
            Width           =   4080
            _ExtentX        =   7197
            _ExtentY        =   476
            Caption         =   ""
            Enabled         =   0   'False
         End
         Begin Project1.CheckBox CheckBox3 
            Height          =   285
            Left            =   45
            TabIndex        =   10
            Top             =   600
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   503
            Caption         =   ""
         End
         Begin Project1.CheckBox CheckBox2 
            Height          =   240
            Left            =   45
            TabIndex        =   9
            Top             =   330
            Width           =   2865
            _ExtentX        =   5054
            _ExtentY        =   423
            Caption         =   ""
            Enabled         =   0   'False
         End
         Begin VB.TextBox Text4 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   45
            TabIndex        =   8
            Top             =   1470
            Width           =   8340
         End
         Begin Project1.CheckBox CheckBox1 
            Height          =   300
            Left            =   45
            TabIndex        =   7
            Top             =   45
            Width           =   3705
            _ExtentX        =   5689
            _ExtentY        =   529
            Caption         =   "Ready"
            Enabled         =   0   'False
         End
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   135
         TabIndex        =   3
         Text            =   "<Select Directory To Create Archive From>"
         Top             =   285
         Width           =   3300
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Build And Compress"
         Height          =   405
         Left            =   120
         TabIndex        =   2
         Top             =   2565
         Width           =   8625
      End
      Begin VB.CommandButton cmdOrdner 
         Height          =   330
         Left            =   3465
         Picture         =   "Form1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   270
         Width           =   330
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7500
      Top             =   75
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Project1.UserControl1 UserControl11 
      Left            =   8205
      Top             =   30
      _ExtentX        =   1085
      _ExtentY        =   1032
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOrdner_Click()
  Dim sOrdner As String
  
  sOrdner = BrowseForFolder("Select a directory:")
  If sOrdner <> "" Then
    Text1.Text = sOrdner
  End If

End Sub

Private Sub Command1_Click()
  On Local Error Resume Next
  With CommonDialog1
    .CancelError = True
    .Filter = "Archive (*.dat)|*.dat"
    .DefaultExt = ".dat"
    .ShowSave
    If Err = 0 Then
        UserControl11.SaveFilesToArchiv Text1.Text, .Filename, True
    End If
  End With
End Sub


Private Sub Command2_Click()
  Dim sPath As String
  On Local Error Resume Next
  With CommonDialog1
    .CancelError = True
    .Filter = "Archiv-Datei (*.dat)|*.dat"
    .ShowOpen
    If Err = 0 Then
      sPath = BrowseForFolder("Please select the directory to extract to:")
      If sPath <> "" Then
        UserControl11.ExtractFilesFromArchiv .Filename, sPath, True
      End If
    End If
  End With
End Sub

Private Sub Command3_Click()
  With CommonDialog1
    .CancelError = True
    .Filter = "Archiv-Datei (*.dat)|*.dat"
    .ShowOpen
    If Err = 0 Then
        UserControl11.ReadFiles .Filename, True
    End If
  End With

End Sub

Private Sub Command4_Click()
  Dim sPath As String
  On Local Error Resume Next
  With CommonDialog1
    .CancelError = True
    .Filter = "Archiv-Datei (*.dat)|*.dat"
    .ShowOpen
    If Err = 0 Then
      sPath = BrowseForFolder("Please select the directory to extract to:")
      If sPath <> "" Then
        UserControl11.ExtractFile .Filename, List1.List(List1.ListIndex), sPath, True
      End If
    End If
  End With
End Sub

Private Sub UserControl11_AddingFile(Filename As String, Filepath As String, FileSize As Long)
Text4.Text = Filename & " (" & FileSize & ")"
End Sub

Private Sub UserControl11_BuildArchiveProgress(Percent As Integer)
CheckBox3.Caption = "Building (" & Percent & "%)"
End Sub

Private Sub UserControl11_CompressArchiveProgress(Percent As Integer)
CheckBox4.Caption = "Compressing (" & Percent & "%)"
End Sub

Private Sub UserControl11_CompressionDone(NormalArchiveSize As Long, CompressedArchiveSize As Long)
Text4.Text = NormalArchiveSize & " has been compressed to " & CompressedArchiveSize & ". Rate: " & 100 - Int((CompressedArchiveSize / NormalArchiveSize) * 100) & "%"
End Sub

Private Sub UserControl11_DeCompressArchiveProgress(Percent As Integer)
CheckBox9.Caption = "Decompressing Archive (" & Percent & "%)"
End Sub

Private Sub UserControl11_ExtractArchiveProgress(Percent As Integer)
CheckBox7.Caption = "Extracting (" & Percent & "%)"
End Sub

Private Sub UserControl11_ExtractSingleDecompressProcess(Percent As Integer)
Text3.Text = Percent & "%"
End Sub

Private Sub UserControl11_ExtractingFileFromArchive(Filename As String, Filepath As String, FileSize As Long)

Text6.Text = Filepath & Filename & " (" & FileSize & ")"
End Sub

Private Sub UserControl11_ReturnContents(Filename As String, FileSize As Long)
List1.AddItem Filename
End Sub

Private Sub UserControl11_ReturnContentsDecryptingProcess(Percent As Integer)
Text3.Text = Percent & "%"
End Sub

Private Sub UserControl11_StatusChange(State As StateStr)

If State = i_Done Then CheckBox5.Caption = "Done!": CheckBox5.Checked = True
If State = i_Ready Then Text2.Text = "Ready"
'---------------------BUILD ARCHIVE-------------------------
If State = i_Building Then CheckBox3.Caption = "Building (0%)"
If State = i_Compressing_Archive Then CheckBox4.Caption = "Compressing (0%)"
If State = i_Done_Building Then CheckBox3.Checked = True: CheckBox3.Caption = "Building (100%)"
If State = i_Done_Compressing Then CheckBox4.Checked = True: CheckBox4.Caption = "Compressing (100%)"
If State = i_Done_Enumerating_Files Then CheckBox2.Checked = True: CheckBox2.Caption = CheckBox2.Caption & "... Done"
If State = i_Enumerating_Files Then CheckBox1.Checked = True: CheckBox1.Caption = "Starting To Build...Done": CheckBox2.Caption = "Enumeration Files"
If State = i_Starting_To_Build Then CheckBox1.Caption = "Starting To Build"
'---------------------EXTRACT ARCHIVE-------------------------
If State = i_Done_Extracting Then CheckBox7.Checked = True: CheckBox7.Caption = "Extracting (100%)": CheckBox6.Caption = "Done!": CheckBox6.Checked = True
If State = i_Extracting Then CheckBox8.Checked = True: CheckBox7.Caption = "Extracting (0%)"
If State = i_Reading_Archive Then CheckBox8.Caption = "Reading Archive"
If State = i_Starting_To_Extract Then CheckBox10.Caption = "Starting To Extract"
If State = i_Decompressing_Archive Then CheckBox10.Checked = True: CheckBox10.Caption = "Starting To Extract...Done": CheckBox9.Caption = "Decompressing Archive (0%)"
If State = i_Done_Decompressing Then CheckBox9.Checked = True: CheckBox9.Caption = "Decompressing Archive (100%)"
'---------------------RETURN CONTENTS-------------------------
If State = i_ReturnContents_Starting Then Text2.Text = "Starting"
If State = i_ReturnContents_Decrypting Then Text2.Text = "Decompressing"
If State = i_ReturnContents_Reading Then Text2.Text = "Reading"
If State = i_ReturnContents_Done Then Text2.Text = "Done"
'---------------------EXTRACT SINGLE-------------------------
If State = i_ExtractSingle_Starting Then Text2.Text = "Starting"
If State = i_ExtractSingle_Decompressing Then Text2.Text = "Decompressing"
If State = i_ExtractSingle_Reading Then Text2.Text = "Reading"
If State = i_ExtractSingle_Done Then Text2.Text = "Done"
End Sub
