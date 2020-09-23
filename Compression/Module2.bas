Attribute VB_Name = "Module2"
Public Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long

Option Explicit
Function InStrR(ByVal sTarget As String, ByVal sFind As String) As Long
Dim P As Long, LastP As Long, Start As Long
  P = InStr(1, sTarget, sFind)
  Do While P
    LastP = P
    P = InStr(LastP + 1, sTarget, sFind)
  Loop
  InStrR = LastP
End Function

'/********************************************/
'/*     Author: Jorge Colaccini              */
'/*             sofware2004@informas.com     */
'/*     Copyright (c) 2004                   */
'/********************************************/
'

Function AllFilesInFolders(ByVal sFolderPath As String, Optional bWithSubFolders As Boolean = False) As String()
'---------------------------------------------------------------------------------------
' Procedimiento : AllFilesInFolders
' FPO           : 02/Feb/2004 16:30
' Autor         : Jorge Colaccini (sofware2004@informas.com)
' Propósito     :
'                 retrieve an array containig all files in a folder,
'                 optionally processing all subfolders too.
'                 Do not use FileSystemObject objects, purely Visual Basic code!!!.
'---------------------------------------------------------------------------------------
'
    Dim sTemp As String
    Dim sDirIn As String
    ReDim sFilelist(0) As String, sSubFolderList(0) As String, sToProcessFolderList(0) As String
    Dim i As Integer, j As Integer
    
    sDirIn = sFolderPath
    If Not (Right$(sDirIn, 1) = "\") Then sDirIn = sDirIn & "\"
    
    On Error Resume Next
    sTemp = Dir$(sDirIn & "*.*")
    
    'Loop on Files that aren't Folders
    Do While sTemp <> ""
      'Add file to the list to return
      AddItem2Array1D sFilelist(), sDirIn & sTemp
      sTemp = Dir
    Loop
    
    'Then, if bWithSubFolders is TRUE
    'Process subfolders after common files
    If bWithSubFolders Then
      
      'Loop on Files that aren't Folders
      sTemp = Dir$(sDirIn & "*.*", vbDirectory)
      Do While sTemp <> ""
         ' Ignore current directory and his father
         If sTemp <> "." And sTemp <> ".." Then
            
            ' check if really is a directory
            If (GetAttr(sDirIn & sTemp) And vbDirectory) = vbDirectory Then
              'Add to temporal array to process later, to avoid problems in recurse DIR
              AddItem2Array1D sToProcessFolderList, sDirIn & sTemp
            End If
         End If
         sTemp = Dir   ' Next entry
      Loop
      
      'Process temporal array containing subfolders of current folder
      If UBound(sToProcessFolderList) > 0 Or UBound(sToProcessFolderList) = 0 And sToProcessFolderList(0) <> "" Then
        For i = 0 To UBound(sToProcessFolderList)
          sSubFolderList = AllFilesInFolders(sToProcessFolderList(i), bWithSubFolders)
          If UBound(sSubFolderList) > 0 Or UBound(sSubFolderList) = 0 And sSubFolderList(0) <> "" Then
            For j = 0 To UBound(sSubFolderList)
              AddItem2Array1D sFilelist(), sSubFolderList(j)
            Next
          End If
        Next
      End If

    End If
        
    AllFilesInFolders = sFilelist

End Function



Public Sub AddItem2Array1D(ByRef VarArray As Variant, ByVal VarValue As Variant)
'---------------------------------------------------------------------------------------
' Procedimiento : AddItem2Array1D
' FPO           : 02/Feb/2004 16:34
' Autor         : Jorge Colaccini (sofware2004@informas.com)
' Propósito     :
'                 Add item to array preserving it content
'
' Uso           :
'                 ReDim FileList(0) As String
'                 FileList = AllFilesInFolders(App.Path, False)
'                 lFLCount = UBound(FileList)
'                 For j = 0 To UBound(FileList)
'                   Debug.Print j, "]"; FileList(j); "["
'                 Next
'---------------------------------------------------------------------------------------
'
    
  Dim i  As Long
  Dim iVarType As Integer
  iVarType = VarType(VarArray) - 8192
  i = UBound(VarArray)
  
  Select Case iVarType
  
    'vbInteger 2 Entero
    'vbLong 3 Entero largo
    'vbSingle 4 Un número de coma flotante de precisión simple
    'vbDouble 5 Un número de coma flotante de precisión doble
    'vbCurrency 6 Valor de moneda
    'vbDecimal 14 Valor decimal
    'vbByte 17 Valor de byte
    Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
      'May fail if First value of VarArray is set to 0
      If VarArray(0) = 0 Then
        i = 0
      Else
        i = i + 1
      End If
    
    'vbDate 7 Valor de fecha
    Case vbDate
      'May fail if First value of VarArray is set to #00:00:00#, >:-(
      If VarArray(0) = "00:00:00" Then
        i = 0
      Else
        i = i + 1
      End If
    
    'vbString 8 Cadena
    Case vbString
      'May fail if First value of VarArray is set to ""
      If VarArray(0) = vbNullString Then
        i = 0
      Else
        i = i + 1
      End If
    
    'vbBoolean 11 Valor booleano
    Case vbBoolean
      'May fail if First value of VarArray is set to False, >:-(
      If VarArray(0) = False Then
        i = 0
      Else
        i = i + 1
      End If
    
    Case Else
      'NO IMPLEMENTED :-p
    
  End Select
  
  ReDim Preserve VarArray(i)
  VarArray(i) = VarValue
  
End Sub




