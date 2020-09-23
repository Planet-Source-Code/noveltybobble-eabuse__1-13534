VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'////////////////////////////////////////////////////////////////////
'//
'// API Declarations
'//
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal length As Long)

'////////////////////////////////////////////////////////////////////
'//
'// Private Type and Const Declarations
'//
Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type


Private Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
   dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
   dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
   dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
   dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
   dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
   dwProductVersionMSl As Integer '  e.g. = &h0003 = 3
   dwProductVersionMSh As Integer '  e.g. = &h0010 = .1
   dwProductVersionLSl As Integer '  e.g. = &h0000 = 0
   dwProductVersionLSh As Integer '  e.g. = &h0031 = .31
   dwFileFlagsMask As Long        '  = &h3F for version "0.42"
   dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
   dwFileType As Long             '  e.g. VFT_DRIVER
   dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long           '  e.g. 0
   dwFileDateLS As Long           '  e.g. 0
End Type

' ----- VS_VERSION.dwFileFlags -----
Private Const VS_FFI_SIGNATURE = &HFEEF04BD
Private Const VS_FFI_STRUCVERSION = &H10000
Private Const VS_FFI_FILEFLAGSMASK = &H3F&

' ----- VS_VERSION.dwFileFlags -----
Private Const VS_FF_DEBUG = &H1
Private Const VS_FF_PRERELEASE = &H2
Private Const VS_FF_PATCHED = &H4
Private Const VS_FF_PRIVATEBUILD = &H8
Private Const VS_FF_INFOINFERRED = &H10
Private Const VS_FF_SPECIALBUILD = &H20

' ----- VS_VERSION.dwFileOS -----
Private Const VOS_UNKNOWN = &H0
Private Const VOS_DOS = &H10000
Private Const VOS_OS216 = &H20000
Private Const VOS_OS232 = &H30000
Private Const VOS_NT = &H40000

Private Const VOS__BASE = &H0
Private Const VOS__WINDOWS16 = &H1
Private Const VOS__PM16 = &H2
Private Const VOS__PM32 = &H3
Private Const VOS__WINDOWS32 = &H4

Private Const VOS_DOS_WINDOWS16 = &H10001
Private Const VOS_DOS_WINDOWS32 = &H10004
Private Const VOS_OS216_PM16 = &H20002
Private Const VOS_OS232_PM32 = &H30003
Private Const VOS_NT_WINDOWS32 = &H40004

' ----- VS_VERSION.dwFileType -----
Private Const VFT_UNKNOWN = &H0
Private Const VFT_APP = &H1
Private Const VFT_DLL = &H2
Private Const VFT_DRV = &H3
Private Const VFT_FONT = &H4
Private Const VFT_VXD = &H5
Private Const VFT_STATIC_LIB = &H7

' ----- VS_VERSION.dwFileSubtype for VFT_WINDOWS_DRV -----
Private Const VFT2_UNKNOWN = &H0
Private Const VFT2_DRV_PRINTER = &H1
Private Const VFT2_DRV_KEYBOARD = &H2
Private Const VFT2_DRV_LANGUAGE = &H3
Private Const VFT2_DRV_DISPLAY = &H4
Private Const VFT2_DRV_MOUSE = &H5
Private Const VFT2_DRV_NETWORK = &H6
Private Const VFT2_DRV_SYSTEM = &H7
Private Const VFT2_DRV_INSTALLABLE = &H8
Private Const VFT2_DRV_SOUND = &H9
Private Const VFT2_DRV_COMM = &HA




'////////////////////////////////////////////////////////////////////
'//
'// Private Data Declarations
'//
Private iHandle As Integer
Private strSourceFilename As String
Private Filename As OPENFILENAME



Private Sub Class_Initialize()
    iHandle = FreeFile
End Sub

Public Function IsItLocked(strSource As String) As Boolean
    Dim iMyHandle As Integer
    
    iMyHandle = FreeFile
    
    On Error Resume Next
    Open strSource For Binary Lock Write As iMyHandle
    IsItLocked = IIf(Err.Number = 0, False, True)
    Close iMyHandle
End Function

Public Function SourceFile(strSource As String) As Boolean
    strSourceFilename = strSource
    
    Close iHandle
    
    If Exists(strSource) = True Then
        On Error Resume Next
        Open strSource For Binary Access Read As iHandle
        SourceFile = IIf(Err.Number <> 0, False, True)
    End If
End Function


Public Function Move(strDest As String) As Boolean
    Dim iDestHandle As Integer
    Dim lLoop  As Long
    Dim lLen As Long
    Dim Buffer As String
    Dim lValue As Long
   
    ' //
    ' //    Build up the destination filepath and destroy any existing copies.
    ' //
    strDest = strDest & Mid(strSourceFilename, InStrRev(strSourceFilename, "\"))
    If Exists(strDest) Then Kill strDest
     
    '//
    '// Determine the length of the file.  Don't forget to check that the
    '// file exists.
    '//
    If Exists(strSourceFilename) = False Then
        Move = False
        Exit Function
    End If
    lLen = FileLen(strSourceFilename)
    
    iDestHandle = FreeFile
    Buffer = " "
   
    '//
    '//  Now make an attempt at opening the destination file
    '//
    On Error Resume Next
    Open strDest For Binary Access Write As iDestHandle Len = Len(Buffer)
    lValue = Err.Number
    On Error GoTo 0
    
    If lValue <> 0 Then
        Move = False
        Exit Function
    End If
    
    On Error GoTo MoveFailed
    
    For lLoop = 1 To lLen
        Get iHandle, lLoop, Buffer
        Put iDestHandle, , Buffer
        DoEvents
    Next
    
    Close iHandle
    Close iDestHandle
    
    Kill strSourceFilename
    SourceFile (strDest)
    
    Move = True
    Exit Function

MoveFailed:
    Move = False
End Function

Private Sub Class_Terminate()
    Close iHandle
End Sub

Public Function Exists(Optional strFilename As String = "-1") As Boolean
    Dim lValue As Long

    If strFilename = "-1" Then strFilename = strSourceFilename

    On Error Resume Next
    lValue = FileLen(strFilename)
    lValue = Err.Number
    On Error GoTo 0

    Exists = IIf(lValue = 0, True, False)
End Function

Public Function GetString(strString As String) As Boolean
    On Error Resume Next
    Input #iHandle, strString
    GetString = IIf(Err.Number = 0, True, False)
End Function


Public Function Copy(strDest As String) As Boolean
    CopyFile strSourceFilename, strDest, 0
End Function

' Default extension for files
Public Property Let DefaultExtension(Extention As String)
    Filename.lpstrDefExt = Extention
End Property

' Default extension for files
Public Property Get DefaultExtension() As String
    DefaultExtension = Filename.lpstrDefExt
End Property

' Object that 'owns' the common dialog.
Public Sub SetObjectOwner(ByRef ParentForm As Variant)
    Filename.hwndOwner = ParentForm.hWnd
End Sub

' File filter. Format : "Text for filter|filter|Next Text|Next filter". Ex : "All Files (*.*)|*.*|Text files (*.txt)|*.txt"
Public Property Let Filter(CustomFilter As String)
    Dim intCount As Integer
    Filename.lpstrFilter = ""
    For intCount = 1 To Len(CustomFilter)
        If Mid(CustomFilter, intCount, 1) = "|" Then Filename.lpstrFilter = Filename.lpstrFilter + Chr(0) Else Filename.lpstrFilter = Filename.lpstrFilter + Mid(CustomFilter, intCount, 1)
    Next intCount
    Filename.lpstrFilter = Filename.lpstrFilter + Chr(0)
End Property

' Window title
Public Property Let WindowTitle(Title As String)
    Filename.lpstrTitle = Title
End Property

' Window title
Public Property Get WindowTitle() As String
    WindowTitle = Filename.lpstrTitle
End Property

' Initial folder
Public Property Let InitialDirectory(InitDir As String)
    Filename.lpstrInitialDir = InitDir
End Property

' Initial folder
Public Property Get InitialDirectory() As String
    InitialDirectory = Filename.lpstrInitialDir
End Property

' Show common dialog open file control and returns file name. Use GetMultiFilename function to
' extract filenames when the multiselect flag is true. The Count function returns the number
' of files selected.
Public Function GetFileOpenName(Optional Multiselect As Boolean = False) As String
    Filename.hInstance = App.hInstance
    'Filename.hwndOwner = hWnd
    Filename.lpstrFile = Chr(0) & Space(259)
    Filename.nMaxFile = 260
    If Multiselect Then Filename.flags = &H80000 Or &H4 Or &H200 Else Filename.flags = &H80000 Or &H4
    Filename.lStructSize = Len(Filename)
    GetOpenFileName Filename
    GetFileOpenName = Filename.lpstrFile
End Function

' Show common dialog save file control and returns file name.
Public Function GetFileSaveName() As String
    Filename.hInstance = App.hInstance
    'Filename.hwndOwner = hWnd
    Filename.lpstrFile = Chr(0) & Space(259)
    Filename.nMaxFile = 260
    Filename.flags = &H80000 Or &H4
    Filename.lStructSize = Len(Filename)
    GetSaveFileName Filename
    GetFileSaveName = Filename.lpstrFile
End Function

' Returns the number of file selected. Use with GetMultiFilename to extract filename when multiselect is true.
Public Function Count() As Integer
    Dim intCount As Integer
    For intCount = 1 To Trim(Len(Filename.lpstrFile))
        If Mid(Trim(Filename.lpstrFile), intCount, 1) = Chr(0) Then Count = Count + 1
    Next intCount
    Count = Count - 2
    If Count < 1 Then Count = Count + 1
End Function

' Returns the filename of the specified filenumber. Use only with open file dialog when multiselect is true.
Public Function GetMultiFilename(Filenumber As Integer) As String
    Dim intCount As Integer
    Dim intOne As Integer
    Dim intFile As Integer
    Dim intNext As Integer
    intOne = InStr(1, Trim(Filename.lpstrFile), Chr(0))
    intFile = 1
    For intCount = 1 To Filenumber
        intFile = InStr(intFile + 1, Trim(Filename.lpstrFile), Chr(0))
    Next intCount
    intNext = InStr(intFile + 1, Trim(Filename.lpstrFile), Chr(0))
    GetMultiFilename = IIf(Right(Mid(Trim(Filename.lpstrFile), 1, intOne - 1), 1) = "\", Mid(Trim(Filename.lpstrFile), 1, intOne - 1), Mid(Trim(Filename.lpstrFile), 1, intOne - 1) + "\") + Mid(Trim(Filename.lpstrFile), intFile + 1, intNext - intFile - 1)
    If Right(GetMultiFilename, 1) = "\" Then GetMultiFilename = Left(GetMultiFilename, Len(GetMultiFilename) - 1)
End Function


Public Function GetFileVersion(ByVal FullPathToFile As String) As String
    Dim rc As Long
    Dim lDummy  As Long
    Dim sBuffer()  As Byte
    Dim lBufferLen  As Long
    Dim lVerPointer  As Long
    Dim udtVerBuffer  As VS_FIXEDFILEINFO
    Dim lVerbufferLen  As Long
    Dim FullFileName As String
    Dim FileVer As String
    Dim StrucVer As String
   
   FullFileName = FullPathToFile

   '*** Get size ****
   lBufferLen = GetFileVersionInfoSize(FullFileName, lDummy)
   If lBufferLen < 1 Then
      GetFileVersion = "No Version Info Available"
      Exit Function
   End If

   '**** Store info to udtVerBuffer struct ****
   ReDim sBuffer(lBufferLen)
   rc = GetFileVersionInfo(FullFileName, 0&, lBufferLen, sBuffer(0))
   rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
   MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)

   '**** Determine Structure Version number - NOT USED ****
   StrucVer = Format$(udtVerBuffer.dwStrucVersionh) & "." & _
      Format$(udtVerBuffer.dwStrucVersionl)

   '**** Determine File Version number ****
    FileVer = Format$(udtVerBuffer.dwFileVersionMSh) & "." & _
      Format$(udtVerBuffer.dwFileVersionMSl) & "." & _
      Format$(udtVerBuffer.dwFileVersionLSh) & "." & _
      Format$(udtVerBuffer.dwFileVersionLSl)

      GetFileVersion = FileVer
End Function





