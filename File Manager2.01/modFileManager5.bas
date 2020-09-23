Attribute VB_Name = "modFileManager5"

'----------
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const OFS_MAXPATHNAME = 260
Public Const LocalFileHeaderSig = &H4034B50
Public Const CentralFileHeaderSig = &H2014B50
Public Const EndCentralDirSig = &H6054B50
'-------------
Public SourcePath   As String
Public Buffer       As String * OFS_MAXPATHNAME
Public f_Type       As String * 80
'--------------
Type WIN32_FIND_DATA
   dwFileAttributes  As Long
   ftCreationTime    As Currency   'As FILETIME
   ftLastAccessTime  As Currency   'As FILETIME
   ftLastWriteTime   As Currency   'As FILETIME
  ' nFileSizeHigh     As Long
  ' nFileSizeLow      As Long
   nBigFileSize      As String * 8
   dwReserved0       As Long
   dwReserved1       As Long
   cFileName         As String * 260
   cAlternate        As String * 14
End Type
Public Type ZipFile
  Version            As Integer
  Flag               As Integer
  CompressionMethod  As Integer
  Time               As Integer
  Date               As Integer
  CRC32              As Long
  CompressedSize     As Long
  UncompressedSize   As Long
  FileNameLength     As Integer
  ExtraFieldLength   As Integer
  FileName           As String
  ExtraField         As String
End Type
Public Declare Function DosDateTimeToFileTime Lib "kernel32" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFileTime As Any) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function GetLogicalDrives Lib "kernel32" () As Long

Public Sub ErrMsgBox(Msg As String)
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , Msg
End Sub
Public Function StripNull(ByVal strIn As String) As String
   On Error GoTo PROC_ERR
   Dim nul As Long
   '
   ' Truncate input string at first null.
   ' If no nulls, perform ordinary Trim.
   '
   nul = InStr(strIn, vbNullChar)
   Select Case nul
      Case Is > 1
         StripNull = Left$(strIn, nul - 1)
      Case 1
         StripNull = ""
      Case 0
         StripNull = Trim$(strIn)
   End Select

PROC_EXIT:
  Exit Function
PROC_ERR:
  ErrMsgBox "mDeclare.StripNull"
  Resume Next

End Function
