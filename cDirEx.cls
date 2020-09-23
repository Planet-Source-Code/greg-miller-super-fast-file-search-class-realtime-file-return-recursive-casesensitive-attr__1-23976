VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDirEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Result Flags to sent to the Result event at end of file search
Public Enum DirExResultFlags
    DirExSuccess = 0
    UserCanceled = 1120
    PathNameNotValid = 1121
    FileSpecNotFound = 1122
    DirExFailed = 1123
End Enum

' Event raised for every folder searched
Public Event CurrentFolder(ByVal FolderPath As String)

' Event raised with each file found
Public Event CurrentFile(ByVal FolderPath As String, ByVal FileName As String, FileAttribute As VbFileAttribute)

' Event raised when all searching is finished
Public Event Finished(ByVal FileCount As Integer, Result As DirExResultFlags)

' Attributes
Private Enum FileAttributes
   FILE_ATTRIBUTE_READONLY = &H1
   FILE_ATTRIBUTE_HIDDEN = &H2
   FILE_ATTRIBUTE_SYSTEM = &H4
   FILE_ATTRIBUTE_DIRECTORY = &H10
   FILE_ATTRIBUTE_ARCHIVE = &H20
   FILE_ATTRIBUTE_DEVICE = &H40
   FILE_ATTRIBUTE_NORMAL = &H80
   FILE_ATTRIBUTE_TEMPORARY = &H100
   FILE_ATTRIBUTE_SPARSE_FILE = &H200
   FILE_ATTRIBUTE_REPARSE_POINT = &H400
   FILE_ATTRIBUTE_COMPRESSED = &H800
   FILE_ATTRIBUTE_OFFLINE = &H1000
   FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = &H2000
   FILE_ATTRIBUTE_ENCRYPTED = &H4000
End Enum

' variables:
Private hFile As Long
Private hItem As Long
Private m_Files As Integer
Private m_FileSpec As String
Private m_Recurse As Boolean
Private m_CaseSensitive As Boolean
Private m_Attributes As FileAttributes
Private m_NotAttributes As FileAttributes
Private m_ResultFlag As DirExResultFlags
Private Const INVALID_HANDLE_VALUE = -1

Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As Currency
        ftLastAccessTime As Currency
        ftLastWriteTime As Currency
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * 260
        cShortFileName As String * 14
End Type
Private WFD As WIN32_FIND_DATA

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" ( _
    ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" ( _
    ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" ( _
    ByVal lpFileName As String) As Long

Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As Long) As Long

' Test File or attribute existence with error trapping
Private Function ExistFile(ByVal File As String, Optional FileAttribute As VbFileAttribute) As Boolean

Dim sFile As String, ExistFileAttribute As FileAttributes
Dim nAttr As FileAttributes

    On Error GoTo ExistError
    If Len(File) = 0 Then GoTo ExistContinue

    sFile = Qualify(File)

    nAttr = GetFileAttributes(sFile)

    Select Case nAttr
        Case -1
            ExistFileAttribute = 0
            ExistFile = False
        Case Else:
            If ((nAttr And FileAttribute) = FileAttribute) Then
                FileAttribute = nAttr And (Not FILE_ATTRIBUTE_NORMAL)
                ExistFile = True
            End If
    End Select

ExistContinue: On Error GoTo 0

Exit Function
ExistError:
ExistFile = False
Resume ExistContinue

Exit Function
End Function

' Get CaseSensitive:
Public Property Get CaseSensitive() As Boolean
    CaseSensitive = m_CaseSensitive
End Property

' RecursiveSearch: Property determines if search path will include sub-folders
Public Property Get RecursiveSearch() As Boolean
    RecursiveSearch = m_Recurse
End Property

' Search CaseSensitive:
Public Property Let CaseSensitive(bVal As Boolean)
    m_CaseSensitive = bVal
End Property

' RecursiveSearch: Property determines if search path will include sub-folders
Public Property Let RecursiveSearch(bVal As Boolean)
    m_Recurse = bVal
End Property

' StopSearch: Do not search for any more files
Public Sub StopSearch()
    m_ResultFlag = UserCanceled
End Sub

' ResultFlag: Value <> 0 Indicates search was canceled or error was raised
Private Property Get ResultFlag() As DirExResultFlags

    DoEvents

    ResultFlag = m_ResultFlag

End Property


' Enum each file matching file spec and attributes
Private Sub FindFiles(sPath As String)
Dim sFile As String, ReturnFile As Boolean
Dim FileFound As Boolean

    On Local Error GoTo EnumErr

    ' Get first file AttributeFound in current directory
    hFile = FindFirstFile(sPath & m_FileSpec, WFD)

    If hFile <> INVALID_HANDLE_VALUE Then

        Do ' Check each file found in current directory
        
            ' Stop search if error detected or search was canceled
            If ResultFlag <> DirExSuccess Then Exit Sub
            
            ReturnFile = False
            
            ' File name with no path info
            sFile = Left(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)

            ' Verify a valid file name
            If sFile <> "." And sFile <> ".." Then

                If m_CaseSensitive And Not (sFile Like m_FileSpec) Then
                    ' Search is case sensitive and doesn't AttributeFound
                    GoTo FindNext
                Else
                    ' Determine if file found has specified search attributes
                    FileFound = (m_Attributes = 0) Or ((WFD.dwFileAttributes And m_Attributes) = m_Attributes)
                    
                    ' Ignore file if specified search attributes don't match or
                    ' any specified NotAttributes are found.
                    ReturnFile = FileFound And (m_NotAttributes = 0 _
                                                            Or Not CBool(WFD.dwFileAttributes And m_NotAttributes))
                End If
                    
                If ReturnFile Then
                    ' Raise event with each file found
                    m_Files = m_Files + 1
                    RaiseEvent CurrentFile(sPath, sFile, WFD.dwFileAttributes And (Not FILE_ATTRIBUTE_NORMAL))
                End If

            End If
            
FindNext:
        Loop While FindNextFile(hFile, WFD)

        Call FindClose(hFile)

    End If

Continue:
    On Local Error GoTo 0
    Exit Sub

EnumErr: m_ResultFlag = DirExFailed
Resume Continue


End Sub


' Find next folder containing files to search
Private Sub FindFolders(sPath As String)

    Dim FilePath As String, dirs As Integer, dirbuf() As String
    Dim dirBuffer As String, i As Integer

    On Local Error GoTo EnumErr

    ' Stop search if error detected or search was canceled
    If ResultFlag <> DirExSuccess Then Exit Sub
    
    FilePath = sPath
    If Right(FilePath, 1) <> "\" Then FilePath = FilePath & "\"

    ' Get handle of first directory and search for any subfolders
    hItem = FindFirstFile(FilePath & "*.*", WFD)

    If hItem <> INVALID_HANDLE_VALUE Then

        Do

            ' Stop search if error detected or search was canceled
            If ResultFlag <> DirExSuccess Then Exit Sub
            
            ' Make sure handle belongs to a directory.
            If m_Recurse And ((WFD.dwFileAttributes And vbDirectory) = vbDirectory) Then

                ' directory path of subfolder
                dirBuffer = Left(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)

                If dirBuffer <> "." And dirBuffer <> ".." Then
                    If (dirs Mod 10) = 0 Then ReDim Preserve dirbuf(dirs + 10)
                    dirs = dirs + 1
                    dirbuf(dirs) = dirBuffer
                  End If

            End If

        ' Get next folder
        Loop While FindNextFile(hItem, WFD)

        Call FindClose(hItem)

    End If

      ' Raise event notifying folder change during search
      RaiseEvent CurrentFolder(Qualify(FilePath))
      
      Call FindFiles(FilePath)
      
    For i = 1 To dirs
        ' Enum subfolders
        FindFolders FilePath & dirbuf(i) & "\"
    Next i

Continue:
    On Local Error GoTo 0
    Exit Sub

EnumErr: m_ResultFlag = DirExFailed
Resume Continue

End Sub

' Search all files matching pathname or attributes
Public Sub SearchFile(ByVal Pathname As String, Optional ByVal Attributes As VbFileAttribute, Optional ByVal NotAttributes As VbFileAttribute)

Dim InitPath As String

On Error GoTo SpecError

    If Len(Pathname) = 0 Then
        m_ResultFlag = PathNameNotValid
        GoTo DirExDone
    End If
    
    ' Reset file count
    m_Files = 0
    
    ' Reset search status
    m_ResultFlag = DirExSuccess
    
    ' specifies a file name or pattern. supports the use of
    ' multiple-character (*) and single-character (?) wildcards.
    ' may also include directory or folder, and drive.
    InitPath = CStr(Pathname)
    
    ' sum specifies file search attributes. If value equals 0, all
    ' files are returned that match pathname.
    m_Attributes = Attributes
    
    ' specifies file attributes to ignore. If value equals 0, all
    ' files are returned that match pathname and search attributes.
    m_NotAttributes = NotAttributes
    
    ' Get file spec
    If FileSpec(InitPath, m_FileSpec) Then
    
        ' Start the search
        Call FindFolders(InitPath)
        
        If m_Files = 0 And (m_ResultFlag = DirExSuccess) Then
            ' no files found
            m_ResultFlag = FileSpecNotFound
        End If
        
    End If
    
    
DirExDone: On Error GoTo 0

' Raise event when finished and send result status for search
RaiseEvent Finished(m_Files, m_ResultFlag)

Exit Sub

SpecError:
m_ResultFlag = DirExFailed
Resume DirExDone

End Sub

' FileSpec: Returns a string representing the file pattern used for searching files
Private Function FileSpec(InitPath As String, Spec As String) As Boolean

Dim sPath As String, SpecPos As Long, sSpec As String

    On Local Error GoTo ErrSpec

    sPath = FullPath(InitPath, SpecPos)

    If Dir$(sPath, vbDirectory) = "." Then
        If Not ExistFile(sPath) Then sPath = Left$(sPath, SpecPos - 1)
        InitPath = Qualify(sPath)
        Spec = "*.*"
    ElseIf ExistFile(sPath) Then
        InitPath = Qualify(Left$(sPath, SpecPos - 1))
        If ExistFile(sPath, vbDirectory) Then
            If SpecPos <= Len(sPath) Then
                ' folder path ending with no backslash
                Spec = Mid$(sPath, SpecPos)
            Else
                ' root - ?:\
                Spec = "*.*"
            End If
        Else
            Spec = Mid$(sPath, SpecPos)
        End If
    Else

        ' file spec may still exist under sub-folders
        If SpecPos <= Len(sPath) Then
            sSpec = Mid$(sPath, SpecPos)
            sPath = Qualify(Left$(sPath, SpecPos - 1))
        Else
            ' root - ?:\
            sPath = Qualify(sPath)
            sSpec = "*.*"
        End If

        ' Verify if spec's folder Path exists
        If ExistFile(sPath) Then
            InitPath = sPath
            Spec = sSpec
        Else
            m_ResultFlag = PathNameNotValid
            FileSpec = False
            GoTo Continue
        End If

    End If

FileSpec = True

Continue: On Error GoTo 0
Exit Function

ErrSpec:
FileSpec = False
m_ResultFlag = DirExFailed
Resume Continue


End Function

' FullPath: Returns fully qualified path of file
Private Function FullPath(sFileName As String, _
                     Optional FilePart As Long, _
                     Optional ExtPart As Long, _
                     Optional DirPart As Long) As String

    Dim c As Long, p As Long, sRet As String
    If sFileName = vbNullString Then Exit Function
    
    ' Get the path size, then create string of that size
    sRet = String(260, 0)
    c = GetFullPathName(sFileName, 260, sRet, p)
    If c <> 0 Then 'success
        sRet = Left$(sRet, c)

        ' Get the directory, file, and extension positions
        PartPos sRet, FilePart, DirPart, ExtPart
        
        FullPath = sRet
    
    End If
    
End Function


' PartPos: Returns file extension, including directory, file, and extension positions
Private Function PartPos(sFull As String, iFilePart As Long, _
                      iDirPart As Long, iExtPart As Long) As Boolean

    Dim iDrv As Long, i As Long, cMax As Long
    
     On Local Error Resume Next
    
    cMax = Len(sFull)

    iDrv = Asc(UCase$(Left$(sFull, 1)))

    ' If in format d:\path\name.ext, return 3
    If iDrv <= 90 Then                          ' Less than Z
        If iDrv >= 65 Then                      ' Greater than A
            If Mid$(sFull, 2, 1) = ":" Then     ' Second character is :
                If Mid$(sFull, 3, 1) = "\" Then ' Third character is \
                    iDirPart = 3
                End If
            End If
        End If
    Else

        ' If in format \\machine\share\path\name.ext, return position of \path
        ' First and second character must be \
        If Mid$(sFull, 1, 2) <> "\\" Then
            PartPos = False
            GoTo FailGetDirExt
        End If

        Dim fFirst As Boolean
        i = 3
        Do
            If Mid$(sFull, i, 1) = "\" Then
                If fFirst Then
                    iDirPart = i
                    Exit Do
                Else
                    fFirst = True
                End If
            End If
            i = i + 1
        Loop Until i = cMax
    End If

    ' Start from end and find extension
    iExtPart = cMax + 1       ' Assume no extension
    fFirst = False
    Dim sChar As String
    For i = cMax To iDirPart Step -1
        sChar = Mid$(sFull, i, 1)
        If Not fFirst Then
            If sChar = "." Then
                iExtPart = i
                fFirst = True
            End If
        End If
        If sChar = "\" Then
            iFilePart = i + 1
            Exit For
        End If
    Next
    
    PartPos = True
    
    On Error GoTo 0
    
    Exit Function
FailGetDirExt:
    iFilePart = 0
    iDirPart = 0
    iExtPart = 0
    Exit Function
End Function

Private Function Qualify(ByVal sPath As String, Optional ByVal WithBackSlash As Boolean) As String
Dim fp As String

    If Len(sPath) = 0 Then Exit Function
    
    ' qualify the root directory path
    If Right(sPath, 1) = ":" Then
        fp = sPath & "\"
    Else
        fp = sPath
    End If
    
    Select Case WithBackSlash
        Case True   ' Make sure path ends with a backslash
            If Right$(fp, 1) <> "\" Then
                Qualify = fp & "\"
            Else
                Qualify = fp
            End If
        Case False  ' Make sure folder path doesn't end with a backslash
            If Len(fp) > 3 And Right$(fp, 1) = "\" Then
                Qualify = Left$(fp, (Len(fp) - 1))
            Else
                Qualify = fp
            End If
    End Select
    
End Function
