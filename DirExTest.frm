VERSION 5.00
Begin VB.Form DirExTEST 
   Caption         =   "DirExSearch Test"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5565
   Icon            =   "DirExTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   5565
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Search Attributes:"
      Height          =   1245
      Left            =   150
      TabIndex        =   7
      Top             =   2220
      Width           =   2955
      Begin VB.CheckBox chkAttributes 
         Caption         =   "Archive"
         Height          =   225
         Index           =   4
         Left            =   1650
         TabIndex        =   12
         Top             =   570
         Width           =   1245
      End
      Begin VB.CheckBox chkAttributes 
         Caption         =   "Read-Only"
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin VB.CheckBox chkAttributes 
         Caption         =   "Not System"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   10
         Tag             =   "2"
         Top             =   570
         Value           =   2  'Grayed
         Width           =   1455
      End
      Begin VB.CheckBox chkAttributes 
         Caption         =   "Not Hidden"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   9
         Tag             =   "2"
         Top             =   300
         Value           =   2  'Grayed
         Width           =   1455
      End
      Begin VB.CheckBox chkAttributes 
         Caption         =   "Directory"
         Height          =   225
         Index           =   3
         Left            =   1650
         TabIndex        =   8
         Top             =   300
         Width           =   1275
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   5505
      TabIndex        =   6
      Top             =   3600
      Width           =   5565
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Case Sensitive"
      Height          =   255
      Left            =   3510
      TabIndex        =   5
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Search Sub-folders"
      Height          =   225
      Left            =   3510
      TabIndex        =   4
      Top             =   2670
      Value           =   1  'Checked
      Width           =   1785
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   150
      TabIndex        =   3
      Top             =   1800
      Width           =   2955
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5565
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Search Now"
      Height          =   405
      Left            =   3480
      TabIndex        =   0
      Top             =   1800
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Search for files or folders"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   1560
      Width           =   1875
   End
End
Attribute VB_Name = "DirExTEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lbLongestItem As String
Dim FilterOption As Integer

Private Const LB_SETHORIZONTALEXTENT = &H194

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                        (ByVal hwnd As Long, ByVal wMsg As Long, _
                         ByVal wParam As Long, lParam As Any) As Long

Private WithEvents DirEx As cDirEx
Attribute DirEx.VB_VarHelpID = -1


Private Sub chkAttributes_Click(Index As Integer)
Static Internal As Boolean

If Not Internal Then
    Internal = True
    If Val(chkAttributes(Index).Tag) = vbChecked Then
        chkAttributes(Index).Value = vbGrayed
        chkAttributes(Index).Caption = "Not " & chkAttributes(Index).Caption
    ElseIf Val(chkAttributes(Index).Tag) = vbGrayed Then
            chkAttributes(Index).Caption = Mid$(chkAttributes(Index).Caption, 5)
    End If
End If
chkAttributes(Index).Tag = CStr(chkAttributes(Index).Value)
Internal = False

End Sub

Private Sub Command1_Click()
Dim Attributes As VbFileAttribute
Dim AttributesNot As VbFileAttribute

If Command1.Caption = "&Search Now" Then

    If Len(Text1.Text) = 0 Then Exit Sub
    
    List1.Clear
    lbLongestItem = ""
    Picture1.Cls
    Picture1.Print "Searching..."
    Command1.Caption = "&Stop Search"
    
    ' Reset scrollbar in listbox
    SendMessage List1.hwnd, LB_SETHORIZONTALEXTENT, 0, 0
    
    With DirEx
    
        ' include search to sub-folders
        .RecursiveSearch = (Check1.Value = vbChecked)
        
        .CaseSensitive = (Check2.Value = vbChecked)

        Attributes = (-((chkAttributes(0).Value = vbChecked) * vbHidden)) _
        Or (-((chkAttributes(1).Value = vbChecked) * vbSystem)) _
        Or (-((chkAttributes(2).Value = vbChecked) * vbReadOnly)) _
        Or (-((chkAttributes(3).Value = vbChecked) * vbDirectory)) _
        Or (-((chkAttributes(4).Value = vbChecked) * vbArchive))
            
        AttributesNot = (-((chkAttributes(0).Value = vbGrayed) * vbHidden)) _
        Or (-((chkAttributes(1).Value = vbGrayed) * vbSystem)) _
        Or (-((chkAttributes(2).Value = vbGrayed) * vbReadOnly)) _
        Or (-((chkAttributes(3).Value = vbGrayed) * vbDirectory)) _
        Or (-((chkAttributes(4).Value = vbGrayed) * vbArchive))
    
        
        Call .SearchFile(Text1.Text, Attributes, AttributesNot)
        
    End With
    
Else
    DirEx.StopSearch
End If

End Sub


Private Sub DirEx_Finished(ByVal FileCount As Integer, Result As DirExResultFlags)
    
    Picture1.Cls
    Command1.Caption = "&Search Now"
    
    Select Case Result
    
        Case DirExFailed: Picture1.Print "Search did not finish due to error."
        Case PathNameNotValid: Picture1.Print "Path name does not exist."
        Case Else: Picture1.Print CStr(FileCount) & " Items found"
    
    End Select

End Sub




' All files or folders matching attributes and filespec are each returned
' immediately after they are found.
Private Sub DirEx_CurrentFile(ByVal FolderPath As String, ByVal FileName As String, FileAttribute As VbFileAttribute)
  
  Dim PixelLength As Long
    
    If Command1.Caption = "&Stop Search" Then
    
        List1.AddItem FolderPath & FileName, 0
        
        If TextWidth(FolderPath & FileName) > TextWidth(lbLongestItem) Then
            lbLongestItem = FolderPath & FileName
            PixelLength = (TextWidth(lbLongestItem) / Screen.TwipsPerPixelX) + 5
            SendMessage List1.hwnd, LB_SETHORIZONTALEXTENT, PixelLength, 0
        End If
        
    End If

End Sub


' Current folder path of files being searched is returned
Private Sub DirEx_CurrentFolder(ByVal FolderPath As String)
    Picture1.Cls
    Picture1.Print "Searching... " & TruncatePath(FolderPath, 50)
End Sub

Private Sub Form_Initialize()
    Set DirEx = New cDirEx
End Sub

Private Sub Form_Load()

    Const LB_INITSTORAGE = &H1A8
    SendMessage List1.hwnd, LB_INITSTORAGE, 30000&, ByVal 30000& * 200
    List1.Move 0, 0, Me.ScaleWidth
    
    Text1.Text = Left$(CurDir, 2) & "\*.txt"
    
End Sub



' Truncate long path names
Private Function TruncatePath(ByVal StrPath As String, Optional ByVal iMax As Integer = 35) As String
Dim i As Integer, LastPart As String
Dim iFinal As Integer
If iMax < 35 Then iMax = 35

    'If lenth is okay then exit function
    If Len(StrPath) <= iMax Then
        TruncatePath = StrPath
        Exit Function
    End If

    'Allow for drive letter,back slash and three periods
    iMax = iMax - 6

    'Find new string length
    For i = Len(StrPath) - iMax To Len(StrPath)
        If Mid(StrPath, i, 1) = "\" Then Exit For
    Next
    LastPart = Right(StrPath, Len(StrPath) - (i - 1))
    
    If Trim$(LastPart) = vbNullString Then
        iFinal = 1: i = 1
        Do Until i = 0
            i = InStr(iFinal + 1, StrPath, "\")
            If (i <> 0) Then iFinal = i
        Loop
        LastPart = Trim$(Mid$(StrPath, iFinal))
    End If
    
    'Send back new string
    TruncatePath = Left(StrPath, 3) + "..." & LastPart

End Function
Private Sub Form_Resize()
    List1.Move 0, 0, Me.ScaleWidth
End Sub


Private Sub Form_Unload(Cancel As Integer)
    DirEx.StopSearch
    Set DirEx = Nothing
End Sub


