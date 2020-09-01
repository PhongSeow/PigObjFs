VERSION 5.00
Begin VB.Form frmPigObjFs 
   Caption         =   "PigObjFs"
   ClientHeight    =   6285
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   12750
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   12750
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMain 
      Height          =   2535
      Left            =   4320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   840
      Width           =   4215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFile_Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnupFileSystemObject 
      Caption         =   "pFileSystemObject"
      Begin VB.Menu mnupFileSystemObject_GetFile 
         Caption         =   "GetFile"
      End
      Begin VB.Menu mnupFileSystemObject_GetFolder 
         Caption         =   "GetFolder"
      End
      Begin VB.Menu mnupFileSystemObject_FileExists 
         Caption         =   "FileExists"
      End
      Begin VB.Menu mnupFileSystemObject_FolderExists 
         Caption         =   "FolderExists"
      End
      Begin VB.Menu mnupFileSystemObject_CreateFolder 
         Caption         =   "CreateFolder"
      End
   End
   Begin VB.Menu mnupTextStream 
      Caption         =   "pTextStream"
      Begin VB.Menu mnupTextStream_ReadFile 
         Caption         =   "Read File"
      End
      Begin VB.Menu mnupTextStream_WriteFile 
         Caption         =   "Write File"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelp_OnlineDoc 
         Caption         =   "Online documentation"
      End
      Begin VB.Menu mnuHelp_SPS 
         Caption         =   "Seow Phong Studio"
      End
   End
End
Attribute VB_Name = "frmPigObjFs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub aaa_Click()

End Sub

Private Sub Form_Load()
    Set goFS = New pFileSystemObject
    goFS.InitLog App.Path, App.Title
    Me.Caption = App.ProductName & App.FileDescription
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With Me.txtMain
        .Top = 50
        .Left = 50
        .Width = Me.ScaleWidth - 100
        .Height = Me.ScaleHeight - 100
    End With
    On Error GoTo 0
End Sub

Private Sub mnuFile_Exit_Click()
    Unload Me
End Sub

Private Sub mnuTest_ErrOpt_Click()
    On Error GoTo ErrOcc:
    Dim oTextStream As New pTextStream
    oTextStream.WriteLine "a"
    On Error GoTo 0
    Exit Sub
ErrOcc:
    MsgBox Err.Description
    On Error GoTo 0
End Sub



Private Sub mnuHelp_OnlineDoc_Click()
    Shell "explorer https://en.seowphong.com/oss/PigObjFs/"
End Sub

Private Sub mnuHelp_SPS_Click()
    Shell "explorer https://en.seowphong.com"
End Sub

Private Sub mnupFileSystemObject_CreateFolder_Click()
    Dim strFolderPath As String
    strFolderPath = InputBox("Enter the Folder path", Me.mnupFileSystemObject_CreateFolder.Caption, "c:\temp")
    With Me.txtMain
        goFS.CreateFolder strFolderPath
        .Text = "CreateFolder:" & vbTab
        If goFS.LastErr = "" Then
            .Text = .Text & "OK" & vbCrLf
        Else
            .Text = .Text & goFS.LastErr & vbCrLf
        End If
    End With
End Sub

Private Sub mnupFileSystemObject_FileExists_Click()
    Dim strFilePath As String
    strFilePath = InputBox("Enter the file path", Me.mnupFileSystemObject_FileExists.Caption, App.Path & "\modFsDeclare.bas")
    With Me.txtMain
        .Text = "FileExists:" & vbTab & goFS.FileExists(strFilePath) & vbCrLf
    End With
End Sub


Private Sub mnupFileSystemObject_FolderExists_Click()
    Dim strFolderPath As String
    strFolderPath = InputBox("Enter the folder path", Me.mnupFileSystemObject_FolderExists.Caption, App.Path)
    With Me.txtMain
        .Text = "FolderExists:" & vbTab & goFS.FolderExists(strFolderPath) & vbCrLf
    End With
End Sub

Private Sub mnupFileSystemObject_GetFile_Click()
    Dim strFilePath As String
    strFilePath = InputBox("Enter the file path", Me.mnupFileSystemObject_GetFile.Caption, App.Path & "\modFsDeclare.bas")
    Dim oFile As pFile
    Set oFile = goFS.GetFile(strFilePath)
    With Me.txtMain
        .Text = "DateCreated:" & vbTab & oFile.DateCreated & vbCrLf
        .Text = .Text & "DateLastModified:" & vbTab & oFile.DateLastModified & vbCrLf
        .Text = .Text & "Name:" & vbTab & oFile.Name & vbCrLf
        .Text = .Text & "Path:" & vbTab & oFile.Path & vbCrLf
    End With
End Sub

Private Sub mnupFileSystemObject_GetFolder_Click()
    Dim strDirPath As String
    strDirPath = InputBox("Enter the folder path", Me.mnupFileSystemObject_GetFolder.Caption, App.Path)
    Dim oFolder  As pFolder
    Set oFolder = goFS.GetFolder(strDirPath)
    With Me.txtMain
        .Text = "DateCreated:" & vbTab & oFolder.DateCreated & vbCrLf
        .Text = .Text & "DateLastModified:" & vbTab & oFolder.DateLastModified & vbCrLf
        .Text = .Text & "Name:" & vbTab & oFolder.Name & vbCrLf
        .Text = .Text & "IsRootFolder:" & vbTab & oFolder.IsRootFolder & vbCrLf
    End With
End Sub

Private Sub mnupTextStream_ReadFile_Click()
    Dim strFilePath As String, strLine As String
    strFilePath = InputBox("Input file path", Me.mnupTextStream_ReadFile.Caption, App.Path & "\modFsDeclare.bas")
    Dim oTextStream As pTextStream
    Set oTextStream = goFS.OpenTextFile(strFilePath, pIOMode.ForReading, True)
    If goFS.LastErr <> "" Then
        MsgBox goFS.LastErr, vbCritical, Me.mnupTextStream_ReadFile.Caption
    Else
        Do While Not oTextStream.AtEndOfStream
            strLine = oTextStream.ReadLine
            With Me.txtMain
                .Text = .Text & strLine & vbCrLf
            End With
            DoEvents
        Loop
        oTextStream.CloseObj
    End If
End Sub

Private Sub mnupTextStream_WriteFile_Click()
    Dim strFilePath As String, strLine As String
    strFilePath = InputBox("Input file path", Me.mnupTextStream_WriteFile.Caption, "c:\temp\TestObjFs.txt")
    Dim oTextStream As pTextStream
    Set oTextStream = goFS.OpenTextFile(strFilePath, pIOMode.ForWriting, True)
    If goFS.LastErr <> "" Then
        MsgBox goFS.LastErr, vbCritical, Me.mnupTextStream_WriteFile.Caption
    Else
        oTextStream.WriteLine "WriteLine"
        oTextStream.WriteBlankLines 2
        oTextStream.WriteStr "WriteStr"
        oTextStream.CloseObj
        Me.txtMain.Text = mnupTextStream_WriteFile & ":" & strFilePath & " OK"
    End If
End Sub
