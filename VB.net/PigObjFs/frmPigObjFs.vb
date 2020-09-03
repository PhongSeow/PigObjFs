'************************************************
'Form name: Obj FileSystemObject Demo
'Author: Seow Phong
'Web site: https://en.seowphong.com
'Description: Reference required PigObjFsLib.dll
'Version: 1.0.1
'Created: September 2, 2020
'************************************************

Imports PigObjFsLib

Public Class frmPigObjFs
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub frmPigObjFs_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        goFS = New pFileSystemObject
        goFS.InitLog(Application.StartupPath, My.Application.Info.AssemblyName)
        Me.Text = My.Application.Info.ProductName
    End Sub

    Private Sub GetFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GetFileToolStripMenuItem.Click
        Dim oFolder As pFolder
        Dim oFile As pFile, strFilePath As String
        strFilePath = InputBox("Enter the file path", sender.ToString, Me.GetCurrDirFristFile)
        oFile = goFS.GetFile(strFilePath)
        If oFile.IsObjReady = True Then
            With Me.tbMain
                .Text = "DateCreated:" & vbTab & oFile.DateCreated & vbCrLf
                .Text = .Text & "DateLastModified:" & vbTab & oFile.DateLastModified & vbCrLf
                .Text = .Text & "Name:" & vbTab & oFile.Name & vbCrLf
                .Text = .Text & "Path:" & vbTab & oFile.Path & vbCrLf
            End With
        Else
            MsgBox("oFile.IsObjReady = " & oFile.IsObjReady, vbCritical, sender.ToString)
        End If
    End Sub
    Public Function GetCurrDirFristFile() As String
        Dim oFolder As pFolder, objFile As Object, oFile As pFile
        oFolder = goFS.GetFolder(Application.StartupPath)
        GetCurrDirFristFile = ""
        If oFolder.IsObjReady = True Then
            For Each objFile In oFolder.Files
                If Not objFile Is Nothing Then
                    oFile = New pFile
                    oFile.Obj = objFile
                    GetCurrDirFristFile = oFile.Path
                    Exit For
                End If
            Next
        End If
    End Function

    Private Sub GetFolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GetFolderToolStripMenuItem.Click
        Dim strDirPath As String
        strDirPath = InputBox("Enter the folder path", sender.ToString, Application.StartupPath)
        Dim oFolder As pFolder
        oFolder = goFS.GetFolder(strDirPath)
        If oFolder.IsObjReady = True Then
            With Me.tbMain
                .Text = "DateCreated:" & vbTab & oFolder.DateCreated & vbCrLf
                .Text = .Text & "DateLastModified:" & vbTab & oFolder.DateLastModified & vbCrLf
                .Text = .Text & "Name:" & vbTab & oFolder.Name & vbCrLf
                .Text = .Text & "IsRootFolder:" & vbTab & oFolder.IsRootFolder & vbCrLf
            End With
        Else
            MsgBox("oFolder.IsObjReady = " & oFolder.IsObjReady, vbCritical, sender.ToString)
        End If
    End Sub

    Private Sub FileExistsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FileExistsToolStripMenuItem.Click
        Dim strFilePath As String
        strFilePath = InputBox("Enter the file path", sender.ToString, Me.GetCurrDirFristFile)
        With Me.tbMain
            .Text = "FileExists:" & vbTab & goFS.FileExists(strFilePath) & vbCrLf
        End With
    End Sub

    Private Sub FolderExistsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FolderExistsToolStripMenuItem.Click
        Dim strFolderPath As String
        strFolderPath = InputBox("Enter the folder path", sender.ToString, Application.StartupPath)
        With Me.tbMain
            .Text = "FolderExists:" & vbTab & goFS.FolderExists(strFolderPath) & vbCrLf
        End With
    End Sub

    Private Sub CreateFolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateFolderToolStripMenuItem.Click
        Dim strFolderPath As String
        strFolderPath = InputBox("Enter the Folder path", sender.ToString, "c:\temp\a\b\c")
        With Me.tbMain
            goFS.CreateFolder(strFolderPath)
            .Text = "CreateFolder:" & vbTab
            If goFS.LastErr = "" Then
                .Text = .Text & "OK" & vbCrLf
            Else
                .Text = .Text & goFS.LastErr & vbCrLf
            End If
        End With
    End Sub

    Private Sub ReadFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReadFileToolStripMenuItem.Click
        Dim strFilePath As String, strLine As String
        strFilePath = Application.StartupPath
        If InStr(strFilePath, "\PigObjFs\bin\Debug") Then strFilePath = Replace(strFilePath, "\PigObjFs\bin\Debug", "")
        strFilePath = InputBox("Input file path", sender.ToString, strFilePath & "\README.cmd")
        If goFS.FileExists(strFilePath) = False Then
            MsgBox(strFilePath & " not found.", vbCritical, sender.ToString)
        Else
            Dim oTextStream As pTextStream
            oTextStream = goFS.OpenTextFile(strFilePath, pIOMode.ForReading, False)
            If goFS.LastErr <> "" Then
                MsgBox(goFS.LastErr, vbCritical, sender.ToString)
            Else
                Do While Not oTextStream.AtEndOfStream
                    strLine = oTextStream.ReadLine
                    With Me.tbMain
                        .Text = .Text & strLine & vbCrLf
                    End With
                    Application.DoEvents()
                Loop
                oTextStream.CloseObj()
            End If

        End If
    End Sub

    Private Sub WriteFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WriteFileToolStripMenuItem.Click
        Dim strFilePath As String
        strFilePath = InputBox("Input file path", sender.ToString, "c:\temp\TestObjFs.txt")
        Dim oTextStream As pTextStream
        oTextStream = goFS.OpenTextFile(strFilePath, pIOMode.ForWriting, True)
        If goFS.LastErr <> "" Then
            MsgBox(goFS.LastErr, vbCritical, sender.ToString)
        Else
            oTextStream.WriteLine("WriteLine")
            oTextStream.WriteBlankLines(2)
            oTextStream.WriteStr("WriteStr")
            oTextStream.CloseObj()
            Me.tbMain.Text = sender.ToString & ":" & strFilePath & " OK"
        End If
    End Sub

    Private Sub SeowPhongStudioToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SeowPhongStudioToolStripMenuItem.Click
        System.Diagnostics.Process.Start("https://en.seowphong.com/")
    End Sub

    Private Sub OnlineDocumentationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OnlineDocumentationToolStripMenuItem.Click
        System.Diagnostics.Process.Start("https://en.seowphong.com/oss/PigObjFs/")
    End Sub
End Class
