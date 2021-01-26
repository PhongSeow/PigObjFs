'**********************************
'* Name: pFileSystemObject
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Amount to Scripting.FileSystemObject of VB6
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.3
'* Create Time: 31/12/2020
'* 1.0.2 15/1/2021   Err.Raise change to Throw New Exception
'* 1.0.3 23/1/2021   pFileSystemObject rename to FileSystemObject
'**********************************
Imports System.IO
Public Class FileSystemObject
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.0.3"

    Private Shadows AppTitle As String

    Public Enum IOMode
        ForAppending = 8
        ForReading = 1
        ForWriting = 2
    End Enum


    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    Public Sub MoveFile(Source As String, Destination As String, Optional OverWriteFiles As Boolean = True)
        Try
            IO.File.Move(Source, Destination)
        Catch ex As Exception
            Me.SetSubErrInf("MoveFile", ex)
        End Try
    End Sub

    Public Sub CopyFile(Source As String, Destination As String, Optional OverWriteFiles As Boolean = True)
        Try
            IO.File.Copy(Source, Destination, OverWriteFiles)
        Catch ex As Exception
            Me.SetSubErrInf("CopyFile", ex)
        End Try
    End Sub

    Public Function CreateFolder(Path As String) As Folder
        Try
            CreateFolder = New Folder
            CreateFolder.Obj = Directory.CreateDirectory(Path)
        Catch ex As Exception
            Me.SetSubErrInf("CreateFolder", ex)
            Return Nothing
        End Try
    End Function

    Public ReadOnly Property FileExists(FileSpec As String) As Boolean
        Get
            Try
                Return IO.File.Exists(FileSpec)
            Catch ex As Exception
                Me.SetSubErrInf("FileExists", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property FolderExists(FolderSpec As String) As Boolean
        Get
            Try
                Return Directory.Exists(FolderSpec)
            Catch ex As Exception
                Me.SetSubErrInf("FolderExists", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public Function GetFile(FilePath As String) As File
        Try
            Dim oFile As New File With {.Obj = New FileInfo(FilePath)}
            Return oFile
        Catch ex As Exception
            Me.SetSubErrInf("GetFile", ex)
            Return Nothing
        End Try
    End Function

    Public Function GetFolder(FolderPath As String) As Folder
        Try
            Dim oFolder As New Folder With {.Obj = New DirectoryInfo(FolderPath)}
            Return oFolder
        Catch ex As Exception
            Me.SetSubErrInf("GetFolder", ex)
            Return Nothing
        End Try
    End Function

    Public Function OpenTextFile(FilePath As String, IOMode As IOMode, Optional Create As Boolean = False) As TextStream
        Try
            OpenTextFile = New TextStream
            OpenTextFile.Init(FilePath, IOMode, Create)
            If OpenTextFile.LastErr <> "" Then Throw New Exception(OpenTextFile.LastErr)
        Catch ex As Exception
            Me.SetSubErrInf("OpenTextFile", ex)
            Return Nothing
        End Try
    End Function

End Class
