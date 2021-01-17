﻿'**********************************
'* Name: pFolder
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Amount to Scripting.Folder of VB6
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.1
'* Create Time: 31/12/2020
'**********************************
Imports System.IO
Public Class pFolder
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.0.1"

    Public Obj As DirectoryInfo

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    Public ReadOnly Property Name() As String
        Get
            Try
                Return Me.Obj.Name
            Catch ex As Exception
                Me.SetSubErrInf("Name", ex)
                Return ""
            End Try
        End Get
    End Property

    Public ReadOnly Property DateLastModified() As DateTime
        Get
            Try
                Return Me.Obj.LastWriteTime
            Catch ex As Exception
                Me.SetSubErrInf("DateLastModified", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property Path() As String
        Get
            Try
                Return Me.Obj.FullName
            Catch ex As Exception
                Me.SetSubErrInf("Path", ex)
                Return ""
            End Try
        End Get
    End Property

    Public ReadOnly Property DateCreated() As DateTime
        Get
            Try
                Return Me.Obj.CreationTime
            Catch ex As Exception
                Me.SetSubErrInf("DateCreated", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Private Function mGetSubDirList(ByRef oDir As DirectoryInfo, ByRef DirArray As String(), ByRef Ret As String) As String
        Try
            If oDir.GetDirectories.LongCount > 0 Then
                For Each oSubDir In oDir.GetDirectories
                    Dim strRet As String = ""
                    strRet = mGetSubDirList(oSubDir, DirArray, Ret)
                    If strRet <> "OK" Then
                        Ret &= oSubDir.FullName & ":Err=" & strRet & vbCrLf
                    Else
                        ReDim Preserve DirArray(DirArray.LongCount)
                        DirArray(DirArray.LongCount - 1) = oSubDir.FullName
                    End If
                Next
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

    Public ReadOnly Property Size() As Long
        Get
            Try
                Size = 0
                For Each oFile In Me.Obj.GetFiles
                    Size += oFile.Length
                Next
                Dim strRet As String = "", saDir(-1) As String
                Me.mGetSubDirList(Me.Obj, saDir, strRet)
                For i = 0 To saDir.LongLength - 1
                    Dim oDir As New DirectoryInfo(saDir(i))
                    For Each oFile In oDir.GetFiles
                        Size += oFile.Length
                    Next
                Next
                Return Size
            Catch ex As Exception
                Me.SetSubErrInf("Size", ex)
                Return -1
            End Try
        End Get
    End Property

    Public ReadOnly Property Files() As pFile()
        Get
            Try
                Dim aFile() As pFile
                ReDim aFile(-1)
                For Each oFileInfo In Me.Obj.GetFiles
                    Dim oFile = New pFile With {.Obj = oFileInfo}
                    ReDim Preserve aFile(aFile.LongLength)
                    aFile(aFile.LongLength - 1) = oFile
                Next
                Return aFile
            Catch ex As Exception
                Me.SetSubErrInf("Files", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property ParentFolder() As pFolder
        Get
            Try
                Dim oFolder As New pFolder With {.Obj = Me.Obj.Parent}
                Return oFolder
            Catch ex As Exception
                Me.SetSubErrInf("ParentFolder", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property IsRootFolder() As Boolean
        Get
            Try
                If Me.Obj.Parent Is Nothing Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                Me.SetSubErrInf("IsRootFolder", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property mSubFolders() As pFolder()
        Get
            Try
                Dim aFolder() As pFolder
                ReDim aFolder(-1)
                For Each oDirectoryInfo In Me.Obj.GetDirectories
                    ReDim Preserve aFolder(aFolder.LongLength)
                    Dim oFolder As New pFolder With {.Obj = oDirectoryInfo}
                    aFolder(aFolder.LongLength - 1) = oFolder
                Next
                Return aFolder
            Catch ex As Exception
                Me.SetSubErrInf("ParentFolder", ex)
                Return Nothing
            End Try
        End Get
    End Property


    Public ReadOnly Property SubFolders() As pFolder()
        Get
            Return Me.mSubFolders
        End Get
    End Property

    Public ReadOnly Property SubFolders(IsIncSub As Boolean) As pFolder()
        Get
            Try
                If IsIncSub = True Then
                    Dim strRet As String = "", saDir(-1) As String
                    Me.mGetSubDirList(Me.Obj, saDir, strRet)
                    Dim aFolder() As pFolder
                    ReDim aFolder(-1)
                    For i = 0 To saDir.LongLength - 1
                        Dim oDir As New DirectoryInfo(saDir(i))
                        ReDim Preserve aFolder(aFolder.LongLength)
                        Dim oFolder As New pFolder With {.Obj = oDir}
                        aFolder(aFolder.LongLength - 1) = oFolder
                    Next
                    Return aFolder
                Else
                    Return Me.mSubFolders
                End If
            Catch ex As Exception
                Me.SetSubErrInf("SubFolders", ex)
                Return Nothing
            End Try
        End Get
    End Property

End Class
