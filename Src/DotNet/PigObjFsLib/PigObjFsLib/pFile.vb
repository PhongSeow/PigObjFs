'**********************************
'* Name: pFile
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Amount to Scripting.File of VB6
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.2
'* Create Time: 30/12/2020
'* 1.0.2 15/1/2021   Err.Raise change to Throw New Exception
'**********************************
Imports System.IO
Public Class pFile
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.0.1"

    Public Obj As FileInfo

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    Public Sub New(FilePath As String)
        MyBase.New(CLS_VERSION)
        Me.Obj = New FileInfo(FilePath)
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

    Public Sub Delete(Optional ByVal Force As Boolean = False)
        Try

            If Me.Obj.Exists Then
                If Force = False Then
                    Throw New Exception("File already exists.")
                End If
            End If
            Obj.Delete()
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Delete", ex)
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

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

    Public ReadOnly Property Size() As Long
        Get
            Try
                Return Me.Obj.Length
            Catch ex As Exception
                Me.SetSubErrInf("Size", ex)
                Return -1
            End Try
        End Get
    End Property

End Class
