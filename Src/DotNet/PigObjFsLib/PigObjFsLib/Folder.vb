'**********************************
'* Name: Folder
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Amount to Scripting.Folder of VB6
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.3
'* Create Time: 31/12/2020
'* 1.0.2 23/1/2021   pFolder rename to Folder
'* 1.1 28/8/2021   Modify mGetSubDirList
'* 1.2 13/3/2021   Modify Obj
'* 1.3 1/5/2023    Add FileCount
'**********************************
Imports System.IO
Public Class Folder
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.3.2"

    Friend Obj As DirectoryInfo

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
            Dim lngCount As Long
#If NET40_OR_GREATER Or NETCOREAPP Then
            lngCount = oDir.GetDirectories.LongCount
#Else
            lngCount = oDir.GetDirectories.Length
#End If

            If lngCount > 0 Then
                For Each oSubDir In oDir.GetDirectories
                    Dim strRet As String = ""
                    strRet = mGetSubDirList(oSubDir, DirArray, Ret)
                    If strRet <> "OK" Then
                        Ret &= oSubDir.FullName & ":Err=" & strRet & vbCrLf
                    Else
                        Dim lngDirCount As Long
#If NET40_OR_GREATER Or NETCOREAPP Then
                        lngDirCount = DirArray.LongCount
#Else
                        lngDirCount = DirArray.Length
#End If
                        ReDim Preserve DirArray(lngDirCount)
                        DirArray(lngDirCount - 1) = oSubDir.FullName
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

    Public ReadOnly Property FileCount() As Integer
        Get
#If NET20 Or NET30 Then
            Return Me.Files.Length
#Else
            Return Me.Files.Count
#End If
        End Get
    End Property


    Public ReadOnly Property Files() As File()
        Get
            Try
                Dim aFile() As File
                ReDim aFile(-1)
                For Each oFileInfo In Me.Obj.GetFiles
                    Dim oFile = New File With {.Obj = oFileInfo}
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

    Public ReadOnly Property ParentFolder() As Folder
        Get
            Try
                Dim oFolder As New Folder With {.Obj = Me.Obj.Parent}
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

    Public ReadOnly Property mSubFolders() As Folder()
        Get
            Try
                Dim aFolder() As Folder
                ReDim aFolder(-1)
                For Each oDirectoryInfo In Me.Obj.GetDirectories
                    ReDim Preserve aFolder(aFolder.LongLength)
                    Dim oFolder As New Folder With {.Obj = oDirectoryInfo}
                    aFolder(aFolder.LongLength - 1) = oFolder
                Next
                Return aFolder
            Catch ex As Exception
                Me.SetSubErrInf("ParentFolder", ex)
                Return Nothing
            End Try
        End Get
    End Property


    Public ReadOnly Property SubFolders() As Folder()
        Get
            Return Me.mSubFolders
        End Get
    End Property

    Public ReadOnly Property SubFolders(IsIncSub As Boolean) As Folder()
        Get
            Try
                If IsIncSub = True Then
                    Dim strRet As String = "", saDir(-1) As String
                    Me.mGetSubDirList(Me.Obj, saDir, strRet)
                    Dim aFolder() As Folder
                    ReDim aFolder(-1)
                    For i = 0 To saDir.LongLength - 1
                        Dim oDir As New DirectoryInfo(saDir(i))
                        ReDim Preserve aFolder(aFolder.LongLength)
                        Dim oFolder As New Folder With {.Obj = oDir}
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
