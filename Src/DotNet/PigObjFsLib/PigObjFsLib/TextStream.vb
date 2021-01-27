'**********************************
'* Name: TextStream
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Amount to Scripting.TextStream of VB6
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.3
'* Create Time: 30/12/2020
'* 1.0.2 15/1/2021   Err.Raise change to Throw New Exception
'* 1.0.3 23/1/2021   pTextStream rename to TextStream
'**********************************
Imports System.IO
Public Class TextStream
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.0.3"

    Private menmIOMode As FileSystemObject.IOMode

    Private srMain As StreamReader
    Private swMain As StreamWriter

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    Friend Sub Init(FilePath As String, IOMode As FileSystemObject.IOMode, Optional Create As Boolean = False)
        Try
            menmIOMode = IOMode
            Select Case menmIOMode
                Case FileSystemObject.IOMode.ForReading
                    If Create = True Then
                        If IO.File.Exists(FilePath) = False Then
                            swMain = New StreamWriter(FilePath)
                            swMain.Close()
                        End If
                    End If
                    srMain = New StreamReader(FilePath)
                Case FileSystemObject.IOMode.ForWriting
                    swMain = New StreamWriter(FilePath, True)
                Case FileSystemObject.IOMode.ForAppending
                    swMain = New StreamWriter(FilePath, False)
                Case Else
                    Throw New Exception("Invalid IOMode " & IOMode.ToString)
            End Select
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Init", ex)
        End Try
    End Sub

    Public ReadOnly Property AtEndOfStream() As Boolean
        Get
            Try
                Select Case menmIOMode
                    Case FileSystemObject.IOMode.ForReading
                        Return srMain.EndOfStream
                    Case Else
                        Return Nothing
                End Select
            Catch ex As Exception
                Me.SetSubErrInf("AtEndOfStream", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public Sub Close()
        Try
            Select Case menmIOMode
                Case FileSystemObject.IOMode.ForReading
                    srMain.Close()
                Case Else
                    swMain.Close()
            End Select
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Close", ex)
        End Try
    End Sub

    Public Function ReadAll() As String
        Try
            Select Case menmIOMode
                Case FileSystemObject.IOMode.ForReading
                    Return srMain.ReadToEnd()
                Case Else
                    Return ""
            End Select
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("ReadAll", ex)
            Return ""
        End Try
    End Function

    Public Function ReadLine() As String
        Try
            Select Case menmIOMode
                Case FileSystemObject.IOMode.ForReading
                    Return srMain.ReadLine
                Case Else
                    Return ""
            End Select
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("ReadLine", ex)
            Return ""
        End Try
    End Function


    Public Sub WriteLine(Text As String)
        Try
            Select Case menmIOMode
                Case FileSystemObject.IOMode.ForReading
                Case Else
                    swMain.WriteLine(Text)
            End Select
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("WriteLine", ex)
        End Try
    End Sub

    Public Sub WriteBlankLines(Lines As Long)
        For i = 1 To Lines
            Me.WriteLine("")
        Next
    End Sub
    Public Sub Write(Text As String)
        Try
            Select Case menmIOMode
                Case FileSystemObject.IOMode.ForReading
                Case Else
                    swMain.Write(Text)
            End Select
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Write", ex)
        End Try
    End Sub

End Class
