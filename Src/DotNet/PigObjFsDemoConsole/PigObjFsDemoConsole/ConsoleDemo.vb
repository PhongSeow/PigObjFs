'**********************************
'* Name: ConsoleDemo
'* Author: Seow Phong
'* License: Copyright (c) 2021-2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.2
'* Create Time: 31/12/202
'* 1.1.8 1/3/2022   Modify New FileSystemObject
'* 1.2.6 13/3/2022   Add TextStreamAsc
'**********************************

Imports PigObjFsLib

Public Class ConsoleDemo

    Private moFS As FileSystemObject

    Public Sub Main()
        moFS = New FileSystemObject()
        moFS.OpenDebug(True)
        Do While True
            Console.WriteLine("*******************")
            Console.WriteLine("Main menu")
            Console.WriteLine("*******************")
            Console.WriteLine("Press Q to Exit")
            Console.WriteLine("Press A to pFileSystemObject")
            Console.WriteLine("Press B to TextStream")
#If NETFRAMEWORK Then
            Console.WriteLine("Press C to TextStreamAsc")
#End If
            Console.WriteLine("*******************")
            Select Case Console.ReadKey(True).Key
                Case ConsoleKey.Q
                    Exit Do
                Case ConsoleKey.A
                    Do While True
                        Console.WriteLine("*******************")
                        Console.WriteLine("Menu pFileSystemObject")
                        Console.WriteLine("*******************")
                        Console.WriteLine("Press Q to Up")
                        Console.WriteLine("Press A to GetFile")
                        Console.WriteLine("Press B to GetFolder")
                        Console.WriteLine("Press C to FileExists")
                        Console.WriteLine("Press D to FolderExists")
                        Console.WriteLine("Press E to CreateFolder")
                        Console.WriteLine("*******************")
                        Select Case Console.ReadKey(True).Key
                            Case ConsoleKey.Q
                                Exit Do
                            Case ConsoleKey.A
                                Dim strFilePath As String
                                Console.WriteLine("#################")
                                Console.WriteLine("Enter the file path, such as " & GetCurrDirFristFile())
                                strFilePath = Console.ReadLine()
                                If strFilePath = "" Then
                                    strFilePath = GetCurrDirFristFile()
                                End If
                                Dim oFile As File = moFS.GetFile(strFilePath)
                                Console.WriteLine("DateCreated: " & oFile.DateCreated)
                                Console.WriteLine("DateLastModified: " & oFile.DateLastModified)
                                Console.WriteLine("Name: " & oFile.Name)
                                Console.WriteLine("Path: " & oFile.Path)
                                Console.WriteLine("#################")
                            Case ConsoleKey.B
                                Dim strFolderPath As String
                                Console.WriteLine("#################")
                                Console.WriteLine("Enter the file folder, such as " & moFS.AppPath)
                                strFolderPath = Console.ReadLine()
                                If strFolderPath = "" Then
                                    strFolderPath = moFS.AppPath
                                End If
                                Dim oFolder As Folder = moFS.GetFolder(strFolderPath)
                                Console.WriteLine("DateCreated: " & oFolder.DateCreated)
                                Console.WriteLine("DateLastModified: " & oFolder.DateLastModified)
                                Console.WriteLine("Name: " & oFolder.Name)
                                Console.WriteLine("Path: " & oFolder.Path)
                                Console.WriteLine("#################")
                            Case ConsoleKey.C
                                Dim strFilePath As String
                                Console.WriteLine("#################")
                                Console.WriteLine("Enter the file path, such as " & GetCurrDirFristFile())
                                strFilePath = Console.ReadLine()
                                If strFilePath = "" Then
                                    strFilePath = GetCurrDirFristFile()
                                End If
                                Console.WriteLine("FileExists: " & moFS.FileExists(strFilePath))
                                Console.WriteLine("#################")
                            Case ConsoleKey.D
                                Dim strFolderPath As String
                                Console.WriteLine("#################")
                                Console.WriteLine("Enter the file folder, such as " & moFS.AppPath)
                                strFolderPath = Console.ReadLine()
                                If strFolderPath = "" Then
                                    strFolderPath = moFS.AppPath
                                End If
                                Console.WriteLine("FolderExists: " & moFS.FolderExists(strFolderPath))
                                Console.WriteLine("#################")
                            Case ConsoleKey.E
                                Dim strFolderPath As String, strDefaPath As String
                                Console.WriteLine("#################")
                                With moFS
                                    strDefaPath = .AppPath & .OsPathSep & "a" & .OsPathSep & "b"
                                    Console.WriteLine("Enter the file folder, such as " & strDefaPath)
                                End With
                                strFolderPath = Console.ReadLine()
                                If strFolderPath = "" Then
                                    strFolderPath = strDefaPath
                                End If
                                moFS.CreateFolder(strFolderPath)
                                Console.Write("CreateFolder: ")
                                If moFS.LastErr = "" Then
                                    Console.WriteLine("OK")
                                Else
                                    Console.WriteLine(moFS.LastErr)
                                End If
                                Console.WriteLine("#################")
                        End Select
                    Loop
                Case ConsoleKey.B
                    Do While True
                        Console.WriteLine("*******************")
                        Console.WriteLine("Menu TextStream")
                        Console.WriteLine("*******************")
                        Console.WriteLine("Press Q to Up")
                        Console.WriteLine("Press A to Read File")
                        Console.WriteLine("Press B to Write File")
                        Console.WriteLine("*******************")
                        Select Case Console.ReadKey().Key
                            Case ConsoleKey.Q
                                Exit Do
                            Case ConsoleKey.A
                                Dim strFilePath As String
                                Console.WriteLine("#################")
                                Console.WriteLine("Enter the file path, such as " & GetCurrDirFristFile())
                                strFilePath = Console.ReadLine()
                                If strFilePath = "" Then
                                    strFilePath = GetCurrDirFristFile()
                                End If
                                If moFS.FileExists(strFilePath) = False Then
                                    Console.WriteLine(strFilePath & " not found.")
                                Else
                                    Dim oTextStream As TextStream
                                    Console.WriteLine("OpenTextFile(" & strFilePath & ")...")
                                    oTextStream = moFS.OpenTextFile(strFilePath, FileSystemObject.IOMode.ForReading, False)
                                    If moFS.LastErr <> "" Then
                                        Console.WriteLine(moFS.LastErr)
                                    Else
                                        Do While Not oTextStream.AtEndOfStream
                                            Dim strLine As String = oTextStream.ReadLine
                                            Console.WriteLine(strLine)
                                        Loop
                                        oTextStream.Close()
                                    End If
                                End If
                                Console.WriteLine("#################")
                            Case ConsoleKey.B
                                Dim strFilePath As String, strDefaFile As String
                                Console.Write("Enter the file path, such as ")
                                If moFS.IsWindows = True Then
                                    strDefaFile = "C:\Temp\TestPigFsDemo.txt"
                                Else
                                    strDefaFile = "/tmp/TestPigFsDemo.txt"
                                End If
                                Console.WriteLine(strDefaFile)
                                strFilePath = Console.ReadLine()
                                If strFilePath = "" Then
                                    strFilePath = strDefaFile
                                End If
                                Dim oTextStream As TextStream
                                Console.WriteLine("OpenTextFile(" & strFilePath & ")...")
                                oTextStream = moFS.OpenTextFile(strFilePath, FileSystemObject.IOMode.ForWriting, True)
                                If moFS.LastErr <> "" Then
                                    Console.WriteLine(moFS.LastErr)
                                Else
                                    oTextStream.WriteLine("WriteLine")
                                    oTextStream.WriteBlankLines(2)
                                    oTextStream.Close()
                                    Console.WriteLine("OK")
                                End If
                        End Select
                    Loop
#If NETFRAMEWORK Then
                Case ConsoleKey.C
                    Do While True
                        Console.WriteLine("*******************")
                        Console.WriteLine("Menu TextStreamAsc")
                        Console.WriteLine("*******************")
                        Console.WriteLine("Press Q to Up")
                        Console.WriteLine("Press A to Read File")
                        Console.WriteLine("Press B to Write File")
                        Console.WriteLine("*******************")
                        Select Case Console.ReadKey().Key
                            Case ConsoleKey.Q
                                Exit Do
                            Case ConsoleKey.A
                                Dim strFilePath As String
                                Console.WriteLine("#################")
                                Console.WriteLine("Enter the file path, such as " & GetCurrDirFristFile())
                                strFilePath = Console.ReadLine()
                                If strFilePath = "" Then
                                    strFilePath = GetCurrDirFristFile()
                                End If
                                If moFS.FileExists(strFilePath) = False Then
                                    Console.WriteLine(strFilePath & " not found.")
                                Else
                                    Dim oTextStream As TextStreamAsc
                                    Console.WriteLine("OpenTextFile(" & strFilePath & ")...")
                                    oTextStream = moFS.OpenTextFileAsc(strFilePath, FileSystemObject.IOMode.ForReading, False)
                                    If moFS.LastErr <> "" Then
                                        Console.WriteLine(moFS.LastErr)
                                    Else
                                        Do While Not oTextStream.AtEndOfStream
                                            Dim strLine As String = oTextStream.ReadLine
                                            Console.WriteLine(strLine)
                                        Loop
                                        oTextStream.Close()
                                    End If
                                End If
                                Console.WriteLine("#################")
                            Case ConsoleKey.B
                                Dim strFilePath As String, strDefaFile As String
                                Console.Write("Enter the file path, such as ")
                                If moFS.IsWindows = True Then
                                    strDefaFile = "C:\Temp\TestPigFsDemo.txt"
                                Else
                                    strDefaFile = "/tmp/TestPigFsDemo.txt"
                                End If
                                Console.WriteLine(strDefaFile)
                                strFilePath = Console.ReadLine()
                                If strFilePath = "" Then
                                    strFilePath = strDefaFile
                                End If
                                Dim oTextStream As TextStreamAsc
                                Console.WriteLine("OpenTextFile(" & strFilePath & ")...")
                                oTextStream = moFS.OpenTextFileAsc(strFilePath, FileSystemObject.IOMode.ForWriting, True)
                                If moFS.LastErr <> "" Then
                                    Console.WriteLine(moFS.LastErr)
                                Else
                                    oTextStream.WriteLine("WriteLine")
                                    oTextStream.WriteBlankLines(2)
                                    oTextStream.Close()
                                    Console.WriteLine("OK")
                                End If
                        End Select
                    Loop
#End If
            End Select
        Loop
    End Sub

    Public Function GetCurrDirFristFile() As String
        Dim oFolder As Folder
        oFolder = moFS.GetFolder(moFS.AppPath)
        GetCurrDirFristFile = ""
        For Each objFile In oFolder.Files
            If Not objFile Is Nothing Then
                GetCurrDirFristFile = objFile.Path
                Exit For
            End If
        Next
    End Function

End Class
