Imports System
Imports PigObjFsLib
Module Program
    Public goFS As New pFileSystemObject
    Sub Main(args As String())
        Do While True
            Console.WriteLine("*******************")
            Console.WriteLine("Main menu")
            Console.WriteLine("*******************")
            Console.WriteLine("Press Q to Exit")
            Console.WriteLine("Press A to pFileSystemObject")
            Console.WriteLine("Press B to pTextStream")
            Console.WriteLine("*******************")
            Select Case Console.ReadKey().Key
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
                                Dim oFile As pFile = goFS.GetFile(strFilePath)
                                Console.WriteLine("DateCreated: " & oFile.DateCreated)
                                Console.WriteLine("DateLastModified: " & oFile.DateLastModified)
                                Console.WriteLine("Name: " & oFile.Name)
                                Console.WriteLine("Path: " & oFile.Path)
                                Console.WriteLine("#################")
                            Case ConsoleKey.B
                                Dim strFolderPath As String
                                Console.WriteLine("#################")
                                Console.WriteLine("Enter the file folder, such as " & goFS.AppPath)
                                strFolderPath = Console.ReadLine()
                                If strFolderPath = "" Then
                                    strFolderPath = goFS.AppPath
                                End If
                                Dim oFolder As pFolder = goFS.GetFolder(strFolderPath)
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
                                Console.WriteLine("FileExists: " & goFS.FileExists(strFilePath))
                                Console.WriteLine("#################")
                            Case ConsoleKey.D
                                Dim strFolderPath As String
                                Console.WriteLine("#################")
                                Console.WriteLine("Enter the file folder, such as " & goFS.AppPath)
                                strFolderPath = Console.ReadLine()
                                If strFolderPath = "" Then
                                    strFolderPath = goFS.AppPath
                                End If
                                Console.WriteLine("FolderExists: " & goFS.FolderExists(strFolderPath))
                                Console.WriteLine("#################")
                            Case ConsoleKey.E
                                Dim strFolderPath As String, strDefaPath As String
                                Console.WriteLine("#################")
                                With goFS
                                    strDefaPath = .AppPath & "a" & .OsPathSep & "b"
                                    Console.WriteLine("Enter the file folder, such as " & strDefaPath)
                                End With
                                strFolderPath = Console.ReadLine()
                                If strFolderPath = "" Then
                                    strFolderPath = strDefaPath
                                End If
                                goFS.CreateFolder(strFolderPath)
                                Console.Write("CreateFolder: ")
                                If goFS.LastErr = "" Then
                                    Console.WriteLine("OK")
                                Else
                                    Console.WriteLine(goFS.LastErr)
                                End If
                                Console.WriteLine("#################")
                        End Select
                    Loop
                Case ConsoleKey.B
                    Do While True
                        Console.WriteLine("*******************")
                        Console.WriteLine("Menu pTextStream")
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
                                If goFS.FileExists(strFilePath) = False Then
                                    Console.WriteLine(strFilePath & " not found.")
                                Else
                                    Dim oTextStream As pTextStream
                                    Console.WriteLine("OpenTextFile(" & strFilePath & ")...")
                                    oTextStream = goFS.OpenTextFile(strFilePath, pFileSystemObject.pIOMode.ForReading, False)
                                    If goFS.LastErr <> "" Then
                                        Console.WriteLine(goFS.LastErr)
                                    Else
                                        Do While Not oTextStream.AtEndOfStream
                                            Console.WriteLine(oTextStream.ReadLine)
                                        Loop
                                        oTextStream.Close()
                                    End If
                                End If
                                Console.WriteLine("#################")
                            Case ConsoleKey.B
                                Dim strFilePath As String, strDefaFile As String
                                Console.Write("Enter the file path, such as ")
                                If goFS.IsWindows = True Then
                                    strDefaFile = "C:\Temp\TestPigFsDemo.txt"
                                Else
                                    strDefaFile = "/Temp/TestPigFsDemo.txt"
                                End If
                                Console.WriteLine(strDefaFile)
                                strFilePath = Console.ReadLine()
                                If strFilePath = "" Then
                                    strFilePath = strDefaFile
                                End If
                                Dim oTextStream As pTextStream
                                Console.WriteLine("OpenTextFile(" & strFilePath & ")...")
                                oTextStream = goFS.OpenTextFile(strFilePath, pFileSystemObject.pIOMode.ForWriting, True)
                                If goFS.LastErr <> "" Then
                                    Console.WriteLine(goFS.LastErr)
                                Else
                                    oTextStream.WriteLine("WriteLine")
                                    oTextStream.WriteBlankLines(2)
                                    oTextStream.Close()
                                    Console.WriteLine("OK")
                                End If
                        End Select
                    Loop
            End Select

        Loop
    End Sub

    Public Function GetCurrDirFristFile() As String
        Dim oFolder As pFolder
        oFolder = goFS.GetFolder(goFS.AppPath)
        GetCurrDirFristFile = ""
        For Each objFile In oFolder.Files
            If Not objFile Is Nothing Then
                GetCurrDirFristFile = objFile.Path
                Exit For
            End If
        Next
    End Function

End Module
