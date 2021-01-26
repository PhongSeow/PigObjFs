## PigObjFsLib Sample code
#### [PigObjFsDemoConsole on GitHub](https://github.com/PhongSeow/PigObjFs/tree/master/Src/DotNet/PigObjFsDemoConsole)

***GetFile Sample code***

```
Dim oFile As File = goFS.GetFile(strFilePath)
Console.WriteLine("DateCreated: " & oFile.DateCreated)
Console.WriteLine("DateLastModified: " & oFile.DateLastModified)
Console.WriteLine("Name: " & oFile.Name)
Console.WriteLine("Path: " & oFile.Path)
```

***Return results***

```
DateCreated: 2021/1/25 12:24:13
DateLastModified: 2021/1/26 17:05:00
Name: PigObjFsDemoConsole.exe
Path: C:\GitHub\Dev\PigObjFs\Src\DotNet\PigObjFsDemoConsole\PigObjFsDemoConsole\bin\Debug\net45\PigObjFsDemoConsole.exe
```

***GetFolder Sample code***

```visual basic
Dim oFolder As Folder = oFS.GetFolder(strFolderPath)
Console.WriteLine("DateCreated: " & oFolder.DateCreated)
Console.WriteLine("DateLastModified: " & oFolder.DateLastModified)
Console.WriteLine("Name: " & oFolder.Name)
Console.WriteLine("Path: " & oFolder.Path)
```

***Return results***	

```
DateCreated: 2021/1/25 11:33:55
DateLastModified: 2021/1/25 12:24:13
Name: net45
Path: C:\GitHub\Dev\PigObjFs\Src\DotNet\PigObjFsDemoConsole\PigObjFsDemoConsole\bin\Debug\net45\
```

***FileExists Sample code***

```
Console.WriteLine("FileExists: " & oFS.FileExists(strFilePath))
```

***Return results***
```
FileExists: True
```

***FolderExists Sample code***

```
Console.WriteLine("FolderExists: " & oFS.FolderExists(strFolderPath))
```

***Return results***
```
FolderExists: True
```


***TextStream Read File  code***

```
Dim oTextStream As TextStream
Console.WriteLine("OpenTextFile(" & strFilePath & ")...")
oTextStream = oFS.OpenTextFile(strFilePath, FileSystemObject.pIOMode.ForReading, False)
If oFS.LastErr <> "" Then
	Console.WriteLine(oFS.LastErr)
Else
    Do While Not oTextStream.AtEndOfStream
        Console.WriteLine(oTextStream.ReadLine)
    Loop
	oTextStream.Close()
End If

```

***TextStream WriteFile  code***

```
Dim oTextStream As TextStream
Console.WriteLine("OpenTextFile(" & strFilePath & ")...")
oTextStream = oFS.OpenTextFile(strFilePath, FileSystemObject.pIOMode.ForWriting, True)
If oFS.LastErr <> "" Then
	Console.WriteLine(oFS.LastErr)
Else
    oTextStream.WriteLine("WriteLine")
    oTextStream.WriteBlankLines(2)
    oTextStream.Close()
    Console.WriteLine("OK")
End If
```
