Explain:
This scheme is suitable for VB6, reference PigObjFsLib.dll The compiled exe can run directly on different versions of windows, but it needs to be registered through regsrv32.exe PigObjFsLib.dll ¡£

Example code:
Open VB6 project file PigObjFsTestDll.vbp , sample code in frmPigObjFs.frm Form.
PigObjFsTestDll.exe By PigObjFsTestDll.vbp It can run directly on different versions of windows, but it needs to be registered through regsrv32.exe PigObjFsLib.dll , the command is: Regsvr32 PigObjFsLib.dll

Class description:

pFileSystemObject
Objectified Scripting.FileSystemObject , encapsulates the main interfaces, such as copyfile, createfolder, deletefile, FileExists, folderexists, GetFile, getfolder, MoveFile, opentextfile, etc.

pFile
Objectified Scripting.File , which is used for file operation and encapsulates some interfaces, such as datecreated, datelastmodified, delete, path, shortname, shortpath, size, etc.

pFolder
Objectified Scripting.Folder , which is used for directory operation and encapsulates some interfaces, such as datecreated, datelastmodified, files, isrootfolder, parentfolder, path, shortname, shortpath, size, subfolders, etc.

pTextStream
Objectified Scripting.TextStream , which is used for reading and writing text files and encapsulates some interfaces, such as atendoffline, atendofstream, readall, readLine, writeblanklines, writeline, writestr, etc.
