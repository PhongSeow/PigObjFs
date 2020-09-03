VB.net
This scheme is applicable to VB.net , reference PigObjFsLib.dll The compiled exe can run directly on different versions of windows, but it needs to be registered by Regsvr32 command PigObjFsLib.dll .

Example code:
Open VB.net Engineering documents PigObjFs.sln In the example code frmPigObjFs.frm Form.
PigObjFs.exe By PigObjFs.sln The target framework is. Net framework 4.5. It can run directly on different versions of windows, but it needs to be registered through regsrv32.exe PigObjFsLib.dll , the command is: Regsvr32 PigObjFsLib.dll

Running instance
Will PigObjFs.exe , PigObjFsLib.dll And PigObjFsSetup.bat Copy to the same directory and run as an administrator on the command line PigObjFsSetup.bat Complete registration PigObjFsLib.dll And call PigObjFs.exe .

Class description:

pFileSystemObject
Objectified Scripting.FileSystemObject , encapsulates the main interfaces, such as copyfile, createfolder, deletefile, FileExists, folderexists, GetFile, getfolder, MoveFile, opentextfile, etc.

pFile
Objectified Scripting.File , which is used for file operation and encapsulates some interfaces, such as datecreated, datelastmodified, delete, path, shortname, shortpath, size, etc.

pFolder
Objectified Scripting.Folder , which is used for directory operation and encapsulates some interfaces, such as datecreated, datelastmodified, files, isrootfolder, parentfolder, path, shortname, shortpath, size, subfolders, etc.

pTextStream
Objectified Scripting.TextStream , which is used for reading and writing text files and encapsulates some interfaces, such as atendoffline, atendofstream, readall, readLine, writeblanklines, writeline, writestr, etc.
