Explain:
	This scheme is suitable for VB6, without reference to DLL modFsDeclare.bas , pFile.cls , pFileSystemObject.cls , pFolder.cls , PigLog.cls , pTextStream.cls These 6 files can be added to your VB6 project.

Example code:
	Open the VB6 project file, PigObjFs.vbp , sample code in frmPigObjFs.frm Form.
	PigObjFs.exe By PigObjFs.vbp You can test whether it can run directly in different windows.
	
Class description:
pFileSystemObject
	Objectified Scripting.FileSystemObject , encapsulates the main interfaces, such as copyfile, createfolder, deletefile, FileExists, folderexists, GetFile, getfolder, MoveFile, opentextfile, etc.

pFile
	Objectified Scripting.File , which is used for file operation and encapsulates some interfaces, such as datecreated, datelastmodified, delete, path, shortname, shortpath, size, etc.

pFolder
	Objectified Scripting.Folder , which is used for directory operation and encapsulates some interfaces, such as datecreated, datelastmodified, files, isrootfolder, parentfolder, path, shortname, shortpath, size, subfolders, etc.

pTextStream
	Objectified Scripting.TextStream , which is used for reading and writing text files and encapsulates some interfaces, such as atendoffline, atendofstream, readall, readLine, writeblanklines, writeline, writestr, etc.	
	