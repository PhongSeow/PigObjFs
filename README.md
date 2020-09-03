# PigObjFs
Piggy Objectified FileSystemObject, For programmers who are accustomed to using Microsoft script runtime, the PigObjFsLib.dll In VB6, VB.net And C ා using familiar interfaces, and PigObjFsLib.dll It can run on different versions of windows.

The functions include file attributes, existence, copy and deletion, folder attributes, existence, file collection and subfolder collection, and text file reading and writing operations.

VB6 compatibility

The EXE file compiled by VB6 can run on almost all windows platforms. Even if it is the new windows 2019, the premise is that DLL is not referenced in VB6 projects.

Microsoft Script Runtime

This is a component often used in VB6 development. The file name is scrrun.dll In general, windows platform will be pre installed, but different versions of Windows version scrrun.dll It is not binary compatible, so if this component is referenced in VB project, it can not be guaranteed to run on different versions of Windows platform.

PigObjFs

This is a project to objectify Microsoft script runtime, which means to use Microsoft script runtime with CreateObject instead of directly referring to VB project. In this way, the compiled exe files can be run on different windows platforms.

How to use PigObjFs

Will modFsDeclare.bas and pFile.cls , pFileSystemObject.cls , pFolder.cls , PigLog.cls , pTextStream.cls These files can be added to your VB6 project.

Open PigObjFs.vbp In the frmpigobjfs form, there will be call example code for each class.

More instructions are available https://en.seowphong.com/oss/PigObjFs/

Catalog description

VB6

This scheme is suitable for VB6, without reference to DLL modFsDeclare.bas , pFile.cls , pFileSystemObject.cls , pFolder.cls , PigLog.cls , pTextStream.cls These 6 files can be added to your VB6 project.

VB6DLL

This scheme is suitable for VB6, reference PigObjFsLib.dll The compiled exe can run directly on different versions of windows, but it needs to be registered through regsrv32.exe PigObjFsLib.dll .

VB.net

This scheme is applicable to VB.NET , reference PigObjFsLib.dll The compiled exe can run directly on different versions of windows.


