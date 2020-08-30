# PigObjFs
Piggy Objectified FileSystemObject, It can run directly in different versions of windows.

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
