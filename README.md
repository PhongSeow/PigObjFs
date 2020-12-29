# PigObjFs
#### [ÖÐÎÄÎÄµµ](https://github.com/PhongSeow/PigObjFs/blob/master/README.CN.md)

This is an object-oriented Microsoft script runtime, that is Microsoft script runtime is used by creating object instead of directly referring to VB6 project. In this way, the compiled EXE file can run on different windows platforms.

There is also a version of .Net 5.0, which can run on Windows and Linux platforms.

## ***Folders and files description***

### Setup

Execute code directory, if you don't want to see the source program, you can use the files in this directory directly in your VB.NET project.

##### Setup\VB6

Add the .cls and .bas files in the directory to your VB6 project, PigObjFs.exe For the compiled sample program.

##### Setup\VB6DLL

Will PigObjFsLib.dll File reference to your VB6 project can be used, but need to register through regsrv32 command, PigObjFsTestDll.exe For the compiled sample program.

##### Setup\DotNet5.0

Add the files in the directory to your. Net project, which can run on Windows, Linux and MacOS platforms.


### Src

Source code directory.

#### Src\VB6

VB6 version of the source code and examples.

#### Src\VB6DLL

VB6 Active DLL version of the source code and examples.

#### Src\DotNet5.0

The .net 5.0 version of the source code and console program examples, compiler can run on Windows, Linux and MacOS platform.

## ***Related technical description***

### What is without components

VB's not Objectified approach is to create an ActiveX object through CreateObject, and then use the interface of the class corresponding to the object variable through this object variable, For example:

Set moFS = CreateObject("Scripting.FileSystemObject")¡£

### What are the benefits of without components

The program that can use VB without component has better compatibility. Ideally, the EXE file compiled can run on almost all versions of Windows platform, and there is no need to register components.

Because Windows has a large number of pre installed ActiveX components, these components can be used without registration, but the versions of these components on different versions of Windows platform are not the same, and the complete binary compatibility can not be guaranteed. If you refer to them in VB project and use such as Dim moFS As Scripting.FileSystemObject If ActiveX If the component binary is not compatible, an error will be reported.

But if it is used by creating object, even if the ActiveX component binary is not compatible, as long as the parameters and types of the interface are compatible, it can be used normally. In this way, the compiled VB program has the best compatibility.


### How to handle without components

Component free needs to get better compatibility, but what CreateObject gets is an object variable, which is more attribute and method in VB GUI development environment, so the development is not friendly enough.

In order to solve this problem, we can create a "without components" class for the original class, which encapsulates the interface of the original class on the "componentless" class. The specific implementation method can be seen in [https://github.com/PhongSeow/PigObjFs/tree/master/Src/VB6](https://github.com/PhongSeow/PigObjFs/tree/master/Src/VB6).
