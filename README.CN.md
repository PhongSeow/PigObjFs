# PigObjFs
#### [English Readme](https://github.com/PhongSeow/PigObjFs/blob/master/README.md)
这是一个对象化的 Microsoft Script Runtime，即是 Microsoft Script Runtime 通过 CreateObject 使用，而不是直接引用到 VB6 工程中。这样，编译后的 EXE 文件就可以在不同的 Windows 平台上运行了。
同时还有一个.net 5.0的版本，可以运行在 Windows 和 Linux 平台。

## ***目录和文件描述***

### Release

发布执行码目录，如果你不想看到源程序，你可以直接使用这个目录中的文件。

##### Release\DotNet

将目录中的类文件添加到您的 .net 工程或使用 NuGet ，可以运行在 Windows、Linux 和 MacOS 平台。

##### Release\VB6

将目录中的 .cls 和 .bas 文件添加到您的 VB6 工程就可以使用，PigObjFs.exe 为编译后的示例程序。


### Src

源码目录。

#### Src\DotNet

支持多种 .net 框架的源码和控制台程序示例，编译程序可以运行在 Windows、Linux 和 MacOS 平台。

#### Src\DotNet\PigObjFsLib

类库目录

#### Src\DotNet\PigObjFsDemo

调用示例目录

#### Src\VB6

VB6版本的源码和示例。


## ***相关技术说明***

### 什么是无组件化

VB的无组件化是指通过 CreateObject 获取一个 ActiveX 对象，然后通过这个对象变量使用该对象变量对应的类的接口，语法例如：

Set moFS = CreateObject("Scripting.FileSystemObject")。

### 无组件有什么好处

无组件化可以使用VB的程序有更好的兼容性，理想情况编译的EXE文件直接可以运行在几乎全部版本的 Windows 平台，而且不需要注册组件。

因为 Windows 有大量预装的 ActiveX 组件，这些组件不用注册就可以使用，但这些组件在不同版本的 Windows 平台上的版本并不相同，不能保证完全的二进制兼容，如果你在VB工程中引用，并使用如 Dim moFS As Scripting.FileSystemObject 的语法定义，如果ActiveX 组件二进制不兼容，就会报错。

但如果通过 CreateObject 来使用，即使 ActiveX 组件二进制不兼容，只要接口的参数和类型兼容，就可以正常使用，这样编译出来的VB程序，兼容性最佳。

### 如何无组件处理

无组件化需要获得更好的兼容性，但 CreateObject 获得的是一个对象变量，在VB的GUI开发环境下是更出属性和方法，开发不够友好。

为了解决这个问题，可以为原有的类创建一个“无组件化“类，该类将原有类的接口封装在“无组件化“类上，具体实现方法可以见 [https://github.com/PhongSeow/PigObjFs/tree/master/Src/VB6](https://github.com/PhongSeow/PigObjFs/tree/master/Src/VB6)。

