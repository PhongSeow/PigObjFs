# PigObjFs
#### [English Readme](https://github.com/PhongSeow/PigObjFs/blob/master/README.md)
����һ�����󻯵� Microsoft Script Runtime������ Microsoft Script Runtime ͨ�� CreateObject ʹ�ã�������ֱ�����õ� VB6 �����С������������� EXE �ļ��Ϳ����ڲ�ͬ�� Windows ƽ̨�������ˡ�
ͬʱ����һ��.net 5.0�İ汾������������ Windows �� Linux ƽ̨��

## ***Ŀ¼���ļ�����***

### Setup

ִ����Ŀ¼������㲻�뿴��Դ���������ֱ��ʹ�����Ŀ¼�е��ļ���

##### Setup\VB6

��Ŀ¼�е� .cls �� .bas �ļ���ӵ����� VB6 ���̾Ϳ���ʹ�ã�PigObjFs.exe Ϊ������ʾ������

##### Setup\VB6DLL

�� PigObjFsLib.dll �ļ����õ����� VB6 ���̾Ϳ���ʹ�ã�����Ҫͨ�� regsrv32 ����ע�ᣬPigObjFsTestDll.exe Ϊ������ʾ������


##### Setup\DotNet

֧�ֶ��ֲ�ͬ�� .net ��ܣ���Ŀ¼�е����ļ���ӵ����� .net ���̾Ϳ���ʹ�ã����������� Windows��Linux �� MacOS ƽ̨��
����ʹ�� NuGet ���𣬼� [https://www.nuget.org/packages/PigObjFsLib/](https://www.nuget.org/packages/PigObjFsLib/)


### Src

Դ��Ŀ¼��

#### Src\VB6

VB6�汾��Դ���ʾ����

#### Src\VB6DLL

VB6 Active DLL �汾��Դ���ʾ����

#### Src\DotNet5.0

.net 5.0 �汾��Դ��Ϳ���̨����ʾ�������������������� Windows��Linux �� MacOS ƽ̨��

#### Src\DotNet

֧�ֶ��� .net �汾��Դ�룬���������������� Windows��Linux �� MacOS ƽ̨��

## ***��ؼ���˵��***

### ʲô���������

VB�����������ָͨ�� CreateObject ��ȡһ�� ActiveX ����Ȼ��ͨ������������ʹ�øö��������Ӧ����Ľӿڣ��﷨���磺

Set moFS = CreateObject("Scripting.FileSystemObject")��

### �������ʲô�ô�

�����������ʹ��VB�ĳ����и��õļ����ԣ�������������EXE�ļ�ֱ�ӿ��������ڼ���ȫ���汾�� Windows ƽ̨�����Ҳ���Ҫע�������

��Ϊ Windows �д���Ԥװ�� ActiveX �������Щ�������ע��Ϳ���ʹ�ã�����Щ����ڲ�ͬ�汾�� Windows ƽ̨�ϵİ汾������ͬ�����ܱ�֤��ȫ�Ķ����Ƽ��ݣ��������VB���������ã���ʹ���� Dim moFS As Scripting.FileSystemObject ���﷨���壬���ActiveX ��������Ʋ����ݣ��ͻᱨ��

�����ͨ�� CreateObject ��ʹ�ã���ʹ ActiveX ��������Ʋ����ݣ�ֻҪ�ӿڵĲ��������ͼ��ݣ��Ϳ�������ʹ�ã��������������VB���򣬼�������ѡ�

### ������������

���������Ҫ��ø��õļ����ԣ��� CreateObject ��õ���һ�������������VB��GUI�����������Ǹ������Ժͷ��������������Ѻá�

Ϊ�˽��������⣬����Ϊԭ�е��ഴ��һ��������������࣬���ཫԭ����Ľӿڷ�װ�ڡ�������������ϣ�����ʵ�ַ������Լ� [https://github.com/PhongSeow/PigObjFs/tree/master/Src/VB6](https://github.com/PhongSeow/PigObjFs/tree/master/Src/VB6)��

