# ComClsIDInterefaceEnum

## 背景

看到篇blog，发现powershell能够枚举对应的clsid下的接口和接口方式，于是用c/c++实现了一下，比如用powershell很简单的就能得到对应接口的方法。

$shell = [Activator]::CreateInstance([type]::GetTypeFromCLSID("72C24DD5-D70A-438B-8A42-98424B88AFB8"))
$shell | gm

## 说明

这项目是枚举当前注册表HKEY_CLASSES_ROOT\Classes\CLSID下的所有的接口和方法名

对应到的方法是std::string EmuComInterfaceFuncs(std::wstring progID, std::wstring strClsid, BOOL isLocal)

第一个参数是ProgID第二个是CLSID两个参数二选一填写，isLocal代表该com对象是否是LocalServer32类型的。

## 为什么不用oleview

不是不用，而是oleview无法获取到接口名和接口提供的方式，所以才重新写了个工具

## 能做什么？

比如利用系统提供的接口创建进程、修改文件、注册表、计划任务什么的，进程树下关联不到你的程序。

## 参考

https://www.mandiant.com/resources/blog/hunting-com-objects