# A 档插件打包指南

本指南用于把源码打包为最终分发产物 `THU-Formatter-Lite.dotm` 和简易安装包 `THU-Formatter-Lite.zip`。

## 一、输入文件
- VBA 模块：`src/vba/THU_Formatter_Addin.bas`
- Ribbon XML：`customUI/customUI14.xml`
- 安装脚本：`INSTALL.ps1`
- 卸载脚本：`UNINSTALL.ps1`
- 包校验脚本：`VALIDATE_PACKAGE.ps1`

## 二、生成 dotm
1. 打开 Word，新建空白文档。
2. 另存为 `Word 启用宏的模板 (*.dotm)`，命名：`THU-Formatter-Lite.dotm`。
3. 按 `Alt + F11` 打开 VBA 编辑器，导入 `src/vba/THU_Formatter_Addin.bas`。
4. 保存并关闭 VBA 编辑器。
5. 使用 Office RibbonX Editor 打开该 dotm。
6. 插入 `Office 2010+ Custom UI Part`。
7. 写入 `customUI/customUI14.xml` 内容并保存。

## 三、放入 package 目录
将打包后的文件放入：
- `package/THU-Formatter-Lite.dotm`

## 四、执行包校验
```powershell
powershell -ExecutionPolicy Bypass -File .\VALIDATE_PACKAGE.ps1
```

通过标准：
- dotm 内包含 `word/vbaProject.bin`
- dotm 内包含 `customUI/customUI14.xml`

## 五、生成简易安装包 zip
执行：
```powershell
powershell -ExecutionPolicy Bypass -File .\PACK_RELEASE.ps1
```

生成内容包括：
- `THU-Formatter-Lite.dotm`
- `INSTALL.ps1`
- `UNINSTALL.ps1`
- `README.zh-CN.md`

输出文件名：
- `package/THU-Formatter-Lite.zip`

## 六、发布前自检
- 先在新建本地用户账号中测试安装
- 再按专家验收清单跑一遍
