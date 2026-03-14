# A 档插件专家验收清单

面向专家评审与内测负责人，用于判断 A 档插件是否达到“可进入真实样本测试”的状态。

## 一、工程结构验收
- [ ] `src/vba/THU_Formatter_Addin.bas` 存在
- [ ] `customUI/customUI14.xml` 存在
- [ ] `INSTALL.ps1`、`UNINSTALL.ps1`、`VALIDATE_PACKAGE.ps1` 存在
- [ ] `package/THU-Formatter-Lite.dotm` 已生成
- [ ] `docs/NEW_USER_ACCOUNT_TEST_CHECKLIST.zh-CN.md` 存在

## 二、打包验收
- [ ] `VALIDATE_PACKAGE.ps1` 通过
- [ ] dotm 中包含 `word/vbaProject.bin`
- [ ] dotm 中包含 `customUI/customUI14.xml`
- [ ] Word 打开后出现 `THU Formatter` Ribbon 标签

## 三、安装与卸载验收
- [ ] `INSTALL.ps1` 可将 dotm 复制到当前用户 Startup
- [ ] 安装后重启 Word，可见插件按钮
- [ ] `UNINSTALL.ps1` 可移除插件
- [ ] 卸载后重启 Word，Ribbon 标签消失

## 四、功能回归验收
- [ ] 一键执行后生成 `*_fixed.docx`
- [ ] 一键执行后生成 `*_fixed.pdf`
- [ ] 一键执行后生成 `*_fix_log.txt`
- [ ] 伪标题转换正常（Heading1/2/3）
- [ ] 题注样式与 SEQ 自动修复正常
- [ ] 附近浮动图转 inline 统计正常
- [ ] 三线表规则执行正常
- [ ] 字段/目录/图表目录更新正常

## 五、风险控制验收
- [ ] 手工直改格式会在套样式前被清理
- [ ] 嵌套表不会被强行修复
- [ ] 浮动图转换失败不会中断主流程
- [ ] 宏安全文档可指导新手完成首次安装

## 六、进入下一阶段标准
- [ ] 3 个极端样本全部可跑通
- [ ] 无致命中断
- [ ] 误伤率可接受
- [ ] 可以进入 20-50 份真实论文样本回归
