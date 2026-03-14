# 新用户账号隔离测试清单（A 档插件）

目标：在不污染主账号 Word 配置的前提下，完成插件安装与回归验证。

## A. 测试前准备
1. 在 Windows 新建本地账号（示例：`word-test`）。
2. 注销主账号，切换到 `word-test`。
3. 在该账号准备测试文件夹，例如：`D:\word-plugin-test\`。
4. 放入：
   - `THU-Formatter-Lite.dotm`（打包产物）
   - 3 类测试文档（标点灾难/跨页长表/域代码污染）

## B. 插件安装
1. 打开 PowerShell（当前用户）。
2. 进入当前插件仓库根目录 `thu-word-plugin-lite`。
3. 执行：
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\INSTALL.ps1
   ```
4. 重启 Word。

## C. 宏安全解锁（首次）
1. 如果按钮不显示，右键 `.dotm` -> 属性 -> 勾选 `解除锁定(Unblock)`。
2. Word -> 选项 -> 信任中心 -> 信任中心设置：
   - 将模板所在目录加入“受信任位置”，或
   - 按组织策略允许已签名宏。
3. 重启 Word，确认出现 `THU Formatter` 标签页。

## D. 功能回归执行
对每份样本执行一次按钮：`一键检测并修复`，检查输出：
- `*_fixed.docx`
- `*_fixed.pdf`
- `*_fix_log.txt`

重点核验日志字段：
- `heading_fixed`
- `caption_fixed`
- `inline_fixed`
- `table_fixed`
- `table_skip`
- `table_warn`

## E. 结果判定
- 通过：流程不报错，输出齐全，核心规则生效。
- 待修：出现 `inline_warn/table_warn`，但主流程可继续。
- 失败：宏无法运行或导出失败。

## F. 回滚与清理（不留痕）
1. 运行：
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\UNINSTALL.ps1
   ```
2. 删除测试目录。
3. 如设置过受信任位置，移除测试路径。
4. 切回主账号，主账号配置不受影响。
