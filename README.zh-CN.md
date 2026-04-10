# THU Formatter Lite（Word 插件版）

面向普通学生的一键式 Word 格式修复插件。目标是把复杂的 VBA/模板加载过程，封装成“安装 -> 打开 Word -> 点按钮”的体验。

## 它是什么
- 基于 `.dotm` 的 Word Ribbon 加载项
- Word 顶部出现 `THU Formatter` 标签页
- 一个核心按钮：`一键检测并修复`

## 它能做什么
- 伪标题 `1 / 1.1 / 1.1.1` 转 Heading 样式
- 正文段落归一
- 题注样式统一与缺失 `SEQ` 自动补齐
- 题注附近浮动图转 inline，降低编号错乱风险
- 三线表规则修复（含嵌套表跳过）
- 字段、目录、图表目录更新
- 导出 `*_fixed.docx`、`*_fixed.pdf`、`*_fix_log.txt`

## 仓库结构
- `src/vba/THU_Formatter_Addin.bas`：插件 VBA 主模块
- `customUI/customUI14.xml`：Ribbon 定义
- `package/`：最终分发产物目录
- `INSTALL.ps1`：安装到当前用户 Word Startup
- `UNINSTALL.ps1`：卸载插件
- `INSTALL.cmd` / `UNINSTALL.cmd`：给普通用户直接双击的一键入口
- `SYNC_VBA_TO_BASE.ps1`：把最新 VBA 源码同步回 `THU-Formatter-Lite.base.dotm`
- `VALIDATE_PACKAGE.ps1`：校验 dotm 包结构
- `PACK_RELEASE.ps1`：基于已校验的 dotm 生成发布 zip
- `scripts/Test-InstalledAddinParity.ps1`：校验 STARTUP 中已安装 add-in 是否与最新 package 完全一致
- `scripts/Test-StartupCompileSmoke.ps1`：通过 Word COM 调用 `AddinSelfCheck`，验证启动后可编译可执行
- `docs/NEW_USER_ACCOUNT_TEST_CHECKLIST.zh-CN.md`：新用户隔离测试清单
- `docs/EXPERT_ACCEPTANCE_CHECKLIST.zh-CN.md`：专家验收清单
- `docs/PACKAGING_GUIDE.zh-CN.md`：dotm 与 zip 打包指南
- `docs/RELEASE_CHECKLIST.zh-CN.md`：发布清单

## 推荐使用流程
1. 如修改过 `src/vba/THU_Formatter_Addin.bas`，先同步 VBA：
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\SYNC_VBA_TO_BASE.ps1
   ```
2. 再按 `docs/PACKAGING_GUIDE.zh-CN.md` 打包 `THU-Formatter-Lite.dotm`
3. 运行包校验：
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\VALIDATE_PACKAGE.ps1
   ```
4. 安装插件：
   方式 A，直接双击：
   - 双击 `INSTALL.cmd`
   - 看到“安装完成”提示后，重启 Word

   方式 B，命令行：
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\INSTALL.ps1
   ```
5. 安装后立即做一致性校验：
   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\Test-InstalledAddinParity.ps1
   ```
6. 做一次 Word 启动编译 smoke：
   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\Test-StartupCompileSmoke.ps1
   ```
7. 重启 Word，确认出现 `THU Formatter`
8. 如需生成最终分发 zip：
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\PACK_RELEASE.ps1
   ```

## 宏安全提示
- 首次测试优先使用“新建本地用户账号”隔离环境
- 对外发布时，优先把 `THU-Formatter-Lite.zip` 给用户；解压后直接双击 `INSTALL.cmd`
- 如果按钮不显示，优先检查：
  - 文件是否 `Unblock`
  - Word Trust Center 是否加入受信任位置

## 相关文档
- 打包指南：`docs/PACKAGING_GUIDE.zh-CN.md`
- 新手隔离测试：`docs/NEW_USER_ACCOUNT_TEST_CHECKLIST.zh-CN.md`
- 专家验收：`docs/EXPERT_ACCEPTANCE_CHECKLIST.zh-CN.md`
- 发布清单：`docs/RELEASE_CHECKLIST.zh-CN.md`

## Optional: Bridge to `thesis-format-engine`
If you have already installed the new Python/CLI engine, the button can call it first:

```powershell
setx THU_ENGINE_MODE cli
setx THU_ENGINE_CMD "thesis-engine"
setx THU_ENGINE_PROFILE "tsinghua-thesis"
setx THU_ENGINE_FIX_MODE "full"
```

Modes:
- `cli`: use only the new engine; fail fast if it errors
- `auto`: try the new engine first, then fall back to legacy VBA
- `legacy`: always use the current VBA pipeline

Fix modes:
- `safe`: default; only LOW-risk repairs
- `full`: includes MEDIUM-risk repairs such as three-line table normalization

Bridge smoke checks:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\Test-WordBridgeChinesePath.ps1
```
Release gate before a plugin release:

```powershell
D:\Program Files\python\python.exe -m pytest tests\unit\test_body_rules.py -q
D:\Program Files\python\python.exe -m pytest tests\integration\test_real_docx_body_fix.py -q
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\Test-InstalledAddinParity.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\Test-StartupCompileSmoke.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\Test-WordBridgeChinesePath.ps1
```

Deferred from this slice:

- `THU-B002`
- `THU-P001`
- `THU-R001`
- `customUI14.xml`
