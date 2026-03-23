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
- `VALIDATE_PACKAGE.ps1`：校验 dotm 包结构
- `PACK_RELEASE.ps1`：基于已校验的 dotm 生成发布 zip
- `docs/NEW_USER_ACCOUNT_TEST_CHECKLIST.zh-CN.md`：新用户隔离测试清单
- `docs/EXPERT_ACCEPTANCE_CHECKLIST.zh-CN.md`：专家验收清单
- `docs/PACKAGING_GUIDE.zh-CN.md`：dotm 与 zip 打包指南
- `docs/RELEASE_CHECKLIST.zh-CN.md`：发布清单

## 推荐使用流程
1. 先按 `docs/PACKAGING_GUIDE.zh-CN.md` 打包 `THU-Formatter-Lite.dotm`
2. 运行包校验：
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\VALIDATE_PACKAGE.ps1
   ```
3. 安装插件：
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\INSTALL.ps1
   ```
4. 重启 Word，确认出现 `THU Formatter`
5. 如需生成最终分发 zip：
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\PACK_RELEASE.ps1
   ```

## 宏安全提示
- 首次测试优先使用“新建本地用户账号”隔离环境
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
