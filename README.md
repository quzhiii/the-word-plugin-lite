# THU Formatter Lite

Word Ribbon add-in for fast, accurate thesis format cleanup.

## What It Is
- A `.dotm`-based Word add-in for non-technical users
- One Ribbon tab: `THU Formatter`
- One main action: detect, fix, export, and log

## What It Fixes
- Pseudo headings (`1`, `1.1`, `1.1.1`) -> Heading styles
- Paragraph normalization for body text
- Caption style normalization and missing `SEQ` rebuild
- Nearby floating shapes -> inline conversion for safer numbering
- Three-line table formatting with nested-table skip strategy
- Field / TOC / list-of-figures refresh

## Repository Layout
- `src/vba/THU_Formatter_Addin.bas` - VBA add-in source
- `customUI/customUI.xml` - Ribbon definition
- `package/` - final distributable `THU-Formatter-Lite.dotm`
- `INSTALL.ps1` / `UNINSTALL.ps1` - install and uninstall scripts
- `VALIDATE_PACKAGE.ps1` - validate packaged `.dotm`
- `PACK_RELEASE.ps1` - build `THU-Formatter-Lite.zip` from the validated package
- `docs/` - packaging, testing, release, and acceptance docs

## Build Template
1. Prepare a base macro-enabled template once:
   - copy `%APPDATA%\Microsoft\Templates\Normal.dotm` to `package/THU-Formatter-Lite.base.dotm`
   - open that copy in Word
   - import `src/vba/THU_Formatter_Addin.bas`
   - save and close
2. Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\BUILD_TEMPLATE.ps1
```

This script:
- copies the base template into `package/THU-Formatter-Lite.dotm`
- injects `customUI/customUI.xml`
- checks the final package still contains `word/vbaProject.bin`

## Validate Package

```powershell
powershell -ExecutionPolicy Bypass -File .\VALIDATE_PACKAGE.ps1
```

## Install
```powershell
powershell -ExecutionPolicy Bypass -File .\BUILD_TEMPLATE.ps1
powershell -ExecutionPolicy Bypass -File .\INSTALL.ps1
```

Or install an already-built `package/THU-Formatter-Lite.dotm` directly:

```powershell
powershell -ExecutionPolicy Bypass -File .\INSTALL.ps1
```

## Build Release Zip
```powershell
powershell -ExecutionPolicy Bypass -File .\PACK_RELEASE.ps1
```

## Uninstall
```powershell
powershell -ExecutionPolicy Bypass -File .\UNINSTALL.ps1
```


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

## Recommended First Test
Use a fresh local Windows account instead of your main account. See:

- `docs/NEW_USER_ACCOUNT_TEST_CHECKLIST.zh-CN.md`
- `docs/EXPERT_ACCEPTANCE_CHECKLIST.zh-CN.md`
