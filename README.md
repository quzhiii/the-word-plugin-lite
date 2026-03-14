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
- `customUI/customUI14.xml` - Ribbon definition
- `package/` - final distributable `THU-Formatter-Lite.dotm`
- `INSTALL.ps1` / `UNINSTALL.ps1` - install and uninstall scripts
- `VALIDATE_PACKAGE.ps1` - validate packaged `.dotm`
- `PACK_RELEASE.ps1` - build `THU-Formatter-Lite.zip` from the validated package
- `docs/` - packaging, testing, release, and acceptance docs

## Package Workflow
1. Build `THU-Formatter-Lite.dotm` from the VBA source.
2. Inject `customUI/customUI14.xml` with RibbonX Editor.
3. Place the final template into `package/THU-Formatter-Lite.dotm`.
4. Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\VALIDATE_PACKAGE.ps1
```

## Install
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

## Recommended First Test
Use a fresh local Windows account instead of your main account. See:

- `docs/NEW_USER_ACCOUNT_TEST_CHECKLIST.zh-CN.md`
- `docs/EXPERT_ACCEPTANCE_CHECKLIST.zh-CN.md`
