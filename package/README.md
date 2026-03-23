# Package Directory

Put the base and final template artifacts here:

- `THU-Formatter-Lite.base.dotm`
  - local build input
  - start from a copy of `Normal.dotm`
  - import `src/vba/THU_Formatter_Addin.bas` in Word once
- `THU-Formatter-Lite.dotm`
  - final distributable template produced by `BUILD_TEMPLATE.ps1`

`INSTALL.ps1` expects the final output filename to stay exactly `THU-Formatter-Lite.dotm`.