Attribute VB_Name = "THU_Formatter_Addin"
Option Explicit

Private gRibbon As IRibbonUI

Private Const ENV_TEMPLATE_PATH As String = "THU_TEMPLATE_PATH"
Private Const ENV_ENGINE_MODE As String = "THU_ENGINE_MODE"
Private Const ENV_ENGINE_CMD As String = "THU_ENGINE_CMD"
Private Const ENV_ENGINE_PROFILE As String = "THU_ENGINE_PROFILE"
Private Const ENV_ENGINE_FIX_MODE As String = "THU_ENGINE_FIX_MODE"
Private Const ENGINE_MODE_AUTO As String = "auto"
Private Const ENGINE_MODE_LEGACY As String = "legacy"
Private Const ENGINE_MODE_CLI As String = "cli"
Private Const FIX_MODE_SAFE As String = "safe"
Private Const FIX_MODE_FULL As String = "full"
Private Const DEFAULT_ENGINE_FIX_MODE As String = FIX_MODE_SAFE
Private Const DEFAULT_ENGINE_PROFILE As String = "tsinghua-thesis"
Private Const BRIDGE_VERSION As String = "2026-03-24.4"
Private Const TABLE_OUTER_WIDTH As Long = wdLineWidth150pt
Private Const TABLE_INNER_WIDTH As Long = wdLineWidth050pt


' ===== Ribbon callbacks =====
Public Sub OnRibbonLoad(ByVal ribbon As IRibbonUI)
    Set gRibbon = ribbon
End Sub

Public Function GetRunButtonLabel(ByVal control As IRibbonControl) As String
    GetRunButtonLabel = "一键检测并修复"
End Function

Public Sub OnRunFormatter(ByVal control As IRibbonControl)
    OneClickDetectAndFix
End Sub

' ===== Main pipeline =====
Public Sub OneClickDetectAndFix()
    Dim srcPath As String
    Dim fixedDocx As String
    Dim fixedPdf As String
    Dim logPath As String
    Dim engineMode As String
    Dim engineCmd As String
    Dim engineProfile As String
    Dim engineFixMode As String
    Dim headingFixCount As Long
    Dim bodyFixCount As Long
    Dim tableFixCount As Long
    Dim tableSkipCount As Long
    Dim tableWarnCount As Long
    Dim captionFixCount As Long
    Dim captionWarnCount As Long
    Dim inlineFixCount As Long
    Dim inlineWarnCount As Long

    srcPath = ActiveDocument.FullName
    If Len(srcPath) = 0 Then
        MsgBox "请先保存文档后再运行格式修复。", vbExclamation
        Exit Sub
    End If

    fixedDocx = BuildOutputPath(srcPath, "_fixed.docx")
    fixedPdf = BuildOutputPath(srcPath, "_fixed.pdf")
    logPath = BuildOutputPath(srcPath, "_fix_log.txt")
    engineMode = ResolveEngineMode()
    engineCmd = ReadSettingValue(ENV_ENGINE_CMD)
    engineProfile = ResolveEngineProfile()
    engineFixMode = ResolveEngineFixMode()

    On Error GoTo Handler
    AppendLog logPath, "START: THU Formatter Add-in"
    AppendLog logPath, "VERSION: " & BRIDGE_VERSION
    AppendLog logPath, "MODE: engine_mode=" & engineMode & ", profile=" & engineProfile & ", fix_mode=" & engineFixMode

    If ShouldTryCli(engineMode, engineCmd) Then
        If RunEngineCliPipeline(srcPath, fixedDocx, fixedPdf, logPath, engineCmd, engineProfile, engineFixMode) Then
            ShowCliSuccess fixedDocx, fixedPdf, BuildOutputPath(srcPath, "_report.html"), BuildOutputPath(srcPath, "_report.txt"), logPath
            Exit Sub
        Else
            AppendLog logPath, "BRIDGE: CLI failed, fallback to legacy VBA pipeline"
        End If
    ElseIf engineMode = ENGINE_MODE_CLI Then
        Err.Raise vbObjectError + 2501, "THU Formatter", "THU_ENGINE_MODE=cli 但未配置 THU_ENGINE_CMD。"
    End If

    TemplateBinder logPath
    DocScanner logPath
    ApplyHeadingStyles logPath, headingFixCount
    NormalizeBody logPath, bodyFixCount
    EnsureCaptions logPath, captionFixCount, captionWarnCount, inlineFixCount, inlineWarnCount
    ApplyThreeLineTables logPath, tableFixCount, tableSkipCount, tableWarnCount
    RefreshFieldsAndToc logPath

    ActiveDocument.SaveAs2 FileName:=fixedDocx, FileFormat:=wdFormatXMLDocument
    ActiveDocument.ExportAsFixedFormat OutputFileName:=fixedPdf, ExportFormat:=wdExportFormatPDF

    AppendLog logPath, "SUMMARY: heading_fixed=" & CStr(headingFixCount) & _
        ", body_fixed=" & CStr(bodyFixCount) & _
        ", table_fixed=" & CStr(tableFixCount) & _
        ", table_skip=" & CStr(tableSkipCount) & _
        ", table_warn=" & CStr(tableWarnCount) & _
        ", caption_fixed=" & CStr(captionFixCount) & _
        ", caption_warn=" & CStr(captionWarnCount) & _
        ", inline_fixed=" & CStr(inlineFixCount) & _
        ", inline_warn=" & CStr(inlineWarnCount)
    AppendLog logPath, "DONE: exported docx/pdf"

    ShowLegacySuccess fixedDocx, fixedPdf, logPath, headingFixCount, bodyFixCount, tableFixCount, captionFixCount, inlineFixCount
    Exit Sub

Handler:
    AppendLog logPath, "ERROR: " & CStr(Err.Number) & " - " & Err.Description
    MsgBox "格式修复失败: " & Err.Description, vbExclamation
End Sub


Private Function BuildOutputPath(ByVal srcPath As String, ByVal suffix As String) As String
    Dim dotPos As Long
    dotPos = InStrRev(srcPath, ".")
    If dotPos > 0 Then
        BuildOutputPath = Left$(srcPath, dotPos - 1) & suffix
    Else
        BuildOutputPath = srcPath & suffix
    End If
End Function

Private Sub TemplateBinder(ByVal logPath As String)
    Dim templatePath As String
    templatePath = ReadSettingValue(ENV_TEMPLATE_PATH)

    If Len(templatePath) = 0 Then
        AppendLog logPath, "TEMPLATE: skip attach (env THU_TEMPLATE_PATH empty)"
        Exit Sub
    End If

    If Dir$(templatePath) = "" Then
        AppendLog logPath, "TEMPLATE: path not found: " & templatePath
        Exit Sub
    End If

    ActiveDocument.AttachedTemplate = templatePath
    ActiveDocument.UpdateStylesOnOpen = True
    AppendLog logPath, "TEMPLATE: attached " & templatePath
End Sub

Private Sub DocScanner(ByVal logPath As String)
    AppendLog logPath, "SCAN: paragraphs=" & CStr(ActiveDocument.Paragraphs.Count) & _
        ", tables=" & CStr(ActiveDocument.Tables.Count) & ", shapes=" & CStr(ActiveDocument.Shapes.Count)
End Sub

Private Sub ApplyHeadingStyles(ByVal logPath As String, ByRef fixCount As Long)
    Dim p As Paragraph
    Dim t As String
    Dim rgx1 As Object
    Dim rgx2 As Object
    Dim rgx3 As Object

    Set rgx1 = CreateObject("VBScript.RegExp")
    rgx1.Pattern = "^\s*[0-9]+\s+"
    Set rgx2 = CreateObject("VBScript.RegExp")
    rgx2.Pattern = "^\s*[0-9]+\.[0-9]+\s+"
    Set rgx3 = CreateObject("VBScript.RegExp")
    rgx3.Pattern = "^\s*[0-9]+\.[0-9]+\.[0-9]+\s+"

    fixCount = 0
    For Each p In ActiveDocument.Paragraphs
        t = NormalizeParagraphText(p.Range.Text)

        If rgx3.Test(t) Then
            If p.Range.Style <> wdStyleHeading3 Then
                ApplyParagraphStyleStrict p, wdStyleHeading3
                fixCount = fixCount + 1
            End If
        ElseIf rgx2.Test(t) Then
            If p.Range.Style <> wdStyleHeading2 Then
                ApplyParagraphStyleStrict p, wdStyleHeading2
                fixCount = fixCount + 1
            End If
        ElseIf rgx1.Test(t) Then
            If p.Range.Style <> wdStyleHeading1 Then
                ApplyParagraphStyleStrict p, wdStyleHeading1
                fixCount = fixCount + 1
            End If
        End If
    Next p

    AppendLog logPath, "HEADING: converted=" & CStr(fixCount)
End Sub

Private Function NormalizeParagraphText(ByVal s As String) As String
    Dim t As String
    t = Replace$(s, vbCr, "")
    t = Replace$(t, Chr$(7), "")
    NormalizeParagraphText = Trim$(t)
End Function

Private Sub ClearManualOverrides(ByVal targetRange As Range)
    On Error Resume Next
    targetRange.Font.Reset
    targetRange.ParagraphFormat.Reset
    On Error GoTo 0
End Sub

Private Sub ApplyParagraphStyleStrict(ByVal p As Paragraph, ByVal styleValue As Variant)
    ClearManualOverrides p.Range
    p.Range.Style = styleValue
End Sub

Private Sub NormalizeBody(ByVal logPath As String, ByRef fixCount As Long)
    Dim p As Paragraph
    fixCount = 0

    For Each p In ActiveDocument.Paragraphs
        If p.Range.Information(wdWithInTable) Then
            GoTo ContinueLoop
        End If

        If p.Range.Style = wdStyleNormal Then
            With p.Format
                .LineSpacingRule = wdLineSpaceMultiple
                .LineSpacing = 24
                .FirstLineIndent = CentimetersToPoints(0.74)
                .SpaceBefore = 0
                .SpaceAfter = 0
            End With
            fixCount = fixCount + 1
        End If
ContinueLoop:
    Next p

    AppendLog logPath, "BODY: normalized=" & CStr(fixCount)
End Sub

Private Sub EnsureCaptions(ByVal logPath As String, ByRef fixCount As Long, ByRef warnCount As Long, ByRef inlineFixCount As Long, ByRef inlineWarnCount As Long)
    Dim p As Paragraph
    Dim t As String
    Dim kindText As String
    Dim titleText As String
    Dim bodyRange As Range
    Dim insertRange As Range

    fixCount = 0
    warnCount = 0
    inlineFixCount = 0
    inlineWarnCount = 0

    For Each p In ActiveDocument.Paragraphs
        t = NormalizeParagraphText(p.Range.Text)

        If p.Range.Information(wdWithInTable) Then
            GoTo ContinueCaptionLoop
        End If

        kindText = DetectCaptionKind(t)
        If Len(kindText) > 0 Then
            If Not TryApplyCaptionStyle(p) Then
                warnCount = warnCount + 1
            End If

            ConvertAnchoredShapesNearCaption p, inlineFixCount, inlineWarnCount

            If Not HasSequenceField(p.Range, kindText) Then
                titleText = ExtractCaptionTitle(t)

                Set bodyRange = p.Range.Duplicate
                bodyRange.End = bodyRange.End - 1
                bodyRange.Text = kindText & " "

                Set insertRange = bodyRange.Duplicate
                insertRange.Collapse wdCollapseEnd
                insertRange.Fields.Add insertRange, wdFieldSequence, kindText
                insertRange.InsertAfter " " & titleText

                fixCount = fixCount + 1
            End If
        End If

ContinueCaptionLoop:
    Next p

    AppendLog logPath, "CAPTION: fixed=" & CStr(fixCount) & ", warning_candidates=" & CStr(warnCount) & _
        ", inline_fixed=" & CStr(inlineFixCount) & ", inline_warn=" & CStr(inlineWarnCount)
End Sub

Private Function TryApplyCaptionStyle(ByVal p As Paragraph) As Boolean
    ClearManualOverrides p.Range

    On Error GoTo Fallback
    p.Range.Style = wdStyleCaption
    TryApplyCaptionStyle = True
    Exit Function

Fallback:
    On Error GoTo Failed
    p.Range.Style = "Caption"
    TryApplyCaptionStyle = True
    Exit Function

Failed:
    TryApplyCaptionStyle = False
End Function

Private Sub ConvertAnchoredShapesNearCaption(ByVal captionParagraph As Paragraph, ByRef fixCount As Long, ByRef warnCount As Long)
    Dim i As Long
    Dim shp As Shape
    Dim captionPos As Long
    Dim anchorPos As Long

    captionPos = captionParagraph.Range.Start

    For i = ActiveDocument.Shapes.Count To 1 Step -1
        Set shp = ActiveDocument.Shapes(i)
        anchorPos = shp.Anchor.Start

        If Abs(anchorPos - captionPos) <= 500 Then
            On Error GoTo ConvertFailed
            shp.ConvertToInlineShape
            fixCount = fixCount + 1
            On Error GoTo 0
        End If
ContinueShapeLoop:
    Next i
    Exit Sub

ConvertFailed:
    warnCount = warnCount + 1
    On Error GoTo 0
    Resume ContinueShapeLoop
End Sub

Private Function DetectCaptionKind(ByVal textLine As String) As String
    Dim rgx As Object

    Set rgx = CreateObject("VBScript.RegExp")
    rgx.IgnoreCase = False
    rgx.Global = False
    rgx.Pattern = "^\s*图\s*[0-9]+([-.][0-9]+)*\s+"
    If rgx.Test(textLine) Then
        DetectCaptionKind = "图"
        Exit Function
    End If

    rgx.Pattern = "^\s*表\s*[0-9]+([-.][0-9]+)*\s+"
    If rgx.Test(textLine) Then
        DetectCaptionKind = "表"
        Exit Function
    End If

    DetectCaptionKind = ""
End Function

Private Function ExtractCaptionTitle(ByVal textLine As String) As String
    Dim rgx As Object
    Dim resultText As String

    Set rgx = CreateObject("VBScript.RegExp")
    rgx.IgnoreCase = False
    rgx.Global = False
    rgx.Pattern = "^\s*[图表]\s*[0-9]+([-.][0-9]+)*\s*"

    resultText = rgx.Replace(textLine, "")
    resultText = Trim$(resultText)
    If Len(resultText) = 0 Then
        resultText = "题注"
    End If

    ExtractCaptionTitle = resultText
End Function

Private Function HasSequenceField(ByVal targetRange As Range, ByVal captionKind As String) As Boolean
    Dim f As Field
    Dim codeText As String

    For Each f In targetRange.Fields
        If f.Type = wdFieldSequence Then
            codeText = LCase$(f.Code.Text)
            If InStr(1, codeText, LCase$(captionKind), vbTextCompare) > 0 Then
                HasSequenceField = True
                Exit Function
            End If
        End If
    Next f

    HasSequenceField = False
End Function

Private Sub ApplyThreeLineTables(ByVal logPath As String, ByRef fixCount As Long, ByRef skipCount As Long, ByRef warnCount As Long)
    Dim tbl As Table
    Dim rowIndex As Long

    fixCount = 0
    skipCount = 0
    warnCount = 0

    For Each tbl In ActiveDocument.Tables
        If ShouldSkipTable(tbl) Then
            skipCount = skipCount + 1
            GoTo ContinueTable
        End If

        On Error GoTo TableFailed

        tbl.Borders.Enable = False

        For rowIndex = 1 To tbl.Rows.Count
            tbl.Rows(rowIndex).Borders(wdBorderTop).LineStyle = wdLineStyleNone
            tbl.Rows(rowIndex).Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        Next rowIndex

        ApplyHorizontalBorder tbl.Rows(1).Borders(wdBorderTop), wdLineStyleSingle, TABLE_OUTER_WIDTH
        ApplyHorizontalBorder tbl.Rows(1).Borders(wdBorderBottom), wdLineStyleSingle, TABLE_INNER_WIDTH
        tbl.Rows(1).HeadingFormat = True
        ApplyHorizontalBorder tbl.Rows(tbl.Rows.Count).Borders(wdBorderBottom), wdLineStyleSingle, TABLE_OUTER_WIDTH

        fixCount = fixCount + 1
        On Error GoTo 0
        GoTo ContinueTable

TableFailed:
        warnCount = warnCount + 1
        On Error GoTo 0
ContinueTable:
    Next tbl

    AppendLog logPath, "TABLE: three_line_applied=" & CStr(fixCount) & _
        ", skipped=" & CStr(skipCount) & ", warnings=" & CStr(warnCount)
End Sub

Private Sub ApplyHorizontalBorder(ByVal borderObj As Border, ByVal lineStyle As WdLineStyle, ByVal lineWidth As WdLineWidth)
    borderObj.LineStyle = lineStyle
    borderObj.LineWidth = lineWidth
End Sub

Private Function ShouldSkipTable(ByVal tbl As Table) As Boolean
    If tbl.Rows.Count = 0 Then
        ShouldSkipTable = True
        Exit Function
    End If

    If tbl.NestingLevel > 1 Then
        ShouldSkipTable = True
        Exit Function
    End If

    ShouldSkipTable = False
End Function

Private Sub RefreshFieldsAndToc(ByVal logPath As String)
    RefreshFieldsAndTocForDocument ActiveDocument, logPath
End Sub

Private Sub RefreshFieldsAndTocForDocument(ByVal targetDoc As Document, ByVal logPath As String)
    Dim toc As TableOfContents
    Dim tof As TableOfFigures

    targetDoc.Fields.Update

    For Each toc In targetDoc.TablesOfContents
        toc.Update
    Next toc

    For Each tof In targetDoc.TablesOfFigures
        tof.Update
    Next tof

    AppendLog logPath, "FIELD: fields/toc/tof updated"
End Sub

Private Function ResolveEngineMode() As String
    Dim modeText As String

    modeText = LCase$(ReadSettingValue(ENV_ENGINE_MODE))
    Select Case modeText
        Case ENGINE_MODE_LEGACY
            ResolveEngineMode = ENGINE_MODE_LEGACY
        Case ENGINE_MODE_CLI
            ResolveEngineMode = ENGINE_MODE_CLI
        Case Else
            ResolveEngineMode = ENGINE_MODE_AUTO
    End Select
End Function

Private Function ResolveEngineProfile() As String
    Dim profileText As String

    profileText = ReadSettingValue(ENV_ENGINE_PROFILE)
    If Len(profileText) = 0 Then
        ResolveEngineProfile = DEFAULT_ENGINE_PROFILE
    Else
        ResolveEngineProfile = profileText
    End If
End Function

Private Function ResolveEngineFixMode() As String
    Dim modeText As String

    modeText = LCase$(ReadSettingValue(ENV_ENGINE_FIX_MODE))
    Select Case modeText
        Case FIX_MODE_FULL
            ResolveEngineFixMode = FIX_MODE_FULL
        Case FIX_MODE_SAFE
            ResolveEngineFixMode = FIX_MODE_SAFE
        Case Else
            ResolveEngineFixMode = DEFAULT_ENGINE_FIX_MODE
    End Select
End Function

Private Function ShouldTryCli(ByVal engineMode As String, ByVal engineCmd As String) As Boolean
    If engineMode = ENGINE_MODE_CLI Then
        ShouldTryCli = True
    ElseIf engineMode = ENGINE_MODE_AUTO And Len(Trim$(engineCmd)) > 0 Then
        ShouldTryCli = True
    Else
        ShouldTryCli = False
    End If
End Function

Private Function RunEngineCliPipeline(ByVal srcPath As String, ByVal fixedDocx As String, ByVal fixedPdf As String, ByVal logPath As String, ByVal engineCmd As String, ByVal engineProfile As String, ByVal engineFixMode As String) As Boolean
    Dim outputDir As String
    Dim reportTextPath As String
    Dim reportJsonPath As String
    Dim reportHtmlPath As String
    Dim cliInputPath As String
    Dim commandText As String
    Dim exitCode As Long

    RunEngineCliPipeline = False

    If Len(Trim$(engineCmd)) = 0 Then
        AppendLog logPath, "BRIDGE: THU_ENGINE_CMD empty"
        Exit Function
    End If

    outputDir = GetParentFolderPath(srcPath)
    reportTextPath = BuildOutputPath(srcPath, "_report.txt")
    reportJsonPath = BuildOutputPath(srcPath, "_report.json")
    reportHtmlPath = BuildOutputPath(srcPath, "_report.html")
    cliInputPath = CreateCliInputSnapshot(srcPath, logPath)

    commandText = NormalizeCommandExecutable(engineCmd) & " fix " & QuoteArg(cliInputPath) & _
        " --profile " & QuoteArg(engineProfile) & _
        " --mode " & QuoteArg(engineFixMode) & " --out " & QuoteArg(outputDir)

    AppendLog logPath, "BRIDGE: start thesis-format-engine " & engineFixMode & " fix"
    AppendLog logPath, "BRIDGE: command=" & engineCmd & " fix <docx> --profile " & engineProfile & " --mode " & engineFixMode & " --out " & outputDir
    AppendLog logPath, "BRIDGE: cli_input=" & cliInputPath
    exitCode = RunCommandCaptureToLog(commandText, logPath)
    AppendLog logPath, "BRIDGE: exit_code=" & CStr(exitCode)

    If exitCode <> 0 Then
        Exit Function
    End If

    If Dir$(fixedDocx) = "" Then
        AppendLog logPath, "BRIDGE: fixed docx missing: " & fixedDocx
        Exit Function
    End If

    If Dir$(reportTextPath) <> "" Then
        AppendExistingTextFile logPath, reportTextPath, "ENGINE_REPORT"
    End If

    ExportPdfFromFixedDocx fixedDocx, fixedPdf, logPath
    AppendLog logPath, "BRIDGE: artifacts docx=" & fixedDocx & ", pdf=" & fixedPdf & ", report_json=" & reportJsonPath & ", report_html=" & reportHtmlPath & ", report_text=" & reportTextPath
    RunEngineCliPipeline = True
End Function

Private Sub ExportPdfFromFixedDocx(ByVal fixedDocx As String, ByVal fixedPdf As String, ByVal logPath As String)
    Dim resultDoc As Document

    On Error GoTo CleanFail
    Set resultDoc = Application.Documents.Open(FileName:=fixedDocx, AddToRecentFiles:=False, ReadOnly:=False, Visible:=False)
    RefreshFieldsAndTocForDocument resultDoc, logPath
    resultDoc.Save
    resultDoc.ExportAsFixedFormat OutputFileName:=fixedPdf, ExportFormat:=wdExportFormatPDF
    resultDoc.Close SaveChanges:=wdSaveChanges
    AppendLog logPath, "BRIDGE: exported pdf " & fixedPdf
    Exit Sub

CleanFail:
    On Error Resume Next
    If Not resultDoc Is Nothing Then
        resultDoc.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function RunCommandCaptureToLog(ByVal commandText As String, ByVal logPath As String) As Long
    Dim shellObj As Object
    Dim execObj As Object
    Dim stdoutText As String
    Dim stderrText As String

    Set shellObj = CreateObject("WScript.Shell")
    Set execObj = shellObj.Exec(commandText)

    Do While execObj.Status = 0
        DoEvents
    Loop

    stdoutText = execObj.StdOut.ReadAll
    stderrText = execObj.StdErr.ReadAll

    If Len(stdoutText) > 0 Then
        AppendCommandOutput logPath, stdoutText
    End If
    If Len(stderrText) > 0 Then
        AppendCommandOutput logPath, stderrText
    End If

    RunCommandCaptureToLog = execObj.ExitCode
End Function

Private Function QuoteArg(ByVal rawText As String) As String
    QuoteArg = Chr$(34) & Replace$(rawText, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Private Function ReadSettingValue(ByVal envName As String) As String
    Dim shellObj As Object
    Dim valueText As String

    On Error Resume Next
    Set shellObj = CreateObject("WScript.Shell")
    valueText = Trim$(CStr(shellObj.RegRead("HKEY_CURRENT_USER\Environment\" & envName)))
    On Error GoTo 0

    If Len(valueText) > 0 Then
        ReadSettingValue = valueText
    Else
        ReadSettingValue = Trim$(Environ$(envName))
    End If
End Function

Private Function ReadTextFileSafe(ByVal filePath As String) As String
    Dim ff As Integer
    Dim lineText As String
    Dim buffer As String

    If Dir$(filePath) = "" Then
        ReadTextFileSafe = ""
        Exit Function
    End If

    On Error GoTo ReadFailed
    ff = FreeFile
    Open filePath For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, lineText
        buffer = buffer & lineText & vbCrLf
    Loop
    Close #ff
    ReadTextFileSafe = buffer
    Exit Function

ReadFailed:
    On Error Resume Next
    If ff > 0 Then
        Close #ff
    End If
    On Error GoTo 0
    ReadTextFileSafe = ""
End Function

Private Function ExtractIntByRegex(ByVal sourceText As String, ByVal patternText As String, ByVal defaultValue As Long) As Long
    Dim rgx As Object
    Dim matches As Object

    If Len(sourceText) = 0 Then
        ExtractIntByRegex = defaultValue
        Exit Function
    End If

    Set rgx = CreateObject("VBScript.RegExp")
    rgx.Pattern = patternText
    rgx.Global = False
    rgx.IgnoreCase = True
    rgx.MultiLine = True

    If rgx.Test(sourceText) Then
        Set matches = rgx.Execute(sourceText)
        ExtractIntByRegex = CLng(matches(0).SubMatches(0))
    Else
        ExtractIntByRegex = defaultValue
    End If
End Function

Private Function ExtractBulletSection(ByVal sourceText As String, ByVal headingText As String, ByVal maxItems As Long) As String
    Dim normalizedText As String
    Dim lines() As String
    Dim currentLine As String
    Dim resultText As String
    Dim foundHeading As Boolean
    Dim itemCount As Long
    Dim i As Long

    If Len(sourceText) = 0 Then
        Exit Function
    End If

    normalizedText = Replace$(sourceText, vbCrLf, vbLf)
    normalizedText = Replace$(normalizedText, vbCr, vbLf)
    lines = Split(normalizedText, vbLf)

    For i = LBound(lines) To UBound(lines)
        currentLine = Trim$(lines(i))
        If Not foundHeading Then
            If StrComp(currentLine, headingText, vbTextCompare) = 0 Then
                foundHeading = True
            End If
        Else
            If Len(currentLine) = 0 Then
                Exit For
            End If
            If Left$(currentLine, 2) <> "- " Then
                Exit For
            End If
            If Len(resultText) > 0 Then
                resultText = resultText & vbCrLf
            End If
            resultText = resultText & currentLine
            itemCount = itemCount + 1
            If itemCount >= maxItems Then
                Exit For
            End If
        End If
    Next i

    ExtractBulletSection = resultText
End Function

Private Function NormalizeCommandExecutable(ByVal rawCommand As String) As String
    Dim trimmedCommand As String

    trimmedCommand = Trim$(rawCommand)
    If Len(trimmedCommand) = 0 Then
        NormalizeCommandExecutable = trimmedCommand
        Exit Function
    End If

    If Left$(trimmedCommand, 1) = Chr$(34) Then
        NormalizeCommandExecutable = trimmedCommand
        Exit Function
    End If

    If InStr(trimmedCommand, " ") > 0 Then
        If Dir$(trimmedCommand) <> "" Then
            NormalizeCommandExecutable = QuoteArg(trimmedCommand)
            Exit Function
        End If
    End If

    NormalizeCommandExecutable = trimmedCommand
End Function

Private Function CreateCliInputSnapshot(ByVal srcPath As String, ByVal logPath As String) As String
    Dim tempRoot As String
    Dim tempDir As String
    Dim snapshotPath As String

    tempRoot = Trim$(Environ$("TEMP"))
    If Len(tempRoot) = 0 Then
        tempRoot = CurDir$
    End If

    tempDir = JoinPathText(tempRoot, "THU-Formatter-CLI")
    EnsureSingleFolder tempDir
    snapshotPath = JoinPathText(tempDir, GetFileNamePart(srcPath))

    On Error Resume Next
    If Len(ActiveDocument.Path) > 0 Then
        ActiveDocument.Save
        If Err.Number <> 0 Then
            AppendLog logPath, "BRIDGE: source_save_warning=" & CStr(Err.Number) & " - " & Err.Description
            Err.Clear
        End If
    End If
    On Error GoTo SnapshotFailed
    If Dir$(snapshotPath) <> "" Then
        Kill snapshotPath
    End If
    ActiveDocument.SaveCopyAs FileName:=snapshotPath
    CreateCliInputSnapshot = snapshotPath
    Exit Function

SnapshotFailed:
    AppendLog logPath, "BRIDGE: cli_snapshot_savecopyas_failed=" & CStr(Err.Number) & " - " & Err.Description
    Err.Clear
    CreateCliInputSnapshot = CreateCliInputSnapshotFromClone(srcPath, snapshotPath, logPath)
    If Len(CreateCliInputSnapshot) = 0 Then
        CreateCliInputSnapshot = srcPath
    End If
End Function

Private Function CreateCliInputSnapshotFromClone(ByVal srcPath As String, ByVal snapshotPath As String, ByVal logPath As String) As String
    Dim tempDoc As Document

    On Error GoTo CloneFailed
    Set tempDoc = Application.Documents.Add(Visible:=False)
    tempDoc.Range.FormattedText = ActiveDocument.Range.FormattedText
    tempDoc.SaveAs2 FileName:=snapshotPath, FileFormat:=wdFormatXMLDocument, AddToRecentFiles:=False
    tempDoc.Close SaveChanges:=wdDoNotSaveChanges
    Set tempDoc = Nothing

    AppendLog logPath, "BRIDGE: cli_snapshot_clone_ok=" & snapshotPath
    CreateCliInputSnapshotFromClone = snapshotPath
    Exit Function

CloneFailed:
    AppendLog logPath, "BRIDGE: cli_snapshot_clone_failed=" & CStr(Err.Number) & " - " & Err.Description
    On Error Resume Next
    If Not tempDoc Is Nothing Then
        tempDoc.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
    CreateCliInputSnapshotFromClone = ""
End Function

Private Sub EnsureSingleFolder(ByVal folderPath As String)
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then
        MkDir folderPath
    End If
End Sub

Private Sub ShowCliSuccess(ByVal fixedDocx As String, ByVal fixedPdf As String, ByVal reportHtmlPath As String, ByVal reportTextPath As String, ByVal logPath As String)
    MsgBox BuildCliSuccessMessage(reportTextPath, fixedDocx, fixedPdf), vbInformation
    OpenCliArtifacts fixedDocx, reportHtmlPath, logPath
End Sub

Private Sub ShowLegacySuccess(ByVal fixedDocx As String, ByVal fixedPdf As String, ByVal logPath As String, ByVal headingFixCount As Long, ByVal bodyFixCount As Long, ByVal tableFixCount As Long, ByVal captionFixCount As Long, ByVal inlineFixCount As Long)
    MsgBox BuildLegacySuccessMessage(fixedDocx, fixedPdf, headingFixCount, bodyFixCount, tableFixCount, captionFixCount, inlineFixCount), vbInformation
    OpenLegacyArtifacts fixedPdf, logPath
End Sub

Private Function BuildCliSuccessMessage(ByVal reportTextPath As String, ByVal fixedDocx As String, ByVal fixedPdf As String) As String
    Dim findingsCount As Long
    Dim fullFixCount As Long
    Dim partialFixCount As Long
    Dim failedFixCount As Long
    Dim manualReviewCount As Long
    Dim reportText As String
    Dim changeSummaryText As String
    Dim changeSection As String

    reportText = ReadTextFileSafe(reportTextPath)
    findingsCount = ExtractIntByRegex(reportText, "Findings:\s+([0-9]+)", 0)
    fullFixCount = ExtractIntByRegex(reportText, "Fix detail:\s+full=([0-9]+)", 0)
    partialFixCount = ExtractIntByRegex(reportText, "Fix detail:\s+full=[0-9]+\s+partial=([0-9]+)", 0)
    failedFixCount = ExtractIntByRegex(reportText, "Fix detail:\s+full=[0-9]+\s+partial=[0-9]+\s+failed=([0-9]+)", 0)
    changeSummaryText = ExtractBulletSection(reportText, "Main changes:", 3)

    manualReviewCount = findingsCount - fullFixCount
    If manualReviewCount < 0 Then
        manualReviewCount = partialFixCount + failedFixCount
    End If

    If Len(changeSummaryText) > 0 Then
        changeSection = "本次主要变化:" & vbCrLf & changeSummaryText & vbCrLf & vbCrLf
    End If

    BuildCliSuccessMessage = "THU Formatter 已完成（thesis-format-engine）。" & vbCrLf & vbCrLf & _
        "发现问题: " & CStr(findingsCount) & " 项" & vbCrLf & _
        "已完全修复: " & CStr(fullFixCount) & " 项" & vbCrLf & _
        "仍需人工确认: " & CStr(manualReviewCount) & " 项" & vbCrLf & _
        "修复失败: " & CStr(failedFixCount) & " 项" & vbCrLf & vbCrLf & _
        changeSection & _
        "点击“确定”后将自动打开修复后文档和可视化报告。" & vbCrLf & vbCrLf & _
        "输出 DOCX: " & fixedDocx & vbCrLf & _
        "输出 PDF: " & fixedPdf
End Function

Private Function BuildLegacySuccessMessage(ByVal fixedDocx As String, ByVal fixedPdf As String, ByVal headingFixCount As Long, ByVal bodyFixCount As Long, ByVal tableFixCount As Long, ByVal captionFixCount As Long, ByVal inlineFixCount As Long) As String
    Dim totalFixed As Long

    totalFixed = headingFixCount + bodyFixCount + tableFixCount + captionFixCount + inlineFixCount

    BuildLegacySuccessMessage = "THU Formatter 已完成（Word 内置修复模式）。" & vbCrLf & vbCrLf & _
        "本次可见修复: " & CStr(totalFixed) & " 处" & vbCrLf & _
        "标题: " & CStr(headingFixCount) & "，正文: " & CStr(bodyFixCount) & "，表格: " & CStr(tableFixCount) & vbCrLf & _
        "题注: " & CStr(captionFixCount) & "，图形转行内: " & CStr(inlineFixCount) & vbCrLf & vbCrLf & _
        "输出 DOCX: " & fixedDocx & vbCrLf & _
        "输出 PDF: " & fixedPdf
End Function

Private Sub OpenCliArtifacts(ByVal fixedDocx As String, ByVal reportHtmlPath As String, ByVal logPath As String)
    On Error Resume Next

    If Dir$(fixedDocx) <> "" Then
        Application.Documents.Open FileName:=fixedDocx, AddToRecentFiles:=False, ReadOnly:=False, Visible:=True
    End If
    If Dir$(reportHtmlPath) <> "" Then
        CreateObject("Shell.Application").Open reportHtmlPath
    End If

    If Err.Number <> 0 Then
        AppendLog logPath, "BRIDGE: open_artifacts_warning=" & CStr(Err.Number) & " - " & Err.Description
        Err.Clear
    End If

    On Error GoTo 0
End Sub

Private Sub OpenLegacyArtifacts(ByVal fixedPdf As String, ByVal logPath As String)
    On Error Resume Next
    If Dir$(fixedPdf) <> "" Then
        CreateObject("Shell.Application").Open fixedPdf
    End If
    If Err.Number <> 0 Then
        AppendLog logPath, "BRIDGE: open_legacy_artifacts_warning=" & CStr(Err.Number) & " - " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Function JoinPathText(ByVal leftPath As String, ByVal rightPath As String) As String
    If Right$(leftPath, 1) = Application.PathSeparator Then
        JoinPathText = leftPath & rightPath
    Else
        JoinPathText = leftPath & Application.PathSeparator & rightPath
    End If
End Function

Private Function GetFileNamePart(ByVal filePath As String) As String
    Dim slashPos As Long

    slashPos = InStrRev(filePath, Application.PathSeparator)
    If slashPos > 0 Then
        GetFileNamePart = Mid$(filePath, slashPos + 1)
    Else
        GetFileNamePart = filePath
    End If
End Function

Private Function GetParentFolderPath(ByVal filePath As String) As String
    Dim slashPos As Long

    slashPos = InStrRev(filePath, Application.PathSeparator)
    If slashPos > 0 Then
        GetParentFolderPath = Left$(filePath, slashPos - 1)
    Else
        GetParentFolderPath = CurDir$
    End If
End Function

Private Sub AppendExistingTextFile(ByVal targetPath As String, ByVal sourcePath As String, ByVal sectionName As String)
    Dim ffIn As Integer
    Dim ffOut As Integer
    Dim lineText As String

    If Dir$(sourcePath) = "" Then
        Exit Sub
    End If

    ffIn = FreeFile
    Open sourcePath For Input As #ffIn

    ffOut = FreeFile
    Open targetPath For Append As #ffOut
    Print #ffOut, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & sectionName & ": begin"

    Do While Not EOF(ffIn)
        Line Input #ffIn, lineText
        Print #ffOut, lineText
    Loop

    Print #ffOut, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & sectionName & ": end"
    Close #ffIn
    Close #ffOut
End Sub

Private Sub AppendLog(ByVal logPath As String, ByVal lineText As String)
    Dim ff As Integer
    ff = FreeFile
    Open logPath For Append As #ff
    Print #ff, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & lineText
    Close #ff
End Sub

Private Sub AppendCommandOutput(ByVal logPath As String, ByVal rawText As String)
    Dim ff As Integer

    ff = FreeFile
    Open logPath For Append As #ff
    Print #ff, rawText;
    If Right$(rawText, 1) <> vbCr And Right$(rawText, 1) <> vbLf Then
        Print #ff, ""
    End If
    Close #ff
End Sub
