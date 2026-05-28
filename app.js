const RULE_ACCENT_PALETTE = [
  { bg: "rgba(180, 83, 9, 0.18)", color: "#7a4418", ring: "rgba(180, 83, 9, 0.12)" },
  { bg: "rgba(29, 78, 216, 0.16)", color: "#18427b", ring: "rgba(29, 78, 216, 0.12)" },
  { bg: "rgba(15, 118, 110, 0.16)", color: "#0f5e57", ring: "rgba(15, 118, 110, 0.12)" },
  { bg: "rgba(126, 34, 206, 0.15)", color: "#6a21a8", ring: "rgba(126, 34, 206, 0.10)" },
];

const FIELD_INSTRUCTION_PREFIX =
  /^(?:=|ask|author|compare|createdate|date|docproperty|filename|form(?:checkbox|dropdown|text)|gotobutton|hyperlink|if|include(?:picture|text)|macrobutton|mergefield|noteref|numpages|page|pageref|quote|ref|seq|set|styleref|symbol|time|title|toc|xe|ta)\b/i;

const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DOCX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const REVISION_AUTHOR = "WordNumHelper";

const TRANSLATIONS = {
  zh: {
    htmlLang: "zh-CN",
    brandSubtitle: "Word 图表编号整理工具",
    metaLocal: "本地处理",
    metaSafeExport: "安全导出 DOCX",
    heroEyebrow: "Production-ready workflow",
    heroTitle: "Word 图表重编号助手",
    heroLead: "上传 `.docx` 后，帮你给杂乱的图表编号重新编号。",
    heroFlowStage1Title: "原文档标号混乱",
    heroFlowStage1Text: "图表编号前后跳号，正文引用也跟着变乱",
    heroFlowConnector1: "扫描识别",
    heroFlowStage2Title: "扫描识别命中项",
    heroFlowStage2Text: "抓出图表标号、正文引用和章号上下文",
    heroFlowConnector2: "自动重排",
    heroFlowStage3Title: "输出正确编号",
    heroFlowStage3Text: "图表与正文引用一起变成连续、正确的新编号",
    heroTag1: "只替换命中的图表标号",
    heroTag2: "正文引用自动同步",
    heroTag3: "规则输入实时高亮",
    heroTag4: "支持动态 `{t数字}` 标题映射",
    workflowEyebrow: "Workflow",
    workflowTitle: "三步完成处理",
    metric1Label: "01 上传文档",
    metric1Value: "导入一个 `.docx` 文件",
    metric2Label: "02 设置规则",
    metric2Value: "组合正则片段、`{t数字}` 与 `{n}`",
    metric3Label: "03 导出结果",
    metric3Value: "预览映射后下载修正文档",
    step1Eyebrow: "Step 1",
    step1Title: "上传 Word 文件",
    docxOnlyChip: "仅支持 .docx",
    uploadTitle: "选择或拖入一个 `.docx` 文件",
    uploadHint: "文件只会在浏览器本地解析",
    fileMetaEmpty: "还没有选择文件",
    fileMetaSelected: "{name} · {size}",
    step2Eyebrow: "Step 2",
    step2Title: "设置图表标注规则",
    regexHelpButtonText: "规则说明",
    addRuleButtonText: "新增规则",
    rulesDescriptionPrimary: "规则里只有大括号占位符会参与自动重排：<code>{n}</code> 表示递增编号，<code>{t数字}</code> 表示按对应级别标题自动编号，比如 <code>{t1}</code>、<code>{t2}</code>、<code>{t3}</code>；其它部分按正则匹配，但不会被改动。",
    rulesDescriptionSecondary: "例如 <code>图{t1}-{n}</code> 会把图号前半段按当前一级标题自动改写；如果写成 <code>图{t3}-{n}</code>，就会按三级标题自动编号，并让该范围内的 <code>{n}</code> 从设定起始值重新递增。",
    patternLegendRegex: "正则匹配部分",
    patternLegendRenumber: "{n} 递增编号",
    patternLegendHeading: "{t数字} 标题编号",
    step3Eyebrow: "Step 3",
    step3Title: "扫描与导出",
    scanDescription: "先扫描确认标题映射和正文引用变化，再导出修正后的 DOCX。",
    headingInferenceLabel: "允许按文本猜测标题级别",
    headingInferenceHint:
      "默认只信任 Word 标题样式或大纲级别。只有旧文档没套用这些设置时再开启，否则可能把 `1.` 或 `1.1` 这类正文误判为标题。",
    documentPagesLabel: "文档总页数",
    documentPagesHint: "上传后会先快速读取页数，再根据总页数校验扫描范围。",
    documentPagesValuePending: "上传后自动读取",
    documentPagesValueUnknown: "暂时无法判定",
    analyzeButtonText: "扫描并预览",
    downloadButtonText: "下载无修订标注版 DOCX",
    downloadRevisionButtonText: "下载修订标注版 DOCX",
    recognitionEyebrow: "识别逻辑",
    recognitionTitle: "如何区分标题与正文引用",
    feature1: "段首且边界更像标题的内容，会优先视为真正的图表标题。",
    feature2: "像“如图3-1所示”这类紧挨正文的内容，会作为正文引用同步更新。",
    feature3: "当规则里出现 `{t数字}` 时，会按对应级别标题的出现顺序自动编号，并使用该占位符自己的起始值。",
    resultsEyebrow: "结果",
    resultsTitle: "识别到的图表编号与引用",
    captionsSummaryLabel: "图表编号",
    referencesSummaryLabel: "同步引用",
    warningsSummaryLabel: "需要留意",
    captionsSectionTitle: "标题重编号",
    referencesSectionTitle: "正文引用同步",
    thCaptionOriginal: "原标号",
    thCaptionNew: "新标号",
    thCaptionExcerpt: "段落内容",
    thReferenceOriginal: "原引用",
    thReferenceNew: "新引用",
    thReferenceExcerpt: "段落内容",
    resultsPanelHint: "点击展开或收起",
    headingReviewTitle: "识别到的标题",
    headingReviewDescription: "按规则中出现的 {t数字} 分组显示，方便核对标题识别是否正确。",
    headingReviewUnused: "当前规则还没有使用标题占位符。",
    headingReviewEmpty: "这一层级还没有识别到标题。",
    headingReviewCount: "共 {count} 条",
    headingReviewGroupTitle: "{placeholder} 对应的 {level} 级标题",
    headingReviewReset: "恢复全部标题",
    headingReviewRemove: "删除",
    warningsTitle: "提示",
    helpEyebrow: "规则说明",
    helpDialogTitle: "图表规则怎么写",
    closeHelpButtonText: "关闭",
    ruleLabel: "规则",
    rulePlaceholder: "例如：图{t1}-{n} 或 图{t3}-{n}",
    placeholderStartsLabel: "起始编号",
    placeholderStartT1Label: "{t1}",
    placeholderStartT2Label: "{t2}",
    placeholderStartNLabel: "{n}",
    removeRule: "删除",
    keepOneRule: "至少保留一条规则。",
    missingPlaceholderError: "规则“{rule}”必须包含一个 {n}。",
    duplicatePlaceholderError: "规则“{rule}”里的 {placeholder} 不能重复出现。",
    unsupportedPlaceholderError: "规则“{rule}”里包含暂不支持的占位符 {placeholder}。",
    invalidRegexError: "规则“{rule}”不是合法的正则表达式。",
    waitingUpload: "等待上传文档。",
    chooseDocx: "请选择一个 `.docx` 文件。",
    onlyDocx: "当前版本只支持 `.docx`。如果是旧版 `.doc`，请先在 Word 中另存为 `.docx`。",
    fileReady: "文件已就绪，可以开始扫描。",
    uploadFirst: "请先上传一个 `.docx` 文件。",
    jszipLoading: "JSZip 仍在加载，请稍后再试。",
    readingDocumentInfo: "正在读取文档页数。",
    ruleInvalid: "图表规则校验未通过。",
    scanning: "正在解析 DOCX 并扫描图表标注。",
    processedStatus: "处理完成：识别到 {captions} 个标题，同步了 {references} 处正文引用{headingMessage}。",
    headingMessage: "；检测到一级标题 {t1Count} 个、二级标题 {t2Count} 个",
    processFailed: "处理失败，请换一个文档再试。",
    xmlParseError: "XML 解析失败：{partName}",
    noProcessableXmlError: "没有找到可处理的 DOCX 文档内容。",
    missingHeadingWarning: "规则里使用了 {placeholder}，但文档里没有识别到对应级别的标题。",
    placeholderFallbackWarning: "{count} 个标题无法为 {placeholder} 找到对应的标题上下文，已保留原来的这部分数字。",
    noCaptionFound: "没有识别到真正的图表标题，请检查规则是否与文档格式一致。",
    noReferenceFound: "没有需要同步的正文引用。",
    noExtraWarnings: "没有额外提示。",
    duplicateWarning: "旧标号“{text}”出现了多次，正文引用已按“就近且优先向前”的规则同步，建议人工复核。",
    emptyCaptionsRow: "还没有识别到图表标题。",
    emptyReferencesRow: "还没有识别到需要同步的正文引用。",
    emptyParagraph: "(空段落)",
    customRangeToggleLabel: "启用自定义范围",
    customRangeToggleHint: "默认关闭。开启后，只对开始标记和结束标记之间的内容进行识别和编号。",
    customRangeTitle: "自定义范围",
    customRangeHint: "分别粘贴两段在正文中只出现一次的内容，作为开始位置和结束位置。输入后不会自动扫描，点击“扫描并预览”时才会校验。",
    customRangeStartLabel: "开始标记",
    customRangeEndLabel: "结束标记",
    customRangeStartPlaceholder: "粘贴一段只出现一次的开始位置内容",
    customRangeEndPlaceholder: "粘贴一段只出现一次的结束位置内容",
    customRangeStartNotFound: "没有在文档正文中找到开始标记内容，或者是标记出现了多次（建议唯一）。",
    customRangeEndNotFound: "没有在文档正文中找到结束标记内容，或者是标记出现了多次（建议唯一）。",
    customRangeMarkerAmbiguous: "标记内容“{text}”在文档中出现了多次，为了准确起见，请使用唯一的内容作为标记。",
    customRangeMarkersInverted: "开始标记在文档中的位置晚于结束标记，请检查输入是否正确。",
    customRangeLabel: "自定义范围内容",
    customRangeValue: "从第 {start} 段到第 {end} 段",
    outputFileSuffix: "-renumbered",
    outputFileRevisionSuffix: "-renumbered-revisions",
  },
  en: {
    htmlLang: "en",
    brandSubtitle: "Production-ready Word figure and table renumbering tool",
    metaLocal: "Local processing",
    metaSafeExport: "Safe DOCX export",
    heroEyebrow: "Production-ready workflow",
    heroTitle: "Word Figure & Table Renumber Helper",
    heroLead: "Upload a `.docx`, renumber real figure and table captions first, then sync matching in-text references while keeping other formatting intact.",
    heroFlowStage1Title: "Messy source labels",
    heroFlowStage1Text: "Caption numbers jump around and in-text references drift with them.",
    heroFlowConnector1: "Scan",
    heroFlowStage2Title: "Matched caption signals",
    heroFlowStage2Text: "The app detects captions, references, and heading context before renumbering.",
    heroFlowConnector2: "Renumber",
    heroFlowStage3Title: "Clean final order",
    heroFlowStage3Text: "Updated caption numbers and references stay aligned after export.",
    heroTag1: "Only matched label text changes",
    heroTag2: "In-text references stay in sync",
    heroTag3: "Realtime rule highlighting",
    heroTag4: "Supports dynamic `{tN}` heading mapping",
    workflowEyebrow: "Workflow",
    workflowTitle: "Finish in three steps",
    metric1Label: "01 Upload file",
    metric1Value: "Import one `.docx` file",
    metric2Label: "02 Configure rules",
    metric2Value: "Combine regex parts with `{tN}` / `{n}`",
    metric3Label: "03 Export result",
    metric3Value: "Preview mappings and download DOCX",
    step1Eyebrow: "Step 1",
    step1Title: "Upload Word File",
    docxOnlyChip: "DOCX only",
    uploadTitle: "Choose or drop a `.docx` file",
    uploadHint: "Files are processed locally in the browser, which makes static deployment on Cloudflare Pages straightforward.",
    fileMetaEmpty: "No file selected yet",
    fileMetaSelected: "{name} · {size}",
    step2Eyebrow: "Step 2",
    step2Title: "Set Caption Rules",
    regexHelpButtonText: "Rule Guide",
    addRuleButtonText: "Add Rule",
    rulesDescriptionPrimary: "Only brace placeholders participate in renumbering: <code>{n}</code> is the incrementing number, and <code>{tN}</code> maps to the heading index of that level, such as <code>{t1}</code>, <code>{t2}</code>, or <code>{t3}</code>. Everything else is regex matching only.",
    rulesDescriptionSecondary: "For example, <code>Fig\\.{t1}-{n}</code> rewrites the chapter part from the current level-1 heading. If you use <code>{t3}</code>, the value comes from the current level-3 heading and <code>{n}</code> restarts inside that scope.",
    patternLegendRegex: "Regex matching part",
    patternLegendRenumber: "{n} incrementing number",
    patternLegendHeading: "{tN} heading index",
    step3Eyebrow: "Step 3",
    step3Title: "Scan and Export",
    scanDescription: "Scan first to confirm caption mappings and reference updates, then export the corrected DOCX.",
    headingInferenceLabel: "Allow text-based heading inference",
    headingInferenceHint:
      "Off by default. Turn this on only for older documents that do not use real Word heading styles or outline levels, because plain text like `1.` or `1.1` may be misread as headings.",
    documentPagesLabel: "Document pages",
    documentPagesHint: "After upload, the app reads the page count first and validates the scan range against the total pages.",
    documentPagesValuePending: "Read after upload",
    documentPagesValueUnknown: "Unavailable",
    analyzeButtonText: "Scan and Preview",
    downloadButtonText: "Download Clean DOCX",
    downloadRevisionButtonText: "Download Review DOCX",
    recognitionEyebrow: "Detection logic",
    recognitionTitle: "How captions and references differ",
    feature1: "Paragraph-start content with caption-like boundaries is treated as a real caption first.",
    feature2: "Inline text such as “as shown in Fig.3-1” is treated as an in-text reference and updated afterward.",
    feature3: "When a rule contains `{tN}`, heading numbers are generated from heading order at that level and each placeholder uses its own configurable start value.",
    resultsEyebrow: "Results",
    resultsTitle: "Detected Captions and References",
    captionsSummaryLabel: "Real captions",
    referencesSummaryLabel: "Synced references",
    warningsSummaryLabel: "Needs review",
    captionsSectionTitle: "Caption Renumbering",
    referencesSectionTitle: "Reference Sync",
    thCaptionOriginal: "Original",
    thCaptionNew: "Updated",
    thCaptionExcerpt: "Paragraph excerpt",
    thReferenceOriginal: "Original",
    thReferenceNew: "Updated",
    thReferenceExcerpt: "Paragraph excerpt",
    resultsPanelHint: "Click to expand or collapse",
    headingReviewTitle: "Detected Headings",
    headingReviewDescription: "Grouped by the `{tN}` placeholders used in the current rules so you can verify heading detection.",
    headingReviewUnused: "The current rules are not using heading placeholders yet.",
    headingReviewEmpty: "No headings were detected for this level.",
    headingReviewCount: "{count} items",
    headingReviewGroupTitle: "{placeholder} mapped level-{level} headings",
    headingReviewReset: "Restore all headings",
    headingReviewRemove: "Remove",
    warningsTitle: "Notes",
    helpEyebrow: "Rule Guide",
    helpDialogTitle: "How to write caption rules",
    closeHelpButtonText: "Close",
    ruleLabel: "Rule",
    rulePlaceholder: "Example: Fig\\.{t1}-{n} or Fig\\.{t3}-{n}",
    placeholderStartsLabel: "Start numbers",
    placeholderStartT1Label: "{t1}",
    placeholderStartT2Label: "{t2}",
    placeholderStartNLabel: "{n}",
    removeRule: "Remove",
    keepOneRule: "Keep at least one rule.",
    missingPlaceholderError: "Rule “{rule}” must include one {n}.",
    duplicatePlaceholderError: "Rule “{rule}” cannot repeat {placeholder}.",
    unsupportedPlaceholderError: "Rule “{rule}” includes an unsupported placeholder {placeholder}.",
    invalidRegexError: "Rule “{rule}” is not a valid regular expression.",
    waitingUpload: "Waiting for a document upload.",
    chooseDocx: "Please choose a `.docx` file.",
    onlyDocx: "This tool supports `.docx` only. Save old `.doc` files as `.docx` first.",
    fileReady: "File is ready. You can start scanning.",
    uploadFirst: "Please upload a `.docx` file first.",
    jszipLoading: "JSZip is still loading. Please try again in a moment.",
    readingDocumentInfo: "Reading document page count.",
    ruleInvalid: "The caption rules are invalid.",
    scanning: "Parsing the DOCX and scanning caption markers.",
    processedStatus: "Done: {captions} captions detected, {references} references synced{headingMessage}.",
    headingMessage: ", with {t1Count} level-1 headings and {t2Count} level-2 headings detected",
    processFailed: "Processing failed. Please try a different document.",
    xmlParseError: "XML parsing failed: {partName}",
    noProcessableXmlError: "No processable DOCX content was found.",
    missingHeadingWarning: "A rule uses {placeholder}, but no matching heading level was detected in the document.",
    placeholderFallbackWarning: "{count} captions could not resolve {placeholder} from heading context, so that part of the original number was kept.",
    noCaptionFound: "No real captions were detected. Please check whether the rules match the document.",
    noReferenceFound: "No in-text references needed syncing.",
    noExtraWarnings: "No additional notes.",
    duplicateWarning: "Old label “{text}” appears multiple times. References were synced by the nearest previous caption and should be reviewed manually.",
    emptyCaptionsRow: "No captions were detected.",
    emptyReferencesRow: "No references needed syncing.",
    emptyParagraph: "(empty paragraph)",
    customRangeToggleLabel: "Enable custom range",
    customRangeToggleHint: "Off by default. When enabled, only content between the start and end markers is renumbered.",
    customRangeTitle: "Custom Range",
    customRangeHint: "Paste two unique text snippets from your document to define the start and end positions. These are validated when you click 'Scan and Preview'.",
    customRangeStartLabel: "Start marker",
    customRangeEndLabel: "End marker",
    customRangeStartPlaceholder: "Paste a unique snippet for the start position",
    customRangeEndPlaceholder: "Paste a unique snippet for the end position",
    customRangeStartNotFound: "Start marker not found in the main document, or it is not unique.",
    customRangeEndNotFound: "End marker not found in the main document, or it is not unique.",
    customRangeMarkerAmbiguous: "Marker '{text}' is not unique in the document. Please use a unique snippet.",
    customRangeMarkersInverted: "The start marker appears after the end marker. Please check your input.",
    customRangeLabel: "Custom range scope",
    customRangeValue: "Paragraph {start} to {end}",
    outputFileSuffix: "-renumbered",
    outputFileRevisionSuffix: "-renumbered-revisions",
  },
};

const HELP_COPY = {
  zh: {
    intro: [
      "规则中只有大括号占位符会被自动改写，其它字符都只是用于匹配。",
      "`{n}` 表示真正递增的编号；`{t数字}` 会按对应级别标题的出现顺序自动取值，比如 `{t1}`、`{t2}`、`{t3}`。",
      "每个占位符都可以单独设置起始编号，默认都是 1。",
    ],
    examplesTitle: "推荐示例",
    tableTitle: "常用正则与占位符说明",
    tableHeaders: ["符号", "含义", "示例"],
  },
  en: {
    intro: [
      "Only brace placeholders are rewritten automatically. Everything else is used for matching only.",
      "`{n}` is the true incrementing number. `{tN}` comes from the order of headings at that level, such as `{t1}`, `{t2}`, or `{t3}`.",
      "Each placeholder has its own configurable start number, and all defaults are 1.",
    ],
    examplesTitle: "Recommended examples",
    tableTitle: "Common regex symbols and placeholders",
    tableHeaders: ["Symbol", "Meaning", "Example"],
  },
};

const HELP_EXAMPLES = {
  zh: [
    { pattern: "图{t1}-{n}", description: "按一级标题自动改章号，每章内的图号重新从 1 递增。" },
    { pattern: "表{t1}-{n}", description: "表格与图片可以分别设置规则与起始编号。" },
    { pattern: "附图\\({t1}\\)-{n}", description: "大括号之外的括号、连接符、转义符等格式都会原样保留。" },
  ],
  en: [
    { pattern: "Fig\\.{t1}-{n}", description: "Rewrite the chapter number from the current level-1 heading and restart figure numbers inside each top-level section." },
    { pattern: "Table{t1}-{n}", description: "Tables and figures can use separate rules and separate start numbers." },
    { pattern: "Appendix\\({t1}\\)-{n}", description: "Parentheses, separators, and escaped characters outside placeholders stay unchanged." },
  ],
};

const HELP_ROWS = {
  zh: [
    { symbol: "(...)", meaning: "把一段内容作为正则分组，只参与匹配，不会被自动改写。", example: "图(\\d+)-{n}" },
    { symbol: "\\d", meaning: "匹配 1 位数字。", example: "(\\d)" },
    { symbol: "\\d+", meaning: "匹配连续的多位数字。", example: "(\\d+)" },
    { symbol: ".", meaning: "匹配任意单个字符；如果要匹配句点本身，请写成 `\\.`。", example: "Fig\\.{t1}-{n}" },
    { symbol: "+", meaning: "前一个正则单元重复 1 次或多次。", example: "\\d+" },
    { symbol: "*", meaning: "前一个正则单元重复 0 次或多次。", example: "\\s*" },
    { symbol: "?", meaning: "前一个正则单元重复 0 次或 1 次。", example: "图表?" },
    { symbol: "|", meaning: "表示二选一。", example: "(图|表){t1}-{n}" },
    { symbol: "[...]", meaning: "匹配方括号内的任一字符。", example: "[图表]{t1}-{n}" },
    { symbol: "\\s", meaning: "匹配一个空白字符。", example: "Fig\\.\\s*{t1}-{n}" },
    { symbol: "\\(", meaning: "匹配左括号本身。", example: "附图\\({t1}\\)-{n}" },
    { symbol: "\\)", meaning: "匹配右括号本身。", example: "附图\\({t1}\\)-{n}" },
    { symbol: "\\.", meaning: "匹配句点本身。", example: "Fig\\.{t1}-{n}" },
    { symbol: "{n}", meaning: "真正需要重新递增的编号。", example: "图{t1}-{n}" },
    { symbol: "{t1}", meaning: "用当前一级标题的顺序号替换，并应用 `{t1}` 的起始编号。", example: "图{t1}-{n}" },
    { symbol: "{t2}", meaning: "用当前二级标题的顺序号替换，并应用 `{t2}` 的起始编号。", example: "图{t1}-{t2}-{n}" },
  ],
  en: [
    { symbol: "(...)", meaning: "Group a regex part. It matches but is not renumbered.", example: "Fig\\.(\\d+)-{n}" },
    { symbol: "\\d", meaning: "Match one digit.", example: "(\\d)" },
    { symbol: "\\d+", meaning: "Match one or more digits.", example: "(\\d+)" },
    { symbol: ".", meaning: "Match any single character. Use `\\.` for a literal dot.", example: "Fig\\.{t1}-{n}" },
    { symbol: "+", meaning: "Repeat the previous regex unit one or more times.", example: "\\d+" },
    { symbol: "*", meaning: "Repeat the previous regex unit zero or more times.", example: "\\s*" },
    { symbol: "?", meaning: "Repeat the previous regex unit zero or one time.", example: "Figure?" },
    { symbol: "|", meaning: "Either-or choice.", example: "(Fig|Table){t1}-{n}" },
    { symbol: "[...]", meaning: "Match any one character inside the brackets.", example: "[FT]{t1}-{n}" },
    { symbol: "\\s", meaning: "Match one whitespace character.", example: "Fig\\.\\s*{t1}-{n}" },
    { symbol: "\\(", meaning: "Match a literal left parenthesis.", example: "Appendix\\({t1}\\)-{n}" },
    { symbol: "\\)", meaning: "Match a literal right parenthesis.", example: "Appendix\\({t1}\\)-{n}" },
    { symbol: "\\.", meaning: "Match a literal dot.", example: "Fig\\.{t1}-{n}" },
    { symbol: "{n}", meaning: "The actual incrementing number to renumber.", example: "Fig\\.{t1}-{n}" },
    { symbol: "{t1}", meaning: "Use the current level-1 heading index, applying the `{t1}` start number.", example: "Fig\\.{t1}-{n}" },
    { symbol: "{t2}", meaning: "Use the current level-2 heading index, applying the `{t2}` start number.", example: "Fig\\.{t1}-{t2}-{n}" },
  ],
};

const XML_FILE_PRIORITY = [
  "word/document.xml",
  "word/header1.xml",
  "word/header2.xml",
  "word/header3.xml",
  "word/footer1.xml",
  "word/footer2.xml",
  "word/footer3.xml",
  "word/footnotes.xml",
  "word/endnotes.xml",
  "word/comments.xml",
];

const state = {
  file: null,
  language: loadInitialLanguage(),
  allowTextHeadingInference: loadInitialTextHeadingInference(),
  enableCustomRange: loadInitialCustomRangeEnabled(),
  customRangeStart: loadInitialCustomRangeValue("wnh-custom-range-start"),
  customRangeEnd: loadInitialCustomRangeValue("wnh-custom-range-end"),
  documentPageCount: null,
  documentPageCountState: "idle",
  rules: createDefaultRules(loadInitialLanguage()),
  preview: null,
  headingExclusions: new Set(),
  statusLevel: "info",
  statusKey: "waitingUpload",
  statusVars: {},
};

const elements = {
  fileInput: document.getElementById("docxFile"),
  fileMeta: document.getElementById("fileMeta"),
  rulesList: document.getElementById("rulesList"),
  addRuleButton: document.getElementById("addRuleButton"),
  addRuleButtonText: document.getElementById("addRuleButtonText"),
  regexHelpButton: document.getElementById("regexHelpButton"),
  regexHelpButtonText: document.getElementById("regexHelpButtonText"),
  closeHelpButton: document.getElementById("closeHelpButton"),
  closeHelpButtonText: document.getElementById("closeHelpButtonText"),
  helpDialog: document.getElementById("helpDialog"),
  helpDialogBackdrop: document.getElementById("helpDialogBackdrop"),
  helpDialogBody: document.getElementById("helpDialogBody"),
  analyzeButton: document.getElementById("analyzeButton"),
  analyzeButtonText: document.getElementById("analyzeButtonText"),
  downloadButton: document.getElementById("downloadButton"),
  downloadButtonText: document.getElementById("downloadButtonText"),
  downloadRevisionButton: document.getElementById("downloadRevisionButton"),
  downloadRevisionButtonText: document.getElementById("downloadRevisionButtonText"),
  langZhButton: document.getElementById("langZhButton"),
  langEnButton: document.getElementById("langEnButton"),
  statusBox: document.getElementById("statusBox"),
  captionsCount: document.getElementById("captionsCount"),
  referencesCount: document.getElementById("referencesCount"),
  warningsCount: document.getElementById("warningsCount"),
  captionsTableBody: document.getElementById("captionsTableBody"),
  referencesTableBody: document.getElementById("referencesTableBody"),
  warningsList: document.getElementById("warningsList"),
  headingReviewTitle: document.getElementById("headingReviewTitle"),
  headingReviewDescription: document.getElementById("headingReviewDescription"),
  headingReviewList: document.getElementById("headingReviewList"),
  uploadZone: document.querySelector(".upload"),
  brandSubtitle: document.getElementById("brandSubtitle"),
  heroEyebrow: document.getElementById("heroEyebrow"),
  heroTitle: document.getElementById("heroTitle"),
  heroLead: document.getElementById("heroLead"),
  heroFlowStage1Title: document.getElementById("heroFlowStage1Title"),
  heroFlowStage1Text: document.getElementById("heroFlowStage1Text"),
  heroFlowConnector1: document.getElementById("heroFlowConnector1"),
  heroFlowStage2Title: document.getElementById("heroFlowStage2Title"),
  heroFlowStage2Text: document.getElementById("heroFlowStage2Text"),
  heroFlowConnector2: document.getElementById("heroFlowConnector2"),
  heroFlowStage3Title: document.getElementById("heroFlowStage3Title"),
  heroFlowStage3Text: document.getElementById("heroFlowStage3Text"),
  heroTag1: document.getElementById("heroTag1"),
  heroTag2: document.getElementById("heroTag2"),
  heroTag3: document.getElementById("heroTag3"),
  heroTag4: document.getElementById("heroTag4"),
  workflowEyebrow: document.getElementById("workflowEyebrow"),
  workflowTitle: document.getElementById("workflowTitle"),
  metric1Label: document.getElementById("metric1Label"),
  metric1Value: document.getElementById("metric1Value"),
  metric2Label: document.getElementById("metric2Label"),
  metric2Value: document.getElementById("metric2Value"),
  metric3Label: document.getElementById("metric3Label"),
  metric3Value: document.getElementById("metric3Value"),
  step1Eyebrow: document.getElementById("step1Eyebrow"),
  step1Title: document.getElementById("step1Title"),
  docxOnlyChip: document.getElementById("docxOnlyChip"),
  uploadTitle: document.getElementById("uploadTitle"),
  uploadHint: document.getElementById("uploadHint"),
  documentScopePanel: document.getElementById("documentScopePanel"),
  documentPagesLabel: document.getElementById("documentPagesLabel"),
  documentPagesHint: document.getElementById("documentPagesHint"),
  documentPagesValue: document.getElementById("documentPagesValue"),
  step2Eyebrow: document.getElementById("step2Eyebrow"),
  step2Title: document.getElementById("step2Title"),
  rulesDescriptionPrimary: document.getElementById("rulesDescriptionPrimary"),
  rulesDescriptionSecondary: document.getElementById("rulesDescriptionSecondary"),
  patternLegendRegex: document.getElementById("patternLegendRegex"),
  patternLegendRenumber: document.getElementById("patternLegendRenumber"),
  patternLegendHeading: document.getElementById("patternLegendHeading"),
  step3Eyebrow: document.getElementById("step3Eyebrow"),
  step3Title: document.getElementById("step3Title"),
  scanDescription: document.getElementById("scanDescription"),
  headingInferenceCheckbox: document.getElementById("enableTextHeadingInference"),
  headingInferenceLabel: document.getElementById("headingInferenceLabel"),
  headingInferenceHint: document.getElementById("headingInferenceHint"),
  customRangeToggleCheckbox: document.getElementById("enableCustomRange"),
  customRangeToggleLabel: document.getElementById("customRangeToggleLabel"),
  customRangeToggleHint: document.getElementById("customRangeToggleHint"),
  customRangeSection: document.getElementById("customRangeSection"),
  customRangeTitle: document.getElementById("customRangeTitle"),
  customRangeHint: document.getElementById("customRangeHint"),
  customRangeStartLabel: document.getElementById("customRangeStartLabel"),
  customRangeEndLabel: document.getElementById("customRangeEndLabel"),
  customRangeStartInput: document.getElementById("customRangeStart"),
  customRangeEndInput: document.getElementById("customRangeEnd"),
  recognitionEyebrow: document.getElementById("recognitionEyebrow"),
  recognitionTitle: document.getElementById("recognitionTitle"),
  feature1: document.getElementById("feature1"),
  feature2: document.getElementById("feature2"),
  feature3: document.getElementById("feature3"),
  resultsEyebrow: document.getElementById("resultsEyebrow"),
  resultsTitle: document.getElementById("resultsTitle"),
  captionsSummaryLabel: document.getElementById("captionsSummaryLabel"),
  referencesSummaryLabel: document.getElementById("referencesSummaryLabel"),
  warningsSummaryLabel: document.getElementById("warningsSummaryLabel"),
  captionsSectionTitle: document.getElementById("captionsSectionTitle"),
  referencesSectionTitle: document.getElementById("referencesSectionTitle"),
  thCaptionOriginal: document.getElementById("thCaptionOriginal"),
  thCaptionNew: document.getElementById("thCaptionNew"),
  thCaptionExcerpt: document.getElementById("thCaptionExcerpt"),
  thReferenceOriginal: document.getElementById("thReferenceOriginal"),
  thReferenceNew: document.getElementById("thReferenceNew"),
  thReferenceExcerpt: document.getElementById("thReferenceExcerpt"),
  captionsPanelHint: document.getElementById("captionsPanelHint"),
  referencesPanelHint: document.getElementById("referencesPanelHint"),
  warningsTitle: document.getElementById("warningsTitle"),
  resetHeadingFiltersButton: document.getElementById("resetHeadingFiltersButton"),
  helpEyebrow: document.getElementById("helpEyebrow"),
  helpDialogTitle: document.getElementById("helpDialogTitle"),
};

boot();

function boot() {
  applyLanguage(state.language, { updateRules: false });
  renderRules();
  bindEvents();
  bindDocumentScopeEvents();
}

function bindEvents() {
  elements.fileInput?.addEventListener("change", (event) => {
    const [file] = event.target.files ?? [];
    handleFileSelection(file ?? null);
  });

  elements.addRuleButton?.addEventListener("click", () => {
    state.rules.push(createEmptyRule());
    renderRules();
    clearPreviewState();
  });

  elements.langZhButton?.addEventListener("click", () => switchLanguage("zh"));
  elements.langEnButton?.addEventListener("click", () => switchLanguage("en"));

  elements.regexHelpButton?.addEventListener("click", openHelpDialog);
  elements.closeHelpButton?.addEventListener("click", closeHelpDialog);
  elements.helpDialogBackdrop?.addEventListener("click", closeHelpDialog);

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && !elements.helpDialog?.hidden) {
      closeHelpDialog();
    }
  });

  elements.rulesList?.addEventListener("input", (event) => {
    const input = event.target;
    if (!(input instanceof HTMLInputElement)) {
      return;
    }

    const ruleId = input.closest("[data-rule-id]")?.dataset.ruleId;
    const rule = state.rules.find((item) => item.id === ruleId);
    if (!rule) {
      return;
    }

    if (input.name === "matchPattern") {
      rule.matchPattern = input.value;
      syncPatternOverlay(input, input.value);
    }

    if (input.name === "renumberPattern") {
      rule.renumberPattern = input.value;
      syncPatternOverlay(input, input.value);
      syncRuleStartFields(input.closest(".rule-card"), rule);
    }

    if (input.name === "placeholderStart") {
      const placeholder = input.dataset.placeholder;
      if (placeholder && getPlaceholderMeta(placeholder)) {
        rule.starts[placeholder] = sanitizeStartAt(input.value);
        input.value = String(rule.starts[placeholder]);
      }
    }

    clearPreviewState();
  });

  elements.rulesList?.addEventListener("focusin", (event) => {
    const input = event.target;
    if (!(input instanceof HTMLInputElement) || input.dataset.overlayInput !== "true") {
      return;
    }

    const shell = input.closest(".input-overlay-shell");
    if (!(shell instanceof HTMLDivElement)) {
      return;
    }

    shell.classList.add("is-editing");
    syncPatternOverlay(input, input.value);
  });

  elements.rulesList?.addEventListener("focusout", (event) => {
    const input = event.target;
    if (!(input instanceof HTMLInputElement) || input.dataset.overlayInput !== "true") {
      return;
    }

    const shell = input.closest(".input-overlay-shell");
    if (!(shell instanceof HTMLDivElement)) {
      return;
    }

    shell.classList.remove("is-editing");
    syncPatternOverlay(input, input.value);
  });

  elements.rulesList?.addEventListener("click", (event) => {
    const button = event.target;
    if (!(button instanceof HTMLButtonElement) || button.dataset.action !== "remove-rule") {
      return;
    }

    if (state.rules.length === 1) {
      setStatus("keepOneRule", "warn");
      return;
    }

    const ruleId = button.closest("[data-rule-id]")?.dataset.ruleId;
    state.rules = state.rules.filter((item) => item.id !== ruleId);
    renderRules();
    clearPreviewState();
  });

  elements.analyzeButton?.addEventListener("click", () => void analyzeDocument());
  elements.downloadButton?.addEventListener("click", () => {
    if (!state.preview?.blob) {
      return;
    }
    downloadProcessedDocument(state.preview.blob, state.file?.name ?? "renumbered.docx", "outputFileSuffix");
  });
  elements.downloadRevisionButton?.addEventListener("click", () => {
    if (!state.preview?.revisionBlob) {
      return;
    }
    downloadProcessedDocument(
      state.preview.revisionBlob,
      state.file?.name ?? "renumbered.docx",
      "outputFileRevisionSuffix",
    );
  });
  elements.resetHeadingFiltersButton?.addEventListener("click", () => void resetHeadingFilters());
  elements.headingReviewList?.addEventListener("click", (event) => {
    const button = event.target;
    if (!(button instanceof HTMLButtonElement) || button.dataset.action !== "remove-heading") {
      return;
    }

    const order = Number.parseInt(button.dataset.headingOrder ?? "", 10);
    if (!Number.isFinite(order)) {
      return;
    }

    state.headingExclusions.add(order);
    void analyzeDocument();
  });

  elements.headingInferenceCheckbox?.addEventListener("change", (event) => {
    const input = event.target;
    if (!(input instanceof HTMLInputElement)) {
      return;
    }

    state.allowTextHeadingInference = input.checked;
    window.localStorage.setItem("wnh-text-heading-inference", String(input.checked));

    if (state.file && state.preview) {
      void analyzeDocument();
      return;
    }

    clearPreviewState();
    if (state.file) {
      setStatus(state.documentPageCountState === "pending" ? "readingDocumentInfo" : "fileReady");
    }
  });

  ["dragenter", "dragover"].forEach((eventName) => {
    elements.uploadZone?.addEventListener(eventName, (event) => {
      event.preventDefault();
      elements.uploadZone.classList.add("is-dragging");
    });
  });

  ["dragleave", "drop"].forEach((eventName) => {
    elements.uploadZone?.addEventListener(eventName, (event) => {
      event.preventDefault();
      elements.uploadZone.classList.remove("is-dragging");
    });
  });

  elements.uploadZone?.addEventListener("drop", (event) => {
    const [file] = event.dataTransfer?.files ?? [];
    handleFileSelection(file ?? null);
  });
}

function bindDocumentScopeEvents() {
  elements.customRangeToggleCheckbox?.addEventListener("change", (event) => {
    const input = event.target;
    if (!(input instanceof HTMLInputElement)) {
      return;
    }

    state.enableCustomRange = input.checked;
    window.localStorage.setItem("wnh-enable-custom-range", String(input.checked));
    syncCustomRangeInputs();

    if (state.file && state.preview) {
      void analyzeDocument();
    } else {
      clearPreviewState();
      if (state.file) {
        setStatus(state.documentPageCountState === "pending" ? "readingDocumentInfo" : "fileReady");
      }
    }
  });

  [elements.customRangeStartInput, elements.customRangeEndInput].forEach((input) => {
    input?.addEventListener("change", (event) => {
      const target = event.target;
      if (!(target instanceof HTMLTextAreaElement)) {
        return;
      }

      if (target === elements.customRangeStartInput) {
        state.customRangeStart = target.value;
        window.localStorage.setItem("wnh-custom-range-start", state.customRangeStart);
      } else {
        state.customRangeEnd = target.value;
        window.localStorage.setItem("wnh-custom-range-end", state.customRangeEnd);
      }

      if (state.enableCustomRange && state.file && state.preview) {
        void analyzeDocument();
      }
    });

    input?.addEventListener("input", (event) => {
      const target = event.target;
      if (!(target instanceof HTMLTextAreaElement)) {
        return;
      }
      clearPreviewState();
      if (state.file) {
        setStatus(state.documentPageCountState === "pending" ? "readingDocumentInfo" : "fileReady");
      }
    });
  });
}

function loadInitialLanguage() {
  return window.localStorage.getItem("wnh-language") === "en" ? "en" : "zh";
}

function loadInitialTextHeadingInference() {
  return window.localStorage.getItem("wnh-text-heading-inference") === "true";
}

function loadInitialCustomRangeEnabled() {
  return window.localStorage.getItem("wnh-enable-custom-range") === "true";
}

function loadInitialCustomRangeValue(key) {
  return window.localStorage.getItem(key) ?? "";
}

function t(key, vars = {}) {
  const dictionary = TRANSLATIONS[state.language] ?? TRANSLATIONS.zh;
  const template = dictionary[key] ?? TRANSLATIONS.zh[key] ?? key;
  return template.replace(/\{(\w+)\}/g, (match, name) =>
    Object.prototype.hasOwnProperty.call(vars, name) ? String(vars[name]) : match,
  );
}

function syncCustomRangeInputs() {
  elements.customRangeSection?.classList.toggle("is-hidden", !state.enableCustomRange);
  if (elements.customRangeToggleCheckbox) {
    elements.customRangeToggleCheckbox.checked = state.enableCustomRange;
  }
  if (elements.customRangeStartInput) {
    elements.customRangeStartInput.value = state.customRangeStart;
    elements.customRangeStartInput.placeholder = t("customRangeStartPlaceholder");
  }
  if (elements.customRangeEndInput) {
    elements.customRangeEndInput.value = state.customRangeEnd;
    elements.customRangeEndInput.placeholder = t("customRangeEndPlaceholder");
  }
}

function renderFileMeta() {
  if (!state.file) {
    setText(elements.fileMeta, t("fileMetaEmpty"));
    return;
  }

  setText(
    elements.fileMeta,
    t("fileMetaSelected", {
      name: state.file.name,
      size: formatFileSize(state.file.size),
    }),
  );
}

function renderDocumentScope() {
  elements.documentScopePanel?.classList.toggle("is-hidden", !state.file);
}

function renderDocumentPageInfo() {
  if (!elements.documentPagesValue) {
    return;
  }

  if (state.documentPageCountState === "ready" && Number.isInteger(state.documentPageCount)) {
    elements.documentPagesValue.textContent = String(state.documentPageCount);
  } else if (state.documentPageCountState === "unknown") {
    elements.documentPagesValue.textContent = t("documentPagesValueUnknown");
  } else {
    elements.documentPagesValue.textContent = t("documentPagesValuePending");
  }
}

function switchLanguage(language) {
  if (language === state.language) {
    return;
  }

  const canReplaceDefaults = areRulesUsingDefaultTemplates(state.rules, state.language);
  state.language = language;
  window.localStorage.setItem("wnh-language", language);

  if (canReplaceDefaults) {
    state.rules = createDefaultRules(language);
    clearPreviewState();
  }

  applyLanguage(language, { updateRules: true });
  syncStatus();
}

function applyLanguage(language, { updateRules }) {
  document.documentElement.lang = TRANSLATIONS[language].htmlLang;
  elements.langZhButton?.classList.toggle("is-active", language === "zh");
  elements.langEnButton?.classList.toggle("is-active", language === "en");

  setText(elements.brandSubtitle, t("brandSubtitle"));
  setText(elements.heroEyebrow, t("heroEyebrow"));
  setText(elements.heroTitle, t("heroTitle"));
  setText(elements.heroLead, t("heroLead"));
  setText(elements.heroFlowStage1Title, t("heroFlowStage1Title"));
  setText(elements.heroFlowStage1Text, t("heroFlowStage1Text"));
  setText(elements.heroFlowConnector1, t("heroFlowConnector1"));
  setText(elements.heroFlowStage2Title, t("heroFlowStage2Title"));
  setText(elements.heroFlowStage2Text, t("heroFlowStage2Text"));
  setText(elements.heroFlowConnector2, t("heroFlowConnector2"));
  setText(elements.heroFlowStage3Title, t("heroFlowStage3Title"));
  setText(elements.heroFlowStage3Text, t("heroFlowStage3Text"));
  setText(elements.heroTag1, t("heroTag1"));
  setText(elements.heroTag2, t("heroTag2"));
  setText(elements.heroTag3, t("heroTag3"));
  setText(elements.heroTag4, t("heroTag4"));
  setText(elements.workflowEyebrow, t("workflowEyebrow"));
  setText(elements.workflowTitle, t("workflowTitle"));
  setText(elements.metric1Label, t("metric1Label"));
  setText(elements.metric1Value, t("metric1Value"));
  setText(elements.metric2Label, t("metric2Label"));
  setText(elements.metric2Value, t("metric2Value"));
  setText(elements.metric3Label, t("metric3Label"));
  setText(elements.metric3Value, t("metric3Value"));
  setText(elements.step1Eyebrow, t("step1Eyebrow"));
  setText(elements.step1Title, t("step1Title"));
  setText(elements.docxOnlyChip, t("docxOnlyChip"));
  setText(elements.uploadTitle, t("uploadTitle"));
  setText(elements.uploadHint, t("uploadHint"));
  setText(elements.documentPagesLabel, t("documentPagesLabel"));
  setText(elements.documentPagesHint, t("documentPagesHint"));
  setText(elements.step2Eyebrow, t("step2Eyebrow"));
  setText(elements.step2Title, t("step2Title"));
  setHtml(elements.rulesDescriptionPrimary, t("rulesDescriptionPrimary"));
  setHtml(elements.rulesDescriptionSecondary, t("rulesDescriptionSecondary"));
  setText(elements.patternLegendRegex, t("patternLegendRegex"));
  setText(elements.patternLegendRenumber, t("patternLegendRenumber"));
  setText(elements.patternLegendHeading, t("patternLegendHeading"));
  setText(elements.regexHelpButtonText, t("regexHelpButtonText"));
  setText(elements.addRuleButtonText, t("addRuleButtonText"));
  setText(elements.step3Eyebrow, t("step3Eyebrow"));
  setText(elements.step3Title, t("step3Title"));
  setText(elements.scanDescription, t("scanDescription"));
  setText(elements.headingInferenceLabel, t("headingInferenceLabel"));
  setText(elements.headingInferenceHint, t("headingInferenceHint"));
  if (elements.headingInferenceCheckbox) {
    elements.headingInferenceCheckbox.checked = state.allowTextHeadingInference;
  }
  syncCustomRangeInputs();

  setText(elements.customRangeToggleLabel, t("customRangeToggleLabel"));
  setText(elements.customRangeToggleHint, t("customRangeToggleHint"));
  setText(elements.customRangeTitle, t("customRangeTitle"));
  setText(elements.customRangeHint, t("customRangeHint"));
  setText(elements.customRangeStartLabel, t("customRangeStartLabel"));
  setText(elements.customRangeEndLabel, t("customRangeEndLabel"));

  setText(elements.analyzeButtonText, t("analyzeButtonText"));
  setText(elements.downloadButtonText, t("downloadButtonText"));
  setText(elements.downloadRevisionButtonText, t("downloadRevisionButtonText"));
  setText(elements.recognitionEyebrow, t("recognitionEyebrow"));
  setText(elements.recognitionTitle, t("recognitionTitle"));
  setText(elements.feature1, t("feature1"));
  setText(elements.feature2, t("feature2"));
  setText(elements.feature3, t("feature3"));
  setText(elements.resultsEyebrow, t("resultsEyebrow"));
  setText(elements.resultsTitle, t("resultsTitle"));
  setText(elements.captionsSummaryLabel, t("captionsSummaryLabel"));
  setText(elements.referencesSummaryLabel, t("referencesSummaryLabel"));
  setText(elements.warningsSummaryLabel, t("warningsSummaryLabel"));
  setText(elements.captionsSectionTitle, t("captionsSectionTitle"));
  setText(elements.referencesSectionTitle, t("referencesSectionTitle"));
  setText(elements.thCaptionOriginal, t("thCaptionOriginal"));
  setText(elements.thCaptionNew, t("thCaptionNew"));
  setText(elements.thCaptionExcerpt, t("thCaptionExcerpt"));
  setText(elements.thReferenceOriginal, t("thReferenceOriginal"));
  setText(elements.thReferenceNew, t("thReferenceNew"));
  setText(elements.thReferenceExcerpt, t("thReferenceExcerpt"));
  setText(elements.captionsPanelHint, t("resultsPanelHint"));
  setText(elements.referencesPanelHint, t("resultsPanelHint"));
  setText(elements.headingReviewTitle, t("headingReviewTitle"));
  setText(elements.headingReviewDescription, t("headingReviewDescription"));
  setText(elements.resetHeadingFiltersButton, t("headingReviewReset"));
  setText(elements.warningsTitle, t("warningsTitle"));
  setText(elements.helpEyebrow, t("helpEyebrow"));
  setText(elements.helpDialogTitle, t("helpDialogTitle"));
  setText(elements.closeHelpButtonText, t("closeHelpButtonText"));
  renderHelpDialogContent();
  renderDocumentScope();
  renderFileMeta();
  renderDocumentPageInfo();

  if (state.preview) {
    renderPreview(state.preview);
  } else {
    renderRows(elements.captionsTableBody, [], t("emptyCaptionsRow"));
    renderRows(elements.referencesTableBody, [], t("emptyReferencesRow"));
    renderHeadingReview({}, []);
    if (elements.warningsList) {
      elements.warningsList.innerHTML = `<li>${escapeHtml(t("noExtraWarnings"))}</li>`;
    }
  }

  if (updateRules) {
    renderRules();
  }
}

function renderHelpDialogContent() {
  if (!elements.helpDialogBody) {
    return;
  }

  const copy = HELP_COPY[state.language] ?? HELP_COPY.zh;
  const examples = HELP_EXAMPLES[state.language] ?? HELP_EXAMPLES.zh;
  const rows = HELP_ROWS[state.language] ?? HELP_ROWS.zh;

  const introHtml = copy.intro.map((line) => `<p>${renderInlineMarkdown(line)}</p>`).join("");
  const examplesHtml = examples
    .map(
      (item) => `
        <article class="modal__example">
          <strong class="pattern-inline">${renderPatternInline(item.pattern)}</strong>
          <p>${escapeHtml(item.description)}</p>
        </article>
      `,
    )
    .join("");
  const rowsHtml = rows
    .map(
      (row) => `
        <tr>
          <td><code>${escapeHtml(row.symbol)}</code></td>
          <td>${renderInlineMarkdown(row.meaning)}</td>
          <td><code>${escapeHtml(row.example)}</code></td>
        </tr>
      `,
    )
    .join("");

  elements.helpDialogBody.innerHTML = `
    ${introHtml}
    <div class="modal__examples">
      <h3>${escapeHtml(copy.examplesTitle)}</h3>
      ${examplesHtml}
    </div>
    <section class="modal__table-section">
      <h3>${escapeHtml(copy.tableTitle)}</h3>
      <div class="table-wrap modal__table-wrap">
        <table>
          <thead>
            <tr>
              ${copy.tableHeaders.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}
            </tr>
          </thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      </div>
    </section>
  `;
}

function renderInlineMarkdown(text) {
  return escapeHtml(text).replace(/`([^`]+)`/g, "<code>$1</code>");
}

function createRuleStarts(overrides = {}) {
  const starts = {
    n: sanitizeStartAt(overrides.n ?? overrides.startAt ?? 1),
  };

  for (const [key, value] of Object.entries(overrides)) {
    if (/^t\d+$/u.test(key)) {
      starts[key] = sanitizeStartAt(value);
    }
  }

  return starts;
}

function normalizeRuleStarts(starts = {}) {
  return createRuleStarts(starts);
}

function getPlaceholderMeta(name) {
  if (name === "n") {
    return { kind: "counter", level: null };
  }

  const match = String(name).match(/^t(\d+)$/u);
  if (!match) {
    return null;
  }

  return {
    kind: "heading",
    level: Number.parseInt(match[1], 10),
  };
}

function isHeadingPlaceholder(name) {
  return getPlaceholderMeta(name)?.kind === "heading";
}

function getPlaceholderSortKey(name) {
  if (name === "n") {
    return Number.MAX_SAFE_INTEGER;
  }
  return getPlaceholderMeta(name)?.level ?? Number.MAX_SAFE_INTEGER - 1;
}

function createDefaultRules(language) {
  if (language === "en") {
    return [
      { id: crypto.randomUUID(), matchPattern: "Fig\\.(\\d+)-(\\d+)", renumberPattern: "Fig\\.{t1}-{n}", starts: createRuleStarts() },
      { id: crypto.randomUUID(), matchPattern: "Table(\\d+)-(\\d+)", renumberPattern: "Table{t1}-{n}", starts: createRuleStarts() },
    ];
  }

  if (language === "zh") {
    return [
      { id: crypto.randomUUID(), matchPattern: "\u56fe(\\d+)-(\\d+)", renumberPattern: "\u56fe{t1}-{n}", starts: createRuleStarts() },
      { id: crypto.randomUUID(), matchPattern: "\u8868(\\d+)-(\\d+)", renumberPattern: "\u8868{t1}-{n}", starts: createRuleStarts() },
    ];
  }

  return [
    { id: crypto.randomUUID(), template: "图{t1}-{n}", starts: createRuleStarts() },
    { id: crypto.randomUUID(), template: "表{t1}-{n}", starts: createRuleStarts() },
  ];
}

function createEmptyRule() {
  return {
    id: crypto.randomUUID(),
    matchPattern: "",
    renumberPattern: "",
    starts: createRuleStarts(),
  };
}

function areRulesUsingDefaultTemplates(rules, language) {
  const defaults = createDefaultRules(language);
  if (rules.length !== defaults.length) {
    return false;
  }

  return rules.every((rule, index) => {
    const current = normalizeRuleForCompare(rule);
    const expected = normalizeRuleForCompare(defaults[index]);
    return JSON.stringify(current) === JSON.stringify(expected);
  });
}

function normalizeRuleForCompare(rule) {
  return {
    matchPattern: rule.matchPattern ?? "",
    renumberPattern: rule.renumberPattern ?? "",
    starts: normalizeRuleStarts(rule.starts),
  };
}

function setText(element, value) {
  if (element) {
    element.textContent = value;
  }
}

function setHtml(element, value) {
  if (element) {
    element.innerHTML = value;
  }
}

function handleFileSelection(file) {
  clearPreviewState();
  state.headingExclusions = new Set();

  if (!file) {
    state.file = null;
    state.documentPageCount = null;
    state.documentPageCountState = "idle";
    renderDocumentScope();
    renderFileMeta();
    renderDocumentPageInfo();
    syncPageRangeInputs();
    setStatus("waitingUpload");
    return;
  }

  if (!file.name.toLowerCase().endsWith(".docx")) {
    state.file = null;
    state.documentPageCount = null;
    state.documentPageCountState = "idle";
    renderDocumentScope();
    setText(elements.fileMeta, t("chooseDocx"));
    renderDocumentPageInfo();
    setStatus("onlyDocx", "warn");
    return;
  }

  state.file = file;
  state.documentPageCount = null;
  state.documentPageCountState = "pending";
  renderDocumentScope();
  renderFileMeta();
  renderDocumentPageInfo();
  setStatus("readingDocumentInfo");
  void inspectUploadedDocument(file);
}

async function inspectUploadedDocument(file) {
  if (typeof JSZip === "undefined") {
    state.documentPageCount = null;
    state.documentPageCountState = "pending";
    renderDocumentPageInfo();
    setStatus("jszipLoading", "warn");
    return;
  }

  try {
    const pageCount = await readDocumentPageCount(await file.arrayBuffer());
    if (state.file !== file) {
      return;
    }

    state.documentPageCount = Number.isInteger(pageCount) && pageCount > 0 ? pageCount : null;
    state.documentPageCountState = state.documentPageCount ? "ready" : "unknown";
    renderDocumentPageInfo();

    try {
      setStatus("fileReady");
    } catch (error) {
      console.warn("Failed setting file ready status", error);
    }
  } catch (error) {
    console.warn("Failed to inspect uploaded document", error);
    if (state.file !== file) {
      return;
    }

    state.documentPageCount = null;
    state.documentPageCountState = "unknown";
    renderDocumentPageInfo();
    setStatus("fileReady");
  }
}

async function readDocumentPageCount(arrayBuffer) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const appXmlFile = zip.file("docProps/app.xml");
  if (appXmlFile) {
    const xmlText = await appXmlFile.async("text");
    const xmlDoc = new DOMParser().parseFromString(xmlText, "application/xml");
    if (!xmlDoc.querySelector("parsererror")) {
      const pagesNode =
        xmlDoc.getElementsByTagName("Pages")[0] ??
        xmlDoc.getElementsByTagNameNS("*", "Pages")[0] ??
        null;
      const pages = parseOptionalPositiveInteger(pagesNode?.textContent ?? "");
      if (Number.isInteger(pages)) {
        return pages;
      }
    }
  }

  const partNames = Object.keys(zip.files)
    .filter(
      (name) =>
        /^word\/.+\.xml$/i.test(name) &&
        !/\/(styles|theme|fontTable|settings|webSettings|numbering)\.xml$/i.test(name),
    )
    .sort((left, right) => comparePartNames(left, right));

  const docs = [];
  let nextParagraphOrder = 0;

  for (const partName of partNames) {
    const file = zip.file(partName);
    if (!file) {
      continue;
    }

    const xmlText = await file.async("text");
    if (!xmlText.includes("<w:p")) {
      continue;
    }

    const xmlDoc = new DOMParser().parseFromString(xmlText, "application/xml");
    if (xmlDoc.querySelector("parsererror")) {
      continue;
    }

    const paragraphs = collectParagraphs(xmlDoc, partName, nextParagraphOrder);
    nextParagraphOrder += paragraphs.length;
    docs.push({ partName, xmlDoc, paragraphs });
  }

  if (!docs.length) {
    return null;
  }

  const pageTracking = assignDocumentPageNumbers(docs);
  return Number.isInteger(pageTracking.estimatedPageCount) ? pageTracking.estimatedPageCount : null;
}

async function resetHeadingFilters() {
  if (!state.headingExclusions.size) {
    return;
  }

  state.headingExclusions = new Set();
  if (state.file) {
    await analyzeDocument();
  } else {
    clearPreviewState();
  }
}

function renderRules() {
  if (!elements.rulesList) {
    return;
  }

  elements.rulesList.innerHTML = state.rules
    .map((rule, index) => {
      const startFields = renderPlaceholderStartFields(rule);
      const matchLabel = state.language === "en" ? "Match pattern" : "匹配规则";
      const renumberLabel = state.language === "en" ? "Renumber pattern" : "重编号规则";
      const matchPlaceholder =
        state.language === "en"
          ? "Example: Fig\\.(\\d+)-(\\d+) or Table(\\d+)"
          : "例如：图(\\d+)-(\\d+) 或 表(\\d+)";
      const renumberPlaceholder =
        state.language === "en"
          ? "Example: Fig\\.{t1}-{n} or Table{n}"
          : "例如：图{t1}-{n} 或 表{n}";

      return `
        <article class="rule-card" data-rule-id="${escapeHtml(rule.id)}">
          <div class="rule-card__patterns field--full">
            <div class="field">
              <label for="match-pattern-${escapeHtml(rule.id)}">${escapeHtml(matchLabel)} ${index + 1}</label>
              <div class="input-overlay-shell">
                <div class="input-overlay" aria-hidden="true">${renderPatternOverlay(rule.matchPattern ?? "")}</div>
                <input
                  id="match-pattern-${escapeHtml(rule.id)}"
                  name="matchPattern"
                  type="text"
                  placeholder="${escapeHtml(matchPlaceholder)}"
                  value="${escapeHtml(rule.matchPattern ?? "")}"
                  autocomplete="off"
                  spellcheck="false"
                  data-overlay-input="true"
                />
              </div>
            </div>
            <div class="field">
              <label for="renumber-pattern-${escapeHtml(rule.id)}">${escapeHtml(renumberLabel)} ${index + 1}</label>
              <div class="input-overlay-shell">
                <div class="input-overlay" aria-hidden="true">${renderPatternOverlay(rule.renumberPattern ?? "")}</div>
                <input
                  id="renumber-pattern-${escapeHtml(rule.id)}"
                  name="renumberPattern"
                  type="text"
                  placeholder="${escapeHtml(renumberPlaceholder)}"
                  value="${escapeHtml(rule.renumberPattern ?? "")}"
                  autocomplete="off"
                  spellcheck="false"
                  data-overlay-input="true"
                />
              </div>
            </div>
          </div>
          <div class="field field--full ${startFields ? "" : "is-hidden"}" data-role="placeholder-starts">
            <label>${escapeHtml(t("placeholderStartsLabel"))}</label>
            <div class="rule-starts">${startFields}</div>
          </div>
          <div class="rule-card__actions">
            <button class="button button--danger" type="button" data-action="remove-rule">
              ${escapeHtml(t("removeRule"))}
            </button>
          </div>
        </article>
      `;
    })
    .join("");
}

function renderPlaceholderStartFields(rule) {
  const placeholders = getRulePlaceholderNames(rule.renumberPattern);
  return placeholders.map((placeholder) => renderPlaceholderStartField(rule, placeholder)).join("");
}

function renderPlaceholderStartField(rule, placeholder) {
  if (!(placeholder in rule.starts)) {
    rule.starts[placeholder] = 1;
  }

  const visual = getPlaceholderVisual(placeholder);

  return `
    <label class="rule-start ${isHeadingPlaceholder(placeholder) ? "rule-start--heading" : "rule-start--counter"}">
      <span class="rule-start__tag" style="${visual.tagStyle}">${escapeHtml(`{${placeholder}}`)}</span>
      <input
        name="placeholderStart"
        data-placeholder="${escapeHtml(placeholder)}"
        type="number"
        min="1"
        step="1"
        value="${escapeHtml(String(rule.starts[placeholder]))}"
      />
    </label>
  `;
}

function syncRuleStartFields(ruleCard, rule) {
  if (!(ruleCard instanceof HTMLElement)) {
    return;
  }
  const wrapper = ruleCard.querySelector('[data-role="placeholder-starts"]');
  const starts = ruleCard.querySelector(".rule-starts");
  if (!(wrapper instanceof HTMLElement) || !(starts instanceof HTMLDivElement)) {
    return;
  }
  const content = renderPlaceholderStartFields(rule);
  starts.innerHTML = content;
  wrapper.classList.toggle("is-hidden", !content);
}

function usesPlaceholder(template, placeholder) {
  return String(template).includes(`{${placeholder}}`);
}

function getRulePlaceholderNames(template) {
  const placeholders = extractPlaceholders(String(template ?? ""))
    .map((item) => item.name)
    .filter((name) => getPlaceholderMeta(name))
    .filter((name, index, values) => values.indexOf(name) === index)
    .sort((left, right) => {
      const sortDiff = getPlaceholderSortKey(left) - getPlaceholderSortKey(right);
      return sortDiff !== 0 ? sortDiff : left.localeCompare(right, "en");
    });

  return placeholders;
}

function clearPreviewState() {
  state.preview = null;
  if (elements.downloadButton) {
    elements.downloadButton.disabled = true;
  }
  if (elements.downloadRevisionButton) {
    elements.downloadRevisionButton.disabled = true;
  }
  setText(elements.captionsCount, "0");
  setText(elements.referencesCount, "0");
  setText(elements.warningsCount, "0");
  renderRows(elements.captionsTableBody, [], t("emptyCaptionsRow"));
  renderRows(elements.referencesTableBody, [], t("emptyReferencesRow"));
  renderHeadingReview({}, []);
  if (elements.warningsList) {
    elements.warningsList.innerHTML = `<li>${escapeHtml(t("noExtraWarnings"))}</li>`;
  }
}

async function analyzeDocument() {
  if (!state.file) {
    setStatus("uploadFirst", "warn");
    return;
  }

  if (typeof JSZip === "undefined") {
    setStatus("jszipLoading", "warn");
    return;
  }

  let compiledRules;
  try {
    compiledRules = compileRules(state.rules);
  } catch (error) {
    setStatus(error instanceof Error ? error.message : t("ruleInvalid"), "warn", { isRaw: true });
    return;
  }

  const customRangeInfo = state.enableCustomRange
    ? { startMarker: state.customRangeStart, endMarker: state.customRangeEnd }
    : null;

  if (elements.analyzeButton) {
    elements.analyzeButton.disabled = true;
  }
  if (elements.downloadButton) {
    elements.downloadButton.disabled = true;
  }
  if (elements.downloadRevisionButton) {
    elements.downloadRevisionButton.disabled = true;
  }
  setStatus("scanning");

  try {
    const preview = await processDocx(await state.file.arrayBuffer(), compiledRules, {
      headingExclusions: state.headingExclusions,
      allowTextHeadingInference: state.allowTextHeadingInference,
      customRange: customRangeInfo,
      documentPageCount: state.documentPageCount,
    });
    state.preview = preview;
    renderPreview(preview);
    if (elements.downloadButton) {
      elements.downloadButton.disabled = false;
    }
    if (elements.downloadRevisionButton) {
      elements.downloadRevisionButton.disabled = false;
    }

    const headingMessage =
      buildHeadingSummary(preview.headingCounts, compiledRules) +
      (preview.customRange ? "；" + buildCustomRangeSummary(preview.customRange) : "");

    setStatus("processedStatus", "info", {
      captions: preview.captions.length,
      references: preview.references.length,
      headingMessage,
    });
  } catch (error) {
    console.error(error);
    setStatus(error instanceof Error ? error.message : t("processFailed"), "warn", { isRaw: true });
  } finally {
    if (elements.analyzeButton) {
      elements.analyzeButton.disabled = false;
    }
  }
}

async function processDocx(arrayBuffer, compiledRules, options = {}) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const customRangeInfo = options.customRange ?? null;
  const stylesInfo = await parseStylesInfo(zip);
  const partNames = Object.keys(zip.files)
    .filter(
      (name) =>
        /^word\/.+\.xml$/i.test(name) &&
        !/\/(styles|theme|fontTable|settings|webSettings|numbering)\.xml$/i.test(name),
    )
    .sort((left, right) => comparePartNames(left, right));

  const docs = [];
  let nextParagraphOrder = 0;

  for (const partName of partNames) {
    const file = zip.file(partName);
    if (!file) {
      continue;
    }

    const xmlText = await file.async("text");
    if (!xmlText.includes("<w:p")) {
      continue;
    }

    const xmlDoc = new DOMParser().parseFromString(xmlText, "application/xml");
    if (xmlDoc.querySelector("parsererror")) {
      throw new Error(t("xmlParseError", { partName }));
    }

    const paragraphs = collectParagraphs(xmlDoc, partName, nextParagraphOrder);
    nextParagraphOrder += paragraphs.length;
    docs.push({ partName, xmlDoc, paragraphs });
  }

  if (docs.length === 0) {
    throw new Error(t("noProcessableXmlError"));
  }

  const customRange = customRangeInfo
    ? resolveCustomRange(docs, customRangeInfo.startMarker, customRangeInfo.endMarker)
    : null;

  const pageTracking = assignDocumentPageNumbers(docs);
  const reportedPageCount = Number.isInteger(options.documentPageCount) ? options.documentPageCount : null;

  const pageInfo = {
    detectedPageBreaks: pageTracking.detectedPageBreaks,
    estimatedPageCount: pageTracking.estimatedPageCount,
    reportedPageCount,
    skippedUnmappedCandidates: 0,
  };

  const headingInfo = buildHeadingContexts(docs, stylesInfo, {
    ignoredHeadingOrders: options.headingExclusions,
    allowTextHeadingInference: options.allowTextHeadingInference,
  });
  const candidates = [];

  for (const doc of docs) {
    for (const paragraph of doc.paragraphs) {
      const paragraphCandidates = selectCandidates(paragraph.fullText, compiledRules);
      for (const candidate of paragraphCandidates) {
        const pageNumber = resolveCandidatePageNumber(paragraph, candidate.start);
        const enrichedCandidate = {
          ...candidate,
          partName: doc.partName,
          paragraphOrder: paragraph.order,
          paragraphKey: paragraph.key,
          paragraph,
          chapterContext: paragraph.chapterContext,
          excerpt: collapseWhitespace(paragraph.fullText),
          classification: classifyCandidate(paragraph.fullText, candidate),
          pageNumber,
        };

        if (customRange && !isCandidateWithinCustomRange(paragraph.order, customRange)) {
          continue;
        }

        candidates.push(enrichedCandidate);
      }
    }
  }

  const counters = new Map();
  const captionMap = new Map();
  const captions = [];

  for (const candidate of candidates) {
    if (candidate.classification !== "caption") {
      continue;
    }

    const renumbered = renumberCaption(candidate, counters);
    const enriched = {
      ...candidate,
      newText: renumbered.newText,
      placeholderFallbacks: renumbered.placeholderFallbacks,
      resolvedPlaceholders: renumbered.resolvedPlaceholders,
    };
    captions.push(enriched);

    const mapKey = makeMappingKey(candidate.rule.id, candidate.text);
    const mapped = captionMap.get(mapKey) ?? [];
    mapped.push(enriched);
    captionMap.set(mapKey, mapped);
  }

  const references = [];
  for (const candidate of candidates) {
    if (candidate.classification === "caption") {
      continue;
    }

    const mapped = captionMap.get(makeMappingKey(candidate.rule.id, candidate.text));
    if (!mapped?.length) {
      continue;
    }

    const target = resolveReferenceTarget(candidate, mapped);
    if (!target || target.newText === candidate.text) {
      continue;
    }

    references.push({ ...candidate, newText: target.newText });
  }

  const replacementsByParagraph = new Map();
  for (const item of [...captions, ...references].filter((entry) => entry.newText !== entry.text)) {
    const replacements = replacementsByParagraph.get(item.paragraphKey) ?? [];
    replacements.push({ start: item.start, end: item.end, newText: item.newText });
    replacementsByParagraph.set(item.paragraphKey, replacements);
  }

  const revisionBlob = await generateRevisionBlob(arrayBuffer, partNames, replacementsByParagraph);

  for (const doc of docs) {
    for (const paragraph of doc.paragraphs) {
      const replacements = replacementsByParagraph.get(paragraph.key);
      if (replacements?.length) {
        applyReplacementsToParagraph(paragraph, replacements);
      }
    }
    zip.file(doc.partName, new XMLSerializer().serializeToString(doc.xmlDoc));
  }

  return {
    blob: await zip.generateAsync({
      type: "blob",
      mimeType: DOCX_MIME_TYPE,
    }),
    revisionBlob,
    captions,
    references,
    headingCounts: headingInfo.counts,
    headingGroups: headingInfo.groups,
    usedHeadingPlaceholders: getUsedHeadingPlaceholders(compiledRules),
    pageInfo,
    customRange: options.customRangeInfo,
    warnings: buildWarnings(captions, references, captionMap, headingInfo, compiledRules, pageInfo),
  };
}

async function generateRevisionBlob(arrayBuffer, partNames, replacementsByParagraph) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const revisionContext = {
    author: REVISION_AUTHOR,
    date: new Date().toISOString(),
    nextId: 1,
  };

  for (const partName of partNames) {
    const file = zip.file(partName);
    if (!file) {
      continue;
    }

    const xmlText = await file.async("text");
    if (!xmlText.includes("<w:p")) {
      continue;
    }

    const xmlDoc = new DOMParser().parseFromString(xmlText, "application/xml");
    if (xmlDoc.querySelector("parsererror")) {
      throw new Error(t("xmlParseError", { partName }));
    }

    const paragraphs = collectParagraphs(xmlDoc, partName, 0);
    for (const paragraph of paragraphs) {
      const replacements = replacementsByParagraph.get(paragraph.key);
      if (replacements?.length) {
        applyTrackedReplacementsToParagraph(paragraph, replacements, revisionContext);
      }
    }

    zip.file(partName, new XMLSerializer().serializeToString(xmlDoc));
  }

  await enableRevisionDisplay(zip);

  return zip.generateAsync({
    type: "blob",
    mimeType: DOCX_MIME_TYPE,
  });
}

function comparePartNames(left, right) {
  const leftIndex = XML_FILE_PRIORITY.indexOf(left);
  const rightIndex = XML_FILE_PRIORITY.indexOf(right);

  if (leftIndex !== -1 || rightIndex !== -1) {
    return (
      (leftIndex === -1 ? Number.MAX_SAFE_INTEGER : leftIndex) -
      (rightIndex === -1 ? Number.MAX_SAFE_INTEGER : rightIndex)
    );
  }

  return left.localeCompare(right, "zh-CN");
}

function compileRules(rules) {
  const cleaned = rules
    .map((rule) => ({
      ...rule,
      matchPattern: String(rule.matchPattern ?? "").trim(),
      renumberPattern: String(rule.renumberPattern ?? "").trim(),
      starts: normalizeRuleStarts(rule.starts),
    }))
    .filter((rule) => rule.matchPattern || rule.renumberPattern);

  if (cleaned.length === 0) {
    throw new Error(t("keepOneRule"));
  }

  return cleaned.map((rule) => {
    if (!rule.matchPattern || !rule.renumberPattern) {
      throw new Error(state.language === "en" ? "Each rule needs both patterns." : "每条规则都需要同时填写匹配规则和重编号规则。");
    }

    const placeholders = extractPlaceholders(rule.renumberPattern);
    const counts = new Map();

    for (const placeholder of placeholders) {
      counts.set(placeholder.name, (counts.get(placeholder.name) ?? 0) + 1);
      if (!getPlaceholderMeta(placeholder.name)) {
        throw new Error(
          t("unsupportedPlaceholderError", {
            rule: rule.renumberPattern,
            placeholder: `{${placeholder.name}}`,
          }),
        );
      }
    }

    if ((counts.get("n") ?? 0) === 0) {
      throw new Error(t("missingPlaceholderError", { rule: rule.renumberPattern }));
    }

    for (const [name, count] of counts.entries()) {
      if (count > 1) {
        throw new Error(
          t("duplicatePlaceholderError", {
            rule: rule.renumberPattern,
            placeholder: `{${name}}`,
          }),
        );
      }
    }

    let regex;
    try {
      regex = new RegExp(rule.matchPattern, "dgu");
    } catch {
      throw new Error(t("invalidRegexError", { rule: rule.matchPattern }));
    }

    const captureGroupCount = countCapturingGroups(rule.matchPattern);
    if (captureGroupCount < placeholders.length) {
      throw new Error(
        state.language === "en"
          ? `Match pattern "${rule.matchPattern}" needs at least ${placeholders.length} capturing groups.`
          : `匹配规则“${rule.matchPattern}”里至少需要有 ${placeholders.length} 个捕获组。`,
      );
    }

    const headingPlaceholders = placeholders
      .map((item) => item.name)
      .filter((name) => isHeadingPlaceholder(name));

    return {
      ...rule,
      regex,
      captureGroupOffset: captureGroupCount - placeholders.length,
      placeholders,
      placeholderOrder: placeholders.map((item) => item.name),
      headingPlaceholders,
    };
  });
}

function extractPlaceholders(template) {
  const result = [];

  for (const match of template.matchAll(/\{([a-z]\d*)\}/g)) {
    result.push({
      name: match[1],
      raw: match[0],
      index: match.index ?? 0,
      length: match[0].length,
    });
  }

  return result;
}

function countCapturingGroups(pattern) {
  let count = 0;
  let escaped = false;
  let inCharClass = false;

  for (let index = 0; index < pattern.length; index += 1) {
    const char = pattern[index];

    if (escaped) {
      escaped = false;
      continue;
    }

    if (char === "\\") {
      escaped = true;
      continue;
    }

    if (inCharClass) {
      if (char === "]") {
        inCharClass = false;
      }
      continue;
    }

    if (char === "[") {
      inCharClass = true;
      continue;
    }

    if (char !== "(") {
      continue;
    }

    if (pattern[index + 1] === "?") {
      continue;
    }

    count += 1;
  }

  return count;
}

function collectParagraphs(xmlDoc, partName, startOrder) {
  return Array.from(xmlDoc.getElementsByTagNameNS("*", "p")).map((node, index) => {
    const cursor = { value: 0 };
    const pieces = [];
    const context = { fieldStack: [], pageBreaks: [] };
    walkParagraph(node, pieces, cursor, context);
    const metadata = getParagraphMetadata(node);

    return {
      node,
      partName,
      order: startOrder + index,
      key: `${partName}::${index}`,
      fullText: pieces.map((item) => item.text).join(""),
      textSegments: pieces.filter((item) => item.kind === "text"),
      styleId: metadata.styleId,
      outlineLevel: metadata.outlineLevel,
      pageBreakBefore: metadata.pageBreakBefore,
      sectionBreakAfter: metadata.sectionBreakAfter,
      pageBreaks: context.pageBreaks,
      pageNumber: null,
      chapterContext: snapshotHeadingContext(new Map()),
    };
  });
}

async function parseStylesInfo(zip) {
  const file = zip.file("word/styles.xml");
  if (!file) {
    return createStylesInfo(new Map());
  }

  const xmlText = await file.async("text");
  const xmlDoc = new DOMParser().parseFromString(xmlText, "application/xml");
  if (xmlDoc.querySelector("parsererror")) {
    return createStylesInfo(new Map());
  }

  const stylesById = new Map();
  for (const styleNode of Array.from(xmlDoc.getElementsByTagNameNS("*", "style"))) {
    if (getWordAttribute(styleNode, "type") !== "paragraph") {
      continue;
    }

    const styleId = getWordAttribute(styleNode, "styleId");
    if (!styleId) {
      continue;
    }

    const nameNode = getDirectChild(styleNode, "name");
    const basedOnNode = getDirectChild(styleNode, "basedOn");
    const pPrNode = getDirectChild(styleNode, "pPr");
    const outlineNode = pPrNode ? getDirectChild(pPrNode, "outlineLvl") : null;

    stylesById.set(styleId, {
      styleId,
      name: getWordAttribute(nameNode, "val") ?? "",
      basedOn: getWordAttribute(basedOnNode, "val") ?? null,
      outlineLevel: parseNullableInteger(getWordAttribute(outlineNode, "val")),
    });
  }

  return createStylesInfo(stylesById);
}

function createStylesInfo(stylesById) {
  const cache = new Map();

  return {
    getHeadingLevel(styleId) {
      if (!styleId) {
        return null;
      }

      if (cache.has(styleId)) {
        return cache.get(styleId);
      }

      const level = resolveStyleHeadingLevel(styleId, stylesById, cache, new Set());
      cache.set(styleId, level);
      return level;
    },
  };
}

function resolveStyleHeadingLevel(styleId, stylesById, cache, visited) {
  if (!styleId || visited.has(styleId)) {
    return null;
  }

  if (cache.has(styleId)) {
    return cache.get(styleId);
  }

  visited.add(styleId);
  const style = stylesById.get(styleId);
  if (!style) {
    cache.set(styleId, null);
    return null;
  }

  if (Number.isInteger(style.outlineLevel) && style.outlineLevel >= 0 && style.outlineLevel <= 8) {
    const level = style.outlineLevel + 1;
    cache.set(styleId, level);
    return level;
  }

  const directLevel = detectHeadingLevelFromStyleLabel(`${style.styleId} ${style.name}`);
  if (directLevel) {
    cache.set(styleId, directLevel);
    return directLevel;
  }

  const inheritedLevel = style.basedOn
    ? resolveStyleHeadingLevel(style.basedOn, stylesById, cache, visited)
    : null;
  cache.set(styleId, inheritedLevel);
  return inheritedLevel;
}

function detectHeadingLevelFromStyleLabel(label) {
  const match = String(label).match(/(?:heading|鏍囬|title)\s*([1-9])/iu);
  return match ? Number.parseInt(match[1], 10) : null;
}

function buildHeadingContexts(docs, stylesInfo, options = {}) {
  const ignoredHeadingOrders = options.ignoredHeadingOrders ?? new Set();
  const allowTextHeadingInference = options.allowTextHeadingInference === true;
  const counts = {};
  const groups = {};
  const currentHeadings = new Map();
  const sequences = new Map();

  for (const doc of docs) {
    for (const paragraph of doc.paragraphs) {
      if (paragraph.partName !== "word/document.xml") {
        paragraph.chapterContext = snapshotHeadingContext(currentHeadings);
        continue;
      }

      const headingLevel =
        !ignoredHeadingOrders.has(paragraph.order)
          ? getParagraphHeadingLevel(paragraph, stylesInfo, allowTextHeadingInference)
          : null;
      if (headingLevel) {
        resetDeeperHeadingLevels(headingLevel, currentHeadings, sequences);
        const nextSequence = (sequences.get(headingLevel) ?? 0) + 1;
        sequences.set(headingLevel, nextSequence);

        const heading = {
          token: `t${headingLevel}`,
          level: headingLevel,
          sequence: nextSequence,
          order: paragraph.order,
          text: collapseWhitespace(paragraph.fullText),
        };

        currentHeadings.set(headingLevel, heading);
        counts[heading.token] = (counts[heading.token] ?? 0) + 1;
        (groups[heading.token] ??= []).push(heading);
      }

      paragraph.chapterContext = snapshotHeadingContext(currentHeadings);
    }
  }

  return { counts, groups };
}

function resetDeeperHeadingLevels(level, currentHeadings, sequences) {
  for (const currentLevel of [...currentHeadings.keys()]) {
    if (currentLevel > level) {
      currentHeadings.delete(currentLevel);
    }
  }

  for (const currentLevel of [...sequences.keys()]) {
    if (currentLevel > level) {
      sequences.delete(currentLevel);
    }
  }
}

function snapshotHeadingContext(currentHeadings) {
  return {
    levels: Object.fromEntries([...currentHeadings.entries()]),
  };
}

function getParagraphHeadingLevel(paragraph, stylesInfo, allowTextHeadingInference = false) {
  if (paragraph.partName !== "word/document.xml") {
    return null;
  }

  if (Number.isInteger(paragraph.outlineLevel) && paragraph.outlineLevel >= 0 && paragraph.outlineLevel <= 8) {
    return paragraph.outlineLevel + 1;
  }

  const styleLevel = stylesInfo.getHeadingLevel(paragraph.styleId);
  if (styleLevel) {
    return styleLevel;
  }

  return allowTextHeadingInference ? inferHeadingLevelFromText(paragraph.fullText) : null;
}

function inferHeadingLevelFromText(text) {
  const value = collapseWhitespace(text);
  if (!value || value.length > 80) {
    return null;
  }

  if (/^(chapter|section)\s+\d+\b/iu.test(value)) {
    return 1;
  }
  if (/^绗?\s*\d+\s*[绔犺妭绡囬儴鍒哴/u.test(value) || /^\d+[銆?)锛嶿\s*[^\d]/u.test(value)) {
    return 1;
  }
  if (/^\d+[.-]\d+[銆?)锛嶿?\s*[^\s]/u.test(value) || /^\d+\.\d+\s+[^\d]/u.test(value)) {
    return 2;
  }

  return null;
}

function getParagraphMetadata(paragraphNode) {
  const pPrNode = getDirectChild(paragraphNode, "pPr");
  const styleNode = pPrNode ? getDirectChild(pPrNode, "pStyle") : null;
  const outlineNode = pPrNode ? getDirectChild(pPrNode, "outlineLvl") : null;
  const pageBreakBeforeNode = pPrNode ? getDirectChild(pPrNode, "pageBreakBefore") : null;
  const sectionNode = pPrNode ? getDirectChild(pPrNode, "sectPr") : null;
  const sectionTypeNode = sectionNode ? getDirectChild(sectionNode, "type") : null;
  const sectionType = getWordAttribute(sectionTypeNode, "val") ?? (sectionNode ? "nextPage" : null);

  return {
    styleId: getWordAttribute(styleNode, "val"),
    outlineLevel: parseNullableInteger(getWordAttribute(outlineNode, "val")),
    pageBreakBefore: getWordOnOffValue(pageBreakBeforeNode),
    sectionBreakAfter: Boolean(sectionType) && sectionType !== "continuous",
  };
}

function getDirectChild(node, localName) {
  return (
    Array.from(node?.childNodes ?? []).find(
      (child) => child.nodeType === Node.ELEMENT_NODE && child.localName === localName,
    ) ?? null
  );
}

function getWordAttribute(node, name) {
  return node?.getAttribute(`w:${name}`) ?? node?.getAttribute(name) ?? null;
}

function getWordOnOffValue(node) {
  if (!node) {
    return false;
  }

  const value = getWordAttribute(node, "val");
  return value === null ? true : !/^(?:0|false|off)$/iu.test(String(value));
}

function parseNullableInteger(value) {
  const parsed = Number.parseInt(String(value), 10);
  return Number.isFinite(parsed) ? parsed : null;
}

function assignDocumentPageNumbers(docs) {
  let currentPage = 1;
  let detectedPageBreaks = 0;

  for (const doc of docs) {
    for (const paragraph of doc.paragraphs) {
      if (paragraph.partName !== "word/document.xml") {
        paragraph.pageNumber = null;
        continue;
      }

      if (paragraph.pageBreakBefore) {
        currentPage += 1;
        detectedPageBreaks += 1;
      }

      paragraph.pageNumber = currentPage;

      for (const pageBreak of paragraph.pageBreaks) {
        pageBreak.pageNumber = currentPage + 1;
        currentPage += 1;
        detectedPageBreaks += 1;
      }

      if (paragraph.sectionBreakAfter) {
        currentPage += 1;
        detectedPageBreaks += 1;
      }
    }
  }

  return {
    detectedPageBreaks,
    estimatedPageCount: currentPage,
  };
}

function resolveCandidatePageNumber(paragraph, offset) {
  if (!Number.isInteger(paragraph.pageNumber)) {
    return null;
  }

  let pageNumber = paragraph.pageNumber;
  for (const pageBreak of paragraph.pageBreaks) {
    if (offset >= pageBreak.position) {
      pageNumber = pageBreak.pageNumber ?? pageNumber + 1;
      continue;
    }
    break;
  }

  return pageNumber;
}



function resolveCustomRange(docs, startMarker, endMarker) {
  if (!startMarker && !endMarker) {
    return null;
  }

  const findOrder = (marker) => {
    if (!marker) return null;
    const normalizedMarker = collapseWhitespace(marker).toLowerCase();
    const matches = [];
    for (const doc of docs) {
      if (doc.partName !== "word/document.xml") continue;
      for (const paragraph of doc.paragraphs) {
        if (collapseWhitespace(paragraph.fullText).toLowerCase().includes(normalizedMarker)) {
          matches.push(paragraph.order);
        }
      }
    }
    if (matches.length === 1) return matches[0];
    if (matches.length > 1) throw new Error(t("customRangeMarkerAmbiguous", { text: marker }));
    return null;
  };

  const startOrder = findOrder(startMarker);
  const endOrder = findOrder(endMarker);

  if (startMarker && startOrder === null) {
    throw new Error(t("customRangeStartNotFound"));
  }
  if (endMarker && endOrder === null) {
    throw new Error(t("customRangeEndNotFound"));
  }

  if (startOrder !== null && endOrder !== null && startOrder > endOrder) {
    throw new Error(t("customRangeMarkersInverted"));
  }

  return { start: startOrder, end: endOrder };
}

function isCandidateWithinCustomRange(candidateOrder, customRange) {
  if (!customRange) {
    return true;
  }

  if (customRange.start !== null && candidateOrder < customRange.start) {
    return false;
  }

  if (customRange.end !== null && candidateOrder > customRange.end) {
    return false;
  }

  return true;
}

function walkParagraph(node, pieces, cursor, context) {
  if (node.nodeType !== Node.ELEMENT_NODE) {
    return;
  }

  if (node.localName === "lastRenderedPageBreak") {
    context.pageBreaks.push({ position: cursor.value, kind: "rendered" });
    return;
  }

  if (node.localName === "fldChar") {
    const fieldType = getWordAttribute(node, "fldCharType");
    if (fieldType === "begin") {
      context.fieldStack.push({ hasSeparate: false });
    } else if (fieldType === "separate") {
      const currentField = context.fieldStack.at(-1);
      if (currentField) {
        currentField.hasSeparate = true;
      }
    } else if (fieldType === "end" && context.fieldStack.length) {
      context.fieldStack.pop();
    }
    return;
  }

  if (node.localName === "t") {
    const text = node.textContent ?? "";
    pieces.push({ kind: "text", node, text, start: cursor.value, end: cursor.value + text.length });
    cursor.value += text.length;
    return;
  }

  if (node.localName === "instrText") {
    const text = node.textContent ?? "";
    if (shouldUseInstrTextAsDisplayText(text, context.fieldStack)) {
      pieces.push({ kind: "text", node, text, start: cursor.value, end: cursor.value + text.length });
      cursor.value += text.length;
    }
    return;
  }

  if (node.localName === "br" && getWordAttribute(node, "type") === "page") {
    context.pageBreaks.push({ position: cursor.value, kind: "manual" });
    return;
  }

  if (node.localName === "tab" || node.localName === "br" || node.localName === "cr") {
    const text = node.localName === "tab" ? "\t" : "\n";
    pieces.push({ kind: "virtual", text, start: cursor.value, end: cursor.value + 1 });
    cursor.value += 1;
    return;
  }

  for (const child of node.childNodes) {
    walkParagraph(child, pieces, cursor, context);
  }
}

function shouldUseInstrTextAsDisplayText(text, fieldStack) {
  if (!text) {
    return false;
  }

  if (fieldStack.some((field) => field.hasSeparate)) {
    return true;
  }

  const normalized = collapseWhitespace(text).trim();
  if (!normalized) {
    return false;
  }

  return !looksLikeFieldInstruction(normalized);
}

function looksLikeFieldInstruction(text) {
  return FIELD_INSTRUCTION_PREFIX.test(text) || /\\[A-Za-z*#@!]/.test(text) || /\bMERGEFORMAT\b/i.test(text);
}

function selectCandidates(fullText, compiledRules) {
  const byStart = new Map();

  for (const rule of compiledRules) {
    rule.regex.lastIndex = 0;

    for (const match of fullText.matchAll(rule.regex)) {
      const start = match.index ?? 0;
      const placeholders = {};
      const captureGroupOffset = rule.captureGroupOffset ?? 0;

      for (const [index, placeholder] of rule.placeholders.entries()) {
        const groupIndex = captureGroupOffset + index + 1;
        const indices = match.indices?.[groupIndex];
        if (!indices) {
          continue;
        }

        placeholders[placeholder.name] = {
          name: placeholder.name,
          start: indices[0],
          end: indices[1],
          value: match[groupIndex] ?? fullText.slice(indices[0], indices[1]),
        };
      }

      if (Object.keys(placeholders).length !== rule.placeholderOrder.length) {
        continue;
      }

      const item = {
        rule,
        start,
        end: start + match[0].length,
        text: match[0],
        placeholders,
      };

      const list = byStart.get(start) ?? [];
      list.push(item);
      byStart.set(start, list);
    }
  }

  const selected = [];
  let cursor = 0;

  for (const start of [...byStart.keys()].sort((left, right) => left - right)) {
    if (start < cursor) {
      continue;
    }

    const winner = (byStart.get(start) ?? []).sort((left, right) => right.text.length - left.text.length)[0];
    selected.push(winner);
    cursor = winner.end;
  }

  return selected;
}

function classifyCandidate(fullText, candidate) {
  const before = fullText.slice(0, candidate.start);
  const after = fullText.slice(candidate.end);
  const gluedBefore = isBodyCharacter(before.slice(-1));
  const gluedAfter = isBodyCharacter(after.charAt(0));
  const looksLikeCaptionHead = before.trim().length === 0 || /^[("鈥溾€榎[锛堛€?\s-]*$/u.test(before);
  const looksLikeCaptionTail = after.length === 0 || /^[\s\t\n:锛?锛?;锛涖€?锛夈€?"鈥??锛侊紵-]/u.test(after);

  if (looksLikeCaptionHead && looksLikeCaptionTail && !gluedBefore) {
    return "caption";
  }

  return gluedBefore || gluedAfter ? "reference" : "reference";
}

function isBodyCharacter(character) {
  return /[\p{L}\p{N}]/u.test(character);
}

function renumberCaption(candidate, counters) {
  const resolvedPlaceholders = {};
  const placeholderFallbacks = [];

  for (const placeholder of candidate.rule.headingPlaceholders) {
    const resolved = resolveHeadingPlaceholder(candidate, placeholder);
    if (resolved === null) {
      placeholderFallbacks.push(placeholder);
      resolvedPlaceholders[placeholder] = candidate.placeholders[placeholder]?.value ?? "";
    } else {
      resolvedPlaceholders[placeholder] = String(resolved);
    }
  }

  const counterScope = candidate.rule.headingPlaceholders
    .map((placeholder) => `${placeholder}:${resolvedPlaceholders[placeholder]}`)
    .join("|");
  const counterKey = `${candidate.rule.id}::${counterScope || "root"}`;
  const nextValue = counters.has(counterKey) ? counters.get(counterKey) : candidate.rule.starts.n;
  counters.set(counterKey, nextValue + 1);
  resolvedPlaceholders.n = String(nextValue);

  return {
    newText: buildRenumberedText(candidate, resolvedPlaceholders),
    placeholderFallbacks,
    resolvedPlaceholders,
  };
}

function resolveHeadingPlaceholder(candidate, placeholder) {
  const meta = getPlaceholderMeta(placeholder);
  if (!meta || meta.kind !== "heading") {
    return null;
  }

  const heading = candidate.chapterContext?.levels?.[meta.level];
  if (!heading) {
    return null;
  }

  return heading.sequence + candidate.rule.starts[placeholder] - 1;
}

function buildRenumberedText(candidate, resolvedPlaceholders) {
  let cursor = 0;
  let result = "";

  for (const placeholder of candidate.rule.placeholders) {
    result += candidate.rule.renumberPattern.slice(cursor, placeholder.index);
    result += resolvedPlaceholders?.[placeholder.name] ?? candidate.placeholders[placeholder.name]?.value ?? "";
    cursor = placeholder.index + placeholder.length;
  }

  return result + candidate.rule.renumberPattern.slice(cursor);
}

function resolveReferenceTarget(reference, mapped) {
  const samePartEarlier = mapped.filter(
    (item) => item.partName === reference.partName && item.paragraphOrder <= reference.paragraphOrder,
  );
  if (samePartEarlier.length) {
    return samePartEarlier.at(-1);
  }

  const anyEarlier = mapped.filter((item) => item.paragraphOrder <= reference.paragraphOrder);
  if (anyEarlier.length) {
    return anyEarlier.at(-1);
  }

  return mapped[0];
}

function applyReplacementsToParagraph(paragraph, replacements) {
  for (const replacement of [...replacements].sort((left, right) => right.start - left.start)) {
    const segments = paragraph.textSegments.filter(
      (segment) => segment.start < replacement.end && segment.end > replacement.start,
    );

    if (!segments.length) {
      continue;
    }

    if (segments.length === 1) {
      const segment = segments[0];
      const localStart = Math.max(0, replacement.start - segment.start);
      const localEnd = Math.min(segment.text.length, replacement.end - segment.start);
      segment.text =
        segment.text.slice(0, localStart) +
        replacement.newText +
        segment.text.slice(localEnd);
      writeSegment(segment);
      continue;
    }

    const first = segments[0];
    const last = segments.at(-1);
    for (const segment of segments) {
      const overlapStart = Math.max(replacement.start, segment.start);
      const overlapEnd = Math.min(replacement.end, segment.end);
      const localStart = Math.max(0, overlapStart - segment.start);
      const localEnd = Math.max(0, overlapEnd - segment.start);

      if (segment === first) {
        segment.text = segment.text.slice(0, localStart) + replacement.newText;
      } else if (segment === last) {
        segment.text = segment.text.slice(localEnd);
      } else {
        segment.text = "";
      }

      writeSegment(segment);
    }
  }
}

function applyTrackedReplacementsToParagraph(paragraph, replacements, revisionContext) {
  const sortedReplacements = [...replacements].sort((left, right) => left.start - right.start);
  const replacedRuns = new Set();

  for (const segment of paragraph.textSegments) {
    const relevantReplacements = sortedReplacements.filter(
      (replacement) => segment.start < replacement.end && segment.end > replacement.start,
    );
    if (!relevantReplacements.length) {
      continue;
    }

    const sourceRun = findAncestorByLocalName(segment.node, "r");
    if (!sourceRun?.parentNode || replacedRuns.has(sourceRun)) {
      continue;
    }

    const parent = sourceRun.parentNode;
    const newNodes = buildTrackedReplacementNodesForSegment(
      sourceRun,
      segment,
      relevantReplacements,
      revisionContext,
    );

    for (const node of newNodes) {
      parent.insertBefore(node, sourceRun);
    }
    parent.removeChild(sourceRun);
    replacedRuns.add(sourceRun);
  }
}

function buildTrackedReplacementNodesForSegment(sourceRun, segment, replacements, revisionContext) {
  const nodes = [];
  let cursor = segment.start;

  for (const replacement of replacements) {
    const overlapStart = Math.max(replacement.start, segment.start);
    const overlapEnd = Math.min(replacement.end, segment.end);

    if (overlapStart > cursor) {
      nodes.push(createTextRunLike(sourceRun, "t", segment.text.slice(cursor - segment.start, overlapStart - segment.start)));
    }

    const deletedText = segment.text.slice(overlapStart - segment.start, overlapEnd - segment.start);
    if (deletedText) {
      nodes.push(createRevisionWrapper("del", sourceRun, deletedText, revisionContext));
    }

    if (replacement.start >= segment.start && replacement.start < segment.end) {
      nodes.push(createRevisionWrapper("ins", sourceRun, replacement.newText, revisionContext));
    }

    cursor = overlapEnd;
  }

  if (cursor < segment.end) {
    nodes.push(createTextRunLike(sourceRun, "t", segment.text.slice(cursor - segment.start)));
  }

  return nodes;
}

function createRevisionWrapper(type, sourceRun, text, revisionContext) {
  const doc = sourceRun.ownerDocument;
  const wrapper = doc.createElementNS(WORD_NS, `w:${type}`);
  wrapper.setAttributeNS(WORD_NS, "w:id", String(revisionContext.nextId++));
  wrapper.setAttributeNS(WORD_NS, "w:author", revisionContext.author);
  wrapper.setAttributeNS(WORD_NS, "w:date", revisionContext.date);
  wrapper.appendChild(createTextRunLike(sourceRun, type === "del" ? "delText" : "t", text));
  return wrapper;
}

function createTextRunLike(sourceRun, textElementName, text) {
  const doc = sourceRun.ownerDocument;
  const run = doc.createElementNS(WORD_NS, "w:r");
  const properties = getDirectChild(sourceRun, "rPr");
  if (properties) {
    run.appendChild(properties.cloneNode(true));
  }

  const textNode = doc.createElementNS(WORD_NS, `w:${textElementName}`);
  textNode.textContent = text;
  if (/^\s|\s$/u.test(text)) {
    textNode.setAttribute("xml:space", "preserve");
  }
  run.appendChild(textNode);
  return run;
}

function findAncestorByLocalName(node, localName) {
  let current = node?.parentNode ?? null;
  while (current) {
    if (current.nodeType === Node.ELEMENT_NODE && current.localName === localName) {
      return current;
    }
    current = current.parentNode;
  }
  return null;
}

async function enableRevisionDisplay(zip) {
  const settingsPath = "word/settings.xml";
  const file = zip.file(settingsPath);
  if (!file) {
    return;
  }

  const xmlText = await file.async("text");
  const xmlDoc = new DOMParser().parseFromString(xmlText, "application/xml");
  if (xmlDoc.querySelector("parsererror")) {
    return;
  }

  const settingsNode = xmlDoc.getElementsByTagNameNS("*", "settings")[0];
  if (!settingsNode) {
    return;
  }

  if (!settingsNode.getElementsByTagNameNS("*", "trackRevisions").length) {
    settingsNode.appendChild(xmlDoc.createElementNS(WORD_NS, "w:trackRevisions"));
  }

  let revisionView = settingsNode.getElementsByTagNameNS("*", "revisionView")[0] ?? null;
  if (!revisionView) {
    revisionView = xmlDoc.createElementNS(WORD_NS, "w:revisionView");
    settingsNode.appendChild(revisionView);
  }

  revisionView.setAttributeNS(WORD_NS, "w:markup", "1");
  revisionView.setAttributeNS(WORD_NS, "w:insDel", "1");
  revisionView.setAttributeNS(WORD_NS, "w:formatting", "1");
  revisionView.setAttributeNS(WORD_NS, "w:comments", "1");

  zip.file(settingsPath, new XMLSerializer().serializeToString(xmlDoc));
}

function writeSegment(segment) {
  segment.node.textContent = segment.text;
  if (/^\s|\s$/u.test(segment.text)) {
    segment.node.setAttribute("xml:space", "preserve");
  } else {
    segment.node.removeAttribute("xml:space");
  }
}

function buildWarnings(captions, references, captionMap, headingInfo, compiledRules, pageInfo) {
  const warnings = [];

  for (const [, mapped] of captionMap.entries()) {
    if (mapped.length < 2) {
      continue;
    }
    if (new Set(mapped.map((item) => item.newText)).size > 1) {
      warnings.push(t("duplicateWarning", { text: mapped[0].text }));
    }
  }

  const usedHeadingPlaceholders = getUsedHeadingPlaceholders(compiledRules);
  for (const placeholder of usedHeadingPlaceholders) {
    if ((headingInfo.counts[placeholder] ?? 0) === 0) {
      warnings.push(t("missingHeadingWarning", { placeholder: `{${placeholder}}` }));
    }

    const fallbackCount = captions.filter((item) =>
      Array.isArray(item.placeholderFallbacks) && item.placeholderFallbacks.includes(placeholder),
    ).length;

    if (fallbackCount > 0) {
      warnings.push(
        t("placeholderFallbackWarning", {
          count: fallbackCount,
          placeholder: `{${placeholder}}`,
        }),
      );
    }
  }

  if (!captions.length) {
    warnings.push(t("noCaptionFound"));
  }
  if (captions.length && !references.length) {
    warnings.push(t("noReferenceFound"));
  }

  return warnings;
}

function getUsedHeadingPlaceholders(compiledRules) {
  return [...new Set(compiledRules.flatMap((rule) => rule.headingPlaceholders))];
}

function hasHeadingPlaceholders(compiledRules) {
  return getUsedHeadingPlaceholders(compiledRules).length > 0;
}

function buildHeadingSummary(headingCounts, compiledRules) {
  const placeholders = getUsedHeadingPlaceholders(compiledRules).sort(
    (left, right) => getPlaceholderSortKey(left) - getPlaceholderSortKey(right),
  );

  if (!placeholders.length) {
    return "";
  }

  const parts = placeholders.map((placeholder) => {
    const count = headingCounts[placeholder] ?? 0;
    return state.language === "en"
      ? `${placeholder}=${count}`
      : `${placeholder}=${count}`;
  });

  return state.language === "en"
    ? `, detected heading counts: ${parts.join(", ")}`
    : `；检测到标题层级计数：${parts.join("、")}`;
}



function buildCustomRangeSummary(customRange) {
  if (!customRange) {
    return "";
  }

  const { start, end } = customRange;
  const startStr = start !== null ? start : "?";
  const endStr = end !== null ? end : "?";

  return t("customRangeLabel") + "：" + t("customRangeValue", { start: startStr, end: endStr });
}

function renderPreview(preview) {
  setText(elements.captionsCount, String(preview.captions.length));
  setText(elements.referencesCount, String(preview.references.length));
  setText(elements.warningsCount, String(preview.warnings.length));
  renderRows(elements.captionsTableBody, preview.captions, t("emptyCaptionsRow"));
  renderRows(elements.referencesTableBody, preview.references, t("emptyReferencesRow"));
  renderHeadingReview(preview.headingGroups, preview.usedHeadingPlaceholders);

  if (!elements.warningsList) {
    return;
  }

  elements.warningsList.innerHTML = preview.warnings.length
    ? preview.warnings.map((warning) => `<li>${escapeHtml(warning)}</li>`).join("")
    : `<li>${escapeHtml(t("noExtraWarnings"))}</li>`;
}

function renderRows(target, rows, emptyText) {
  if (!target) {
    return;
  }

  if (!rows.length) {
    target.innerHTML = `<tr><td colspan="3" class="empty-row">${escapeHtml(emptyText)}</td></tr>`;
    return;
  }

  target.innerHTML = rows
    .slice(0, 200)
    .map(
      (item) => `
        <tr style="${getRuleAccentStyle(item)}">
          <td>${escapeHtml(item.text)}</td>
          <td>${escapeHtml(item.newText)}</td>
          <td class="excerpt">${renderHighlightedExcerpt(item)}</td>
        </tr>
      `,
    )
    .join("");
}

function renderHeadingReview(groups = {}, usedPlaceholders = []) {
  if (!elements.headingReviewList) {
    return;
  }

  elements.resetHeadingFiltersButton?.classList.toggle("is-hidden", state.headingExclusions.size === 0);

  const placeholders = [...new Set(usedPlaceholders.filter(isHeadingPlaceholder))].sort(
    (left, right) => getPlaceholderSortKey(left) - getPlaceholderSortKey(right),
  );

  if (!placeholders.length) {
    elements.headingReviewList.innerHTML = `<div class="heading-review-empty">${escapeHtml(t("headingReviewUnused"))}</div>`;
    return;
  }

  elements.headingReviewList.innerHTML = placeholders
    .map((placeholder) => {
      const meta = getPlaceholderMeta(placeholder);
      const items = groups[placeholder] ?? [];
      const visual = getPlaceholderVisual(placeholder);

      const rows = items.length
        ? items
          .map(
            (item) => `
                <li class="heading-review-item">
                  <div class="heading-review-item__main">
                    <span class="heading-review-item__index">${escapeHtml(String(item.sequence))}</span>
                    <span class="heading-review-item__text">${escapeHtml(
              normalizePreviewWhitespace(item.text) || t("emptyParagraph"),
            )}</span>
                  </div>
                  <button
                    class="button button--ghost heading-review-item__remove"
                    type="button"
                    data-action="remove-heading"
                    data-heading-order="${escapeHtml(String(item.order))}"
                  >
                    ${escapeHtml(t("headingReviewRemove"))}
                  </button>
                </li>
              `,
          )
          .join("")
        : `<div class="heading-review-empty">${escapeHtml(t("headingReviewEmpty"))}</div>`;

      return `
        <article class="heading-review-group">
          <div class="heading-review-group__header">
            <div class="heading-review-group__title-row">
              <span class="pattern-chip pattern-chip--heading" style="${visual.chipStyle}">${escapeHtml(`{${placeholder}}`)}</span>
              <h4 class="heading-review-group__title">${escapeHtml(
        t("headingReviewGroupTitle", {
          placeholder: `{${placeholder}}`,
          level: meta?.level ?? "?",
        }),
      )}</h4>
            </div>
            <span class="heading-review-group__count">${escapeHtml(
        t("headingReviewCount", { count: items.length }),
      )}</span>
          </div>
          ${items.length ? `<ul class="heading-review-group__list">${rows}</ul>` : rows}
        </article>
      `;
    })
    .join("");
}

function downloadProcessedDocument(blob, originalName, suffixKey = "outputFileSuffix") {
  const suffixIndex = originalName.toLowerCase().endsWith(".docx")
    ? originalName.length - 5
    : originalName.length;
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${originalName.slice(0, suffixIndex)}${t(suffixKey)}.docx`;
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function setStatus(message, level = "info", options = {}) {
  state.statusLevel = level;
  state.statusKey = options.isRaw ? null : message;
  state.statusVars = options.isRaw ? {} : options;

  if (!elements.statusBox) {
    return;
  }

  elements.statusBox.textContent = options.isRaw ? message : t(message, options);
  elements.statusBox.style.background =
    level === "warn" ? "rgba(199, 109, 44, 0.08)" : "rgba(15, 108, 91, 0.07)";
  elements.statusBox.style.borderColor =
    level === "warn" ? "rgba(199, 109, 44, 0.18)" : "rgba(15, 108, 91, 0.14)";
  elements.statusBox.style.color = level === "warn" ? "#7b4b24" : "var(--accent-strong)";
}

function syncStatus() {
  if (state.statusKey) {
    setStatus(state.statusKey, state.statusLevel, state.statusVars);
  }
}

function sanitizeStartAt(value) {
  const parsed = Number.parseInt(String(value), 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : 1;
}

function parseOptionalPositiveInteger(value) {
  const normalized = String(value ?? "").trim();
  if (!normalized) {
    return null;
  }

  const parsed = Number.parseInt(normalized, 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : Number.NaN;
}

function makeMappingKey(ruleId, originalText) {
  return `${ruleId}::${originalText}`;
}

function renderPatternOverlay(pattern) {
  return renderPatternTokens(pattern, { mode: "overlay", allowPlaceholders: true });
}

function renderPatternInline(pattern) {
  return renderPatternTokens(pattern, { mode: "inline", allowPlaceholders: true });
}

function renderPatternTokens(pattern, { mode, allowPlaceholders }) {
  const tokens = tokenizePattern(pattern, { allowPlaceholders });
  return tokens
    .map((token) => {
      if (mode === "inline") {
        if (token.type === "literal") {
          return `<span>${escapeHtml(token.text)}</span>`;
        }

        if (token.type === "regex") {
          return `<span class="pattern-chip pattern-chip--regex">${escapeHtml(token.text)}</span>`;
        }

        const visual = getPlaceholderVisual(token.name);
        return `<span class="pattern-chip ${getPatternChipClass(token.name)}" style="${visual.chipStyle}">${escapeHtml(token.text)}</span>`;
      }

      if (token.type === "placeholder") {
        const visual = getPlaceholderVisual(token.name);
        return `<span class="${getPatternTokenClass(token)}" style="${visual.overlayStyle}">${escapeHtml(token.text)}</span>`;
      }

      return `<span class="${getPatternTokenClass(token)}">${escapeHtml(token.text)}</span>`;
    })
    .join("");
}

function getPatternTokenClass(token) {
  if (token.type === "regex") {
    return "pattern-overlay__token pattern-overlay__token--regex";
  }
  if (token.type === "placeholder") {
    return `pattern-overlay__token pattern-overlay__token--placeholder pattern-overlay__token--${token.name}`;
  }
  return "pattern-overlay__token pattern-overlay__token--literal";
}

function getPatternChipClass(name) {
  if (name === "n") {
    return "pattern-chip--renumber";
  }
  if (isHeadingPlaceholder(name)) {
    return "pattern-chip--heading";
  }
  return "pattern-chip--regex";
}

function getPlaceholderVisual(name) {
  if (name === "n") {
    return {
      chipStyle: "color:#8c4d18;background:rgba(180, 83, 9, 0.13);border-color:rgba(180, 83, 9, 0.18);",
      overlayStyle: "color:#8c4d18;background:rgba(180, 83, 9, 0.13);",
      tagStyle: "color:#8c4d18;background:rgba(180, 83, 9, 0.13);",
    };
  }

  const meta = getPlaceholderMeta(name);
  const level = meta?.level ?? 1;
  const palette = [
    { color: "#0f5e57", bg: "rgba(15, 118, 110, 0.13)", border: "rgba(15, 118, 110, 0.18)" },
    { color: "#6a21a8", bg: "rgba(126, 34, 206, 0.12)", border: "rgba(126, 34, 206, 0.16)" },
    { color: "#18427b", bg: "rgba(29, 78, 216, 0.12)", border: "rgba(29, 78, 216, 0.16)" },
    { color: "#9a3412", bg: "rgba(234, 88, 12, 0.12)", border: "rgba(234, 88, 12, 0.16)" },
  ][(Math.max(level, 1) - 1) % 4];

  return {
    chipStyle: `color:${palette.color};background:${palette.bg};border-color:${palette.border};`,
    overlayStyle: `color:${palette.color};background:${palette.bg};`,
    tagStyle: `color:${palette.color};background:${palette.bg};`,
  };
}

function tokenizePattern(pattern, { allowPlaceholders } = {}) {
  if (!pattern) {
    return [];
  }

  if (!allowPlaceholders) {
    return tokenizeRegexAndLiteral(pattern);
  }

  const tokens = [];
  let cursor = 0;

  for (const match of pattern.matchAll(/\{([a-z]\d*)\}/g)) {
    const start = match.index ?? 0;
    if (start > cursor) {
      tokens.push(...tokenizeRegexAndLiteral(pattern.slice(cursor, start)));
    }
    tokens.push({ type: "placeholder", name: match[1], text: match[0] });
    cursor = start + match[0].length;
  }

  if (cursor < pattern.length) {
    tokens.push(...tokenizeRegexAndLiteral(pattern.slice(cursor)));
  }

  return tokens;
}

function tokenizeRegexAndLiteral(segment) {
  const tokens = [];
  const regexToken = /\([^)]*\)|\[[^\]]*\]|\\.[+*?]?/g;
  let cursor = 0;

  for (const match of segment.matchAll(regexToken)) {
    const start = match.index ?? 0;
    if (start > cursor) {
      tokens.push({ type: "literal", text: segment.slice(cursor, start) });
    }
    tokens.push({ type: "regex", text: match[0] });
    cursor = start + match[0].length;
  }

  if (cursor < segment.length) {
    tokens.push({ type: "literal", text: segment.slice(cursor) });
  }

  return tokens.filter((token) => token.text.length > 0);
}

function syncPatternOverlay(input, pattern) {
  const shell = input.closest(".input-overlay-shell");
  const overlay = shell?.querySelector(".input-overlay");
  if (!(overlay instanceof HTMLDivElement)) {
    return;
  }

  if (shell?.classList.contains("is-editing")) {
    overlay.innerHTML = "";
    return;
  }

  overlay.innerHTML = pattern ? renderPatternOverlay(pattern) : "";
}

function openHelpDialog() {
  if (!elements.helpDialog) {
    return;
  }
  elements.helpDialog.hidden = false;
  document.body.style.overflow = "hidden";
}

function closeHelpDialog() {
  if (!elements.helpDialog) {
    return;
  }
  elements.helpDialog.hidden = true;
  document.body.style.overflow = "";
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function renderHighlightedExcerpt(item) {
  const sourceText = item.paragraph?.fullText ?? item.excerpt ?? "";
  if (!sourceText) {
    return t("emptyParagraph");
  }

  const maxContext = 42;
  const start = Math.max(0, item.start ?? 0);
  const end = Math.min(sourceText.length, item.end ?? start);
  const excerptStart = Math.max(0, start - maxContext);
  const excerptEnd = Math.min(sourceText.length, end + maxContext);
  const prefixEllipsis = excerptStart > 0 ? "..." : "";
  const suffixEllipsis = excerptEnd < sourceText.length ? "..." : "";

  const before = normalizePreviewWhitespace(sourceText.slice(excerptStart, start));
  const highlighted = normalizePreviewWhitespace(sourceText.slice(start, end));
  const after = normalizePreviewWhitespace(sourceText.slice(end, excerptEnd));

  return `${escapeHtml(prefixEllipsis + before)}<mark>${escapeHtml(
    highlighted || item.text,
  )}</mark>${escapeHtml(after + suffixEllipsis)}`;
}

function getRuleAccentStyle(item) {
  const index = Math.max(0, state.rules.findIndex((rule) => rule.id === item.rule?.id));
  const palette = RULE_ACCENT_PALETTE[index % RULE_ACCENT_PALETTE.length];
  return `--excerpt-mark-bg:${palette.bg};--excerpt-mark-color:${palette.color};--excerpt-mark-ring:${palette.ring};`;
}

function normalizePreviewWhitespace(value) {
  return value.replace(/\s+/g, " ");
}

function collapseWhitespace(value) {
  return String(value).replace(/\s+/g, " ").trim();
}

function formatFileSize(size) {
  if (size < 1024) {
    return `${size} B`;
  }
  if (size < 1024 * 1024) {
    return `${(size / 1024).toFixed(1)} KB`;
  }
  return `${(size / (1024 * 1024)).toFixed(2)} MB`;
}
