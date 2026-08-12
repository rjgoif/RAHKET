; ============================================================================
; MODULE: PartyMacro
; Written by Reece J. Goiffon, MD, PhD
; AutoHotkey v2
; Can be run standalone or included in RAHKET_Main.ahk
;
; Converts a HARM-templated report into a "party macro" style report:
; auto-detects exam type from the pasted report header, applies universal +
; exam-specific cleanup rules, standardizes measurement units/precision,
; reflows long sections, shows a track-changes style diff, and puts the
; result on the clipboard as BOTH plain text and RTF (with the party
; template's real PowerScribe fields intact) so a normal paste in
; PowerScribe reconstructs live fields.
;
; The field-substitution and clipboard mechanism here are ported directly
; from the working pattern in ThyroidNodules.ahk (TN_XmlSetDefault /
; TN_SetClipboardTextAndRTF) rather than reinvented.
;
; Folder layout (all under this module's own folder):
;   PartyMacro\Rules\universal_rules.json   - cleanup rules run for every exam
;   PartyMacro\Configs\<Exam>.json          - one config per exam type
;   PartyMacro\Templates\<Exam>.rtf         - the real party autotext RTF+XML
; ============================================================================

#Requires AutoHotkey v2.0

if (A_LineFile = A_ScriptFullPath) {
    #SingleInstance Force
    Persistent

    A_TrayMenu.Delete()
    A_TrayMenu.Add("Party Macro", (*) => Show_PartyMacro())
    A_TrayMenu.Add("Exit", (*) => ExitApp())
}

; This module's own directory, independent of whether it's included or run
; standalone (A_ScriptDir would point at RAHKET's root dir when included).
PartyMacro_SelfDir := ""
SplitPath(A_LineFile, , &PartyMacro_SelfDir)
PartyMacro_ConfigDir   := PartyMacro_SelfDir "\PartyMacro\Configs"
PartyMacro_TemplateDir := PartyMacro_SelfDir "\PartyMacro\Templates"
PartyMacro_RulesFile    := PartyMacro_SelfDir "\PartyMacro\Rules\universal_rules.json"


; ============================================================================
; MODULE CODE
; ============================================================================

; ---------------------------------------------------------------------------
; Minimal self-contained JSON parser (config files only). Kept separate from
; any other RAHKET JSON handling since that code isn't available here.
; ---------------------------------------------------------------------------
class PartyMacro_JSON {
    static Parse(text) {
        this.str := text
        this.pos := 1
        this.SkipWs()
        return this.ParseValue()
    }

    static SkipWs() {
        len := StrLen(this.str)
        while (this.pos <= len && InStr(" `t`r`n", SubStr(this.str, this.pos, 1)))
            this.pos++
    }

    static ParseValue() {
        this.SkipWs()
        c := SubStr(this.str, this.pos, 1)
        if (c = "{")
            return this.ParseObject()
        if (c = "[")
            return this.ParseArray()
        if (c = "`"")
            return this.ParseString()
        if (c = "t" || c = "f")
            return this.ParseBool()
        if (c = "n") {
            this.pos += 4
            return ""
        }
        return this.ParseNumber()
    }

    static ParseObject() {
        obj := Map()
        this.pos++
        this.SkipWs()
        if (SubStr(this.str, this.pos, 1) = "}") {
            this.pos++
            return obj
        }
        loop {
            this.SkipWs()
            key := this.ParseString()
            this.SkipWs()
            this.pos++ ; colon
            val := this.ParseValue()
            obj[key] := val
            this.SkipWs()
            c := SubStr(this.str, this.pos, 1)
            this.pos++
            if (c = "}")
                break
        }
        return obj
    }

    static ParseArray() {
        arr := []
        this.pos++
        this.SkipWs()
        if (SubStr(this.str, this.pos, 1) = "]") {
            this.pos++
            return arr
        }
        loop {
            arr.Push(this.ParseValue())
            this.SkipWs()
            c := SubStr(this.str, this.pos, 1)
            this.pos++
            if (c = "]")
                break
        }
        return arr
    }

    static ParseString() {
        this.pos++ ; opening quote
        out := ""
        loop {
            c := SubStr(this.str, this.pos, 1)
            if (c = "`"") {
                this.pos++
                break
            }
            if (c = "\") {
                nc := SubStr(this.str, this.pos + 1, 1)
                switch nc {
                    case "n": out .= "`n"
                    case "t": out .= "`t"
                    case "r": out .= "`r"
                    case "`"": out .= "`""
                    case "\": out .= "\"
                    case "/": out .= "/"
                    default:  out .= nc
                }
                this.pos += 2
            } else {
                out .= c
                this.pos++
            }
        }
        return out
    }

    static ParseNumber() {
        start := this.pos
        while (InStr("-+0123456789.eE", SubStr(this.str, this.pos, 1)))
            this.pos++
        return SubStr(this.str, start, this.pos - start) + 0
    }

    static ParseBool() {
        if (SubStr(this.str, this.pos, 4) = "true") {
            this.pos += 4
            return true
        }
        this.pos += 5
        return false
    }
}


; ---------------------------------------------------------------------------
; Config / rules loading
; ---------------------------------------------------------------------------

; Returns array of {name, path, config} for every config in Configs\
PartyMacro_ScanConfigs() {
    global PartyMacro_ConfigDir
    result := []
    if !DirExist(PartyMacro_ConfigDir)
        return result
    loop files, PartyMacro_ConfigDir "\*.json" {
        try {
            cfg := PartyMacro_JSON.Parse(FileRead(A_LoopFileFullPath, "UTF-8"))
            displayName := cfg.Has("name") ? cfg["name"] : A_LoopFileName
            result.Push({ name: displayName, path: A_LoopFileFullPath, config: cfg })
        }
    }
    return result
}

; Flattens the universal_rules.json category structure into one ordered list
PartyMacro_LoadUniversalRules() {
    global PartyMacro_RulesFile
    if !FileExist(PartyMacro_RulesFile)
        return []
    data := PartyMacro_JSON.Parse(FileRead(PartyMacro_RulesFile, "UTF-8"))
    flat := []
    if !data.Has("categories")
        return flat
    for cat in data["categories"] {
        for rule in cat["rules"]
            flat.Push(rule)
    }
    return flat
}

; Tries each loaded config's examMatch patterns against the pasted text.
; Returns the matching entry from configList, or "" if none matched.
PartyMacro_DetectConfig(sourceText, configList) {
    for entry in configList {
        cfg := entry.config
        if !cfg.Has("examMatch")
            continue
        for pat in cfg["examMatch"]["patterns"] {
            if RegExMatch(sourceText, "i)" pat)
                return entry
        }
    }
    return ""
}

; Grabs the leading modality token from the report's first line (e.g. "CT"
; from "CT ABDOMEN/PELVIS WITH CONTRAST"). Best-effort heuristic based on
; MGB's header convention -- not a guaranteed match for every report type.
PartyMacro_DetectModality(sourceText) {
    lines := StrSplit(sourceText, "`n", "`r")
    firstLine := lines.Length ? lines[1] : ""
    if RegExMatch(firstLine, "i)^\s*([A-Za-z]{2,4})\b", &m)
        return StrUpper(m[1])
    return ""
}


; ---------------------------------------------------------------------------
; Generic regex-replace-with-callback helper (AHK's RegExReplace has no
; callback support, so this fills that gap for measurement conversion)
; ---------------------------------------------------------------------------
PartyMacro_RegExReplaceCB(haystack, needle, cb) {
    result := ""
    pos := 1
    len := StrLen(haystack)
    while (foundPos := RegExMatch(haystack, needle, &m, pos)) {
        result .= SubStr(haystack, pos, foundPos - pos)
        result .= cb(m)
        matchLen := m.Len(0)
        pos := foundPos + (matchLen > 0 ? matchLen : 1)
        if (pos > len + 1)
            break
    }
    result .= SubStr(haystack, pos)
    return result
}


; ---------------------------------------------------------------------------
; Cleanup rules (literal / regex find-replace)
; ---------------------------------------------------------------------------
PartyMacro_ApplyCleanupRules(text, rules) {
    for rule in rules {
        if (rule["type"] = "literal")
            text := StrReplace(text, rule["find"], rule["replace"])
        else if (rule["type"] = "regex")
            text := RegExReplace(text, rule["find"], rule["replace"])
    }
    return text
}


; ---------------------------------------------------------------------------
; Measurement standardization: converts cm<->mm to the config's target unit
; and harmonizes decimal precision within a dimension pair (replaces what
; used to be two separate manual warnings with an automatic fix).
;   cm target -> 1 decimal place
;   mm target -> whole number (1 decimal only if value < 1)
; ---------------------------------------------------------------------------
PartyMacro_FormatMeasurement(val, unit) {
    if (unit = "mm")
        return (val < 1) ? Format("{:.1f}", val) : Format("{:d}", Round(val))
    return Format("{:.1f}", val)
}

PartyMacro_ConvertVal(val, fromUnit, toUnit) {
    if (fromUnit = toUnit)
        return val
    if (fromUnit = "cm" && toUnit = "mm")
        return val * 10
    if (fromUnit = "mm" && toUnit = "cm")
        return val / 10
    return val
}

PartyMacro_ProcessMeasurements(text, targetUnit) {
    needle := "i)(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)\s*(cm|mm)\b|(\d+(?:\.\d+)?)\s*(cm|mm)\b"
    cb := (m) => PartyMacro_MeasurementCallback(m, targetUnit)
    return PartyMacro_RegExReplaceCB(text, needle, cb)
}

PartyMacro_MeasurementCallback(m, targetUnit) {
    if (m.Value(1) != "") {
        fromUnit := StrLower(m.Value(3))
        a := PartyMacro_ConvertVal(m.Value(1) + 0, fromUnit, targetUnit)
        b := PartyMacro_ConvertVal(m.Value(2) + 0, fromUnit, targetUnit)
        return PartyMacro_FormatMeasurement(a, targetUnit) " x " PartyMacro_FormatMeasurement(b, targetUnit) " " targetUnit
    }
    fromUnit := StrLower(m.Value(5))
    a := PartyMacro_ConvertVal(m.Value(4) + 0, fromUnit, targetUnit)
    return PartyMacro_FormatMeasurement(a, targetUnit) " " targetUnit
}


; ---------------------------------------------------------------------------
; Section extraction / reflow
; ---------------------------------------------------------------------------
PartyMacro_ExtractSection(sourceText, label) {
    needle := "im)^" PartyMacro_EscapeRegex(label) ":[ \t]*(.*)$"
    if !RegExMatch(sourceText, needle, &m)
        return ""
    return { whole: m.Value(0), content: Trim(m.Value(1)) }
}

; sourceLabel in config may be a single string or an array of fallback
; labels tried in order (e.g. "Lungs" then "Lungs/airways") -- mirrors the
; fallback pattern in your own QM script.
PartyMacro_ExtractSectionMulti(sourceText, labelOrList) {
    if (Type(labelOrList) = "Array") {
        for lbl in labelOrList {
            r := PartyMacro_ExtractSection(sourceText, lbl)
            if (r != "" && Trim(r.content) != "")
                return r
        }
        return ""
    }
    return PartyMacro_ExtractSection(sourceText, labelOrList)
}

PartyMacro_EscapeRegex(s) {
    return RegExReplace(s, "([.^$|()\[\]{}*+?\\/])", "\$1")
}

PartyMacro_CountMatches(text, needle) {
    count := 0
    pos := 1
    while (foundPos := RegExMatch(text, needle, &m, pos)) {
        count++
        pos := foundPos + Max(m.Len(0), 1)
    }
    return count
}

PartyMacro_ReflowSection(text, threshold) {
    static splitRE := "(?<!\d\.)(?<=\d|[a-zA-Z])([\)\].!?]{1,2})\s+(?=[A-Z\d])"
    if (PartyMacro_CountMatches(text, splitRE) >= threshold)
        return RegExReplace(text, splitRE, "$1`n")
    return text
}


; ---------------------------------------------------------------------------
; Word-level diff (LCS) for the track-changes view
; ---------------------------------------------------------------------------
PartyMacro_SplitWords(text) {
    words := []
    for w in StrSplit(text, " ") {
        if (w != "")
            words.Push(w)
    }
    return words
}

PartyMacro_WordDiff(origText, finalText) {
    a := PartyMacro_SplitWords(origText)
    b := PartyMacro_SplitWords(finalText)
    n := a.Length, m := b.Length

    dp := []
    loop n + 1
        dp.Push(PartyMacro_ZeroArray(m + 1))

    loop n {
        i := A_Index
        loop m {
            j := A_Index
            if (a[i] = b[j])
                dp[i + 1][j + 1] := dp[i][j] + 1
            else
                dp[i + 1][j + 1] := Max(dp[i][j + 1], dp[i + 1][j])
        }
    }

    ops := []
    i := n, j := m
    while (i > 0 || j > 0) {
        if (i > 0 && j > 0 && a[i] = b[j]) {
            ops.InsertAt(1, { type: "same", word: a[i] })
            i--, j--
        } else if (j > 0 && (i = 0 || dp[i + 1][j] >= dp[i][j + 1])) {
            ops.InsertAt(1, { type: "add", word: b[j] })
            j--
        } else {
            ops.InsertAt(1, { type: "del", word: a[i] })
            i--
        }
    }
    return ops
}

PartyMacro_ZeroArray(len) {
    arr := []
    loop len
        arr.Push(0)
    return arr
}

PartyMacro_HtmlEscape(s) {
    s := StrReplace(s, "&", "&amp;")
    s := StrReplace(s, "<", "&lt;")
    s := StrReplace(s, ">", "&gt;")
    return s
}

PartyMacro_DiffToHtml(ops) {
    html := ""
    for op in ops {
        word := PartyMacro_HtmlEscape(op.word)
        if (op.type = "del") {
            html .= '<span class="del">' word '</span> '
        } else {
            isSimilar := RegExMatch(op.word, "i)^similar")
            cls := (op.type = "add") ? "add" : ""
            if (isSimilar)
                cls := Trim(cls " similar")
            if (cls != "")
                html .= '<span class="' cls '">' word '</span> '
            else
                html .= word ' '
        }
    }
    return html
}


; ---------------------------------------------------------------------------
; RTF field template handling (ported pattern from ThyroidNodules.ahk)
; ---------------------------------------------------------------------------

; Escapes text for insertion into the XML <defaultvalue> element
PartyMacro_XmlEscape(str) {
    str := StrReplace(str, "&", "&amp;")
    str := StrReplace(str, "<", "&lt;")
    str := StrReplace(str, ">", "&gt;")
    return str
}

; Replaces the <defaultvalue>...</defaultvalue> for a given <name>fieldName</name>
; in the template's trailing XML block. Does NOT touch start/length or the
; visible RTF body -- PowerScribe reconstructs fields by name, not offset.
PartyMacro_XmlSetDefault(rtf, fieldName, newVal) {
    newVal := PartyMacro_XmlEscape(newVal)
    pattern := "(<name>" PartyMacro_EscapeRegex(fieldName) "</name><defaultvalue>)(.*?)(</defaultvalue>)"
    return RegExReplace(rtf, pattern, "$1" newVal "$3")
}

; Strips a field's entire <field>...</field> block from the trailing XML,
; matched by <name>. Used for conditionalOmit sections (e.g. Coronary).
PartyMacro_RemoveXmlField(rtf, fieldName) {
    pattern := "<field[^>]*><name>" PartyMacro_EscapeRegex(fieldName) "</name>.*?</field>"
    return RegExReplace(rtf, pattern, "")
}

; Strips one literal line/fragment from the visible RTF body. The caller
; supplies the exact text (from inspecting the real template file) since
; RTF body layout isn't derivable generically from a field name alone.
PartyMacro_RemoveRtfLine(rtf, literalLine) {
    return StrReplace(rtf, literalLine, "")
}

; Clipboard write: CF_UNICODETEXT + registered "Rich Text Format", ported
; directly from TN_SetClipboardTextAndRTF in ThyroidNodules.ahk.
PartyMacro_SetClipboardTextAndRTF(plain, rtf) {
    if !DllCall("OpenClipboard", "ptr", 0, "int") {
        MsgBox "Could not open clipboard."
        return
    }

    DllCall("EmptyClipboard")

    ; --- CF_UNICODETEXT ---
    lenW := (StrLen(plain) + 1) * 2
    hText := DllCall("GlobalAlloc", "uint", 0x2, "uptr", lenW, "ptr")
    if (hText) {
        pText := DllCall("GlobalLock", "ptr", hText, "ptr")
        StrPut(plain, pText, "UTF-16")
        DllCall("GlobalUnlock", "ptr", hText)
        DllCall("SetClipboardData", "uint", 13, "ptr", hText)  ; CF_UNICODETEXT
    }

    ; --- Rich Text Format (RTF) ---
    cfRtf := DllCall("RegisterClipboardFormat", "str", "Rich Text Format", "uint")
    lenA := StrLen(rtf) + 1
    hRtf := DllCall("GlobalAlloc", "uint", 0x2, "uptr", lenA, "ptr")
    if (hRtf) {
        pRtf := DllCall("GlobalLock", "ptr", hRtf, "ptr")
        StrPut(rtf, pRtf, "CP0")
        DllCall("GlobalUnlock", "ptr", hRtf)
        DllCall("SetClipboardData", "uint", cfRtf, "ptr", hRtf)
    }

    DllCall("CloseClipboard")
}


; ---------------------------------------------------------------------------
; Core pipeline
; ---------------------------------------------------------------------------
PartyMacro_Process(sourceText, config, universalRules) {
    targetUnit := config.Has("measurementUnit") ? config["measurementUnit"] : "cm"
    threshold  := config.Has("reflowThreshold") ? config["reflowThreshold"] : 3

    allRules := []
    for r in universalRules
        allRules.Push(r)
    if config.Has("additionalRules") {
        for r in config["additionalRules"]
            allRules.Push(r)
    }

    ; Pass 1: extract + clean all "extract" mode sections
    fieldText := Map()      ; outputField -> final cleaned text
    sectionResults := []    ; for diff display, in config order

    for sec in config["sections"] {
        if (sec["mode"] != "extract")
            continue

        extracted := PartyMacro_ExtractSectionMulti(sourceText, sec["sourceLabel"])
        if (extracted = "") {
            sectionResults.Push({ label: sec["outputField"], found: false, diffHtml: "", flag: "" })
            continue
        }

        cleaned := PartyMacro_ApplyCleanupRules(extracted.content, allRules)
        cleaned := PartyMacro_ProcessMeasurements(cleaned, targetUnit)
        cleaned := Trim(RegExReplace(cleaned, " {2,}", " "))
        reflowed := PartyMacro_ReflowSection(cleaned, threshold)

        fieldText[sec["outputField"]] := StrReplace(reflowed, "`n", " ")

        diffOps := PartyMacro_WordDiff(extracted.content, reflowed)
        sectionResults.Push({ label: sec["outputField"], found: true, diffHtml: PartyMacro_DiffToHtml(diffOps), flag: "" })
    }

    ; Pass 2: keywordScan sections (e.g. Stones) -- flag only, don't overwrite
    for sec in config["sections"] {
        if (sec["mode"] != "keywordScan")
            continue

        scanText := fieldText.Has(sec["scanField"]) ? fieldText[sec["scanField"]] : ""
        hit := ""
        for kw in sec["keywords"] {
            if RegExMatch(scanText, "i)\b" kw) {
                hit := kw
                break
            }
        }

        if (hit != "") {
            sectionResults.Push({ label: sec["outputField"], found: true,
                diffHtml: '<span class="flag">Possible ' PartyMacro_HtmlEscape(sec["outputField"]) '-related language ("' PartyMacro_HtmlEscape(hit) '") found in ' sec["scanField"] ' -- review and move manually. Field left at template default.</span>',
                flag: hit })
        } else {
            sectionResults.Push({ label: sec["outputField"], found: true,
                diffHtml: '<span class="missing">(no ' PartyMacro_HtmlEscape(sec["outputField"]) '-related language detected -- left at template default)</span>', flag: "" })
        }
        ; not written to fieldText -- template default stays in the RTF
    }

    ; Pass 3: conditionalOmit sections (e.g. Coronary) -- if keywords are
    ; found, the field gets stripped entirely at copy time. If not found,
    ; nothing happens: the field is simply never touched, so it stays in
    ; the output as its original live picklist/default.
    omitFields := []
    for sec in config["sections"] {
        if (sec["mode"] != "conditionalOmit")
            continue

        scanText := (sec.Has("scanWholeSource") && sec["scanWholeSource"])
            ? sourceText
            : (fieldText.Has(sec["scanField"]) ? fieldText[sec["scanField"]] : "")

        hit := ""
        for kw in sec["keywords"] {
            if InStr(scanText, kw) {
                hit := kw
                break
            }
        }

        if (hit != "") {
            omitFields.Push({ outputField: sec["outputField"], rtfLine: sec["rtfLineToRemove"] })
            sectionResults.Push({ label: sec["outputField"], found: true,
                diffHtml: '<span class="flag">"' PartyMacro_HtmlEscape(hit) '" found in report -- field omitted from party output.</span>' })
        } else {
            sectionResults.Push({ label: sec["outputField"], found: true,
                diffHtml: '<span class="missing">(no matching language found -- field kept as-is, untouched)</span>' })
        }
    }

    return { fieldText: fieldText, sections: sectionResults, omitFields: omitFields }
}


; ---------------------------------------------------------------------------
; HTML assembly for the ActiveX diff view
; ---------------------------------------------------------------------------
PartyMacro_BuildDiffDocument(result) {
    html := '<html><head><meta charset="utf-8"><style>'
        . 'body{font-family:Segoe UI,Arial,sans-serif;font-size:13px;margin:8px;}'
        . 'h4{margin:10px 0 2px 0;}'
        . '.del{color:#b00020;text-decoration:line-through;background:#ffe5e5;}'
        . '.add{color:#0a7d33;background:#e6ffe9;}'
        . '.similar{background:#fff3b0;}'
        . '.flag{background:#ffd9a0;border:1px solid #cc8400;padding:2px 4px;}'
        . '.missing{color:#888;font-style:italic;}'
        . '</style></head><body>'

    for r in result.sections {
        html .= '<h4>' PartyMacro_HtmlEscape(r.label) '</h4>'
        if !r.found {
            html .= '<div class="missing">(section not found in source text)</div>'
            continue
        }
        html .= '<div>' r.diffHtml '</div>'
    }

    html .= '</body></html>'
    return html
}

PartyMacro_SetWebHtml(webCtrl, html) {
    webCtrl.Value.Navigate("about:blank")
    start := A_TickCount
    while (webCtrl.Value.ReadyState != 4 && (A_TickCount - start) < 2000)
        Sleep(10)
    doc := webCtrl.Value.Document
    doc.Write(html)
    doc.Close()
}


; ---------------------------------------------------------------------------
; GUI
; ---------------------------------------------------------------------------
Show_PartyMacro() {
    global PartyMacro_TemplateDir

    configList := PartyMacro_ScanConfigs()
    if (configList.Length = 0) {
        MsgBox("No config files found.", "Party Macro", 48)
        return
    }
    universalRules := PartyMacro_LoadUniversalRules()

    names := []
    for entry in configList
        names.Push(entry.name)

    pmGui := Gui("+Resize", "Party Macro")
    pmGui.SetFont("s10", "Segoe UI")

    statusText := pmGui.Add("Text", "xm y10 w780", "Paste the full report (including header) below, then click Process. Exam type will be auto-detected.")

    pmGui.Add("Text", "xm y+10", "Exam type (auto-detected -- override if wrong):")
    ddl := pmGui.Add("DropDownList", "xm y+2 w300", names)
    ddl.Choose(1)

    pmGui.Add("Text", "xm y+10", "Paste full report text:")
    srcEdit := pmGui.Add("Edit", "xm y+2 w780 h150")

    btnProcess := pmGui.Add("Button", "xm y+10 w150", "Process")

    pmGui.Add("Text", "xm y+12", "Changes -- red/strikethrough = removed, green = added, yellow = flagged term, orange = needs manual review:")
    webCtrl := pmGui.Add("ActiveX", "xm y+2 w780 h230", "Shell.Explorer.2")

    btnCopy := pmGui.Add("Button", "xm y+10 w220", "Copy to Clipboard (Text + RTF Fields)")

    btnProcess.OnEvent("Click", (*) => PartyMacro_OnProcess(configList, universalRules, ddl, srcEdit, webCtrl, statusText))
    btnCopy.OnEvent("Click", (*) => PartyMacro_OnCopy(configList, ddl, statusText))

    pmGui.OnEvent("Close", (*) => pmGui.Destroy())
    pmGui.Show()

    ; Auto-run from clipboard -- this is the primary trigger path (Quick
    ; Launcher / tray flyout): report is already copied, no manual paste
    ; or button click needed. The paste box + Process button remain as a
    ; manual override for editing text and re-running.
    clipText := A_Clipboard
    if (Trim(clipText) != "") {
        srcEdit.Value := clipText
        PartyMacro_OnProcess(configList, universalRules, ddl, srcEdit, webCtrl, statusText)
    } else {
        statusText.Text := "Clipboard is empty -- paste report text below and click Process."
    }
}

; Holds the last processed result + template RTF so the Copy button can use it
PartyMacro_LastResult := ""
PartyMacro_LastTemplateRtf := ""

PartyMacro_OnProcess(configList, universalRules, ddl, srcEdit, webCtrl, statusText) {
    sourceText := srcEdit.Value
    if (Trim(sourceText) = "") {
        MsgBox("Paste report text first.", "Party Macro", 48)
        return
    }

    detected := PartyMacro_DetectConfig(sourceText, configList)
    if (detected != "") {
        ddl.Text := detected.name
        statusText.Text := "Auto-detected: " detected.name
        PartyMacro_RunPipeline(detected, sourceText, universalRules, webCtrl, statusText)
        return
    }

    ; No exact match -- narrow to same modality if we can tell what it is,
    ; and let the user pick from that shortlist rather than guessing.
    modality := PartyMacro_DetectModality(sourceText)
    candidates := []
    for e in configList {
        if (modality != "" && e.config.Has("modality") && e.config["modality"] = modality)
            candidates.Push(e)
    }
    if (candidates.Length = 0)
        candidates := configList

    PartyMacro_ShowChooser(modality, candidates, sourceText, universalRules, webCtrl, statusText, ddl)
}

; Runs extraction/cleanup/diff for a chosen config entry and updates the GUI
PartyMacro_RunPipeline(entry, sourceText, universalRules, webCtrl, statusText) {
    global PartyMacro_LastResult, PartyMacro_LastTemplateRtf, PartyMacro_TemplateDir

    config := entry.config
    templatePath := PartyMacro_TemplateDir "\" config["templateFile"]
    if !FileExist(templatePath) {
        MsgBox("Template file not found:`n" templatePath, "Party Macro", 48)
        return
    }
    templateRtf := FileRead(templatePath, "UTF-8")

    result := PartyMacro_Process(sourceText, config, universalRules)

    PartyMacro_LastResult := result
    PartyMacro_LastTemplateRtf := templateRtf

    html := PartyMacro_BuildDiffDocument(result)
    PartyMacro_SetWebHtml(webCtrl, html)
}

; Popup shown when no exam-type match is found. Lists only configs sharing
; the detected modality (falls back to the full list if modality is unknown
; or nothing shares it).
PartyMacro_ShowChooser(modality, candidates, sourceText, universalRules, webCtrl, statusText, ddl) {
    names := []
    for e in candidates
        names.Push(e.name)

    cGui := Gui("+AlwaysOnTop", "Select Party Macro")
    cGui.SetFont("s10", "Segoe UI")

    label := (modality != "")
        ? "No exact match found. Showing " modality " party macros -- pick one to use for this case:"
        : "No exact match found and modality could not be determined. Pick a party macro to use:"
    cGui.Add("Text", "xm y10 w360", label)

    lb := cGui.Add("ListBox", "xm y+8 w360 r8", names)
    lb.Choose(1)

    btnUse := cGui.Add("Button", "xm y+10 w120", "Use This")
    btnCancel := cGui.Add("Button", "x+10 w120", "Cancel")

    btnUse.OnEvent("Click", (*) => PartyMacro_ChooserConfirm(cGui, candidates, lb, sourceText, universalRules, webCtrl, statusText, ddl))
    btnCancel.OnEvent("Click", (*) => cGui.Destroy())
    cGui.OnEvent("Close", (*) => cGui.Destroy())
    cGui.Show()
}

PartyMacro_ChooserConfirm(cGui, candidates, lb, sourceText, universalRules, webCtrl, statusText, ddl) {
    entry := candidates[lb.Value]
    ddl.Text := entry.name
    statusText.Text := "Manually selected: " entry.name " (no auto-match found)"
    cGui.Destroy()
    PartyMacro_RunPipeline(entry, sourceText, universalRules, webCtrl, statusText)
}

PartyMacro_OnCopy(configList, ddl, statusText) {
    global PartyMacro_LastResult, PartyMacro_LastTemplateRtf

    if (PartyMacro_LastResult = "" || PartyMacro_LastTemplateRtf = "") {
        MsgBox("Run Process first.", "Party Macro", 48)
        return
    }

    rtf := PartyMacro_LastTemplateRtf
    plainParts := []
    for fieldName, val in PartyMacro_LastResult.fieldText {
        rtf := PartyMacro_XmlSetDefault(rtf, fieldName, val)
        plainParts.Push(fieldName ": " val)
    }
    for omit in PartyMacro_LastResult.omitFields {
        rtf := PartyMacro_RemoveXmlField(rtf, omit.outputField)
        rtf := PartyMacro_RemoveRtfLine(rtf, omit.rtfLine)
    }
    plain := ""
    for idx, part in plainParts
        plain .= (idx = 1 ? "" : "`n") part

    PartyMacro_SetClipboardTextAndRTF(plain, rtf)
    statusText.Text := "Copied to clipboard (text + RTF fields). Paste into PowerScribe."
}
