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
    if (m[1] != "") {
        fromUnit := StrLower(m[3])
        a := PartyMacro_ConvertVal(m[1] + 0, fromUnit, targetUnit)
        b := PartyMacro_ConvertVal(m[2] + 0, fromUnit, targetUnit)
        return PartyMacro_FormatMeasurement(a, targetUnit) " x " PartyMacro_FormatMeasurement(b, targetUnit) " " targetUnit
    }
    fromUnit := StrLower(m[5])
    a := PartyMacro_ConvertVal(m[4] + 0, fromUnit, targetUnit)
    return PartyMacro_FormatMeasurement(a, targetUnit) " " targetUnit
}


; ---------------------------------------------------------------------------
; Section extraction / reflow
; ---------------------------------------------------------------------------
PartyMacro_ExtractSection(sourceText, label) {
    needle := "im)^" PartyMacro_EscapeRegex(label) ":[ \t]*(.*)$"
    if !RegExMatch(sourceText, needle, &m)
        return ""
    return { whole: m[0], content: Trim(m[1], " `t`r`n") }
}

; sourceLabel in config may be a single string or an array of fallback
; labels tried in order (e.g. "Lungs" then "Lungs/airways") -- mirrors the
; fallback pattern in your own QM script.
PartyMacro_ExtractSectionMulti(sourceText, labelOrList) {
    if (Type(labelOrList) = "Array") {
        for lbl in labelOrList {
            r := PartyMacro_ExtractSection(sourceText, lbl)
            if (r != "" && Trim(r.content, " `t`r`n") != "")
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
; Two-level diff (sentence-first, then word-level within changed sentences)
; for the track-changes view. Doing word-level LCS across a whole paragraph
; at once is prone to nonsensical alignments when a word repeats (e.g. two
; "Unchanged"s) -- diffing sentences first, and only descending to word
; level for sentences that actually differ, avoids that.
; ---------------------------------------------------------------------------
PartyMacro_SplitWords(text) {
    words := []
    for w in StrSplit(text, " ") {
        if (w != "")
            words.Push(w)
    }
    return words
}

; Splits text into sentences, keeping each sentence's own trailing
; punctuation attached to it. Uses the same boundary regex as reflow.
PartyMacro_SplitSentences(text) {
    static splitRE := "(?<!\d\.)(?<=\d|[a-zA-Z])([\)\].!?]{1,2})\s+(?=[A-Z\d])"
    sentences := []
    pos := 1
    lastEnd := 1
    while (foundPos := RegExMatch(text, splitRE, &m, pos)) {
        boundaryEnd := foundPos + m.Len(0)
        sentences.Push(SubStr(text, lastEnd, foundPos + m.Len(1) - lastEnd))
        lastEnd := boundaryEnd
        pos := boundaryEnd
        if (m.Len(0) = 0)
            pos += 1
    }
    tail := SubStr(text, lastEnd)
    if (Trim(tail, " `t`r`n") != "")
        sentences.Push(tail)
    return sentences
}

; Generic LCS diff over two arrays of string tokens (words or sentences)
PartyMacro_TokenDiff(a, b) {
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
            ops.InsertAt(1, { type: "same", token: a[i] })
            i--, j--
        } else if (j > 0 && (i = 0 || dp[i + 1][j] >= dp[i][j + 1])) {
            ops.InsertAt(1, { type: "add", token: b[j] })
            j--
        } else {
            ops.InsertAt(1, { type: "del", token: a[i] })
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

; Renders word-level ops (same/add/del) to HTML spans
PartyMacro_WordOpsToHtml(ops) {
    html := ""
    for op in ops {
        word := PartyMacro_HtmlEscape(op.token)
        if (op.type = "del") {
            html .= '<span class="del">' word '</span> '
        } else {
            isSimilar := RegExMatch(op.token, "i)^similar")
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

; Top-level entry point: sentence-level diff first; consecutive del+add
; sentence pairs (i.e. a sentence that was modified, not removed/added
; outright) get refined into a word-level diff scoped to just that pair.
PartyMacro_TextDiffHtml(origText, finalText) {
    sentA := PartyMacro_SplitSentences(origText)
    sentB := PartyMacro_SplitSentences(finalText)
    ops := PartyMacro_TokenDiff(sentA, sentB)

    html := ""
    idx := 1
    while (idx <= ops.Length) {
        op := ops[idx]

        if (op.type = "same") {
            html .= PartyMacro_HtmlEscape(op.token) " "
            idx++
            continue
        }

        ; a "del" immediately followed by an "add" is treated as one
        ; modified sentence -- refine with a word-level diff scoped to
        ; just that pair, instead of showing the whole sentence swapped
        if (op.type = "del" && idx < ops.Length && ops[idx + 1].type = "add") {
            wordOps := PartyMacro_TokenDiff(PartyMacro_SplitWords(op.token), PartyMacro_SplitWords(ops[idx + 1].token))
            html .= PartyMacro_WordOpsToHtml(wordOps)
            idx += 2
            continue
        }

        if (op.type = "del") {
            html .= '<span class="del">' PartyMacro_HtmlEscape(op.token) '</span> '
            idx++
            continue
        }

        ; add with no paired del -- a genuinely new sentence
        html .= '<span class="add">' PartyMacro_HtmlEscape(op.token) '</span> '
        idx++
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

; Parses every field's name/start/length from the pristine template XML, in
; document order. The declared "length" always matches that field's actual
; plain-text span (confirmed against real templates -- for picklist fields
; like Coronary this includes trailing text like ":none/mild/moderate/severe",
; not just the visible colored word).
PartyMacro_ParseFieldPositions(rtf) {
    fields := []
    pos := 1
    needle := 'field type="\d+" start="(\d+)" length="(\d+)"[^>]*><name>(.*?)</name>'
    while (foundPos := RegExMatch(rtf, needle, &m, pos)) {
        fields.Push({ name: m[3], start: m[1] + 0, length: m[2] + 0 })
        pos := foundPos + m.Len(0)
    }
    return fields
}

; Updates just the start="" length="" attributes on one field's tag,
; identified by its <name>
PartyMacro_XmlSetFieldPosition(rtf, fieldName, newStart, newLength) {
    pattern := '(<field type="\d+") start="\d+" length="\d+"([^>]*><name>' PartyMacro_EscapeRegex(fieldName) '</name>)'
    replacement := '$1 start="' newStart '" length="' newLength '"$2'
    return RegExReplace(rtf, pattern, replacement)
}

; Minimal RTF control-character escaping for inserted plain text. Embedded
; line breaks become \line (a soft break within the same paragraph, not
; \par which would start a new one) so multi-line content -- nodule lists,
; merged continuation text -- actually renders as separate lines rather
; than being flattened to spaces.
PartyMacro_RtfEscape(s) {
    s := StrReplace(s, "\", "\\")
    s := StrReplace(s, "{", "\{")
    s := StrReplace(s, "}", "\}")
    s := StrReplace(s, "`r`n", "\line ")
    s := StrReplace(s, "`n", "\line ")
    return s
}

; This is what actually controls what PowerScribe displays on paste.
; Most fields use "\cf2 Name\cf1" (e.g. "\cf2 Lines\cf1"), but some
; (Impression, RECOMMENDATION) go straight to "\cf2 Name\par" with no
; \cf1 at all -- confirmed against the real template. Matching whichever
; terminator is actually present, instead of assuming \cf1 always exists,
; is what makes Impression carry-forward actually work.
PartyMacro_SetRtfFieldText(rtf, fieldName, newVal) {
    pattern := "\\cf2 " PartyMacro_EscapeRegex(fieldName) "(\\cf1|\\par)"
    replacement := "\cf2 " PartyMacro_RtfEscape(newVal) "$1"
    return RegExReplace(rtf, pattern, replacement)
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
; Isolates the substring between "FINDINGS:" and "IMPRESSION:" (case
; insensitive). TECHNIQUE/COMPARISON/indication text above it is expected
; to not match any field.
PartyMacro_ExtractFindingsBlock(sourceText) {
    startPos := RegExMatch(sourceText, "i)FINDINGS:", &m1)
    if !startPos
        return sourceText
    blockStart := startPos + m1.Len(0)
    endPos := RegExMatch(sourceText, "i)IMPRESSION:", &m2, blockStart)
    if !endPos
        return SubStr(sourceText, blockStart)
    return SubStr(sourceText, blockStart, endPos - blockStart)
}

; Isolates the IMPRESSION block: from right after "IMPRESSION:" to right
; before "RECOMMENDATION:"/"RECOMMENDATIONS:" (or end of text if neither
; is present).
PartyMacro_ExtractImpressionBlock(sourceText) {
    startPos := RegExMatch(sourceText, "i)IMPRESSION:", &m1)
    if !startPos
        return ""
    blockStart := startPos + m1.Len(0)
    endPos := RegExMatch(sourceText, "i)RECOMMENDATIONS?:", &m2, blockStart)
    block := endPos ? SubStr(sourceText, blockStart, endPos - blockStart) : SubStr(sourceText, blockStart)
    return Trim(block, " `t`r`n")
}

; Isolates the RECOMMENDATION block: from right after "RECOMMENDATION:" or
; "RECOMMENDATIONS:" to the end of the text (no further section marker to
; bound it against currently -- if reports commonly have something after
; it, that boundary would need adding).
PartyMacro_ExtractRecommendationBlock(sourceText) {
    startPos := RegExMatch(sourceText, "i)RECOMMENDATIONS?:", &m1)
    if !startPos
        return ""
    blockStart := startPos + m1.Len(0)
    return Trim(SubStr(sourceText, blockStart), " `t`r`n")
}

; Builds a regex alternation of every known extract-mode source label,
; used to find where one field's span ends and the next begins.
PartyMacro_BuildLabelAlternation(config) {
    parts := []
    for sec in config["sections"] {
        if (sec["mode"] != "extract")
            continue
        if (Type(sec["sourceLabel"]) = "Array") {
            for lbl in sec["sourceLabel"]
                parts.Push(PartyMacro_EscapeRegex(lbl))
        } else {
            parts.Push(PartyMacro_EscapeRegex(sec["sourceLabel"]))
        }
    }
    alt := ""
    for idx, p in parts
        alt .= (idx = 1 ? "" : "|") p
    return alt
}

; Extracts a field's full multi-line span: everything from right after
; "Label:" up to (not including) the next recognized label-starting-line,
; or the end of the findings block. This is the key change that makes
; continuation lines -- list items, stray unlabeled notes, etc. -- belong
; to the preceding field automatically, instead of needing separate
; "orphan" handling: they were never actually separate, they're just text
; that comes after a label and before the next one.
PartyMacro_ExtractFieldSpan(findingsBlock, label, labelAlternation) {
    myPattern := "im)^" PartyMacro_EscapeRegex(label) ":[ \t]*"
    foundPos := RegExMatch(findingsBlock, myPattern, &m)
    if !foundPos
        return ""
    contentStart := foundPos + m.Len(0)

    nextPattern := "im)^(?:" labelAlternation "):"
    nextPos := RegExMatch(findingsBlock, nextPattern, &m2, contentStart)
    spanEnd := nextPos ? nextPos : StrLen(findingsBlock) + 1

    return Trim(SubStr(findingsBlock, contentStart, spanEnd - contentStart), " `t`r`n")
}

PartyMacro_ExtractFieldSpanMulti(findingsBlock, labelOrList, labelAlternation) {
    if (Type(labelOrList) = "Array") {
        for lbl in labelOrList {
            r := PartyMacro_ExtractFieldSpan(findingsBlock, lbl, labelAlternation)
            if (r != "")
                return r
        }
        return ""
    }
    return PartyMacro_ExtractFieldSpan(findingsBlock, labelOrList, labelAlternation)
}

; Text before the FIRST recognized field has no preceding field to belong
; to -- that's the only case still surfaced as genuinely orphaned.
PartyMacro_ExtractLeadingOrphan(findingsBlock, labelAlternation) {
    pattern := "im)^(?:" labelAlternation "):"
    foundPos := RegExMatch(findingsBlock, pattern, &m)
    if !foundPos
        return Trim(findingsBlock, " `t`r`n")
    return Trim(SubStr(findingsBlock, 1, foundPos - 1), " `t`r`n")
}

; The RTF value keeps embedded line breaks now (PartyMacro_RtfEscape turns
; them into \line when inserted into the RTF body) rather than flattening
; to spaces. StrLen() already counts each embedded \n as exactly 1
; character, which is what the offset math needs to stay consistent with
; the same 1-character-per-break assumption used for Coronary's \par.
PartyMacro_NormalizeForRtf(text) {
    return Trim(RegExReplace(text, " {2,}", " "), " `t`r`n")
}

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

    findingsBlock := PartyMacro_ExtractFindingsBlock(sourceText)
    labelAlternation := PartyMacro_BuildLabelAlternation(config)

    fieldText := Map()      ; outputField -> RTF-safe value (for scan lookups + RTF building)
    fieldOrder := []        ; [{field, value}] RTF-safe values, in config order
    plainOrder := []        ; [{field, value}] plain values (line breaks preserved), in config order
    sectionResults := []    ; diff display, in config order
    rtfRemovalMap := Map()

    ; Single pass through sections in config order, handling every mode
    ; inline -- this is what keeps the diff view in the same order as the
    ; template (e.g. Coronary right after Mediastinum) instead of grouped
    ; by mode.
    for sec in config["sections"] {
        mode := sec["mode"]

        if (mode = "extract") {
            spanText := PartyMacro_ExtractFieldSpanMulti(findingsBlock, sec["sourceLabel"], labelAlternation)
            if (spanText = "") {
                sectionResults.Push({ label: sec["outputField"], found: false, diffHtml: "", flag: "" })
                continue
            }

            cleaned := PartyMacro_ApplyCleanupRules(spanText, allRules)
            cleaned := PartyMacro_ProcessMeasurements(cleaned, targetUnit)
            cleaned := RegExReplace(cleaned, " {2,}", " ")
            cleaned := RegExReplace(cleaned, "(?:[ \t]*\r?\n){3,}", "`n`n")
            cleaned := Trim(cleaned, " `t`r`n")
            reflowed := PartyMacro_ReflowSection(cleaned, threshold)

            plainVal := reflowed
            rtfVal := PartyMacro_NormalizeForRtf(reflowed)

            fieldText[sec["outputField"]] := rtfVal
            fieldOrder.Push({ field: sec["outputField"], value: rtfVal })
            plainOrder.Push({ field: sec["outputField"], value: plainVal })

            diffHtml := PartyMacro_TextDiffHtml(spanText, reflowed)
            sectionResults.Push({ label: sec["outputField"], found: true, diffHtml: diffHtml, flag: "" })

        } else if (mode = "impressionBlock") {
            impText := PartyMacro_ExtractImpressionBlock(sourceText)
            if (impText = "") {
                sectionResults.Push({ label: sec["outputField"], found: false, diffHtml: "", flag: "" })
                continue
            }

            ; strip a leading "1." / "1.\t" -- the template's own
            ; auto-numbering already provides that for the first bullet
            impText := RegExReplace(impText, "^\s*\d+\.\s*", "")

            cleaned := PartyMacro_ApplyCleanupRules(impText, allRules)
            cleaned := PartyMacro_ProcessMeasurements(cleaned, targetUnit)
            cleaned := Trim(RegExReplace(cleaned, " {2,}", " "), " `t`r`n")

            plainVal := cleaned
            rtfVal := PartyMacro_NormalizeForRtf(cleaned)

            fieldText[sec["outputField"]] := rtfVal
            fieldOrder.Push({ field: sec["outputField"], value: rtfVal })
            plainOrder.Push({ field: sec["outputField"], value: plainVal })

            diffHtml := PartyMacro_TextDiffHtml(impText, cleaned)
            sectionResults.Push({ label: sec["outputField"], found: true, diffHtml: diffHtml, flag: "" })

        } else if (mode = "recommendationBlock") {
            recText := PartyMacro_ExtractRecommendationBlock(sourceText)
            if (recText = "") {
                sectionResults.Push({ label: sec["outputField"], found: false, diffHtml: "", flag: "" })
                continue
            }

            cleaned := PartyMacro_ApplyCleanupRules(recText, allRules)
            cleaned := PartyMacro_ProcessMeasurements(cleaned, targetUnit)
            cleaned := Trim(RegExReplace(cleaned, " {2,}", " "), " `t`r`n")

            plainVal := cleaned
            rtfVal := PartyMacro_NormalizeForRtf(cleaned)

            fieldText[sec["outputField"]] := rtfVal
            fieldOrder.Push({ field: sec["outputField"], value: rtfVal })
            plainOrder.Push({ field: sec["outputField"], value: plainVal })

            diffHtml := PartyMacro_TextDiffHtml(recText, cleaned)
            sectionResults.Push({ label: sec["outputField"], found: true, diffHtml: diffHtml, flag: "" })

        } else if (mode = "keywordScan") {
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

        } else if (mode = "conditionalOmit") {
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

            if (hit != "" && sec.Has("rtfContentToRemove")) {
                rtfRemovalMap[sec["outputField"]] := sec["rtfContentToRemove"]
                sectionResults.Push({ label: sec["outputField"], found: true,
                    diffHtml: '<span class="flag">"' PartyMacro_HtmlEscape(hit) '" found in report -- field removed from party output.</span>' })
            } else if (hit != "") {
                sectionResults.Push({ label: sec["outputField"], found: true,
                    diffHtml: '<span class="flag">"' PartyMacro_HtmlEscape(hit) '" found in report -- review and remove/edit this field manually in PowerScribe (left untouched here).</span>' })
            } else {
                sectionResults.Push({ label: sec["outputField"], found: true,
                    diffHtml: '<span class="missing">(no matching language found -- field kept as-is, untouched)</span>' })
            }
        }
    }

    orphanText := PartyMacro_ExtractLeadingOrphan(findingsBlock, labelAlternation)

    return { fieldText: fieldText, fieldOrder: fieldOrder, plainOrder: plainOrder,
             sections: sectionResults, orphanText: orphanText, rtfRemovalMap: rtfRemovalMap }
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
        . '.orphan{background:#ffe0e0;border:2px solid #b00020;padding:8px;white-space:pre-wrap;}'
        . '</style></head><body>'

    if (result.orphanText != "") {
        html .= '<h4 style="color:#b00020;">Unmatched text before the first recognized field -- no preceding field to attach to, add manually</h4>'
        html .= '<div class="orphan">' PartyMacro_HtmlEscape(result.orphanText) '</div>'
    }

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

; Builds the final RTF+XML from the pristine template: edits the visible
; RTF body text for each field, AND recomputes every field's start/length
; XML offsets cascading through the document, so fields after the first
; edited one still land in the right place. This is required -- editing
; only the visible text left every later field's offset stale, since
; PowerScribe uses those offsets (not just the visible span) to carve up
; the pasted plain text on paste.
;
; rtfRemovalMap (fieldName -> literal RTF content) is for fields being
; fully removed (e.g. Coronary): the exact content span is stripped
; directly, plus its own trailing \par (to avoid leaving a blank line
; behind). The content span's length is verified against the field's
; declared XML length. The \par's contribution to PowerScribe's plain-text
; offset count is NOT independently verified -- it's inferred from
; comparing "par removed" (corrupted every later field) against "par
; preserved" (worked correctly) in actual testing, which points to
; exactly 1 character. Worth specifically re-checking that fields after a
; removed one still land correctly.
PartyMacro_BuildFinalRtf(templateRtf, valueMap, rtfRemovalMap) {
    fields := PartyMacro_ParseFieldPositions(templateRtf)

    rtf := templateRtf
    cumulativeDelta := 0

    for f in fields {
        newStart := f.start + cumulativeDelta

        if rtfRemovalMap.Has(f.name) {
            rtf := StrReplace(rtf, rtfRemovalMap[f.name], "")
            rtf := PartyMacro_RemoveXmlField(rtf, f.name)
            cumulativeDelta += -(f.length + 1)
        } else if valueMap.Has(f.name) {
            newVal := valueMap[f.name]
            newLen := StrLen(newVal)
            rtf := PartyMacro_SetRtfFieldText(rtf, f.name, newVal)
            rtf := PartyMacro_XmlSetDefault(rtf, f.name, newVal)
            rtf := PartyMacro_XmlSetFieldPosition(rtf, f.name, newStart, newLen)
            cumulativeDelta += newLen - f.length
        } else {
            ; untouched field (Coronary when kept, RECOMMENDATION) --
            ; still needs its start shifted if any earlier field's
            ; length changed
            rtf := PartyMacro_XmlSetFieldPosition(rtf, f.name, newStart, f.length)
        }
    }

    return rtf
}

PartyMacro_OnCopy(configList, ddl, statusText) {
    global PartyMacro_LastResult, PartyMacro_LastTemplateRtf

    if (PartyMacro_LastResult = "" || PartyMacro_LastTemplateRtf = "") {
        MsgBox("Run Process first.", "Party Macro", 48)
        return
    }

    valueMap := Map()
    for pair in PartyMacro_LastResult.fieldOrder
        valueMap[pair.field] := pair.value

    rtf := PartyMacro_BuildFinalRtf(PartyMacro_LastTemplateRtf, valueMap, PartyMacro_LastResult.rtfRemovalMap)

    plainParts := []
    for pair in PartyMacro_LastResult.plainOrder {
        if (pair.value != "")
            plainParts.Push(pair.field ": " pair.value)
    }
    if (PartyMacro_LastResult.orphanText != "")
        plainParts.Push("--- UNMATCHED TEXT (no preceding field to attach to -- add manually) ---`n" PartyMacro_LastResult.orphanText)
    plain := ""
    for idx, part in plainParts
        plain .= (idx = 1 ? "" : "`n") part

    PartyMacro_SetClipboardTextAndRTF(plain, rtf)
    statusText.Text := "Copied to clipboard (text + RTF fields). Paste into PowerScribe."
}
