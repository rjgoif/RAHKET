; ============================================================================
; MODULE: PowerScribe Text Formatter
; Written by Reece J. Goiffon, MD, PhD
; AutoHotkey v2
; Can be run standalone or included in RAHKET_Main.ahk
; ============================================================================

#Requires AutoHotkey v2.0

; Global variable to track if Report Editor is enabled
global LineJoinerEnabled := false
global LineJoiner_PromptDeclined := false  ; if true, don't ask again this session

; Only set these if running standalone
if (A_LineFile = A_ScriptFullPath) {
    #SingleInstance Force
    Persistent
    
    ; Enable Report Editor by default when running standalone
    LineJoinerEnabled := true
    
    ; Create standalone tray menu
    A_TrayMenu.Delete()
    A_TrayMenu.Add("Text Formatter (Enabled)"
        , (*) => MsgBox(
        "Text Formatter is active`n`n"
        . "Ctrl + J         = Join lines (merge paragraphs, add periods)`n"
        . "Ctrl + Shift + J = Split into single-sentence lines`n"
        . "Ctrl + Shift + D = Convert dates in selection to YYYY-MMM-DD`n"
        . "Ctrl + Shift + S = Sort measurements in selection (descending)`n"
        . "Ctrl + Shift + U = Convert units (cm to mm, or mm to cm)",
        "PowerScribe Text Formatter", 64))
    A_TrayMenu.Add()
    A_TrayMenu.Add("Exit", (*) => ExitApp())
    
    ; Show welcome message
    MsgBox(
        "Text Formatter Enabled`n`n"
        . "Highlight text in PowerScribe:`n`n"
        . "Ctrl + J         = Join lines (merge paragraphs, add periods)`n"
        . "Ctrl + Shift + J = Split into single-sentence lines`n"
        . "Ctrl + Shift + D = Convert dates in selection to YYYY-MMM-DD`n"
        . "Ctrl + Shift + S = Sort measurements in selection (descending)`n"
        . "Ctrl + Shift + U = Convert units (cm to mm, or mm to cm)",
        "Text Formatter", 64)
}


; ============================================================================
; MODULE CODE
; ============================================================================

; Ctrl+J: Merge all lines into single paragraph, removing blank lines
; Adds periods at end of lines if missing
; Only active when PowerScribe is the active window
#HotIf WinActive("ahk_exe Nuance.PowerScribeOne.exe") and LineJoinerEnabled
^j up::
{
    ; Store original clipboard
    ClipSaved := ClipboardAll()
    
    ; Clear clipboard and copy selected text
    A_Clipboard := ""
    Send("^c")
    
    ; Wait for clipboard to contain data
    if !ClipWait(1) {
        A_Clipboard := ClipSaved
        return
    }
    
    selectedText := A_Clipboard
    
    ; First normalize line endings to just `n
    selectedText := StrReplace(selectedText, "`r`n", "`n")
    selectedText := StrReplace(selectedText, "`r", "`n")
    
    ; Split on double newlines (blank lines)
    paragraphs := StrSplit(selectedText, "`n`n")
    modifiedText := ""
    
    for index, paragraph in paragraphs {
        paragraph := Trim(paragraph)
        
        ; Skip completely empty paragraphs
        if (paragraph = "") {
            continue
        }
        
        ; Process each paragraph: join its lines into one line
        lines := StrSplit(paragraph, "`n")
        processedParagraph := ""
        
        for lineIndex, line in lines {
            line := Trim(line)
            if (line = "") {
                continue
            }
            
            ; Add period at end if not already there
            if !RegExMatch(line, "[.!?]$") {
                line .= "."
            }
            
            ; Add to paragraph with space separator (except for first line)
            if (processedParagraph != "") {
                processedParagraph .= " "
            }
            processedParagraph .= line
        }
        
        ; Add paragraph to result with single newline (no blank line)
        if (processedParagraph != "") {
            if (modifiedText != "") {
                modifiedText .= "`r`n"
            }
            modifiedText .= processedParagraph
        }
    }
    
    ; Put modified text in clipboard and paste
    A_Clipboard := modifiedText
    Send("^v")

    ; Restore original clipboard after a short delay,
    ; only if clipboard still matches our text
    SetTimer(RestoreClipboard_PSF.Bind(ClipSaved, modifiedText), -1000)

}

; Ctrl+Shift+U: Convert units. cm -> mm if any cm present, otherwise mm -> cm
^+u up::
{
    ClipSaved := ClipboardAll()

    A_Clipboard := ""
    Send("{Shift up}")
    Send("^c")

    if !ClipWait(1) {
        A_Clipboard := ClipSaved
        return
    }

    selectedText := A_Clipboard
    modifiedText := ConvertUnitsInText(selectedText)

    if (modifiedText = selectedText) {
        A_Clipboard := ClipSaved
        return
    }

    A_Clipboard := modifiedText
    Send("^v")

    SetTimer(RestoreClipboard_PSF.Bind(ClipSaved, modifiedText), -1000)
}

; Ctrl+Shift+S: Sort dimensions in the selection in descending order
; e.g. "3.3 x 6.9 x 8.2 cm" -> "8.2 x 6.9 x 3.3 cm"
^+s up::
{
    ClipSaved := ClipboardAll()

    A_Clipboard := ""
    Send("{Shift up}")
    Send("^c")

    if !ClipWait(1) {
        A_Clipboard := ClipSaved
        return
    }

    selectedText := A_Clipboard
    modifiedText := SortMeasurementsInText(selectedText)

    A_Clipboard := modifiedText
    Send("^v")

    SetTimer(RestoreClipboard_PSF.Bind(ClipSaved, modifiedText), -1000)
}

; Ctrl+Shift+J: Break lines at periods (opposite of Ctrl+J)
; Adds newline at start, then breaks each sentence to its own line
; Only active when PowerScribe is the active window
^+j up::
{
    ; Store original clipboard
    ClipSaved := ClipboardAll()
    
    ; Clear clipboard and copy selected text
    A_Clipboard := ""
	Send("{Shift up}")
    Send("^c")
    
    ; Wait for clipboard to contain data
    if !ClipWait(1) {
        A_Clipboard := ClipSaved
        return
    }
    
    selectedText := A_Clipboard
    
    ; Split at periods (keeping the period with the sentence)
    ; Match period followed by space and capture both parts
    modifiedText := RegExReplace(selectedText, "\.\s+", ".`r`n")
    
    ; Add newline at the start
    modifiedText := "`r`n" . modifiedText
    
    ; Put modified text in clipboard and paste
    A_Clipboard := modifiedText
    Send("^v")

    ; Restore original clipboard after a short delay,
    ; only if clipboard still matches our text
    SetTimer(RestoreClipboard_PSF.Bind(ClipSaved, modifiedText), -1000)

}

; Ctrl+Shift+D: Convert common date formats to YYYY-MMM-DD in the selection
^+d up::
{
    ; Store original clipboard
    ClipSaved := ClipboardAll()
    
    ; Clear clipboard and copy selected text
    A_Clipboard := ""
	Send("{Shift up}")
    Send("^c")
    
    ; Wait for clipboard to contain data
    if !ClipWait(1) {
        A_Clipboard := ClipSaved
        return
    }
    
    selectedText := A_Clipboard
    
    ; Convert dates within the selected text
    modifiedText := FormatDatesInText(selectedText)
    
    ; Put modified text in clipboard and paste
    A_Clipboard := modifiedText
    Send("^v")

    ; Restore original clipboard after a short delay,
    ; only if clipboard still matches our text
    SetTimer(RestoreClipboard_PSF.Bind(ClipSaved, modifiedText), -1000)

}
#HotIf


; When PowerScribe is active but Report Editor is disabled, ask once per
; session whether to turn it on. Yes (default/Enter) enables and re-fires
; the same shortcut. No suppresses further prompts until re-enabled.
#HotIf WinActive("ahk_exe Nuance.PowerScribeOne.exe") and !LineJoinerEnabled
^j up::
{
    if LineJoiner_PromptAndEnable()
        Send("^j")
}

^+j up::
{
    if LineJoiner_PromptAndEnable()
        Send("^+j")
}

^+d up::
{
    if LineJoiner_PromptAndEnable()
        Send("^+d")
}

^+s up::
{
    if LineJoiner_PromptAndEnable()
        Send("^+s")
}

^+u up::
{
    if LineJoiner_PromptAndEnable()
        Send("^+u")
}

#HotIf


; ============================================================================
; HELPER FUNCTIONS
; ============================================================================

; Enable Report Editor hotkeys
LineJoiner_Enable() {
    global LineJoinerEnabled, LineJoiner_PromptDeclined
    LineJoinerEnabled := true
    LineJoiner_PromptDeclined := false
    MsgBox(
        "Report Editor Enabled`n`n"
        . "Highlight text in PowerScribe:`n`n"
        . "Ctrl + J         = Join lines (merge paragraphs, add periods)`n"
        . "Ctrl + Shift + J = Split into single-sentence lines`n"
        . "Ctrl + Shift + D = Convert dates in selection to YYYY-MMM-DD`n"
        . "Ctrl + Shift + S = Sort measurements in selection (descending)`n"
        . "Ctrl + Shift + U = Convert units (cm to mm, or mm to cm)",
        "Report Editor", 64)
}

; Disable Report Editor hotkeys
LineJoiner_Disable() {
    global LineJoinerEnabled
    LineJoinerEnabled := false
}

; Prompt (once per session) to enable Report Editor. Returns true if
; enabled (either already, or just now), false if declined/still disabled.
LineJoiner_PromptAndEnable() {
    global LineJoinerEnabled, LineJoiner_PromptDeclined

    if (LineJoiner_PromptDeclined) {
        return false
    }

    result := MsgBox("Report Editor is currently disabled.`n`nEnable it now?"
                    , "Enable Report Editor", "YesNo 32")

    if (result = "No") {
        LineJoiner_PromptDeclined := true
        return false
    }

    LineJoinerEnabled := true
    return true
}

; ---- Date formatting helpers ----

; Walk the text, replacing any recognized date string with YYYY-MMM-DD
FormatDatesInText(text) {
    ; Match:
    ;  - 12/31/25, 12-31-2025
    ;  - 2025-12-31, 2025/12/31
    ;  - December 31, 2025 / Dec 31 25
    ;  - 31 December 2025 / 31 Dec 25
    pattern := "\b(?:"
        . "\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}"                              ; MM/DD/YY or MM-DD-YYYY
        . "|"
        . "\d{4}[\/-]\d{1,2}[\/-]\d{1,2}"                                ; YYYY-MM-DD or YYYY/MM/DD
        . "|"
        . "(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May"
            . "|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?"
            . "|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)"
            . "\s+\d{1,2}(?:st|nd|rd|th)?(?:,)?\s+\d{2,4}"               ; Month DD, YYYY
        . "|"
        . "\d{1,2}(?:st|nd|rd|th)?\s+"
            . "(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May"
            . "|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?"
            . "|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)"
            . ",?\s+\d{2,4}"                                             ; DD Month YYYY
        . ")\b"

    result   := ""
    startPos := 1

    loop {
        foundPos := RegExMatch(text, pattern, &m, startPos)
        if (!foundPos)
            break

        ; Text before this match
        result .= SubStr(text, startPos, foundPos - startPos)

        original := m[0]
        replacement := FormatDateString(original)
        if (replacement = "")
            replacement := original

        result   .= replacement
        startPos := foundPos + m.Len(0)
    }

    ; Tail after last match
    result .= SubStr(text, startPos)
    return result
}

; Convert a single date string into yyyy-MMM-dd (or "" if not valid)
FormatDateString(str) {
    str := Trim(str)
    local m, month, day, year

    ; MM/DD/YY or MM-DD-YY(YY)
    if RegExMatch(str, "^(?<m>\d{1,2})[\/-](?<d>\d{1,2})[\/-](?<y>\d{2,4})$", &m) {
        month := Integer(m["m"])
        day   := Integer(m["d"])
        year  := NormalizeYear(m["y"])
    }
    ; YYYY-MM-DD or YYYY/MM/DD
    else if RegExMatch(str, "^(?<y>\d{4})[\/-](?<m>\d{1,2})[\/-](?<d>\d{1,2})$", &m) {
        year  := Integer(m["y"])
        month := Integer(m["m"])
        day   := Integer(m["d"])
    }
    ; Month DD, YYYY (or YY)
    else if RegExMatch(str
        , "^(?<monthName>Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May"
          . "|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?"
          . "|Nov(?:ember)?|Dec(?:ember)?)"
          . "\s+(?<d>\d{1,2})(?:st|nd|rd|th)?(?:,)?\s+(?<y>\d{2,4})$"
        , &m) {
        month := MonthNameToNumber(m["monthName"])
        day   := Integer(m["d"])
        year  := NormalizeYear(m["y"])
    }
    ; DD Month YYYY (or YY)
    else if RegExMatch(str
        , "^(?<d>\d{1,2})(?:st|nd|rd|th)?\s+"
          . "(?<monthName>Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May"
          . "|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?"
          . "|Nov(?:ember)?|Dec(?:ember)?)"
          . ",?\s+(?<y>\d{2,4})$"
        , &m) {
        day   := Integer(m["d"])
        month := MonthNameToNumber(m["monthName"])
        year  := NormalizeYear(m["y"])
    }
    else {
        return ""
    }

    ; Basic sanity
    if (month < 1 || month > 12 || day < 1 || day > 31)
        return ""

    dateStr := Format("{:04}{:02}{:02}000000", year, month, day)
    try {
        return FormatTime(dateStr, "yyyy-MMM-dd")
    } catch {
        ; Invalid date (e.g. Feb 30)
        return ""
    }
}

MonthNameToNumber(name) {
    name := StrLower(name)
    name := SubStr(name, 1, 3)
    switch name {
        case "jan": return 1
        case "feb": return 2
        case "mar": return 3
        case "apr": return 4
        case "may": return 5
        case "jun": return 6
        case "jul": return 7
        case "aug": return 8
        case "sep": return 9
        case "oct": return 10
        case "nov": return 11
        case "dec": return 12
        default: return 0
    }
}

NormalizeYear(y) {
    y := Integer(y)
    if (y < 100) {
        if (y < 50)
            y += 2000
        else
            y += 1900
    }
    return y
}

; ---- Measurement sorting helpers ----

; Find dimension groups (2+ numbers joined by "x") and sort the numbers descending
SortMeasurementsInText(text) {
    numPat  := "\d+(?:\.\d+)?"
    unitPat := "(?:\s*(?:mm|cm))?"
    pattern := "i)\b" . numPat . unitPat
        . "(?:\s*[x\x{D7}]\s*" . numPat . unitPat . ")+"

    result   := ""
    startPos := 1

    while (foundPos := RegExMatch(text, pattern, &m, startPos)) {
        result .= SubStr(text, startPos, foundPos - startPos)
        result .= SortMeasurementGroup(m[0])
        startPos := foundPos + m.Len(0)
    }

    result .= SubStr(text, startPos)
    return result
}

; Sort the numbers inside one matched group, leaving separators and units in place.
; If the group mixes units (e.g. "3 mm x 1.2 cm"), it is returned unchanged.
SortMeasurementGroup(grp) {
    nums  := []
    units := Map()
    pos   := 1

    while (p := RegExMatch(grp, "i)(\d+(?:\.\d+)?)(?:\s*(mm|cm))?", &n, pos)) {
        nums.Push(n[1])
        if (n[2] != "")
            units[StrLower(n[2])] := true
        pos := p + n.Len(0)
    }

    if (units.Count > 1)
        return grp

    ; Insertion sort, descending (keeps original number text, e.g. "3.0")
    sorted := nums.Clone()
    loop sorted.Length - 1 {
        i := A_Index + 1
        v := sorted[i]
        j := i - 1
        while (j >= 1 && Number(sorted[j]) < Number(v)) {
            sorted[j + 1] := sorted[j]
            j--
        }
        sorted[j + 1] := v
    }

    ; Rebuild: replace each number in order, keep everything else
    out := ""
    pos := 1
    idx := 0
    while (p := RegExMatch(grp, "\d+(?:\.\d+)?", &n, pos)) {
        idx++
        out .= SubStr(grp, pos, p - pos) . sorted[idx]
        pos := p + n.Len(0)
    }
    out .= SubStr(grp, pos)
    return out
}

WaitForClipboard(expectedText, timeoutMs := 1000) {
    start := A_TickCount
    while (A_TickCount - start < timeoutMs) {
        if (A_Clipboard = expectedText)
            return true
        Sleep(10)
    }
    return false
}

; ---- Unit conversion helpers ----

; If any cm is present, convert cm -> mm (mm values untouched).
; Otherwise, if mm is present, convert mm -> cm.
ConvertUnitsInText(text) {
    hasCm := RegExMatch(text, "i)\d\s*cm(?![A-Za-z])")
    hasMm := RegExMatch(text, "i)\d\s*mm(?![A-Za-z])")

    if (!hasCm && !hasMm)
        return text

    toMm := hasCm

    numPat  := "\d+(?:\.\d+)?"
    unitPat := "(?:\s*(?:mm|cm)(?![A-Za-z]))?"
    pattern := "i)\b" . numPat . unitPat
        . "(?:\s*[x\x{D7}]\s*" . numPat . unitPat . ")*"

    result   := ""
    startPos := 1

    while (foundPos := RegExMatch(text, pattern, &m, startPos)) {
        result .= SubStr(text, startPos, foundPos - startPos)
        result .= ConvertUnitGroup(m[0], toMm)
        startPos := foundPos + m.Len(0)
    }

    result .= SubStr(text, startPos)
    return result
}

; Convert one number group. Numbers without their own unit take the next
; unit found later in the group (e.g. "3.2 x 1.1 cm" -> both are cm).
ConvertUnitGroup(grp, toMm) {
    tokenPat := "i)(\d+(?:\.\d+)?)(\s*(mm|cm)(?![A-Za-z]))?"
    srcUnit  := toMm ? "cm" : "mm"
    dstUnit  := toMm ? "mm" : "cm"

    ; Pass 1: collect numbers and their explicit units
    nums  := []
    units := []
    pos   := 1
    while (p := RegExMatch(grp, tokenPat, &n, pos)) {
        nums.Push(n[1])
        units.Push(n[3] != "" ? StrLower(n[3]) : "")
        pos := p + n.Len(0)
    }

    ; Back-fill effective units from the right
    eff      := []
    nextUnit := ""
    loop nums.Length {
        i := nums.Length - A_Index + 1
        if (units[i] != "")
            nextUnit := units[i]
        eff.InsertAt(1, nextUnit)
    }

    ; Pass 2: rebuild
    out := ""
    pos := 1
    idx := 0
    while (p := RegExMatch(grp, tokenPat, &n, pos)) {
        idx++
        out .= SubStr(grp, pos, p - pos)

        if (eff[idx] = srcUnit) {
            out .= ConvertNumber(n[1], toMm)
            if (n[2] != "")
                out .= RegExReplace(n[2], "i)(mm|cm)", dstUnit)
        } else {
            out .= n[0]
        }

        pos := p + n.Len(0)
    }
    out .= SubStr(grp, pos)
    return out
}

; cm -> mm: whole numbers shown without decimals (3.2 -> 32)
; mm -> cm: always at least one decimal (10 -> 1.0)
ConvertNumber(numStr, toMm) {
    v := Number(numStr)
    s := Round(toMm ? v * 10 : v / 10, 3)
    s := String(s)

    if InStr(s, ".") {
        s := RegExReplace(s, "(\.\d*?)0+$", "$1")
        s := RegExReplace(s, "\.$", "")
    }

    if (!toMm && !InStr(s, "."))
        s .= ".0"

    return s
}

RestoreClipboard_PSF(savedClip, expectedText) {
    ; Only restore if nothing else has changed the clipboard
    if (A_Clipboard = expectedText) {
        A_Clipboard := savedClip
    }
}
