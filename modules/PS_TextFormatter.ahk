; ============================================================================
; MODULE: PowerScribe Text Formatter
; Written by Reece J. Goiffon, MD, PhD
; AutoHotkey v2
; Can be run standalone or included in RAHKET_Main.ahk
; ============================================================================

#Requires AutoHotkey v2.0

; Global variable to track if Report Editor is enabled
global LineJoinerEnabled := false

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
        . "Ctrl + Shift + D = Convert dates in selection to YYYY-MMM-DD",
        "PowerScribe Text Formatter", 64))
    A_TrayMenu.Add()
    A_TrayMenu.Add("Exit", (*) => ExitApp())
    
    ; Show welcome message
    MsgBox(
        "Text Formatter Enabled`n`n"
        . "Highlight text in PowerScribe:`n`n"
        . "Ctrl + J         = Join lines (merge paragraphs, add periods)`n"
        . "Ctrl + Shift + J = Split into single-sentence lines`n"
        . "Ctrl + Shift + D = Convert dates in selection to YYYY-MMM-DD",
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


; ============================================================================
; HELPER FUNCTIONS
; ============================================================================

; Enable Report Editor hotkeys
LineJoiner_Enable() {
    global LineJoinerEnabled
    LineJoinerEnabled := true
    MsgBox(
        "Report Editor Enabled`n`n"
        . "Highlight text in PowerScribe:`n`n"
        . "Ctrl + J         = Join lines (merge paragraphs, add periods)`n"
        . "Ctrl + Shift + J = Split into single-sentence lines`n"
        . "Ctrl + Shift + D = Convert dates in selection to YYYY-MMM-DD",
        "Report Editor", 64)
}

; Disable Report Editor hotkeys
LineJoiner_Disable() {
    global LineJoinerEnabled
    LineJoinerEnabled := false
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

WaitForClipboard(expectedText, timeoutMs := 1000) {
    start := A_TickCount
    while (A_TickCount - start < timeoutMs) {
        if (A_Clipboard = expectedText)
            return true
        Sleep(10)
    }
    return false
}

RestoreClipboard_PSF(savedClip, expectedText) {
    ; Only restore if nothing else has changed the clipboard
    if (A_Clipboard = expectedText) {
        A_Clipboard := savedClip
    }
}
