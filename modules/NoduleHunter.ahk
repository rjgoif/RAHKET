; ============================================================================
; MODULE: NODULE HUNTER
; Written by Reece J. Goiffon, MD, PhD
; AutoHotkey v2
; Can be run standalone or included in RAHK_Main.ahk
; ============================================================================

#Requires AutoHotkey v2.0

; Only set these if running standalone
if (A_LineFile = A_ScriptFullPath) {
    #SingleInstance Force
    Persistent
    
    ; Create standalone tray menu
    A_TrayMenu.Delete()
    A_TrayMenu.Add("Open Nodule Hunter", (*) => Show_NoduleHunter())
    A_TrayMenu.Add()
    A_TrayMenu.Add("Exit", (*) => ExitApp())
    
    ; Auto-show the GUI when running standalone
    Show_NoduleHunter()
}

Show_NoduleHunter() {
    NH := Gui(, "Nodule Hunter")
    NH.OnEvent("Close", (*) => NH.Destroy())

    ; --- Handlers (closures) ---------------------------------------
    NH_ProcessText(*) {
        input := NH_InputBox.Value
        output := ProcessNodules(input)
        NH_OutputBox.Value := output
    }

    NH_ClearInput(*) {
        NH_InputBox.Value := ""
        NH_ProcessText()
    }

    NH_CopyOutput(*) {
        A_Clipboard := NH_OutputBox.Value
    }

    NH_ShowInstructions(*) {
        instText := "
        (
        Nodule Hunter will summarize your list of pulmonary nodules. Paste a list of nodules, one per line, with valid formats including:
        3 mm right lower lobe (3:265)
        5 x 2 mm left lower lobe (3:210)
        6 mm groundglass left lung (3:102)
        12 x 9 mm lingula ground glass (3:302)

        Any other text in a line (e.g. describing change, solid and cystic, etc) will be preserved, but that nodule won't be lumped into the summary statements.
        )"
        MsgBox(instText, "Nodule Hunter Instructions")
    }

    NH_CloseGui(*) {
        NH.Destroy()
    }

	; --- Main GUI --------------------------------------------------
	; Instructions button at top-left
	NH_InstrBtn := NH.Add("Button", "x10 y10 w110 h25", "Instructions")
	NH_InstrBtn.OnEvent("Click", NH_ShowInstructions)

	; Shift text boxes down slightly
	NH.Add("Text", "x10 y45", "Input Text:")
	NH_InputBox := NH.Add("Edit", "x10 y65 w300 h235")
	NH_InputBox.OnEvent("Change", NH_ProcessText)

	NH.Add("Text", "x320 y45", "Output Text:")
	NH_OutputBox := NH.Add("Edit", "x320 y65 w300 h235 ReadOnly")

	; --- Bottom buttons --------------------------------------------
	btnY := 315

	NH_CopyBtn := NH.Add("Button", "x10 y" btnY " w80 h25", "Copy")
	NH_CopyBtn.OnEvent("Click", NH_CopyOutput)

	NH_ClearBtn := NH.Add("Button", "x460 y" btnY " w80 h25", "Clear")
	NH_ClearBtn.OnEvent("Click", NH_ClearInput)

	NH_CloseBtn := NH.Add("Button", "x550 y" btnY " w70 h25", "Close")
	NH_CloseBtn.OnEvent("Click", NH_CloseGui)


    NH.Show("w630 h380")
}


ProcessNodules(inputText) {
    ; Define lobe/region order (added right/left lung)
    lobeOrder := [
        "right upper lobe", "right middle lobe", "right lower lobe",
        "left upper lobe", "lingula", "left lower lobe",
        "right lung", "left lung"
    ]
    
    ; Initialize data structures
    solidNodules := Map()
    groundglassNodules := Map()
    seriesNumbers := Map()
    unmatchedLines := []
    
    ; Parse input line by line
    Loop Parse, inputText, "`n", "`r" {
        line := Trim(A_LoopField)
        if (line = "")
            continue

        lineMatched := false

        ; Extract series/image: (N:M) or (N M:M)
        series := "", image := ""
        if RegExMatch(line, "\((\d+)\s*:\s*(\d+)\)", &mSI) {
            series := mSI[1], image := mSI[2]
            if !seriesNumbers.Has(series)
                seriesNumbers[series] := []
            seriesNumbers[series].Push(image)
        } else if RegExMatch(line, "\((\d+)\s+(\d+):(\d+)\)", &mAlt) {
            series := mAlt[1] mAlt[2], image := mAlt[3]
            if !seriesNumbers.Has(series)
                seriesNumbers[series] := []
            seriesNumbers[series].Push(image)
        }

        ; Detect ground glass anywhere
        isGroundglass := RegExMatch(line, "i)\bground\s*glass\b")

        ; Extract size anywhere: X mm or X x Y mm (average if XxY)
        avgSize := ""
        sizeText := ""   ; exact matched text, e.g. "4 mm" or "4 x 6 mm"
        if RegExMatch(line, "i)(\d+)(?:\s*x\s*(\d+))?\s*mm", &mSz) {
            sizeText := mSz[0]  ; exact substring to remove during strict check
            if (mSz[2] != "")
                avgSize := Round((Integer(mSz[1]) + Integer(mSz[2])) / 2.0)
            else
                avgSize := Integer(mSz[1])
        }

        ; Find lobe anywhere in the line
        lobeName := ""
        lowerLine := StrLower(line)
        for lobe in lobeOrder {
            if InStr(lowerLine, StrLower(lobe)) {
                lobeName := lobe
                break
            }
        }

        ; If we have all three core components, verify no extra tokens
        if (avgSize != "" && lobeName != "" && series != "" && image != "") {
            if IsPureNoduleLine(line, lobeName, sizeText, isGroundglass) {
                targetMap := isGroundglass ? groundglassNodules : solidNodules
                if !targetMap.Has(lobeName)
                    targetMap[lobeName] := []
                targetMap[lobeName].Push({size: avgSize, slice: series ":" image})
                lineMatched := true
            }
        }

        if (!lineMatched) {
            unmatchedLines.Push(line)  ; preserve as-is for bottom output
        }
    }

    ; Correct N M -> N0M pattern if needed
    CorrectSliceNumbers(&solidNodules, seriesNumbers)
    CorrectSliceNumbers(&groundglassNodules, seriesNumbers)
    
    ; Generate output: ALL solids first (by lobe order), then ALL ground glass (by lobe order)
    output := []

    ; 1) Solids across all lobes
    for lobe in lobeOrder {
        if solidNodules.Has(lobe)
            output.Push(FormatLobeOutput(lobe, solidNodules[lobe], false))
    }

    ; 2) Ground glass across all lobes
    for lobe in lobeOrder {
        if groundglassNodules.Has(lobe)
            output.Push(FormatLobeOutput(lobe, groundglassNodules[lobe], true))
    }

    ; 3) As-is lines at the bottom, original order
    for u in unmatchedLines
        output.Push(u)

    ; Join and add trailing blank line to output window
    result := (output.Length > 0) ? JoinArray(output, "`n") : ""
    if (result != "")
        result .= "`n"   ; ensures a blank line at the end
    return result
}

CorrectSliceNumbers(&noduleMap, seriesNumbers) {
    for lobe, nodules in noduleMap {
        for nodule in nodules {
            if RegExMatch(nodule.slice, "(\d+):(\d+)", &match) {
                series := match[1], image := match[2]
                if (StrLen(image) = 2 && seriesNumbers.Has(series)) {
                    correctedImage := SubStr(image, 1, 1) "0" SubStr(image, 2, 1)
                    for existingImage in seriesNumbers[series] {
                        if (existingImage = correctedImage) {
                            nodule.slice := series ":" correctedImage
                            break
                        }
                    }
                }
            }
        }
    }
}

FormatLobeOutput(lobeName, nodules, isGroundglass) {
    ; Sort by size descending
    SortBySize(&nodules)
    
    maxSize := nodules[1].size
    
    ; Count distinct sizes to decide stars and "up to"
    distinct := Map()
    for n in nodules
        distinct[String(n.size)] := true
    hasMultipleSizes := (distinct.Count > 1)
    
    ; Split nodules into max-size and others
    maxNodules := []
    otherNodules := []
    for nodule in nodules {
        if (nodule.size = maxSize)
            maxNodules.Push(nodule)
        else
            otherNodules.Push(nodule)
    }
    
    ; Sort slices ascending
    SortBySlice(&maxNodules)
    SortBySlice(&otherNodules)
    
    ; Build slice list (star only if multiple distinct sizes)
    slices := []
    for nodule in maxNodules {
        if RegExMatch(nodule.slice, ":(\d+)", &m)
            slices.Push(m[1] (hasMultipleSizes ? "*" : ""))
    }
    for nodule in otherNodules {
        if RegExMatch(nodule.slice, ":(\d+)", &m)
            slices.Push(m[1])
    }
    
    ; Get series number from first nodule
    seriesNum := ""
    if RegExMatch(nodules[1].slice, "(\d+):", &m2)
        seriesNum := m2[1]
    
    sizeText := hasMultipleSizes ? ("up to " maxSize) : maxSize
    gg := isGroundglass ? " groundglass" : ""
    return lobeName gg " " sizeText " mm (" seriesNum ":" JoinArray(slices, ", ") ")"

}

SortBySize(&arr) {
    n := arr.Length
    Loop n - 1 {
        i := A_Index
        Loop n - i {
            j := A_Index
            if (arr[j].size < arr[j + 1].size) {
                temp := arr[j]
                arr[j] := arr[j + 1]
                arr[j + 1] := temp
            }
        }
    }
}

SortBySlice(&arr) {
    n := arr.Length
    Loop n - 1 {
        i := A_Index
        Loop n - i {
            j := A_Index
            slice1 := 0, slice2 := 0
            if RegExMatch(arr[j].slice, ":(\d+)", &m1)
                slice1 := Integer(m1[1])
            if RegExMatch(arr[j + 1].slice, ":(\d+)", &m2)
                slice2 := Integer(m2[1])
            if (slice1 > slice2) {
                temp := arr[j]
                arr[j] := arr[j + 1]
                arr[j + 1] := temp
            }
        }
    }
}

JoinArray(arr, delimiter) {
    result := ""
    for index, value in arr {
        result .= value
        if (index < arr.Length)
            result .= delimiter
    }
    return result
}

; Returns true only if the line is *exactly* the formula:
; "<size mm> [ground glass] <lobe> (series:image)"  OR  "<lobe> [ground glass] <size mm> (series:image)"
; Extra tokens => false. Order of "ground glass" is flexible.
IsPureNoduleLine(line, lobeName, sizeText, hasGroundGlass) {
    s := StrLower(Trim(line))

    ; Remove exactly one (series:image)
    c1 := 0
    s := RegExReplace(s, "\(\s*\d+\s*:\s*\d+\s*\)", "", &c1, 1)

    ; Remove exactly one occurrence of the size text (treat literally, case-insensitive)
    if (sizeText != "") {
        ; escape literal text
        patSize := "i)\Q" StrLower(sizeText) "\E"
        c2 := 0
        s := RegExReplace(s, patSize, "", &c2, 1)
    }

    ; Remove exactly one occurrence of the lobe name (treat literally, case-insensitive)
    if (lobeName != "") {
        patLobe := "i)\Q" StrLower(lobeName) "\E"
        c3 := 0
        s := RegExReplace(s, patLobe, "", &c3, 1)
    }

    ; Remove any "ground glass" tokens (optional, order-agnostic)
    if (hasGroundGlass) {
        c4 := 0
        s := RegExReplace(s, "i)\bground\s*glass\b", "", &c4, -1)
    }

    ; Strip leftover punctuation / separators / whitespace
    s := RegExReplace(s, "[\s,;:\-]+", "")

    ; Pure only if nothing remains
    return (s = "")
}
