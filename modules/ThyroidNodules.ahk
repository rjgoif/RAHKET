; ============================================================================
; MODULE: THYROID NODULES
; Written by Reece J. Goiffon, MD, PhD
; AutoHotkey v2
; Can be run standalone or included in RAHK_Main.ahk
; made from thyroid_nodules_14.ahk
; ============================================================================

#Requires AutoHotkey v2.0

; Initialize global variables BEFORE the standalone check
TN_Selections := Map()
TN_GuiObj := ""
TN_SplitterPos := 720  ; Initial splitter position
TN_Dragging := false
TN_MouseOffset := 0
TN_MinTop := 400  ; will be updated dynamically after building the tabs

; ---- layout constants (tweak to taste) ----
TN_MinOutput      := 60   ; minimum height for the output Edit (80 default)
TN_ButtonH        := 30   ; actual button height (30 default)
TN_BottomMargin   := 40   ; breathing room at window bottom (20 default)
TN_GapAfterSplit  := 30   ; splitter -> top of output box (30 default)
TN_GapAfterOutput := 10   ; output box -> buttons (10 default)


; highlight colors
TN_HighlightColor      := "FFFFAA"  ; pale yellow for single-select
TN_MultiHighlightColor := "AAFFAA"  ; pale green for multi-select (Echogenic Foci)


; Derived: total vertical overhead beneath the splitter
TN_StackOverhead := TN_GapAfterSplit + TN_GapAfterOutput + TN_ButtonH + TN_BottomMargin

; Path to resized thumbnail images (relative to the script folder)
; Default when included from RAHKET_Main or another caller:
TN_ImgDir := A_ScriptDir "\modules\thyroid_nodules\assets\resized\"

; Only set these if running standalone
if (A_LineFile = A_ScriptFullPath) {
    ; Override image dir when launched standalone (or compiled EXE)
    TN_ImgDir := A_ScriptDir "\thyroid_nodules\assets\resized\"

    #SingleInstance Force
    Persistent
    
    ; Create standalone tray menu
    A_TrayMenu.Delete()
    A_TrayMenu.Add("Open Thyroid Nodules", (*) => Show_ThyroidNodules())
    A_TrayMenu.Add()
    A_TrayMenu.Add("Exit", (*) => ExitApp())
    
    ; Auto-show the GUI when running standalone
    Show_ThyroidNodules()
}

; Initialize selection structure for a nodule
InitNoduleSelections(noduleNum) {
    global TN_Selections
    
    TN_Selections["Nodule" noduleNum] := Map(
        "Location", "",
        "Shape", "",
        "CentralVascularity", "",
        "ComparedToPrior", "",
        "ChangeInFeatures", "",
        "PriorFNA", "",
        "ChangeInTIRADS", "",
        "Composition", "",
        "Echogenicity", "",
        "EchogenicFoci", [],
        "Margin", "",
        ; NEW: size-related fields
        "SizeDims", [ "", "", "" ],
        "SizeCalcText", ""
    )
}


TN_ComputePriorThreshold(curr) {
    ; curr = current size in mm (numeric)
    ; We want the largest whole-mm prior size p such that:
    ;   curr >= p + 2   AND   curr >= 1.2 * p
    ; If no such p >= 1 exists, return "NA".
    if (curr = "" || curr < 0.1)
        return "NA"
    
    maxPrior := Floor( Min(curr - 2, curr / 1.2) )
    if (maxPrior < 1)
        return "NA"
    return maxPrior
}


TN_SizeEditChange(noduleNum, dimIndex, ctrl, *) {
    global TN_Selections, TN_GuiObj
    
    val := Trim(ctrl.Value)
    num := ""
    if (val != "" && RegExMatch(val, "^\s*(\d+(\.\d+)?)\s*$", &m))
        num := m[1] + 0  ; numeric
    
    selections := TN_Selections["Nodule" noduleNum]
    arr := selections["SizeDims"]
    arr[dimIndex] := num  ; updates in-place
    
    ; Recompute thresholds for all 3 dimensions
    results := []
    Loop 3 {
        curr := arr[A_Index]
        if (curr = "" || curr < 0.1)
            result := "NA"
        else
            result := TN_ComputePriorThreshold(curr)
        results.Push(result)
    }
    
    calcStr := "Prior size bounds (mm): " results[1] " / " results[2] " / " results[3]
    selections["SizeCalcText"] := calcStr
    
    ; Update on-screen text below "Compared to Prior"
    try TN_GuiObj["TN" noduleNum "_Txt_SizeCalc"].Text := calcStr
    
    ; immediately refresh the report so Size: ... mm appears
    TN_UpdateDisplay()
}




global TN_Points := Map(
    "Composition", Map(
        "Simple cystic/almost cystic", 0,
        "Spongiform", 0,
        "Mixed cystic/solid (>50% cystic)", 1,
        "Mixed cystic/solid (<50% cystic)", 1,
        "Solid/almost completely solid", 2,
        "Composition obscured", 2
    ),
    "Echogenicity", Map(
        "Anechoic", 0,
        "Hyperechoic", 1,
        "Isoechoic", 1,
        "Hypoechoic", 2,
        "Very Hypoechoic", 3,
        "Echogenicity obscured", 1
    ),
    "Shape", Map(
        '"wider-than-tall"', 0,
        '"taller-than-wide"', 3
    ),
	"Margin", Map(
		"Smooth margin", 0,
		"Ill-defined margin", 0,
		"Lobulated margin", 2,
		"Irregular margin", 2,
		"Extra-thyroid extension", 3
	),


    "Foci", Map(
        "None", 0,
        "Comet-tail artifacts", 0,
        "Macrocalcifications", 1,
        "Peripheral (rim) calcifications", 2,
        "Punctate echogenic foci", 3
    )
)


TN_ComputeNodulePoints(selections) {
    global TN_Points

    total := 0
    breakdown := []

    ; Composition + Echogenicity obscured combined rule
    comp := selections["Composition"]
    echo := selections["Echogenicity"]

    if (comp = "Composition obscured" && echo = "Echogenicity obscured") {
        total += 3
        breakdown.Push("composition and echogenicity obscured (3 points)")
    } else {
        ; Composition
        if (comp != "" && TN_Points["Composition"].Has(comp)) {
            p := TN_Points["Composition"][comp]
            total += p
            breakdown.Push(StrLower(comp) " (" p " point" (p = 1 ? "" : "s") ")")
        }

        ; Echogenicity
        if (echo != "" && TN_Points["Echogenicity"].Has(echo)) {
            p := TN_Points["Echogenicity"][echo]
            total += p
            breakdown.Push(StrLower(echo) " (" p " point" (p = 1 ? "" : "s") ")")
        }
    }

    ; Shape (keep capitalization exactly as your stored string)
    shape := selections["Shape"]
    if (shape != "" && TN_Points["Shape"].Has(shape)) {
        p := TN_Points["Shape"][shape]
        total += p
        breakdown.Push(StrLower(shape) " (" p " point" (p = 1 ? "" : "s") ")")
    }

    ; Margin
    margin := selections["Margin"]
    if (margin != "" && TN_Points["Margin"].Has(margin)) {
        p := TN_Points["Margin"][margin]
        total += p
        breakdown.Push(StrLower(margin) " (" p " point" (p = 1 ? "" : "s") ")")
    }

    ; Echogenic foci (multiple allowed)
    noneOn := false
    if selections.Has("EchogenicFociNone")
        noneOn := selections["EchogenicFociNone"]

    if (noneOn) {
        breakdown.Push("no echogenic foci (0 points)")
    } else {
        ef := selections["EchogenicFoci"]
        if (IsObject(ef)) {
            Loop ef.Length {
                f := ef[A_Index]
                if (TN_Points["Foci"].Has(f)) {
                    p := TN_Points["Foci"][f]
                    total += p
                    breakdown.Push(StrLower(f) " (" p " point" (p = 1 ? "" : "s") ")")
                }
            }
        }
    }


    return { Total: total, Breakdown: breakdown }
}



TN_ShouldShowTIRADS(points, sizeDims, force) {
    ; If the user explicitly forces reporting, always show TI-RADS
    if (force)
        return true

    ; Otherwise, apply the size/points thresholds
    maxSize := 0

    if IsObject(sizeDims) {
        Loop sizeDims.Length {
            val := sizeDims[A_Index]
            if (val != "") {
                num := val + 0  ; numeric cast
                if (num > maxSize)
                    maxSize := num
            }
        }
    }

    ; Threshold rules:
    ; - â‰¥7 points AND at least one dimension â‰¥5 mm
    ; - 4â€“6 points AND at least one dimension â‰¥10 mm
    ; - 3 points AND at least one dimension â‰¥15 mm
    if (points >= 7 && maxSize >= 5)
        return true

    if (points >= 4 && points <= 6 && maxSize >= 10)
        return true

    if (points = 3 && maxSize >= 15)
        return true

    return false
}


Show_ThyroidNodules() {
    global TN_GuiObj, TN_Selections, TN_SplitterPos
    global TN_NoduleType                ; <--- ADD THIS

    TN_NoduleType := Map()
    Loop 5
        TN_NoduleType["Right #" A_Index] := "Right"
    Loop 5
        TN_NoduleType["Left #" A_Index] := "Left"
    Loop 5
        TN_NoduleType["Isthmus #" A_Index] := "Isthmus"



    ; Initialize selections for 15 nodules
    Loop 15
        InitNoduleSelections(A_Index)

    
    ; Create GUI
    TN := Gui("+Resize", "Thyroid Nodules - ACR TI-RADS")
    TN_GuiObj := TN  ; Store globally
    TN.OnEvent("Close", TN_Close)
    TN.OnEvent("Size", TN_OnResize)
    TN.SetFont("s9", "Segoe UI")

    ; ================================================================
    ; TOP-ROW GLAND MEASUREMENTS
    ; ================================================================
	TN.Add("Text", "x20 y10 w120 h25 +Right vTN_IsthmusLbl", "Isthmus thickness:")
	istEdit := TN.Add("Edit", "x145 y10 w60 h23 vTN_IsthmusThickness")
	istEdit.OnEvent("Change", (*) => TN_UpdateDisplay())
	TN.Add("Text", "x210 y10 w35 h25", "mm")

	TN.Add("Text", "x260 y10 w120 h25 +Right", "Right lobe length:")
	rightEdit := TN.Add("Edit", "x385 y10 w60 h23 vTN_RightLobeLength")
	rightEdit.OnEvent("Change", (*) => TN_UpdateDisplay())
	TN.Add("Text", "x450 y10 w35 h25", "mm")

	TN.Add("Text", "x500 y10 w120 h25 +Right", "Left lobe length:")
	leftEdit := TN.Add("Edit", "x625 y10 w60 h23 vTN_LeftLobeLength")
	leftEdit.OnEvent("Change", (*) => TN_UpdateDisplay())
	TN.Add("Text", "x690 y10 w35 h25", "mm")

    ; Figure out where to place the tabs so they sit just under the top row
    TN["TN_IsthmusLbl"].GetPos(, &rowTop, , &rowH)
    tabsY := rowTop + rowH + 12  ; a little extra padding



    ; ================================================================
    ; TAB CONTROL (currently 3 tabs as before)
    ; ================================================================
    global TN_Tab := TN.Add("Tab3"
        , "x20 y" tabsY " w1160 h780 vTN_TabControl"
        , ["Right #1", "Right #2", "Right #3", "Right #4", "Right #5"
           , "Left #1",  "Left #2",  "Left #3",  "Left #4",  "Left #5"
           , "Isthmus #1", "Isthmus #2", "Isthmus #3", "Isthmus #4", "Isthmus #5"])





    ; Build each tabâ€™s content
    Loop 15 {
        TN_Tab.UseTab(A_Index)
        BuildNoduleTab(TN, A_Index)
    }
    TN_Tab.UseTab()  ; End tab context

    ; --- Measure how tall the top content needs to be (tab 1 is representative) ---
    ; Bottom of the right margin row (Extra-thyroid image)
    TN["TN1_Img_ExtraThyroid"].GetPos(, &mY, , &mH)
    bottomRight := mY + mH

    ; Bottom of the left sidebar (Clear nodule button)
    leftBottom := 0
    try {
        TN["TN1_Btn_ClearNodule"].GetPos(, &cY, , &cH)
        leftBottom := cY + cH
    }

    ; Take the deeper of the two and add padding
    requiredTop := Max(bottomRight, leftBottom) + 40

    TN_MinTop := requiredTop
    TN_SplitterPos := TN_MinTop

    ; Ensure the tab control matches the new splitter position
    TN["TN_TabControl"].Move(,, 1200, TN_SplitterPos - 20)


    ; --- Splitter bar (draggable) ---
    TN.Add("Text", "x10 y" TN_SplitterPos " w1200 h6 BackgroundGray vTN_Splitter +0x100")

    ; --- Output section ---
    initialOutputHeight := 150
    TN.Add("Text", "x10 y" (TN_SplitterPos + 10) " w100 vTN_OutputLabel", "Report Output:")
    TN.Add("Edit", "x10 y" (TN_SplitterPos + 30) " w1200 h" initialOutputHeight " +VScroll ReadOnly vTN_OutputBox")

    ; --- Buttons at bottom ---
    buttonY := TN_SplitterPos + TN_GapAfterSplit + initialOutputHeight + TN_GapAfterOutput

    TN.Add("Button", "x10 y"  buttonY " w100 h30 vTN_CopyBtn", "Copy Text")
        .OnEvent("Click", TN_CopyReport)

    TN.Add("Button", "x1000 y" buttonY " w100 h30 vTN_ClearBtn", "Clear All")
        .OnEvent("Click", (*) => TN_ClearAll())

    TN.Add("Button", "x120 y" buttonY " w140 h30 vTN_RecsBtn", "Recommendations")
        .OnEvent("Click", TN_ShowRecommendations)

    TN.Add("Button", "x1110 y" buttonY " w100 h30 vTN_CloseBtn", "Close")
        .OnEvent("Click", TN_Close)

    ; --- Ensure the initial window is tall enough so OnResize doesn't push the splitter up ---
    requiredHeight := TN_MinTop + TN_StackOverhead + TN_MinOutput
    windowHeight   := buttonY + TN_BottomMargin  ; whatever you were using as padding

    if (windowHeight < requiredHeight)
        windowHeight := requiredHeight

    ; optional global floor so it doesn't open tiny
    if (windowHeight < 1000)
        windowHeight := 1000

    TN.Show("w1220 h" windowHeight)
	SetTimer(() => TN_UpdateFieldColors(), -10)




    ; --- Show window with a tall enough default height ---
    windowHeight := buttonY + 60          ; padding at the bottom
    if (windowHeight < 1050)              ; enforce a generous minimum
        windowHeight := 1050

    TN.Show("w1220 h" windowHeight)
	SetTimer(() => TN_UpdateFieldColors(), -10)

    ; --- Set up splitter drag using GUI events ---
    OnMessage(0x201, TN_WM_LBUTTONDOWN)  ; WM_LBUTTONDOWN
    OnMessage(0x200, TN_WM_MOUSEMOVE)    ; WM_MOUSEMOVE

    TN_UpdateDisplay()
}


TN_Close(*) {
    global TN_GuiObj, TN_Dragging

    ; Unhook the mouse message handlers so they don't fire after the GUI is gone
    OnMessage(0x201, TN_WM_LBUTTONDOWN, 0)  ; remove WM_LBUTTONDOWN handler
    OnMessage(0x200, TN_WM_MOUSEMOVE, 0)    ; remove WM_MOUSEMOVE handler

    TN_Dragging := false

    if IsObject(TN_GuiObj) {
        try TN_GuiObj.Destroy()
    }

    TN_GuiObj := ""  ; mark as no active thyroid GUI
}




TN_WM_MOUSEMOVE(wParam, lParam, msg, hwnd) {
    global TN_GuiObj
    if (!TN_GuiObj)
        return

    ; Which control are we over?
    ; 2 = retrieve control hwnd
    MouseGetPos(, , &winHwnd, &ctrlHwnd, 2)
    ; If itâ€™s not our GUI, ignore
    if (winHwnd != TN_GuiObj.Hwnd)
        return

    ; If weâ€™re hovering the splitter, set the NS-resize cursor
    if (ctrlHwnd = TN_GuiObj["TN_Splitter"].Hwnd) {
        ; 32645 = IDC_SIZENS (vertical resize)
        static hCur := DllCall("LoadCursor", "ptr", 0, "ptr", 32645, "ptr")
        DllCall("SetCursor", "ptr", hCur)
        ; Donâ€™t consume the message; just return
    }
}


TN_SplitterCursor(gui, ctrl, *) {
    ; 32645 = V-SIZE cursor
    if (ctrl && ctrl.Hwnd = gui["TN_Splitter"].Hwnd) {
        DllCall("SetCursor", "ptr", DllCall("LoadCursor", "ptr", 0, "ptr", 32645, "ptr"))
    }
}

; Handle left mouse button down anywhere in the GUI
TN_WM_LBUTTONDOWN(wParam, lParam, msg, hwnd) {
    global TN_GuiObj, TN_Dragging, TN_MouseOffset
    if (!TN_GuiObj)
        return

    ; Work in client coords of the main GUI
    CoordMode "Mouse", "Client"
    MouseGetPos(, &mY, &winHwnd, &ctrlHwnd, 2)

    ; Only react if we clicked inside this GUI
    if (winHwnd != TN_GuiObj.Hwnd)
        return

    ; And specifically on the splitter control
    if (ctrlHwnd != TN_GuiObj["TN_Splitter"].Hwnd)
        return

    ; Splitter position (client coords)
    TN_GuiObj["TN_Splitter"].GetPos(, &sY, , &sH)

    ; Remember where inside the splitter we grabbed
    TN_MouseOffset := mY - sY

    TN_Dragging := true
    ; Capture mouse so drag continues even if cursor leaves the splitter
    DllCall("SetCapture", "ptr", TN_GuiObj.Hwnd)
    SetTimer(TN_SplitterDrag, 10)
    return 0
}


TN_GetNoduleLabel(noduleNum) {
    if (noduleNum <= 5) {
        region := "Right"
        ord := noduleNum
    } else if (noduleNum <= 10) {
        region := "Left"
        ord := noduleNum - 5
    } else {
        region := "Isthmus"
        ord := noduleNum - 10
    }
    return region " nodule #" ord
}

TN_NoduleHasContent(selections) {
    dims := selections["SizeDims"]
    hasSize := false
    if IsObject(dims) {
        Loop dims.Length {
            if (dims[A_Index] != "") {
                hasSize := true
                break
            }
        }
    }

    hasFoci := false
    ef := selections["EchogenicFoci"]
    if IsObject(ef) && ef.Length > 0
        hasFoci := true

    return (
           hasSize
        || selections["Location"]          != ""
        || selections["Shape"]             != ""
        || selections["Composition"]       != ""
        || selections["Echogenicity"]      != ""
        || selections["Margin"]            != ""
        || selections["CentralVascularity"]!= ""
        || selections["ComparedToPrior"]   != ""
        || selections["ChangeInFeatures"]  != ""
        || selections["PriorFNA"]          != ""
        || selections["ChangeInTIRADS"]    != ""
        || hasFoci
    )
}

TN_GetRecommendation(risk, maxSizeMm) {
    ; risk: integer 1â€“5 from TN_MapTIRADSCategory
    ; maxSizeMm: largest dimension in mm (float or int)

    if (maxSizeMm <= 0)
        return ""

    if (risk = 3) {
        if (maxSizeMm >= 25)
            return "Tissue sampling"
        if (maxSizeMm >= 15)
            return "1, 3, and 5 year ultrasound"
        return ""
    }

    if (risk = 4) {
        if (maxSizeMm >= 15)
            return "Tissue sampling"
        if (maxSizeMm >= 10)
            return "1, 2, 3, and 5 year ultrasound"
        return ""
    }

    if (risk = 5) {
        if (maxSizeMm >= 10)
            return "Tissue sampling"
        if (maxSizeMm >= 5)
            return "Annual ultrasound for 5 years"
        return ""
    }

    ; TR1â€“2: no structured follow-up here
    return ""
}

TN_RecommendationRank(recText) {
    if (recText = "Tissue sampling")
        return 3
    if (recText = "Annual ultrasound for 5 years")
        return 2
    if (recText = "1, 3 and 5 year ultrasound")
        return 1
    return 0
}

TN_ShowRecommendations(*) {
    global TN_Selections

    ; Bucket: recText => [labels...]
    recBuckets := Map()

    Loop 15 {
        n := A_Index
        selections := TN_Selections["Nodule" n]

        ; Skip nodules with nothing meaningful
        if !TN_NoduleHasContent(selections)
            continue

        ; Compute TI-RADS points & category
        score := TN_ComputeNodulePoints(selections)
        points := score.Total
        risk   := TN_MapTIRADSCategory(points)

        ; Compute max size in mm
        dims := selections["SizeDims"]
        maxSize := 0
        if IsObject(dims) {
            Loop dims.Length {
                val := dims[A_Index]
                if (val != "") {
                    val := val + 0  ; force numeric
                    if (val > maxSize)
                        maxSize := val
                }
            }
        }

        recText := TN_GetRecommendation(risk, maxSize)
        if (recText = "")
            continue

        label := TN_GetNoduleLabel(n)

        if !recBuckets.Has(recText)
            recBuckets[recText] := []

        recBuckets[recText].Push(label)
    }

    if (recBuckets.Count = 0) {
        MsgBox "No nodules meet TI-RADS follow-up or tissue sampling thresholds."
        return
    }

	; Convert buckets to sortable items
	items := []
	for recText, labels in recBuckets {
		rank := TN_RecommendationRank(recText)
		items.Push({ Rec: recText, Labels: labels, Rank: rank })
	}

	; --- sort items in-place by Rank (descending) ---
	for i, _ in items {
		for j, _ in items {
			if (i < j && items[i].Rank < items[j].Rank) {
				tmp := items[i]
				items[i] := items[j]
				items[j] := tmp
			}
		}
	}


	; Build display text
	msg := ""
	Loop items.Length {
		item := items[A_Index]

		; Join labels with comma + space
		labStr := ""
		Loop item.Labels.Length {
			if (A_Index > 1)
				labStr .= ", "
			labStr .= item.Labels[A_Index]
		}

		msg .= labStr ": " item.Rec
		if (A_Index < items.Length)
			msg .= "`n"
	}


    MsgBox msg, "TI-RADS Recommendations"
}



TN_SplitterDrag() {
    global TN_Dragging, TN_SplitterPos, TN_GuiObj
    global TN_MinTop, TN_MinOutput, TN_StackOverhead, TN_GapAfterSplit, TN_GapAfterOutput, TN_ButtonH

    if (!TN_Dragging)
        return

    if (!GetKeyState("LButton", "P")) {
        TN_Dragging := false
        SetTimer(TN_SplitterDrag, 0)
        DllCall("ReleaseCapture")
        return
    }

    CoordMode "Mouse", "Client"
    MouseGetPos(, &mY)

    TN_GuiObj.GetPos(, , &guiW, &guiH)

    ; Max splitter position that still leaves room for (min output + buttons + bottom margin)
    maxPos := guiH - (TN_StackOverhead + TN_MinOutput)

    ; If the window is so short that maxPos would be above TN_MinTop,
    ; pin maxPos at TN_MinTop instead of letting the splitter go higher.
    if (maxPos < TN_MinTop)
        maxPos := TN_MinTop

    ; Proposed new position from mouse
    newPos := mY - TN_MouseOffset

    ; Constrain strictly between TN_MinTop and maxPos
    if (newPos < TN_MinTop)
        newPos := TN_MinTop
    else if (newPos > maxPos)
        newPos := maxPos


    if (newPos = TN_SplitterPos)
        return

    TN_SplitterPos := newPos

    ; Output height: fill whatever space remains, but not below TN_MinOutput
    outputBoxHeight := guiH - TN_SplitterPos - TN_StackOverhead
    if (outputBoxHeight < TN_MinOutput)
        outputBoxHeight := TN_MinOutput

    ; Move controls
    TN_GuiObj["TN_TabControl"].Move(,, , TN_SplitterPos - 20)
    TN_GuiObj["TN_Splitter"].Move(, TN_SplitterPos)
    TN_GuiObj["TN_OutputLabel"].Move(, TN_SplitterPos + 10)
    TN_GuiObj["TN_OutputBox"].Move(, TN_SplitterPos + TN_GapAfterSplit, , outputBoxHeight)

    buttonY := TN_SplitterPos + TN_GapAfterSplit + outputBoxHeight + TN_GapAfterOutput
    TN_GuiObj["TN_CopyBtn"].Move(, buttonY)
    TN_GuiObj["TN_ClearBtn"].Move(, buttonY)
    TN_GuiObj["TN_RecsBtn"].Move(, buttonY)
    TN_GuiObj["TN_CloseBtn"].Move(, buttonY)
}




; Handle window resize
TN_OnResize(GuiObj, MinMax, Width, Height) {
    global TN_SplitterPos, TN_MinTop, TN_MinOutput, TN_StackOverhead, TN_GapAfterSplit, TN_GapAfterOutput, TN_ButtonH

    if (MinMax = -1)
        return

    ; Max splitter allowed for current window height
    maxPos := Height - (TN_StackOverhead + TN_MinOutput)

    ; If the window is tiny, don't ever push the splitter above TN_MinTop.
    if (maxPos < TN_MinTop)
        maxPos := TN_MinTop

    ; Clamp splitter strictly between TN_MinTop and maxPos
    if (TN_SplitterPos < TN_MinTop)
        TN_SplitterPos := TN_MinTop
    else if (TN_SplitterPos > maxPos)
        TN_SplitterPos := maxPos


    ; Recompute output height
    outputBoxHeight := Height - TN_SplitterPos - TN_StackOverhead
    if (outputBoxHeight < TN_MinOutput)
        outputBoxHeight := TN_MinOutput

    ; Resize/move controls
    GuiObj["TN_TabControl"].Move(,, Width - 20, TN_SplitterPos - 20)
    GuiObj["TN_Splitter"].Move(,, Width - 20)
    GuiObj["TN_OutputBox"].Move(,, Width - 20, outputBoxHeight)

    buttonY := TN_SplitterPos + TN_GapAfterSplit + outputBoxHeight + TN_GapAfterOutput

    ; Keep X positions; only adjust Y
    GuiObj["TN_CopyBtn"].Move(, buttonY)
    GuiObj["TN_ClearBtn"].Move(, buttonY)
    GuiObj["TN_RecsBtn"].Move(, buttonY)
    GuiObj["TN_CloseBtn"].Move(Width - 110, buttonY)

}


BuildNoduleTab(TN, noduleNum) {
    global TN_Tab  ; we only need the tab control to read its label

    ; LEFT SIDEBAR (entire column: Location â†’ Clear this nodule)
    startY := 70
    sideX  := 30      ; was 20 â€” add a little gutter from left edge
    sideW  := 190     ; was 180 â€” slightly wider so text isn't cramped

    ; ------------------------------------------------------------
    ; LOCATION SECTION (Differs for Isthmus vs others)
    ; ------------------------------------------------------------
    ; Tabs 1â€“3: Right, 4â€“6: Left, 7â€“9: Isthmus
    isIsthmus := (noduleNum >= 11)

    TN.Add("Text"
        , "x" sideX " y" startY " w" sideW " h25 +Center +Border BackgroundFFB366"
        , "Location")

    if (isIsthmus) {
        ; Isthmus nodules: location = Right / Medial / Left
        TN.Add("Button"
            , "x" sideX " y" (startY+25) " w" sideW " h25 vTN" noduleNum "_Btn_I_Right"
            , "Right")
            .OnEvent("Click", TN_ButtonClick.Bind(noduleNum, "Location", "right"))

        TN.Add("Button"
            , "x" sideX " y" (startY+55) " w" sideW " h25 vTN" noduleNum "_Btn_I_Medial"
            , "Medial")
            .OnEvent("Click", TN_ButtonClick.Bind(noduleNum, "Location", "medial"))

        TN.Add("Button"
            , "x" sideX " y" (startY+85) " w" sideW " h25 vTN" noduleNum "_Btn_I_Left"
            , "Left")
            .OnEvent("Click", TN_ButtonClick.Bind(noduleNum, "Location", "left"))
    } else {
        ; Right or Left lobes: Upper / Mid / Lower pole
        TN.Add("Button"
            , "x" sideX " y" (startY+25) " w" sideW " h25 vTN" noduleNum "_Btn_Upper"
            , "Upper")
            .OnEvent("Click", TN_ButtonClick.Bind(noduleNum, "Location", "upper pole"))

        TN.Add("Button"
            , "x" sideX " y" (startY+55) " w" sideW " h25 vTN" noduleNum "_Btn_Mid"
            , "Mid")
            .OnEvent("Click", TN_ButtonClick.Bind(noduleNum, "Location", "midgland"))

        TN.Add("Button"
            , "x" sideX " y" (startY+85) " w" sideW " h25 vTN" noduleNum "_Btn_Lower"
            , "Lower")
            .OnEvent("Click", TN_ButtonClick.Bind(noduleNum, "Location", "lower pole"))
    }



    ; ------------------------------------------------------------
    ; CURRENT SIZE (mm)
    ; ------------------------------------------------------------
    startY += 120
    TN.Add("Text"
        , "x" sideX " y" startY " w" sideW " h25 +Center +Border BackgroundAqua"
        , "Current size (mm)")

    ; three size edits in a row
    size1X := sideX
    size2X := sideX + 60
    size3X := sideX + 120

    sizeCtrl1 := TN.Add("Edit"
        , "x" size1X " y" (startY+30) " w50 h23 vTN" noduleNum "_Edit_Size1")
    sizeCtrl1.OnEvent("Change", TN_SizeEditChange.Bind(noduleNum, 1))

    sizeCtrl2 := TN.Add("Edit"
        , "x" size2X " y" (startY+30) " w50 h23 vTN" noduleNum "_Edit_Size2")
    sizeCtrl2.OnEvent("Change", TN_SizeEditChange.Bind(noduleNum, 2))

    sizeCtrl3 := TN.Add("Edit"
        , "x" size3X " y" (startY+30) " w50 h23 vTN" noduleNum "_Edit_Size3")
    sizeCtrl3.OnEvent("Change", TN_SizeEditChange.Bind(noduleNum, 3))

    ; "mm" label at right edge of the size row
    TN.Add("Text"
		, "x" (sideX + sideW - 35) " y" (startY+34) " w60 h20"
		, "mm")


    ; ------------------------------------------------------------
    ; COMPARED TO PRIOR
    ; ------------------------------------------------------------
    startY += 60
    TN.Add("Text"
        , "x" sideX " y" startY " w" sideW " h25 +Center +Border BackgroundAqua"
        , "Compared to Prior*")

    TN.Add("Button"
        , "x" sideX " y" (startY+25) " w" sideW " h25 vTN" noduleNum "_Btn_Stable"
        , "Stable")
        .OnEvent("Click", TN_ButtonClick.Bind(noduleNum, "ComparedToPrior", "Stable"))

    TN.Add("Button"
        , "x" sideX " y" (startY+50) " w" sideW " h25 vTN" noduleNum "_Btn_NotSeen"
        , "Not Seen")
        .OnEvent("Click", TN_ButtonClick.Bind(noduleNum, "ComparedToPrior", "Not Seen"))

    TN.Add("Button"
        , "x" sideX " y" (startY+75) " w" sideW " h25 vTN" noduleNum "_Btn_Larger"
        , "Larger")
        .OnEvent("Click", TN_ButtonClick.Bind(noduleNum, "ComparedToPrior", "Larger"))

    TN.Add("Button"
        , "x" sideX " y" (startY+100) " w" sideW " h25 vTN" noduleNum "_Btn_Smaller"
        , "Smaller")
        .OnEvent("Click", TN_ButtonClick.Bind(noduleNum, "ComparedToPrior", "Smaller"))

    ; Growth thresholds text
    TN.Add("Text"
        , "x" sideX " y" (startY+130) " w" sideW " h40 vTN" noduleNum "_Txt_SizeCalc"
        , "Prior size bounds (mm): NA / NA / NA")

    ; ------------------------------------------------------------
    ; SHAPE
    ; ------------------------------------------------------------
    startY += 170
    TN.Add("Text"
        , "x" sideX " y" startY " w" sideW " h25 +Center +Border BackgroundLime"
        , "Shape")
    TN.Add("Text"
        , "x" sideX " y" (startY+25) " w" sideW " h20 +Center"
        , "(based on measurement)")

    TN.Add("Button"
        , "x" sideX " y" (startY+45) " w" sideW " h25 vTN" noduleNum "_Btn_WiderThanTall"
        , '"Wider-than-tall"')
        .OnEvent("Click", TN_ButtonClick.Bind(noduleNum, "Shape", '"wider-than-tall"'))

    TN.Add("Button"
        , "x" sideX " y" (startY+70) " w" sideW " h25 vTN" noduleNum "_Btn_TallerThanWide"
        , '"Taller-than-wide"')
        .OnEvent("Click", TN_ButtonClick.Bind(noduleNum, "Shape", '"taller-than-wide"'))

    ; ------------------------------------------------------------
    ; CENTRAL VASCULARITY
    ; ------------------------------------------------------------
    startY += 110
    TN.Add("Text"
        , "x" sideX " y" startY " w" sideW " h25 +Center +Border BackgroundLime"
        , "Central Vascularity")

    TN.Add("Button"
        , "x" sideX " y" (startY+25) " w" sideW " h25 vTN" noduleNum "_Btn_Absent"
        , "Absent")
        .OnEvent("Click", TN_ButtonClick.Bind(noduleNum, "CentralVascularity", "Absent"))

    TN.Add("Button"
        , "x" sideX " y" (startY+50) " w" sideW " h25 vTN" noduleNum "_Btn_Minimal"
        , "Minimal")
        .OnEvent("Click", TN_ButtonClick.Bind(noduleNum, "CentralVascularity", "Minimal"))

    TN.Add("Button"
        , "x" sideX " y" (startY+75) " w" sideW " h25 vTN" noduleNum "_Btn_Moderate"
        , "Moderate")
        .OnEvent("Click", TN_ButtonClick.Bind(noduleNum, "CentralVascularity", "Moderate"))

    TN.Add("Button"
        , "x" sideX " y" (startY+100) " w" sideW " h25 vTN" noduleNum "_Btn_Marked"
        , "Marked")
        .OnEvent("Click", TN_ButtonClick.Bind(noduleNum, "CentralVascularity", "Marked"))

    ; ------------------------------------------------------------
    ; CHANGE IN FEATURES?
    ; ------------------------------------------------------------
    startY += 140
    TN.Add("Text"
        , "x" sideX " y" startY " w" sideW " h25 +Center +Border BackgroundAqua"
        , "Change in Features?")
    TN.Add("Button"
        , "x" sideX " y" (startY+25) " w85 h25 vTN" noduleNum "_Btn_ChangeFeatYes"
        , "Yes")
        .OnEvent("Click", TN_ButtonClick.Bind(noduleNum, "ChangeInFeatures", "Yes"))
    TN.Add("Button"
        , "x" (sideX + 95) " y" (startY+25) " w85 h25 vTN" noduleNum "_Btn_ChangeFeatNo"
        , "No")
        .OnEvent("Click", TN_ButtonClick.Bind(noduleNum, "ChangeInFeatures", "No"))

    ; ------------------------------------------------------------
    ; PRIOR FNA/BIOPSY?
    ; ------------------------------------------------------------
    startY += 60
    TN.Add("Text"
        , "x" sideX " y" startY " w" sideW " h25 +Center +Border BackgroundAqua"
        , "Prior FNA/Biopsy?")
    TN.Add("Button"
        , "x" sideX " y" (startY+25) " w85 h25 vTN" noduleNum "_Btn_FNAYes"
        , "Yes")
        .OnEvent("Click", TN_ButtonClick.Bind(noduleNum, "PriorFNA", "Yes"))
    TN.Add("Button"
        , "x" (sideX + 95) " y" (startY+25) " w85 h25 vTN" noduleNum "_Btn_FNANo"
        , "No")
        .OnEvent("Click", TN_ButtonClick.Bind(noduleNum, "PriorFNA", "No"))

    ; ------------------------------------------------------------
    ; TIRADS CATEGORY CHANGE?
    ; ------------------------------------------------------------
    startY += 60
    TN.Add("Text"
        , "x" sideX " y" startY " w" sideW " h25 +Center +Border BackgroundAqua"
        , "TIRADS category change?")
    TN.Add("Button"
        , "x" sideX " y" (startY+25) " w85 h25 vTN" noduleNum "_Btn_TIRADSYes"
        , "Yes")
        .OnEvent("Click", TN_ButtonClick.Bind(noduleNum, "ChangeInTIRADS", "Yes"))
    TN.Add("Button"
        , "x" (sideX + 95) " y" (startY+25) " w85 h25 vTN" noduleNum "_Btn_TIRADSNo"
        , "No")
        .OnEvent("Click", TN_ButtonClick.Bind(noduleNum, "ChangeInTIRADS", "No"))

    ; ------------------------------------------------------------
    ; FOOTNOTE + TI-RADS POINTS + FORCE CHECKBOX + CLEAR BUTTON
    ; ------------------------------------------------------------
    footY := startY + 70
    TN.Add("Text"
        , "x" sideX " y" footY " w" sideW " h40"
        , "* Growth = 2 measurements ↑ by 20% AND ≥2 mm.")

    ptsY := footY + 45
    TN.Add("Text"
        , "x" sideX " y" ptsY " w" sideW " h20 vTN" noduleNum "_Txt_TIRADSPoints"
        , "TI-RADS points: 0")

    forceY := ptsY + 25
    TN.Add("Checkbox"
        , "x" sideX " y" forceY " w" sideW " h20 vTN" noduleNum "_Chk_ForceTIRADS"
        , "Force TI-RADS reporting")
        .OnEvent("Click", (*) => TN_UpdateDisplay())

    clearY := forceY + 30
    TN.Add("Button"
    , "x" sideX " y" clearY " w" sideW " h25 vTN" noduleNum "_Btn_ClearNodule"
    , "Clear this nodule")
    .OnEvent("Click", TN_ClearSingleNodule.Bind(noduleNum))






    ; MAIN CONTENT AREA ----------------------------------------------------
    ; Shared layout constants
    headerH := 30
    gapAfterHeader := 20
    rowGap := 70

    imgW := 150
    imgH := 120

    ; Horizontal padding inside card labels
    labelPad := 6
    textW   := imgW - 2*labelPad

    leftMargin := 240
    startX := leftMargin


;
    ; ============================================================
    ; COMPOSITION
    ; ============================================================
    compHeaderY := 70
	TN.Add("Text"
		, "x235 y" compHeaderY " w960 h" headerH " +Center +Border BackgroundLime vTN" noduleNum "_Hdr_Composition"
		, "Composition"
	).SetFont("s11 Bold")
    TN.SetFont("s9", "Segoe UI")

    compImgY := compHeaderY + headerH + gapAfterHeader

	TN.Add("Text"
		, "x" startX " y" (compImgY + imgH + 5) " w" imgW " +Center vTN" noduleNum "_Lbl_Cystic"
		, "Simple cystic/almost cystic")
	TN.Add("Text"
		, "x" (startX + (imgW + 10)) " y" (compImgY + imgH + 5) " w" imgW " +Center vTN" noduleNum "_Lbl_Spongiform"
		, "Spongiform")
	TN.Add("Text"
		, "x" (startX + 2*(imgW + 10)) " y" (compImgY + imgH + 5) " w" imgW " +Center vTN" noduleNum "_Lbl_MixedMore50"
		, "Mixed cystic/solid (>50% cystic)")
	TN.Add("Text"
		, "x" (startX + 3*(imgW + 10)) " y" (compImgY + imgH + 5) " w" imgW " +Center vTN" noduleNum "_Lbl_MixedLess50"
		, "Mixed cystic/solid (<50% cystic)")
	TN.Add("Text"
		, "x" (startX + 4*(imgW + 10)) " y" (compImgY + imgH + 5) " w" imgW " +Center vTN" noduleNum "_Lbl_Solid"
		, "Solid/almost completely solid")
	TN.Add("Text"
		, "x" (startX + 5*(imgW + 10)) " y" (compImgY + imgH + 5) " w" imgW " +Center vTN" noduleNum "_Lbl_CalcLimit"
		, "Limited assessment`n(calcs, etc.)")


	TN.Add("Picture", "x" startX " y" compImgY " w" imgW " h" imgH " +Border vTN" noduleNum "_Img_Cystic"
		, TN_ImgDir . "_Img_Cystic.jpg")
		.OnEvent("Click", TN_ImageClick.Bind(noduleNum, "Composition", "Simple cystic/almost cystic"))
	TN.Add("Picture", "x" (startX + (imgW + 10)) " y" compImgY " w" imgW " h" imgH " +Border vTN" noduleNum "_Img_Spongiform"
		, TN_ImgDir . "_Img_Spongiform.jpg")
		.OnEvent("Click", TN_ImageClick.Bind(noduleNum, "Composition", "Spongiform"))
	TN.Add("Picture", "x" (startX + 2*(imgW + 10)) " y" compImgY " w" imgW " h" imgH " +Border vTN" noduleNum "_Img_MixedMore50"
		, TN_ImgDir . "_Img_MixedMore50.jpg")
		.OnEvent("Click", TN_ImageClick.Bind(noduleNum, "Composition", "Mixed cystic/solid (>50% cystic)"))
	TN.Add("Picture", "x" (startX + 3*(imgW + 10)) " y" compImgY " w" imgW " h" imgH " +Border vTN" noduleNum "_Img_MixedLess50"
		, TN_ImgDir . "_Img_MixedLess50.jpg")
		.OnEvent("Click", TN_ImageClick.Bind(noduleNum, "Composition", "Mixed cystic/solid (<50% cystic)"))
	TN.Add("Picture", "x" (startX + 4*(imgW + 10)) " y" compImgY " w" imgW " h" imgH " +Border vTN" noduleNum "_Img_Solid"
		, TN_ImgDir . "_Img_Solid.jpg")
		.OnEvent("Click", TN_ImageClick.Bind(noduleNum, "Composition", "Solid/almost completely solid"))
	TN.Add("Picture", "x" (startX + 5*(imgW + 10)) " y" compImgY " w" imgW " h" imgH " +Border vTN" noduleNum "_Img_CalcLimit"
		, TN_ImgDir . "_Img_CalcLimit.jpg")
		.OnEvent("Click", TN_ImageClick.Bind(noduleNum, "Composition", "Composition obscured"))

    ; ============================================================
    ; ECHOGENICITY
    ; ============================================================
    echoHeaderY := compImgY + imgH + rowGap
	TN.Add("Text"
		, "x235 y" echoHeaderY " w960 h" headerH " +Center +Border BackgroundLime vTN" noduleNum "_Hdr_Echogenicity"
		, "Echogenicity"
	).SetFont("s11 Bold")
    TN.SetFont("s9", "Segoe UI")

    echoImgY := echoHeaderY + headerH + gapAfterHeader

    ; 6 columns, 150px wide
	TN.Add("Text"
		, "x" startX " y" (echoImgY + 120 + 5) " w150 +Center vTN" noduleNum "_Lbl_Anechoic"
		, "Anechoic")
	TN.Add("Text"
		, "x" (startX + 160) " y" (echoImgY + 120 + 5) " w150 +Center vTN" noduleNum "_Lbl_Hyperechoic"
		, "Hyperechoic")
	TN.Add("Text"
		, "x" (startX + 320) " y" (echoImgY + 120 + 5) " w150 +Center vTN" noduleNum "_Lbl_Isoechoic"
		, "Isoechoic")
	TN.Add("Text"
		, "x" (startX + 480) " y" (echoImgY + 120 + 5) " w150 +Center vTN" noduleNum "_Lbl_Hypoechoic"
		, "Hypoechoic")
	TN.Add("Text"
		, "x" (startX + 640) " y" (echoImgY + 120 + 5) " w150 +Center vTN" noduleNum "_Lbl_VeryHypoechoic"
		, "Very hypoechoic")
	TN.Add("Text"
		, "x" (startX + 800) " y" (echoImgY + 120 + 5) " w150 +Center vTN" noduleNum "_Lbl_EchoLimited"
		, "Echogenicity obscured")

	
	TN.Add("Picture", "x" startX " y" echoImgY " w150 h120 +Border vTN" noduleNum "_Img_Anechoic"
		, TN_ImgDir . "_Img_Anechoic.jpg")
		.OnEvent("Click", TN_ImageClick.Bind(noduleNum, "Echogenicity", "Anechoic"))
	TN.Add("Picture", "x" (startX + 160) " y" echoImgY " w150 h120 +Border vTN" noduleNum "_Img_Hyperechoic"
		, TN_ImgDir . "_Img_Hyperechoic.jpg")
		.OnEvent("Click", TN_ImageClick.Bind(noduleNum, "Echogenicity", "Hyperechoic"))
	TN.Add("Picture", "x" (startX + 320) " y" echoImgY " w150 h120 +Border vTN" noduleNum "_Img_Isoechoic"
		, TN_ImgDir . "_Img_Isoechoic.jpg")
		.OnEvent("Click", TN_ImageClick.Bind(noduleNum, "Echogenicity", "Isoechoic"))
	TN.Add("Picture", "x" (startX + 480) " y" echoImgY " w150 h120 +Border vTN" noduleNum "_Img_Hypoechoic"
		, TN_ImgDir . "_Img_Hypoechoic.jpg")
		.OnEvent("Click", TN_ImageClick.Bind(noduleNum, "Echogenicity", "Hypoechoic"))
	TN.Add("Picture", "x" (startX + 640) " y" echoImgY " w150 h120 +Border vTN" noduleNum "_Img_VeryHypoechoic"
		, TN_ImgDir . "_Img_VeryHypoechoic.jpg")
		.OnEvent("Click", TN_ImageClick.Bind(noduleNum, "Echogenicity", "Very Hypoechoic"))

    ; Limited assessment option for Echogenicity
	TN.Add("Picture", "x" (startX + 800) " y" echoImgY " w150 h120 +Border vTN" noduleNum "_Img_EchoLimited"
		, TN_ImgDir . "_Img_EchoLimited.jpg")
		.OnEvent("Click", TN_ImageClick.Bind(noduleNum, "Echogenicity", "Echogenicity obscured"))
    TN.Add("Text"
        , "x" (startX + 800) " y" (echoImgY + imgH + 5) " w150 +Center"
        , "Limited assessment`n(calcs, etc.)")

    ; ============================================================
    ; ECHOGENIC FOCI
    ; ============================================================
    fociHeaderY := echoImgY + imgH + rowGap
	TN.Add("Text"
		, "x235 y" fociHeaderY " w960 h" headerH " +Center +Border BackgroundLime vTN" noduleNum "_Hdr_EchogenicFoci"
		, "Echogenic Foci"
	).SetFont("s11 Bold")
    TN.SetFont("s9", "Segoe UI")

    fociImgY := fociHeaderY + headerH + gapAfterHeader



	; labels under the foci tiles (NOT including None â€“ that text is on the tile itself)
	TN.Add("Text"
		, "x" (startX + 190) " y" (fociImgY + 120 + 5) " w180 +Center vTN" noduleNum "_Lbl_CometTail"
		, "Comet-tail artifacts")
	TN.Add("Text"
		, "x" (startX + 380) " y" (fociImgY + 120 + 5) " w180 +Center vTN" noduleNum "_Lbl_Macrocalc"
		, "Macrocalcifications")
	TN.Add("Text"
		, "x" (startX + 570) " y" (fociImgY + 120 + 5) " w180 +Center vTN" noduleNum "_Lbl_Peripheral"
		, "Peripheral (rim) calcifications")
	TN.Add("Text"
		, "x" (startX + 760) " y" (fociImgY + 120 + 5) " w180 +Center vTN" noduleNum "_Lbl_Punctate"
		, "Punctate echogenic foci")
		
		
    ; NONE â€” first position (fully centered text)
	TN.Add("Picture"
		, "x" startX " y" fociImgY
		  " w180 h120 +Border BackgroundBlack vTN" noduleNum "_Img_None"
		, "")
		.OnEvent("Click", TN_ImageClick.Bind(noduleNum, "EchogenicFoci", "None"))

	; Big "NONE" text on top of the black tile
	noneLbl := TN.Add("Text"
		, "x" startX " y" (fociImgY + 40) " w180 h40 +Center vTN" noduleNum "_Lbl_None"
		, "NONE")
	noneLbl.SetFont("s16 Bold cBlack")


	TN.Add("Picture", "x" (startX + 190) " y" fociImgY " w180 h120 +Border vTN" noduleNum "_Img_CometTail"
		, TN_ImgDir . "_Img_CometTail.jpg")
		.OnEvent("Click", TN_MultiImageClick.Bind(noduleNum, "EchogenicFoci", "Comet-tail artifacts"))
	TN.Add("Picture", "x" (startX + 380) " y" fociImgY " w180 h120 +Border vTN" noduleNum "_Img_Macrocalc"
		, TN_ImgDir . "_Img_Macrocalc.jpg")
		.OnEvent("Click", TN_MultiImageClick.Bind(noduleNum, "EchogenicFoci", "Macrocalcifications"))
	TN.Add("Picture", "x" (startX + 570) " y" fociImgY " w180 h120 +Border vTN" noduleNum "_Img_Peripheral"
		, TN_ImgDir . "_Img_Peripheral.jpg")
		.OnEvent("Click", TN_MultiImageClick.Bind(noduleNum, "EchogenicFoci", "Peripheral (rim) calcifications"))
	TN.Add("Picture", "x" (startX + 760) " y" fociImgY " w180 h120 +Border vTN" noduleNum "_Img_Punctate"
		, TN_ImgDir . "_Img_Punctate.jpg")
		.OnEvent("Click", TN_MultiImageClick.Bind(noduleNum, "EchogenicFoci", "Punctate echogenic foci"))


    TN.SetFont("s9", "Segoe UI")

	; ============================================================
	; MARGIN
	; ============================================================
	marginHeaderY := fociImgY + 120 + rowGap   ; 120 because foci images are 120px tall
	TN.Add("Text"
		, "x235 y" marginHeaderY " w960 h" headerH " +Center +Border BackgroundLime vTN" noduleNum "_Hdr_Margin"
		, "Margin"
	).SetFont("s11 Bold")
	TN.SetFont("s9", "Segoe UI")

	marginImgY := marginHeaderY + headerH + gapAfterHeader

	; ----- text labels (5 columns) -----
	TN.Add("Text"
		, "x" startX " y" (marginImgY + 120 + 5) " w180 +Center vTN" noduleNum "_Lbl_Smooth"
		, "Smooth")

	TN.Add("Text"
		, "x" (startX + 190) " y" (marginImgY + 120 + 5) " w180 +Center vTN" noduleNum "_Lbl_IllDefined"
		, "Ill-defined")

	TN.Add("Text"
		, "x" (startX + 380) " y" (marginImgY + 120 + 5) " w180 +Center vTN" noduleNum "_Lbl_Irregular"
		, "Irregular")

	TN.Add("Text"
		, "x" (startX + 570) " y" (marginImgY + 120 + 5) " w180 +Center vTN" noduleNum "_Lbl_Lobulated"
		, "Lobulated")

	TN.Add("Text"
		, "x" (startX + 760) " y" (marginImgY + 120 + 5) " w180 +Center vTN" noduleNum "_Lbl_ExtraThyroid"
		, "Extra-thyroid extension")


	; ----- 5 margin images (3:2, 180Ã—120) -----
	TN.Add("Picture"
		, "x" startX " y" marginImgY " w180 h120 +Border vTN" noduleNum "_Img_Smooth"
		, TN_ImgDir . "_Img_Smooth.jpg")
		.OnEvent("Click", TN_ImageClick.Bind(noduleNum, "Margin", "Smooth margin"))

	TN.Add("Picture"
		, "x" (startX + 190) " y" marginImgY " w180 h120 +Border vTN" noduleNum "_Img_IllDefined"
		, TN_ImgDir . "_Img_IllDefined.jpg")
		.OnEvent("Click", TN_ImageClick.Bind(noduleNum, "Margin", "Ill-defined margin"))

	TN.Add("Picture"
		, "x" (startX + 380) " y" marginImgY " w180 h120 +Border vTN" noduleNum "_Img_Irregular"
		, TN_ImgDir . "_Img_Irregular.jpg")
		.OnEvent("Click", TN_ImageClick.Bind(noduleNum, "Margin", "Irregular margin"))

	TN.Add("Picture"
		, "x" (startX + 570) " y" marginImgY " w180 h120 +Border vTN" noduleNum "_Img_Lobulated"
		, TN_ImgDir . "_Img_Lobulated.jpg")
		.OnEvent("Click", TN_ImageClick.Bind(noduleNum, "Margin", "Lobulated margin"))

	TN.Add("Picture"
		, "x" (startX + 760) " y" marginImgY " w180 h120 +Border vTN" noduleNum "_Img_ExtraThyroid"
		, TN_ImgDir . "_Img_ExtraThyroid.jpg")
		.OnEvent("Click", TN_ImageClick.Bind(noduleNum, "Margin", "Extra-thyroid extension"))


}
   
 
 
 
 

; Handle button clicks (single selection)
TN_ButtonClick(noduleNum, category, value, *) {
    global TN_Selections
    selections := TN_Selections["Nodule" noduleNum]
    
    ; Toggle selection
    if (selections[category] = value) {
        selections[category] := ""
    } else {
        selections[category] := value
    }
    
    TN_UpdateDisplay()
    TN_UpdateHighlights()
	TN_UpdateFieldColors()
}


; Handle image clicks (single selection)
TN_ImageClick(noduleNum, category, value, ctrl, *) {
    global TN_Selections

    selections := TN_Selections["Nodule" noduleNum]

    if (category = "EchogenicFoci" && value = "None") {
        ; --- Toggle the None flag, and clear others when turning it ON ---
        noneOn := false
        if selections.Has("EchogenicFociNone")
            noneOn := selections["EchogenicFociNone"]

		if !noneOn {
			; turning None ON
			selections["EchogenicFociNone"] := true

			; clear all other foci (just reset to a new empty array)
			selections["EchogenicFoci"] := []
		} else {
			; turning None OFF
			selections["EchogenicFociNone"] := false
		}

    } else {
        ; -------- normal single-selection categories --------
        if (selections[category] = value) {
            selections[category] := ""
        } else {
            selections[category] := value
        }
    }

    TN_UpdateDisplay()
    TN_UpdateHighlights()
	TN_UpdateFieldColors()
}




; Handle image clicks for multi-selection (Echogenic Foci)
TN_MultiImageClick(noduleNum, category, value, ctrl, *) {
    global TN_Selections
    selections := TN_Selections["Nodule" noduleNum]

    ; --- lazy init of foci array ---
    arr := selections["EchogenicFoci"]
    if !IsObject(arr)
        arr := selections["EchogenicFoci"] := []

    ; Toggle membership in the array
    found := false
    foundIndex := 0
    for index, item in arr {
        if (item = value) {
            found := true
            foundIndex := index
            break
        }
    }

    if (found) {
        ; Remove it
        arr.RemoveAt(foundIndex)
    } else {
        ; Add it
        arr.Push(value)
        ; If any microcalc is selected, None cannot be true
        selections["EchogenicFociNone"] := false
    }

    TN_UpdateDisplay()
    TN_UpdateHighlights()
	TN_UpdateFieldColors()
}



; Clear all selections
TN_ClearAll() {
    global TN_Selections, TN_GuiObj

    ; reset all 15 nodules (data + size text + TI-RADS label + Force checkbox)
    Loop 15 {
        n := A_Index
        InitNoduleSelections(n)

        if (TN_GuiObj) {
            try TN_GuiObj["TN" n "_Edit_Size1"].Value := ""
            try TN_GuiObj["TN" n "_Edit_Size2"].Value := ""
            try TN_GuiObj["TN" n "_Edit_Size3"].Value := ""
            try TN_GuiObj["TN" n "_Txt_SizeCalc"].Text := "Prior size bounds (mm): NA / NA / NA"
            try TN_GuiObj["TN" n "_Txt_TIRADSPoints"].Text := "TI-RADS points: 0"
            try TN_GuiObj["TN" n "_Chk_ForceTIRADS"].Value := 0   ; <-- uncheck Force
        }
    }

    ; clear top-row gland measurements (Isthmus + both lobes)
    if (TN_GuiObj) {
        try TN_GuiObj["TN_IsthmusThickness"].Value := ""
        try TN_GuiObj["TN_RightLobeLength"].Value  := ""
        try TN_GuiObj["TN_LeftLobeLength"].Value   := ""
    }

    ; *** NEW: jump back to the first tab (Right #1) ***
    if (TN_GuiObj) {
        try TN_GuiObj["TN_TabControl"].Value := 1
    }

    TN_UpdateDisplay()
    TN_UpdateHighlights()
	TN_UpdateFieldColors()
}



TN_ClearSingleNodule(noduleNum, *) {
    global TN_Selections, TN_GuiObj

    ; Reset data model
    InitNoduleSelections(noduleNum)

    ; Clear size edits
    try {
        TN_GuiObj["TN" noduleNum "_Edit_Size1"].Value := ""
        TN_GuiObj["TN" noduleNum "_Edit_Size2"].Value := ""
        TN_GuiObj["TN" noduleNum "_Edit_Size3"].Value := ""
    }

    ; Reset prior size bounds helper text
    try TN_GuiObj["TN" noduleNum "_Txt_SizeCalc"].Text := "Prior size bounds (mm): NA / NA / NA"

    ; Reset TI-RADS points label
    try TN_GuiObj["TN" noduleNum "_Txt_TIRADSPoints"].Text := "TI-RADS points: 0"

    ; Turn off â€œForce TI-RADSâ€ checkbox
    try TN_GuiObj["TN" noduleNum "_Chk_ForceTIRADS"].Value := 0

    ; Rebuild report + refresh highlights
    TN_UpdateDisplay()
    TN_UpdateHighlights()
	TN_UpdateFieldColors()
}





TN_ComputeTIRADSPoints(selections) {
    total := 0

    ; ----- Composition -----
    comp := StrLower(selections["Composition"])
    if (comp != "") {
        ; Check mixed first so it isn't swallowed by a generic "cystic" match
        if InStr(comp, "mixed cystic/solid") || InStr(comp, "mixed cystic and solid") {
            total += 1  ; mixed cystic and solid
        } else if InStr(comp, "solid/almost") || InStr(comp, "almost completely solid") {
            total += 2  ; solid or almost completely solid
        } else if InStr(comp, "composition obscured") || InStr(comp, "limited assessment") {
            total += 2  ; limited assessment
        } else if InStr(comp, "spongiform") {
            ; 0 points
        } else if InStr(comp, "simple cystic") || InStr(comp, "almost cystic") {
            ; 0 points
        }
    }


    ; ----- Echogenicity -----
    echo := StrLower(selections["Echogenicity"])
    if (echo != "") {
        if InStr(echo, "anechoic") {
            ; 0 points
        } else if InStr(echo, "hyperechoic") || InStr(echo, "isoechoic") {
            total += 1
        } else if InStr(echo, "hypoechoic") && !InStr(echo, "very") {
            total += 2
        } else if InStr(echo, "very hypoechoic") {
            total += 3
        } else if InStr(echo, "obscured") || InStr(echo, "limited assessment") {
            total += 1  ; limited assessment
        }
    }

    ; ----- Shape -----
    shape := StrLower(selections["Shape"])
    if (shape != "") {
        if InStr(shape, "taller-than-wide") || InStr(shape, "taller") {
            total += 3
        } else if InStr(shape, "wider-than-tall") || InStr(shape, "wider") {
            ; 0 points
        }
    }

    ; ----- Margin -----
    margin := StrLower(selections["Margin"])
    if (margin != "") {
        if InStr(margin, "smooth") || InStr(margin, "ill-defined") {
            ; 0 points
        } else if InStr(margin, "lobulated") || InStr(margin, "irregular") {
            total += 2
		} else if InStr(margin, "extra-thyroid") {
			total += 3
		}

    }

    ; ----- Echogenic foci (sum all selected) -----
    ef := selections["EchogenicFoci"]
    if IsObject(ef) && ef.Length > 0 {
        for _, f in ef {
            fl := StrLower(f)
            if (fl = "none") {
                ; 0
            } else if InStr(fl, "comet-tail") {
                ; 0
            } else if InStr(fl, "macrocalcifications") {
                total += 1
            } else if InStr(fl, "peripheral (rim) calcifications") || InStr(fl, "rim calcifications") {
                total += 2
            } else if InStr(fl, "punctate echogenic foci") || InStr(fl, "punctate") {
                total += 3
            }
        }
    }

    return total
}


TN_MapTIRADSCategory(points) {
    if (points >= 7)
        return 5
    if (points >= 4)
        return 4
    if (points >= 3)
        return 3
    if (points >= 2)
        return 2
    if (points >= 0)
        return 1
    return ""
}


; ======================================================================
; Update the display text (CLEAN, VALIDATED VERSION)
; ======================================================================
TN_UpdateDisplay() {
    global TN_Selections, TN_GuiObj
    global TN_RightText, TN_LeftText, TN_IsthmusText

    if (!TN_GuiObj)
        return

    bullet := "* "    ; filled bullet + space
    output := ""

    ; Per-region text buffers
    outRight   := ""
    outLeft    := ""
    outIsthmus := ""


    ; --------------------------------------------------------
    ; First pass: assign compact ordinals per region
    ; --------------------------------------------------------
    ordMap := Map()           ; key: nodule index (1â€“9) â†’ ordinal (1â€“3)
    rightCount := 0
    leftCount  := 0
    istCount   := 0

    Loop 15 {
        n := A_Index
        selections := TN_Selections["Nodule" n]

        if !TN_NoduleHasContent(selections)
            continue

        ; Region from index: 1â€“3=Right, 4â€“6=Left, 7â€“9=Isthmus
        if (n <= 5) {
            rightCount++
            ordMap[n] := rightCount
        } else if (n <= 10) {
            leftCount++
            ordMap[n] := leftCount
        } else {
            istCount++
            ordMap[n] := istCount
        }
    }

    ; --------------------------------------------------------
    ; Second pass: build report in tab order, but using compact
    ; numbering per region
    ; --------------------------------------------------------
    Loop 15 {
        n := A_Index
        selections := TN_Selections["Nodule" n]

        if !TN_NoduleHasContent(selections)
            continue

        ; mark the start of this noduleâ€™s block in the output string
        startPos := StrLen(output)

        ; TI-RADS points & breakdown
        score := TN_ComputeNodulePoints(selections)
        points := score.Total
        breakdown := score.Breakdown

        ; Update the per-tab points label (if exists)
        try TN_GuiObj["TN" n "_Txt_TIRADSPoints"].Text := "TI-RADS points: " points

        ; TI-RADS category & show/hide rules
        risk  := TN_MapTIRADSCategory(points)
        force := 0
        try force := TN_GuiObj["TN" n "_Chk_ForceTIRADS"].Value

        showTIRADS := TN_ShouldShowTIRADS(points, selections["SizeDims"], force)

        ; --- Size string, sorted largest â†’ smallest ---
        dims := selections["SizeDims"]
        sizeStr := ""
        maxSize := 0

        if IsObject(dims) {
            sizeVals := []
            Loop dims.Length {
                val := dims[A_Index]
                if (val != "") {
                    num := val + 0
                    sizeVals.Push(num)
                    if (num > maxSize)
                        maxSize := num
                }
            }
            if (sizeVals.Length > 1) {
                Loop sizeVals.Length - 1 {
                    i := A_Index
                    Loop sizeVals.Length - i {
                        j := A_Index
                        if (sizeVals[j] < sizeVals[j+1]) {
                            tmp := sizeVals[j]
                            sizeVals[j] := sizeVals[j+1]
                            sizeVals[j+1] := tmp
                        }
                    }
                }
            }
            Loop sizeVals.Length {
                if (A_Index > 1)
                    sizeStr .= " x "
                sizeStr .= Format("{:g}", sizeVals[A_Index])
            }
        }

        ; --- Region + compact ordinal label ---
        if (n <= 5) {
            regionName := "Right"
        } else if (n <= 10) {
            regionName := "Left"
        } else {
            regionName := "Isthmus"
        }

        isIsthmus := (regionName = "Isthmus")
        ordNum := (ordMap.Has(n) ? ordMap[n] : 0)

        loc := selections["Location"]

        ; First line: e.g. "Right nodule #1: 10 x 7 x 6 mm, upper pole"
        line1 := regionName " nodule #" ordNum ":"
        if (sizeStr != "") {
            line1 .= " " sizeStr " mm"
            if (loc != "")
                line1 .= ", " (isIsthmus ? loc : StrLower(loc))
        } else if (loc != "") {
            line1 .= " " (isIsthmus ? loc : StrLower(loc))
        }
        output .= line1 "`n"


        ; --- Bullet 1: features line (with or without TI-RADS points) ---

        ; central vascularity text
        cv := selections["CentralVascularity"]
        cvStr := ""
        if (cv != "")
            cvStr := "central vascularity " StrLower(cv)

        if (showTIRADS) {
            ; Use TI-RADS breakdown + vascularity
            parts := []

            Loop breakdown.Length
                parts.Push(breakdown[A_Index])

            if (cvStr != "")
                parts.Push(cvStr)

            if (parts.Length > 0) {
                firstPart := parts[1]
                firstPart := StrUpper(SubStr(firstPart, 1, 1)) . SubStr(firstPart, 2)
                txt := bullet firstPart
                Loop parts.Length - 1
                    txt .= ", " parts[A_Index+1]
                txt .= "."
                output .= txt "`n"
            }
        } else {
            ; Original non-point feature line
            features := []

            comp := selections["Composition"]
            if (comp != "")
                features.Push(comp)

            echo := selections["Echogenicity"]
            if (echo != "")
                features.Push(StrLower(echo))

            shape := selections["Shape"]
            if (shape != "")
                features.Push(shape)

            margin := selections["Margin"]
            if (margin != "")
                features.Push(StrLower(margin))

            ; --- Echogenic foci text ---
            fociArr := []
            if selections.Has("EchogenicFoci") && IsObject(selections["EchogenicFoci"])
                fociArr := selections["EchogenicFoci"]

            noneOn := false
            if selections.Has("EchogenicFociNone")
                noneOn := selections["EchogenicFociNone"]

            fociPhrase := ""
            if (noneOn) {
                fociPhrase := "no echogenic foci"
            } else if (fociArr.Length > 0) {
                ; optional: lowercase the list
                lowerList := []
                for _, v in fociArr
                    lowerList.Push(StrLower(v))
                fociPhrase := "echogenic foci: " . TN_JoinList(lowerList)
            }

            if (fociPhrase != "")
                features.Push(fociPhrase)

            ; central vascularity text
            if (cvStr != "")
                features.Push(cvStr)

            if (features.Length > 0) {
                lineFeatures := bullet features[1]
                Loop features.Length - 1
                    lineFeatures .= ", " features[A_Index+1]
                lineFeatures .= "."
                output .= lineFeatures "`n"
            }
        }


        ; --- Bullet 2: size compared to prior; change in features ---
        ctPrior    := selections["ComparedToPrior"]
        changeFeat := selections["ChangeInFeatures"]
        cfParts := []

        if (ctPrior != "")
            cfParts.Push("Size compared to prior: " StrLower(ctPrior))
        if (changeFeat != "")
            cfParts.Push("Change in features: " StrLower(changeFeat))

        if (cfParts.Length > 0) {
            lineCF := bullet cfParts[1]
            Loop cfParts.Length - 1
                lineCF .= "; " cfParts[A_Index+1]
            lineCF .= "."
            output .= lineCF "`n"
        }

        ; --- Bullet 3: prior biopsy ---
        priorFNA := selections["PriorFNA"]
        if (priorFNA != "") {
            output .= bullet "Prior biopsy: " StrLower(priorFNA) ".`n"
        }

        ; --- Bullet 4: TI-RADS summary (only if weâ€™re showing it) ---
        if (showTIRADS) {
            output .= bullet "TI-RADS total points: " points
               . "; TI-RADS risk category: " risk ".`n"
        }

        ; --- Bullet 5: change in TI-RADS risk category (only if selected) ---
        tirChange := selections["ChangeInTIRADS"]
        if (tirChange != "") {
            changeText := StrLower(tirChange)
            output .= bullet "Change in TI-RADS risk category: " changeText ".`n"
        }

        ; blank line between nodules (keep existing formatting)
        output .= "`n"

        ; Extract just this noduleâ€™s block from the global output
        block := SubStr(output, startPos + 1)
        output := SubStr(output, 1, startPos)

        ; Route the block by region
        if (regionName = "Right")
            outRight .= block
        else if (regionName = "Left")
            outLeft  .= block
        else  ; Isthmus
            outIsthmus .= block
    }

    ; --------------------------------------------------------
    ; Build final findings text from the plain-text template
    ; --------------------------------------------------------
    ; Save per-region text for the RTF copy function
    TN_RightText   := outRight
    TN_LeftText    := outLeft
    TN_IsthmusText := outIsthmus

    ; After youâ€™ve built outRight, outLeft, outIsthmus:
    ; (these should contain ONLY the nodule sentences, NOT the lengths)

    global TN_IsthmusText, TN_RightText, TN_LeftText
    TN_IsthmusText := outIsthmus
    TN_RightText   := outRight
    TN_LeftText    := outLeft
	
    ; At this point, outIsthmus, outRight, and outLeft should contain
    ; just the nodule text for each region (however you've formatted it).
    ; Save them globally so TN_CopyReport can see them.
    global TN_IsthmusText, TN_RightText, TN_LeftText
    TN_IsthmusText := outIsthmus
    TN_RightText   := outRight
    TN_LeftText    := outLeft

	; ------------- Build plain-text findings for the output box -------------
	istThick := Trim(TN_GuiObj["TN_IsthmusThickness"].Value)
	rightLen := Trim(TN_GuiObj["TN_RightLobeLength"].Value)
	leftLen  := Trim(TN_GuiObj["TN_LeftLobeLength"].Value)
    if (outIsthmus = "" && istThick != "")
        outIsthmus := "No significant nodules."
    if (outRight = "" && rightLen != "")
        outRight := "No significant nodules."
    if (outLeft = "" && leftLen != "")
        outLeft := "No significant nodules."
	
	; detect when weâ€™re using the auto â€œno nodulesâ€ default text
    istDefault  := (outIsthmus = "No significant nodules.")
    rightDefault := (outRight   = "No significant nodules.")
    leftDefault  := (outLeft    = "No significant nodules.")



    if (outRight = "" && outLeft = "" && outIsthmus = "" && istThick = "" && rightLen = "" && leftLen = "") {
        finalText := "(No selections made)`n"
    } else {
        finalText := "FINDINGS:`n`n"

        ; Isthmus block
        finalText .= "Isthmus:`n"
        if (istThick != "")
            finalText .= "Isthmus thickness " istThick " mm.`n"
        if (outIsthmus != "") {
            if (istDefault)
                ; no extra blank line when using default text
                finalText .= outIsthmus "`n"
            else
                ; real nodule text: pad with blank lines
                finalText .= "`n" outIsthmus "`n"
        }
        finalText .= "`n"

        ; Right lobe block
        finalText .= "Right thyroid:`n"
        if (rightLen != "")
            finalText .= "Right lobe length " rightLen " mm (sagittal).`n"
        if (outRight != "") {
            if (rightDefault)
                finalText .= outRight "`n"
            else
                finalText .= "`n" outRight "`n"
        }
        finalText .= "`nRight nodes`n`n"

        ; Left lobe block
        finalText .= "Left thyroid:`n"
        if (leftLen != "")
            finalText .= "Left lobe length " leftLen " mm (sagittal).`n"
        if (outLeft != "") {
            if (leftDefault)
                finalText .= outLeft "`n"
            else
                finalText .= "`n" outLeft "`n"
        }
        finalText .= "`nLeft nodes`n"

        finalText .= "Parathyroid: none/parathyroid`n"
    }

	; Save *just* the nodule blocks for the RTF function
	global TN_IsthmusText, TN_RightText, TN_LeftText
	TN_IsthmusText := outIsthmus
	TN_RightText   := outRight
	TN_LeftText    := outLeft
	
	TN_GuiObj["TN_OutputBox"].Value := finalText
	TN_UpdateFieldColors()
}







; --------------------------------------------------------------------
; HIGHLIGHT HELPERS
; --------------------------------------------------------------------

TN_JoinList(arr) {
    out := ""
    for i, v in arr {
        out .= (i > 1 ? ", " : "") . v
    }
    return out
}



TN_ArrayContains(arr, value) {
    if !IsObject(arr)
        return false
    for item in arr {
        if (item = value)
            return true
    }
    return false
}

TN_SetButtonHighlight(ctrlName, isOn) {
    global TN_GuiObj
    if !TN_GuiObj
        return
    try ctrl := TN_GuiObj[ctrlName]
    catch
        return
    if (isOn)
        ctrl.SetFont("Bold cBlue")
    else
        ctrl.SetFont("Norm cDefault")
}

TN_SetImageHighlight(ctrlName, isOn, isMulti := false) {
    global TN_GuiObj
    if !TN_GuiObj
        return
    try ctrl := TN_GuiObj[ctrlName]
    catch
        return

    if (isOn) {
        ; multi-select (echogenic foci) can be a different color if you want
        if (isMulti)
            ctrl.SetFont("Bold cGreen")
        else
            ctrl.SetFont("Bold cBlue")
    } else {
        ctrl.SetFont("Norm cDefault")
    }
}


TN_UpdateHighlights() {
    global TN_Selections, TN_GuiObj
    if !TN_GuiObj
        return

    Loop 15 {
        n := A_Index
        key := "Nodule" n
        if !TN_Selections.Has(key)
            continue
        selections := TN_Selections[key]

        ; --------------- SIDEBAR BUTTONS ---------------

        ; Location â€“ lobes
        loc := selections["Location"]
        for val, suffix in Map(
            "Right lobe", "_Btn_RightLobe",
            "Left lobe",  "_Btn_LeftLobe",
            "Upper pole", "_Btn_Upper",
            "Mid",        "_Btn_Mid",
            "Lower pole", "_Btn_Lower"
        ) {
            TN_SetButtonHighlight("TN" n suffix, loc = val)
        }
        ; Location â€“ isthmus (Right / Medial / Left)
        for val, suffix in Map(
            "Right",  "_Btn_I_Right",
            "Medial", "_Btn_I_Medial",
            "Left",   "_Btn_I_Left"
        ) {
            TN_SetButtonHighlight("TN" n suffix, loc = val)
        }


        ; Shape
        shp := selections["Shape"]
        for val, suffix in Map(
            '"Wider-than-tall"',  "_Btn_WiderThanTall",
            '"Taller-than-wide"', "_Btn_TallerThanWide"
        ) {
            TN_SetButtonHighlight("TN" n suffix, shp = val)
        }

        ; Central Vascularity
        cv := selections["CentralVascularity"]
        for val, suffix in Map(
            "Absent",   "_Btn_Absent",
            "Minimal",  "_Btn_Minimal",
            "Moderate", "_Btn_Moderate",
            "Marked",   "_Btn_Marked"
        ) {
            TN_SetButtonHighlight("TN" n suffix, cv = val)
        }

        ; Compared to Prior
        ctp := selections["ComparedToPrior"]
        for val, suffix in Map(
            "Stable",   "_Btn_Stable",
            "Not Seen", "_Btn_NotSeen",
            "Larger",   "_Btn_Larger",
            "Smaller",  "_Btn_Smaller"
        ) {
            TN_SetButtonHighlight("TN" n suffix, ctp = val)
        }

        ; Change in Features?
        cif := selections["ChangeInFeatures"]
        TN_SetButtonHighlight("TN" n "_Btn_ChangeFeatYes", cif = "Yes")
        TN_SetButtonHighlight("TN" n "_Btn_ChangeFeatNo" , cif = "No")

        ; Prior FNA/Biopsy?
        fna := selections["PriorFNA"]
        TN_SetButtonHighlight("TN" n "_Btn_FNAYes", fna = "Yes")
        TN_SetButtonHighlight("TN" n "_Btn_FNANo" , fna = "No")

        ; TI-RADS category change?
        tir := selections["ChangeInTIRADS"]
        TN_SetButtonHighlight("TN" n "_Btn_TIRADSYes", tir = "Yes")
        TN_SetButtonHighlight("TN" n "_Btn_TIRADSNo" , tir = "No")

        ; --------------- RIGHT-PANEL IMAGES ---------------

        ; Composition
        comp := selections["Composition"]
		for val, suffix in Map(
			"Simple cystic/almost cystic",           "_Lbl_Cystic",
			"Spongiform",                            "_Lbl_Spongiform",
			"Mixed cystic/solid (>50% cystic)",      "_Lbl_MixedMore50",
			"Mixed cystic/solid (<50% cystic)",      "_Lbl_MixedLess50",
			"Solid/almost completely solid",         "_Lbl_Solid",
			"Composition obscured",                  "_Lbl_CalcLimit"
		) {
			TN_SetImageHighlight("TN" n suffix, comp = val)
		}


        ; Echogenicity
        echo := selections["Echogenicity"]
		for val, suffix in Map(
			"Anechoic",              "_Lbl_Anechoic",
			"Isoechoic",             "_Lbl_Isoechoic",
			"Hyperechoic",           "_Lbl_Hyperechoic",
			"Hypoechoic",            "_Lbl_Hypoechoic",
			"Very hypoechoic",       "_Lbl_VeryHypoechoic",
			"Echogenicity obscured", "_Lbl_EchoLimited"
		) {
			TN_SetImageHighlight("TN" n suffix, echo = val)
		}


		; Margin
		mrg := selections["Margin"]
		for val, suffix in Map(
			"Smooth margin",           "_Lbl_Smooth",
			"Ill-defined margin",      "_Lbl_IllDefined",
			"Irregular margin",        "_Lbl_Irregular",
			"Lobulated margin",        "_Lbl_Lobulated",
			"Extra-thyroid extension", "_Lbl_ExtraThyroid"
		) {
			TN_SetImageHighlight("TN" n suffix, mrg = val)
		}



        ; Echogenic Foci (multi-select + separate None flag)
        fociArr := []
        if selections.Has("EchogenicFoci") && IsObject(selections["EchogenicFoci"])
            fociArr := selections["EchogenicFoci"]

        noneOn := false
        if selections.Has("EchogenicFociNone")
            noneOn := selections["EchogenicFociNone"]

        ; Highlight "None" label when the flag is true
        TN_SetImageHighlight("TN" n "_Lbl_None", !!noneOn, true)

        ; Highlight the other foci if theyâ€™re in the array
        for val, suffix in Map(
            "Comet-tail artifacts",            "_Lbl_CometTail",
            "Macrocalcifications",             "_Lbl_Macrocalc",
            "Peripheral (rim) calcifications", "_Lbl_Peripheral",
            "Punctate echogenic foci",         "_Lbl_Punctate"
        ) {
            TN_SetImageHighlight("TN" n suffix, TN_ArrayContains(fociArr, val), true)
        }
    }
	TN_UpdateFieldColors()
}

TN_UpdateFieldColors() {
    global TN_Selections, TN_GuiObj
    if !TN_GuiObj
        return

    orangeBg := "BackgroundEB8D52"
    defaultBg := "BackgroundDefault"

    Loop 15 {
        n := A_Index
        selections := TN_Selections["Nodule" n]
        
        ; Location buttons
        hasLocation := (selections["Location"] != "")
        locColor := hasLocation ? defaultBg : orangeBg
        
        ; Check if isthmus or lobe nodule for different button sets
        isIsthmus := (n >= 11)
        
        if (isIsthmus) {
            try TN_GuiObj["TN" n "_Btn_I_Right"].Opt(locColor)
            try TN_GuiObj["TN" n "_Btn_I_Medial"].Opt(locColor)
            try TN_GuiObj["TN" n "_Btn_I_Left"].Opt(locColor)
        } else {
            try TN_GuiObj["TN" n "_Btn_Upper"].Opt(locColor)
            try TN_GuiObj["TN" n "_Btn_Mid"].Opt(locColor)
            try TN_GuiObj["TN" n "_Btn_Lower"].Opt(locColor)
        }
        
        ; Size edit fields
        dims := selections["SizeDims"]
        hasSize := false
        if IsObject(dims) {
            Loop dims.Length {
                if (dims[A_Index] != "") {
                    hasSize := true
                    break
                }
            }
        }
        sizeColor := hasSize ? defaultBg : orangeBg
        try TN_GuiObj["TN" n "_Edit_Size1"].Opt(sizeColor)
        try TN_GuiObj["TN" n "_Edit_Size2"].Opt(sizeColor)
        try TN_GuiObj["TN" n "_Edit_Size3"].Opt(sizeColor)
        
        ; Category headers - change background color like image labels
        compHasSelection := (selections["Composition"] != "")
        echoHasSelection := (selections["Echogenicity"] != "")
        marginHasSelection := (selections["Margin"] != "")
        
        ; Echogenic foci: check if None flag OR any selections
        noneOn := false
        if selections.Has("EchogenicFociNone")
            noneOn := selections["EchogenicFociNone"]
        hasFoci := noneOn || (IsObject(selections["EchogenicFoci"]) && selections["EchogenicFoci"].Length > 0)
        
        ; Update header backgrounds (these are Text controls with names like "TN1_Hdr_Composition")
        try TN_GuiObj["TN" n "_Hdr_Composition"].Opt(compHasSelection ? "BackgroundLime" : orangeBg)
        try TN_GuiObj["TN" n "_Hdr_Echogenicity"].Opt(echoHasSelection ? "BackgroundLime" : orangeBg)
        try TN_GuiObj["TN" n "_Hdr_EchogenicFoci"].Opt(hasFoci ? "BackgroundLime" : orangeBg)
        try TN_GuiObj["TN" n "_Hdr_Margin"].Opt(marginHasSelection ? "BackgroundLime" : orangeBg)
    }
    
    ; Top lobe length fields
    istThick := Trim(TN_GuiObj["TN_IsthmusThickness"].Value)
    rightLen := Trim(TN_GuiObj["TN_RightLobeLength"].Value)
    leftLen := Trim(TN_GuiObj["TN_LeftLobeLength"].Value)
    
    TN_GuiObj["TN_IsthmusThickness"].Opt((istThick != "") ? defaultBg : orangeBg)
    TN_GuiObj["TN_RightLobeLength"].Opt((rightLen != "") ? defaultBg : orangeBg)
    TN_GuiObj["TN_LeftLobeLength"].Opt((leftLen != "") ? defaultBg : orangeBg)
}

; --------------------------------------------------------------------
; COPY FUNCTIONS
; --------------------------------------------------------------------


TN_XmlEscape(str) {
    ; Minimal XML escaping: & < >
    str := StrReplace(str, "&", "&amp;")
    str := StrReplace(str, "<", "&lt;")
    str := StrReplace(str, ">", "&gt;")
    return str
}

TN_XmlSetDefault(rtf, fieldName, newVal) {
    ; Replace the <defaultvalue>...</defaultvalue> for a given <name>fieldName</name>
    newVal := TN_XmlEscape(newVal)
    pattern := "(<name>" fieldName "</name><defaultvalue>)(.*?)(</defaultvalue>)"
    return RegExReplace(rtf, pattern, "$1" newVal "$3")
}


; ====================================================================
; FULL Nuance RTF template (with XML + fields) embedded in script
; ====================================================================
TN_RtfTemplate := '
(
{\rtf1\ansi\ansicpg1252\deff0\nouicompat\deflang1033{\fonttbl{\f0\fnil\fcharset0 Franklin Gothic Medium;}{\f1\fnil\fcharset1 Cambria Math;}}
{\colortbl ;\red234\green255\blue255;\red178\green34\blue34;\red0\green255\blue255;}
{\*\generator Riched20 10.0.19041}{\*\mmathPr\mmathFont1\mwrapIndent1440 }\viewkind4\uc1 
\pard\cf1\b\f0\fs20 FINDINGS:\b0\par
\par
Isthmus:\par
\cf2 isthmus\cf1  thick.\cf2\par
isthmus nodules\par
\cf1\par
\par
Right thyroid:\par
\cf2 right\cf1  lobe length in sagittal plane.\cf2\par
right nodules\par
Right nodes\par
\cf1\par
\par
Left thyroid:\par
\cf2 left\cf1  lobe length in sagittal plane.\cf2\par
left nodules\par
Left nodes\par
\cf1\par
\par
\cf2 parathyroid\cf1 :none/parathyroid\cf2\par
\cf1\par
\b IMPRESSION:\b0\par
\cf2 IMPRESSION\cf1 :0-1 benign, no FNA/2 not susp, no FNA/3&\f1\u8805?\f0 1.5cm mildly susp, +1/3/5y/3&\f1\u8805?\f0 2.5cm mildly susp, FNA/4-6&\f1\u8805?\f0 1cm mod susp, +1/2/3/5y/4-6&\f1\u8805?\f0 1.5cm mod susp, FNA/\f1\u8805?\f0 7&\f1\u8805?\f0 0.5cm highly susp, +1/2/3/5y/\f1\u8805?\f0 7&\f1\u8805?\f0 1cm highly susp, +FNA\cf2\lang1033\par
\cf1\par
\cf2 RECOMMENDATIONS:\par
Verbal Prelim\par
\cf1\par
Glossary of terms and other information on TI-RADS (Thyroid Imaging Reporting and Data System) can be found at \cf3 https://www.acr.org/Clinical-Resources/Clinical-Tools-and-Reference/Reporting-and-Data-Systems/TI-RADS\cf1\par
\fs20\par
}
 {\xml}<?xml version="1.0" encoding="utf8"?><autotext version="2" editMode="2"><fields><field type="1" start="20" length="7"><name>isthmus</name><defaultvalue>4 mm</defaultvalue><customproperties><property><name>AllCaps</name><value>False</value></property><property><name>AllowEmpty</name><value>False</value></property><property><name>ImpressionField</name><value>False</value></property><property><name>DoesNotIndicateFindings</name><value>True</value></property><property><name>FindingsCodes</name><value></value></property><property><name>EnforcePickList</name><value>False</value></property></customproperties></field><field type="1" start="35" length="15"><name>isthmus nodules</name><defaultvalue>isthmus nodules</defaultvalue><customproperties><property><name>AllCaps</name><value>False</value></property><property><name>AllowEmpty</name><value>False</value></property><property><name>ImpressionField</name><value>False</value></property><property><name>DoesNotIndicateFindings</name><value>False</value></property><property><name>FindingsCodes</name><value></value></property><property><name>EnforcePickList</name><value>False</value></property></customproperties></field><field type="1" start="68" length="5"><name>right</name><defaultvalue>36 mm</defaultvalue><customproperties><property><name>AllCaps</name><value>False</value></property><property><name>AllowEmpty</name><value>False</value></property><property><name>ImpressionField</name><value>False</value></property><property><name>DoesNotIndicateFindings</name><value>True</value></property><property><name>FindingsCodes</name><value></value></property><property><name>EnforcePickList</name><value>False</value></property></customproperties></field><field type="1" start="105" length="13"><name>right nodules</name><defaultvalue>right nodules</defaultvalue><customproperties><property><name>AllCaps</name><value>False</value></property><property><name>AllowEmpty</name><value>False</value></property><property><name>ImpressionField</name><value>False</value></property><property><name>DoesNotIndicateFindings</name><value>False</value></property><property><name>FindingsCodes</name><value></value></property><property><name>EnforcePickList</name><value>False</value></property></customproperties></field><field type="1" start="119" length="11"><name>Right nodes</name><defaultvalue>No right lymphadenopathy.</defaultvalue><customproperties><property><name>ImpressionField</name><value>False</value></property><property><name>IncludeInImpression</name><value>False</value></property><property><name>FindingsCodes</name><value></value></property></customproperties></field><field type="1" start="147" length="4"><name>left</name><defaultvalue>49 mm</defaultvalue><customproperties><property><name>AllCaps</name><value>False</value></property><property><name>AllowEmpty</name><value>False</value></property><property><name>ImpressionField</name><value>False</value></property><property><name>DoesNotIndicateFindings</name><value>True</value></property><property><name>FindingsCodes</name><value></value></property><property><name>EnforcePickList</name><value>False</value></property></customproperties></field><field type="1" start="183" length="12"><name>left nodules</name><defaultvalue>left nodules</defaultvalue><customproperties><property><name>AllCaps</name><value>False</value></property><property><name>AllowEmpty</name><value>False</value></property><property><name>ImpressionField</name><value>False</value></property><property><name>DoesNotIndicateFindings</name><value>False</value></property><property><name>FindingsCodes</name><value></value></property><property><name>EnforcePickList</name><value>False</value></property></customproperties></field><field type="1" start="196" length="10"><name>Left nodes</name><defaultvalue>No left lymphadenopathy.</defaultvalue><customproperties><property><name>AllCaps</name><value>False</value></property><property><name>AllowEmpty</name><value>False</value></property><property><name>ImpressionField</name><value>False</value></property><property><name>DoesNotIndicateFindings</name><value>True</value></property><property><name>FindingsCodes</name><value></value></property><property><name>EnforcePickList</name><value>False</value></property></customproperties></field><field type="3" start="209" length="28"><name>parathyroid</name><choices><choice name="none"></choice><choice name="parathyroid">Parathyroid: No candidate adenoma.</choice></choices><customproperties><property><name>AllCaps</name><value>False</value></property><property><name>AllowEmpty</name><value>False</value></property><property><name>ImpressionField</name><value>False</value></property><property><name>DoesNotIndicateFindings</name><value>False</value></property><property><name>FindingsCodes</name><value></value></property><property><name>EnforcePickList</name><value>False</value></property></customproperties></field><field type="3" start="251" length="217"><name>IMPRESSION</name><choices><choice name="0-1 benign, no FNA"></choice><choice name="2 not susp, no FNA"></choice><choice name="3&amp;&#x2265;1.5cm mildly susp, +1/3/5y"></choice><choice name="3&amp;&#x2265;2.5cm mildly susp, FNA"></choice><choice name="4-6&amp;&#x2265;1cm mod susp, +1/2/3/5y"></choice><choice name="4-6&amp;&#x2265;1.5cm mod susp, FNA"></choice><choice name="&#x2265;7&amp;&#x2265;0.5cm highly susp, +1/2/3/4/5y"></choice><choice name="&#x2265;7&amp;&#x2265;1cm highly susp, +FNA"></choice></choices><customproperties><property><name>AllCaps</name><value>False</value></property><property><name>AllowEmpty</name><value>False</value></property><property><name>ImpressionField</name><value>True</value></property><property><name>DoesNotIndicateFindings</name><value>True</value></property><property><name>FindingsCodes</name><value></value></property><property><name>EnforcePickList</name><value>False</value></property></customproperties></field><field type="1" start="470" length="16"><name>RECOMMENDATIONS:</name><customproperties><property><name>AllCaps</name><value>False</value></property><property><name>ImpressionField</name><value>False</value></property><property><name>DoesNotIndicateFindings</name><value>True</value></property><property><name>FindingsCodes</name><value></value></property></customproperties></field><field type="4" start="487" length="13" mergeid="10000" mergename="Verbal Prelim"><name>Verbal Prelim</name></field></fields><links><link target="at https://www.acr.org/Clinical-Resources/Clinical-Tools-and-Reference/Reporting-and-Data-Systems/TI-RADS" detected="1" start="613" length="73"><name>link</name></link></links><textSource><range type="1" start="20" length="7" /><range type="1" start="35" length="15" /><range type="1" start="68" length="5" /><range type="1" start="105" length="13" /><range type="1" start="119" length="11" /><range type="1" start="147" length="4" /><range type="1" start="183" length="12" /><range type="1" start="196" length="10" /><range type="3" start="209" length="11" /><range type="3" start="251" length="10" /><range type="1" start="470" length="16" /><range type="4" start="487" length="13" /></textSource></autotext>
 )'





TN_CopyReport(*) {
    global TN_GuiObj
    global TN_RtfTemplate
    global TN_IsthmusText, TN_RightText, TN_LeftText

    ; Plain text for non-RTF paste targets
    plain := TN_GuiObj["TN_OutputBox"].Value

    ; Start from the embedded template (your NEW DELETE TEMP.rtf content)
    rtf := TN_RtfTemplate

    ; ------------------------------------------------------------
    ; 1) Length fields: "X mm" from the top of the GUI
    ;    XML field names: isthmus, right, left
    ; ------------------------------------------------------------
    istThick := Trim(TN_GuiObj["TN_IsthmusThickness"].Value)
    rightLen := Trim(TN_GuiObj["TN_RightLobeLength"].Value)
    leftLen  := Trim(TN_GuiObj["TN_LeftLobeLength"].Value)

    if (istThick != "")
        rtf := TN_XmlSetDefault(rtf, "isthmus", istThick " mm")
    else
        rtf := TN_XmlSetDefault(rtf, "isthmus", "")

    if (rightLen != "")
        rtf := TN_XmlSetDefault(rtf, "right", rightLen " mm")
    else
        rtf := TN_XmlSetDefault(rtf, "right", "")

    if (leftLen != "")
        rtf := TN_XmlSetDefault(rtf, "left", leftLen " mm")
    else
        rtf := TN_XmlSetDefault(rtf, "left", "")

    ; ------------------------------------------------------------
    ; 2) Nodule fields:
    ;    XML field names: "isthmus nodules", "right nodules", "left nodules"
    ;
    ;    Rule: IF there are nodules, wrap with a newline before and after,
    ;    and those blank lines should NOT have bullets.
    ; ------------------------------------------------------------

    defaultText := "No significant nodules."

    istText := RTrim(TN_IsthmusText, "`r`n")
    rtText  := RTrim(TN_RightText,   "`r`n")
    ltText  := RTrim(TN_LeftText,    "`r`n")

    ; Isthmus
    if (istText != "") {
        if (istText = defaultText)
            istField := istText             ; no extra blank lines for default text
        else
            istField := "`n" istText "`n"   ; blank line before & after real nodules
    } else {
        istField := ""
    }

    ; Right
    if (rtText != "") {
        if (rtText = defaultText)
            rtField := rtText
        else
            rtField := "`n" rtText "`n"
    } else {
        rtField := ""
    }

    ; Left
    if (ltText != "") {
        if (ltText = defaultText)
            ltField := ltText
        else
            ltField := "`n" ltText "`n"
    } else {
        ltField := ""
    }

    rtf := TN_XmlSetDefault(rtf, "isthmus nodules", istField)
    rtf := TN_XmlSetDefault(rtf, "right nodules",   rtField)
    rtf := TN_XmlSetDefault(rtf, "left nodules",    ltField)

    ; ------------------------------------------------------------
    ; 3) Put both plain text AND full RTF (with updated defaults) on clipboard
    ; ------------------------------------------------------------
    TN_SetClipboardTextAndRTF(plain, rtf)
}




TN_RtfEscape(str) {
    ; Escape backslashes and braces for RTF
    str := StrReplace(str, "\", "\\")
    str := StrReplace(str, "{", "\{")
    str := StrReplace(str, "}", "\}")
    return str
}

TN_BuildRTF(plain) {
    ; used to be lines starting with "Â·<TAB>" become list-style bullet paragraphs
    ; Non-bullet lines (including blanks) explicitly reset paragraph formatting
    ; so the list does not "spill over" between nodules.
	; note that I am changing to this character: *

    lines := StrSplit(plain, "`n", "`r")

    rtf  := "{\rtf1\ansi\deff0"
    ; f0 = body font, f1 = Symbol for bullets
    rtf .= "{\fonttbl{\f0 Segoe UI;}{\f1\fnil\fcharset2 Symbol;}}"
    rtf .= "\viewkind4\uc1" . "`n"

    for index, line in lines {
        line := RTrim(line, "`r")

        ; Blank line: ensure we are NOT in a list
        if (line = "") {
            ; reset formatting, then plain blank paragraph
            rtf .= "\pard\plain\par" . "`n"
            continue
        }

        ; Bullet-style paragraph: Bullet (was middot) + TAB prefix in plain text
        if (SubStr(line, 1, 2) = "*`t") {
            ; Strip "*<TAB>"
            text := SubStr(line, 3)
            textEsc := TN_RtfEscape(text)

            ; List-style bullet paragraph:
            ; \pard{\pntext\f1\'B7\tab}{\*\pn\pnlvlblt\pnf1\pnindent0{\pntxtb\'B7}}\fi-360\li360 <text>\par
            rtf .= "\pard"
            rtf .= "{\pntext\f1\'B7\tab}"
            rtf .= "{\*\pn\pnlvlblt\pnf1\pnindent0{\pntxtb\'B7}}"
            rtf .= "\fi-360\li360 " textEsc "\par" . "`n"
        } else {
            ; Normal paragraph: reset formatting so it's not part of the list
            textEsc := TN_RtfEscape(line)
            rtf .= "\pard\plain\f0 " textEsc "\par" . "`n"
        }
    }

    rtf .= "}"
    return rtf
}




TN_SetClipboardTextAndRTF(plain, rtf) {
    ; plain : Unicode string
    ; rtf   : ASCII/RTF string
    
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

TN_PlainToRtfInline(plain) {
    ; Convert a plain-text block (with `n line breaks) into
    ; RTF paragraphs suitable to drop in where [[...]] sits.
    ; Uses current default formatting from the template.
    plain := RTrim(plain, "`r`n")
    if (plain = "")
        return ""

    lines := StrSplit(plain, "`n", "`r")
    frag := ""
    for _, line in lines {
        line := RTrim(line, "`r")
        if (line = "") {
            frag .= "\par "
        } else {
            esc := TN_RtfEscape(line)
            frag .= esc "\par "
        }
    }
    return frag
}


