; ============================================================================
; MODULE: VISAGE SERIES:IMAGE INSERTER (v2)
; Written by Reece J. Goiffon, MD, PhD
; AutoHotkey v2
; Can be run standalone or included in RAHKET_Main.ahk
; ============================================================================

#Requires AutoHotkey v2.0

; --- Enhanced DPI Awareness and Coordinate System Setup ---
; Try multiple DPI awareness levels for maximum compatibility
try {
    ; First try Per-Monitor V2 (best for mixed DPI setups)
    DllCall("User32.dll\SetProcessDpiAwarenessContext", "ptr", -4, "ptr")
} catch {
    try {
        ; Fallback to Per-Monitor V1
        DllCall("User32.dll\SetProcessDpiAwarenessContext", "ptr", -3, "ptr")
    } catch {
        try {
            ; Fallback to System DPI Aware
            DllCall("User32.dll\SetProcessDpiAwarenessContext", "ptr", -2, "ptr")
        } catch {
            ; If all fail, continue without (Windows will use default)
        }
    }
}

; Set coordinate modes to Screen for consistent positioning across monitors
CoordMode "Mouse",   "Screen"
CoordMode "Pixel",   "Screen"
CoordMode "ToolTip", "Screen"

; Global variables
global capturedImage := ""
global startX := 0
global startY := 0
global isDragging := false
global overlayGui := ""
global selectionBox := ""
global selectionMode := false
global annotationColor := ""
global ImageInserter_Active := false
global learnedCropWidth := 0
global learnedCropHeight := 0
global cropLearned := false
global activeFlashGuis := []  ; Track active flash message GUIs
global ControlCropMap := Map()  ; key: "ctrlWxctrlH" → { w: cropWidth, h: cropHeight }
global g_SearchGui := ""        ; progress popup while searching
global g_SearchText := ""       ; text control inside that GUI

; --- OCR crop bounds ---
global MAX_OCR_WIDTH  := 320      ; hard ceiling in pixels
global MAX_OCR_HEIGHT := 220      ; header is a short strip
global MAX_OCR_MONITOR_FRACTION := 0.30  ; 30% of monitor pixel area

; --- Visage version ---
global VISAGE_EXE_PATH := "C:\Program Files\Visage Imaging\Visage 7.1\bin\arch-Win\vsclient.exe"
global g_VisageIsNewUI := ""  ; "" = not yet checked, true/false once resolved



; ============================================================================
; MODULE CONTROL FUNCTIONS (for RAHKET integration)
; ============================================================================

global ImageInserter_IsActive := false

; Register the backtick hotkey ONCE
Hotkey("``", ImageInserter_HotkeyRouter)


GetVisageVersion() {
    ; Returns the file version of the Visage client exe as an object with
    ; Major, Minor, Build, Patch integer fields, or "" on failure.
    global VISAGE_EXE_PATH
    ver := ""
    try {
        ver := FileGetVersion(VISAGE_EXE_PATH)
    } catch {
        return ""
    }
    if (ver = "")
        return ""
    parts := StrSplit(ver, ".")
    return {
        Major: (parts.Length >= 1 ? Integer(parts[1]) : 0),
        Minor: (parts.Length >= 2 ? Integer(parts[2]) : 0),
        Build: (parts.Length >= 3 ? Integer(parts[3]) : 0),
        Patch: (parts.Length >= 4 ? Integer(parts[4]) : 0),
        Full:  ver
    }
}


IsNewVisageUI() {
    ; Returns true if Visage version is >= 7.1.19 (new DICOM dialog UI).
    ; Caches the result so the exe is only queried once per session.
    global g_VisageIsNewUI
    if (g_VisageIsNewUI != "")
        return g_VisageIsNewUI
    v := GetVisageVersion()
    if (v = "") {
        g_VisageIsNewUI := false
        return false
    }
    ; Compare: major > 7, or major=7 and minor > 1, or major=7 minor=1 and build >= 19
    if (v.Major > 7)
        result := true
    else if (v.Major = 7 && v.Minor > 1)
        result := true
    else if (v.Major = 7 && v.Minor = 1 && v.Build >= 19)
        result := true
    else
        result := false
    g_VisageIsNewUI := result
    return result
}


ResetCropLearningState() {
    global ControlCropMap, learnedCropWidth, learnedCropHeight, cropLearned
    global g_SearchGui, g_SearchText

    ; Forget all stored crop sizes
    ControlCropMap := Map()

    ; Reset any legacy/extra learning flags
    learnedCropWidth  := 0
    learnedCropHeight := 0
    cropLearned       := false

    ; Kill any leftover search popup
    if IsObject(g_SearchGui) {
        try g_SearchGui.Destroy()
    }
    g_SearchGui  := ""
    g_SearchText := ""
}



ImageInserter_Enable() {
    global ImageInserter_IsActive

    if (ImageInserter_IsActive) {
        ImageInserter_Disable()
        Sleep(100)
    }

    ; If we're being run as a module (i.e., from RAHKET_Main),
    ; forget learned crop sizes and relearn from scratch.
    if (A_LineFile != A_ScriptFullPath) {
        ResetCropLearningState()
    }

    ImageInserter_Initialize()
}


ImageInserter_Disable() {
    global ImageInserter_IsActive, overlayGui, selectionBox
    global selectionMode, isDragging

    ; Stop the module
    ImageInserter_IsActive := false

    ; --- NEW: kill any running timers ---
    SetTimer(CheckForShift, 0)
    SetTimer(MonitorMouse, 0)

    ; --- NEW: hard reset selection state ---
    selectionMode := false
    isDragging := false

    ; Clean up GUIs
    try {
        if overlayGui
            overlayGui.Destroy()
        overlayGui := ""
    }
    
    try {
        if IsObject(selectionBox) {
            selectionBox.Destroy()
            selectionBox := ""
        }
    }
    
    TrayTip("Image Inserter disabled", "RAHKET", 1)
}


ImageInserter_Initialize() {
    global ImageInserter_IsActive, selectionMode, capturedImage

    ImageInserter_IsActive := true
    selectionMode := false
    capturedImage := ""
    ImageInserter_Active := false

    ; New Visage UI (>= 7.1.19) uses the DICOM parameters dialog — no calibration needed.
    if IsNewVisageUI() {
        v := GetVisageVersion()
        MsgBox("Visage " . v.Full . " detected.`n`nNo calibration needed.`nPress `` while over a Visage viewer to capture series/image info.`nPress `` in PowerScribe to reformat pairs.",
               "Image Inserter Ready", "64")
        return
    }

    ; Legacy path: OCR-based capture requires color calibration.
    result := MsgBox("This works best with colored window text. Once set to color, calibrate text color by:`n`n" .
                     "1) opening a study`n" .
                     "2) adjusting image position and/or window and level to put DICOM information against black.`n" .
                     "3) when ready, hold down Shift and drag the box around the DICOM text in the corner of the image", 
                     "DICOM Header Calibration", "OK 64")
    
    ; Wait for Shift to be pressed to enter selection mode
    SetTimer(CheckForShift, 50)
}

; ============================================================================
; AUTO-START (only if running standalone)
; ============================================================================

if (A_LineFile = A_ScriptFullPath) {
    ; Running standalone
    #SingleInstance Force
    ImageInserter_Initialize()
}

CheckForShift() {
    global selectionMode, ImageInserter_IsActive

    ; --- NEW: if disabled, stop this timer and exit ---
    if !ImageInserter_IsActive {
        SetTimer(CheckForShift, 0)
        return
    }

    if GetKeyState("Shift", "P") && !selectionMode {
        selectionMode := true
        SetTimer(CheckForShift, 0)
        StartSelectionMode()
    }
}

StartSelectionMode() {
    global overlayGui, isDragging, startX, startY, selectionMode
    
    ; Get virtual screen dimensions (all monitors combined)
    virtualScreenLeft := SysGet(76)    ; SM_XVIRTUALSCREEN
    virtualScreenTop := SysGet(77)     ; SM_YVIRTUALSCREEN
    virtualScreenWidth := SysGet(78)   ; SM_CXVIRTUALSCREEN
    virtualScreenHeight := SysGet(79)  ; SM_CYVIRTUALSCREEN
    
    ; Create a fullscreen transparent overlay to capture mouse events
    overlayGui := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20")
    overlayGui.BackColor := "White"
    WinSetTransparent(1, overlayGui)
    overlayGui.Show("X" . virtualScreenLeft . " Y" . virtualScreenTop . " W" . virtualScreenWidth . " H" . virtualScreenHeight)
    
    ; Create selection box GUI (initially hidden)
    global selectionBox
    selectionBox := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20")
    selectionBox.BackColor := "Red"
    WinSetTransparent(100, selectionBox)
    
    ; Set up mouse monitoring
    SetTimer(MonitorMouse, 10)
}

MonitorMouse() {
    global isDragging, startX, startY, overlayGui, selectionBox, selectionMode, capturedImage
    global ImageInserter_IsActive

    ; --- if disabled, teardown and stop monitoring ---
    if !ImageInserter_IsActive {
        SetTimer(MonitorMouse, 0)
        try if IsObject(overlayGui) overlayGui.Destroy()
        try if IsObject(selectionBox) selectionBox.Destroy()
        overlayGui := ""
        selectionBox := ""
        selectionMode := false
        isDragging := false
        return
    }

    ; Check if Shift is still held
    if !GetKeyState("Shift", "P") {
        ; Shift released
        SetTimer(MonitorMouse, 0)
        
        if IsObject(overlayGui) {
            overlayGui.Destroy()
            overlayGui := ""
        }
        
        if IsObject(selectionBox) {
            selectionBox.Destroy()
            selectionBox := ""
        }
        
        isDragging := false
        selectionMode := false
        
        ; If we captured something, show confirmation
        if capturedImage && FileExist(capturedImage) {
            ShowConfirmation()
        } else {
            ; No capture, restart
            SetTimer(CheckForShift, 50)
        }
        return
    }
    
    ; Check left button state
    if GetKeyState("LButton", "P") {
        if !isDragging {
            ; Start new drag
            isDragging := true
            MouseGetPos(&startX, &startY)
        } else {
            ; Update selection box
            MouseGetPos(&currentX, &currentY)
            
            selectionX := Min(startX, currentX)
            selectionY := Min(startY, currentY)
            selectionW := Abs(currentX - startX)
            selectionH := Abs(currentY - startY)
            
            if (selectionW > 0 && selectionH > 0) {
                selectionBox.Show("NA X" . selectionX . " Y" . selectionY . " W" . selectionW . " H" . selectionH)
            }
        }
    } else if isDragging {
        ; Button released - capture
        isDragging := false
        MouseGetPos(&endX, &endY)
        
        ; Hide selection box
        if IsObject(selectionBox) {
            selectionBox.Hide()
        }
        
        captureX := Min(startX, endX)
        captureY := Min(startY, endY)
        captureW := Abs(endX - startX)
        captureH := Abs(endY - startY)
        
        if (captureW > 5 && captureH > 5) {
            CaptureArea(captureX, captureY, captureW, captureH)
        }
    }
}

CaptureArea(x, y, w, h) {
    global capturedImage
    
    ; Create a temporary file path
    tempPath := "C:\temp"
    if !DirExist(tempPath) {
        DirCreate(tempPath)
    }
    
    tempFile := tempPath . "\dicom_header_temp.png"
    
    ; Capture the screen area
    try {
        pToken := Gdip_Startup()
        pBitmap := Gdip_BitmapFromScreen(x . "|" . y . "|" . w . "|" . h)
        Gdip_SaveBitmapToFile(pBitmap, tempFile)
        Gdip_DisposeImage(pBitmap)
        Gdip_Shutdown(pToken)
        
        capturedImage := tempFile
    } catch as err {
        MsgBox("Error capturing screen: " . err.Message, "Error", "16")
    }
}

Gdip_CloneBitmapArea(pBitmap, x, y, w, h) {
    pClone := 0
    DllCall("gdiplus\GdipCloneBitmapArea", "Float", x, "Float", y, "Float", w, "Float", h, "Int", 0x26200A, "Ptr", pBitmap, "PtrP", &pClone)
    return pClone
}

ShowConfirmation() {
    global capturedImage
    
    if !capturedImage || !FileExist(capturedImage) {
        MsgBox("No area was captured. Please try again.", "Error", "16")
        SetTimer(CheckForShift, 50)
        return
    }
    
    ; Create confirmation GUI with image preview
    confirmGui := Gui("+AlwaysOnTop", "Confirm DICOM Header Area")
    confirmGui.OnEvent("Escape", (*) => ExitApp())
    confirmGui.OnEvent("Close", (*) => ExitApp())
    
    ; Add the captured image
    try {
        confirmGui.Add("Picture", "w400 h-1", capturedImage)
    }
    
    confirmGui.Add("Text", "w400 Center", "Accept DICOM header area?")
    
    ; Add buttons
    btnYes := confirmGui.Add("Button", "w80 x100", "&Yes")
    btnNo := confirmGui.Add("Button", "w80 x200 yp", "&No")
    btnCancel := confirmGui.Add("Button", "w80 x300 yp", "&Cancel")
    
    btnYes.OnEvent("Click", (*) => OnConfirmYes(confirmGui))
    btnNo.OnEvent("Click", (*) => OnConfirmNo(confirmGui))
    btnCancel.OnEvent("Click", (*) => ExitApp())
    
    confirmGui.Show()
}

OnConfirmYes(guiObj) {
    guiObj.Destroy()
    AnalyzeColors()
}

OnConfirmNo(guiObj) {
    global capturedImage
    
    guiObj.Destroy()
    
    ; Clean up temp file
    if FileExist(capturedImage) {
        try FileDelete(capturedImage)
    }
    
    capturedImage := ""
    
    ; Go back to selection
    SetTimer(CheckForShift, 50)
}

AnalyzeColors() {
    global capturedImage, annotationColor
    
    if !FileExist(capturedImage) {
        MsgBox("Error: Captured image not found.", "Error", "16")
        ExitApp()
    }
    
    ; Load the image and analyze pixel colors
    pToken := Gdip_Startup()
    pBitmap := Gdip_CreateBitmapFromFile(capturedImage)
    
    if !pBitmap {
        MsgBox("Error: Could not load captured image.", "Error", "16")
        Gdip_Shutdown(pToken)
        ExitApp()
    }
    
    ; Get image dimensions
    width := Gdip_GetImageWidth(pBitmap)
    height := Gdip_GetImageHeight(pBitmap)
    
    ; Count color frequencies
    colorCounts := Map()
    
    loop height {
        y := A_Index - 1
        loop width {
            x := A_Index - 1
            color := Gdip_GetPixel(pBitmap, x, y)
            
            ; Convert ARGB to RGB (ignore alpha)
            rgb := color & 0xFFFFFF
            
            if colorCounts.Has(rgb) {
                colorCounts[rgb] := colorCounts[rgb] + 1
            } else {
                colorCounts[rgb] := 1
            }
        }
    }
    
    ; Sort colors by frequency
    colorArray := []
    for color, count in colorCounts {
        colorArray.Push({color: color, count: count})
    }
    
    ; Sort by count descending
    colorArray := SortByCount(colorArray)
    
    ; Get 2nd most frequent color (index 2, since index 1 is most frequent)
    if colorArray.Length < 2 {
        MsgBox("Error: Not enough distinct colors found in image.", "Error", "16")
        Gdip_DisposeImage(pBitmap)
        Gdip_Shutdown(pToken)
        ExitApp()
    }
    
    secondColor := colorArray[2].color
    annotationColor := secondColor
    
    ; Extract RGB components
    r := (secondColor >> 16) & 0xFF
    g := (secondColor >> 8) & 0xFF
    b := secondColor & 0xFF
    
    ; Create filtered image showing only the annotation color
    filteredPath := "C:\temp\dicom_filtered_temp.png"
    pFilteredBitmap := CreateFilteredBitmap(pBitmap, width, height, secondColor)
    Gdip_SaveBitmapToFile(pFilteredBitmap, filteredPath)
    
    ; Clean up
    Gdip_DisposeImage(pFilteredBitmap)
    Gdip_DisposeImage(pBitmap)
    Gdip_Shutdown(pToken)
    
    ; Show color confirmation GUI
    ShowColorConfirmation(secondColor, r, g, b, filteredPath)
}

SortByCount(arr) {
    ; Simple bubble sort by count (descending)
    n := arr.Length
    loop n - 1 {
        i := A_Index
        loop n - i {
            j := A_Index
            if arr[j].count < arr[j + 1].count {
                temp := arr[j]
                arr[j] := arr[j + 1]
                arr[j + 1] := temp
            }
        }
    }
    return arr
}

CreateFilteredBitmap(pSourceBitmap, width, height, targetColor) {
    ; Create new bitmap for filtered image
    pNewBitmap := Gdip_CreateBitmap(width, height)
    pGraphics := Gdip_GraphicsFromImage(pNewBitmap)
    
    ; Fill with black background
    Gdip_GraphicsClear(pGraphics, 0xFF000000)
    
    ; Extract target RGB
    targetR := (targetColor >> 16) & 0xFF
    targetG := (targetColor >> 8) & 0xFF
    targetB := targetColor & 0xFF
    
    ; Color tolerance (adjust if needed)
    tolerance := 10
    
    ; Copy pixels that match the target color (with tolerance)
    loop height {
        y := A_Index - 1
        loop width {
            x := A_Index - 1
            pixelColor := Gdip_GetPixel(pSourceBitmap, x, y) & 0xFFFFFF
            
            ; Extract pixel RGB
            pixelR := (pixelColor >> 16) & 0xFF
            pixelG := (pixelColor >> 8) & 0xFF
            pixelB := pixelColor & 0xFF
            
            ; Check if within tolerance
            if (Abs(pixelR - targetR) <= tolerance 
                && Abs(pixelG - targetG) <= tolerance 
                && Abs(pixelB - targetB) <= tolerance) {
                ; Set pixel in new bitmap (use original pixel color)
                Gdip_SetPixel(pNewBitmap, x, y, 0xFF000000 | pixelColor)
            }
        }
    }
    
    Gdip_DeleteGraphics(pGraphics)
    
    return pNewBitmap
}

ShowColorConfirmation(color, r, g, b, filteredPath) {
    global annotationColor, capturedImage
    
    ; Check if color is grayscale
    isGrayscale := (r = g && g = b)
    
    ; Create confirmation GUI
    colorGui := Gui("+AlwaysOnTop", "Confirm Annotation Color")
    colorGui.OnEvent("Escape", (*) => ExitApp())
    colorGui.OnEvent("Close", (*) => ExitApp())
    
    ; Show warning if grayscale
    if isGrayscale {
        warningText := colorGui.Add("Text", "w500 cRed", 
            "CAUTION: the proposed color is grayscale. This will probably break the program.`n" .
            "Consider setting your annotation colors in File → Preferences → Properties`n" .
            "and searching for `"Text Overlays in Viewer`" and changing your `"Text and Line Color.`"")
        warningText.SetFont("bold")
        colorGui.Add("Text", "w500", "") ; Spacer
    }
    
    ; Add color swatch
    colorGui.Add("Text", "w500", "Detected annotation color (2nd most frequent):")
    
    ; Create color swatch (100x100 box)
    hexColor := Format("{:06X}", color)
    colorGui.Add("Progress", "w100 h100 Background" . hexColor . " -Smooth")
    colorGui.Add("Text", "yp xp+110", "RGB: " . r . ", " . g . ", " . b . "`nHex: #" . hexColor)
    
    ; Add filtered image preview
    colorGui.Add("Text", "w500 Section", "`nFiltered image (showing only annotation color pixels):")
    try {
        colorGui.Add("Picture", "w400 h-1", filteredPath)
    }
    
    colorGui.Add("Text", "w500 Center", "`nAccept annotation color?")
    
    ; Add buttons
    btnYes := colorGui.Add("Button", "w80 x140", "&Yes")
    btnNo := colorGui.Add("Button", "w80 x240 yp", "&No")
    btnCancel := colorGui.Add("Button", "w80 x340 yp", "&Cancel")
    
    btnYes.OnEvent("Click", (*) => OnColorAccepted(colorGui, filteredPath))
    btnNo.OnEvent("Click", (*) => OnColorRejected(colorGui, filteredPath))
    btnCancel.OnEvent("Click", (*) => ExitApp())
    
    colorGui.Show()
}

OnColorAccepted(guiObj, filteredPath) {
    global capturedImage, annotationColor, ImageInserter_Active
    
    guiObj.Destroy()
    
    ; Clean up temp files
    if FileExist(capturedImage) {
        try FileDelete(capturedImage)
    }
    if FileExist(filteredPath) {
        try FileDelete(filteredPath)
    }
    
    ; Extract RGB for display
    r := (annotationColor >> 16) & 0xFF
    g := (annotationColor >> 8) & 0xFF
    b := annotationColor & 0xFF
    
    ; FIX #2: Properly escape backtick in MsgBox
    MsgBox("Annotation color accepted!`nRGB: " . r . ", " . g . ", " . b . "`n`nScript is now active.`nPress `` to capture series/image info.`nPress `` in PowerScribe to reformat pairs.", "Success", "64")
    

}

OnColorRejected(guiObj, filteredPath) {
    global capturedImage
    
    guiObj.Destroy()
    
    ; Clean up temp files
    if FileExist(capturedImage) {
        try FileDelete(capturedImage)
    }
    if FileExist(filteredPath) {
        try FileDelete(filteredPath)
    }
    
    capturedImage := ""
    
    ; Go back to selection
    SetTimer(CheckForShift, 50)
}


ImageInserter_HotkeyRouter(*) {
    global ImageInserter_IsActive

    ; If the module is disabled via menu, do nothing
    if !ImageInserter_IsActive {
        ShowFlashMessage("Image Inserter is disabled in the menu.", 400, 300)
        return
    }

    ; If PowerScribe is active, reformat pairs
    if WinActive("ahk_exe Nuance.PowerScribeOne.exe") {
        ReformatSeriesImage()
        return
    }

    ; Route to correct capture method based on Visage version
    if IsNewVisageUI() {
        CaptureSeriesImage_NewUI()
    } else {
        CaptureSeriesImage()
    }
}


CaptureSeriesImage_NewUI(*) {
    ; Used for Visage >= 7.1.19.
    ; Opens the DICOM Parameters dialog via Ctrl+D, reads the clipboard content,
    ; then parses series (0020-0011) and instance (0020-0013) numbers from it.

    MouseGetPos(&mouseX, &mouseY, &mouseWinID)

    ; If a Visage window is already active, use it directly without touching focus.
    ; If not, check if the mouse is over a Visage window and click to activate it.
    activeWinID := WinGetID("A")
    if InStr(WinGetProcessName("ahk_id " . activeWinID), "vsclient.exe") {
        visageWinID := activeWinID
    } else {
        if !InStr(WinGetProcessName("ahk_id " . mouseWinID), "vsclient.exe") {
            ShowFlashMessage("No Visage window is active or under the mouse.", mouseX, mouseY, "Red")
            return
        }
        ; Click to activate the Visage window under the mouse
        Click(mouseX, mouseY)
        Sleep(100)
        visageWinID := WinGetID("A")
        if !InStr(WinGetProcessName("ahk_id " . visageWinID), "vsclient.exe") {
            ShowFlashMessage("Could not activate Visage.", mouseX, mouseY, "Red")
            return
        }
    }

    ; Block mouse from Ctrl+D send through paste completion
    BlockInput("Mouse")
    BlockInput("MouseMove")
    safetyTimer := () => (BlockInput("Off"), BlockInput("MouseMoveOff"))
    SetTimer(safetyTimer, -3000)
    escUnblock := (*) => (BlockInput("Off"), BlockInput("MouseMoveOff"), SetTimer(safetyTimer, 0), Hotkey("Escape", escUnblock, "Off"))
    Hotkey("Escape", escUnblock, "On")

    ; Send Ctrl+D to the already-active window without re-activating it
    Send("^d")

    ; Wait for the DICOM Parameters dialog
    if !WinWait("DICOM Parameters of Active Viewer ahk_exe vsclient.exe", , 5) {
        SetTimer(safetyTimer, 0)
        Hotkey("Escape", escUnblock, "Off")
        BlockInput("Off")
        BlockInput("MouseMoveOff")
        ShowFlashMessage("DICOM Parameters dialog did not open.", mouseX, mouseY, "Red")
        return
    }
    WinActivate("DICOM Parameters of Active Viewer ahk_exe vsclient.exe")
    WinWaitActive("DICOM Parameters of Active Viewer ahk_exe vsclient.exe", , 2)

    ; Save current clipboard so it can be restored
    oldClipboard := A_Clipboard
    A_Clipboard := ""

    ; Tab 5 times to focus the content area, then select all and copy
    Send("{Tab 5}")
    Sleep(100)
    Send("^a")
    Sleep(50)
    Send("^c")

    ; Wait for clipboard to be populated
    if !ClipWait(3) {
        Send("{Escape}")
        A_Clipboard := oldClipboard
        SetTimer(safetyTimer, 0)
        Hotkey("Escape", escUnblock, "Off")
        BlockInput("Off")
        BlockInput("MouseMoveOff")
        ShowFlashMessage("No data copied from DICOM dialog.", mouseX, mouseY, "Red")
        return
    }

    dicomText := A_Clipboard

    ; Close the dialog
    WinClose("DICOM Parameters of Active Viewer ahk_exe vsclient.exe")
    WinWaitActive("ahk_exe vsclient.exe", , 2)

    ; Restore clipboard after a short delay
    SetTimer(RestoreClipboard.Bind(oldClipboard, dicomText), -500)

    ; ------------------------------------------------------------------
    ; Parse series number: find "0020-0011", then "SeriesNumber", then
    ; the next non-empty line is the value.
    ; Same pattern for instance number: "0020-0013" / "InstanceNumber".
    ; ------------------------------------------------------------------
    seriesNum := ParseDicomField(dicomText, "0020-0011", "SeriesNumber")
    imaNum    := ParseDicomField(dicomText, "0020-0013", "InstanceNumber")

    if (seriesNum = "" || imaNum = "") {
        ShowFlashMessage("Could not parse Series/IMA from DICOM data.", mouseX, mouseY, "Red")
        return
    }

    ; Insert into PowerScribe, same as the legacy path
    textToSend := " (" . seriesNum . ":" . imaNum . ")"

    flashHandles := ShowFlashMessage("Series " . seriesNum . " IMA " . imaNum,
                                     mouseX, mouseY, "Green")

    currentWinID := WinGetID("A")

    try {
        if WinExist("ahk_exe Nuance.PowerScribeOne.exe") {
            BlockInput("Mouse")
            BlockInput("MouseMove")
            safetyTimer := () => (BlockInput("Off"), BlockInput("MouseMoveOff"))
            SetTimer(safetyTimer, -3000)
            escUnblock := (*) => (BlockInput("Off"), BlockInput("MouseMoveOff"), SetTimer(safetyTimer, 0), Hotkey("Escape", escUnblock, "Off"))
            Hotkey("Escape", escUnblock, "On")

            WinActivate("ahk_exe Nuance.PowerScribeOne.exe")
            WinWaitActive("ahk_exe Nuance.PowerScribeOne.exe", , 1)
            Sleep(20)

            A_Clipboard := textToSend
            Sleep(20)
            SendInput("^v")

            SetTimer(safetyTimer, 0)
            Hotkey("Escape", escUnblock, "Off")
            BlockInput("Off")
            BlockInput("MouseMoveOff")

            SetTimer(RestoreClipboard.Bind(oldClipboard, textToSend), -500)

            for flashGui in flashHandles {
                try flashGui.Destroy()
            }

            Sleep(20)
            WinActivate("ahk_id " . currentWinID)
        } else {
            ShowFlashMessage("PowerScribe window not found.", mouseX, mouseY, "Red")
        }
    } catch as err {
        BlockInput("Off")
        BlockInput("MouseMoveOff")
        ShowFlashMessage("Error activating PowerScribe: " . err.Message, mouseX, mouseY, "Red")
    }
}


ParseDicomField(text, tagLine, labelLine) {
    ; Searches text for tagLine, then labelLine on the next non-empty line,
    ; then returns the value on the line after that.
    ; Strip CR first so CRLF line endings don't leave \r on trimmed values.
    text  := StrReplace(text, "`r", "")
    lines := StrSplit(text, "`n")
    i := 1
    while i <= lines.Length {
        if (Trim(lines[i]) = tagLine) {
            ; Find labelLine on the immediately following non-empty line
            j := i + 1
            while (j <= lines.Length && Trim(lines[j]) = "")
                j++
            if (j <= lines.Length && Trim(lines[j]) = labelLine) {
                ; Value is the next non-empty line after the label
                k := j + 1
                while (k <= lines.Length && Trim(lines[k]) = "")
                    k++
                if (k <= lines.Length) {
                    val := Trim(lines[k])
                    if (val != "")
                        return val
                }
            }
        }
        i++
    }
    return ""
}


VerifyOcrWithLargerCrop(pBitmap, width, height, baseCropW, baseCropH, annotationColor
                      , &seriesNum, &imaNum) {
    ; Try a slightly larger crop to see if we were clipping digits.
    ; If we get a better (longer) IMA, keep it.

	biggerW := Min(width,  Round(baseCropW * 1.35), MAX_OCR_WIDTH)
	biggerH := Min(height, Round(baseCropH * 1.20), MAX_OCR_HEIGHT)


    tmpSeries := ""
    tmpIma    := ""

    if TryOcrOnBitmapRegion(pBitmap, width, height, biggerW, biggerH
                          , annotationColor, &tmpSeries, &tmpIma) {

        ; If the bigger crop yields same series and a longer IMA, accept it.
        if (tmpSeries = seriesNum && StrLen(tmpIma) > StrLen(imaNum)) {
            seriesNum := tmpSeries
            imaNum    := tmpIma
            return { ok: true, w: biggerW, h: biggerH, improved: true }
        }

        ; Even if not longer, but still valid and consistent, prefer bigger crop for stability.
        if (tmpSeries = seriesNum && tmpIma = imaNum) {
            return { ok: true, w: biggerW, h: biggerH, improved: false }
        }
    }

    return { ok: false }
}


CaptureSeriesImage(*) {
    global annotationColor, ControlCropMap
    global g_SearchGui, g_SearchText

    if (annotationColor = "") {
        ShowFlashMessage(
            "Image Inserter not calibrated yet." . "`n" .
            "Hold Shift and drag a box over a header to calibrate.",
            400, 330
        )
        return
    }

    MouseGetPos(&mouseX, &mouseY, &winID)

    winTitle   := ""
    winProcess := ""
    try {
        winTitle   := WinGetTitle("ahk_id " . winID)
        winProcess := WinGetProcessName("ahk_id " . winID)
    }

    isOverVisage := (InStr(winProcess, "vsclient.exe") || InStr(winTitle, "Visage Client"))
    if !isOverVisage {
        ShowFlashMessage("No series/image captured. Mouse not over Visage.", mouseX, mouseY)
        return
    }

    controlID := ""
    try {
        MouseGetPos(, , , &controlID, 2)
    } catch {
        ShowFlashMessage("No series/image captured. Could not get control info.", mouseX, mouseY, "Red")
        return
    }

    ctrlX := 0, ctrlY := 0, ctrlW := 0, ctrlH := 0
    try {
        ControlGetPos(&ctrlX, &ctrlY, &ctrlW, &ctrlH, controlID, "ahk_id " . winID)
    } catch {
        ShowFlashMessage("No series/image captured. Could not get control info.", mouseX, mouseY, "Red")
        return
    }

    winX := 0, winY := 0, winW := 0, winH := 0
    try {
        WinGetPos(&winX, &winY, &winW, &winH, "ahk_id " . winID)
    } catch {
        ShowFlashMessage("No series/image captured. Could not get window position.", mouseX, mouseY, "Red")
        return
    }

    try {
        WinGetClientPos(&clientX, &clientY, , , "ahk_id " . winID)
        offsetX := clientX - winX
        offsetY := clientY - winY
    } catch {
        offsetX := 0
        offsetY := 0
    }

    screenX := winX + offsetX + ctrlX
    screenY := winY + offsetY + ctrlY

    if (ctrlW <= 0 || ctrlH <= 0) {
        ShowFlashMessage("No series/image captured. Invalid control dimensions.", mouseX, mouseY, "Red")
        return
    }

    virtualLeft   := SysGet(76)
    virtualTop    := SysGet(77)
    virtualWidth  := SysGet(78)
    virtualHeight := SysGet(79)
    virtualRight  := virtualLeft + virtualWidth
    virtualBottom := virtualTop + virtualHeight

    if (screenX < virtualLeft)
        screenX := virtualLeft
    if (screenY < virtualTop)
        screenY := virtualTop
    if (screenX + ctrlW > virtualRight)
        ctrlW := virtualRight - screenX
    if (screenY + ctrlH > virtualBottom)
        ctrlH := virtualBottom - screenY

    pToken := 0
    pBitmap := 0

    try {
        pToken  := Gdip_Startup()
        pBitmap := Gdip_BitmapFromScreen(screenX . "|" . screenY . "|" . ctrlW . "|" . ctrlH)

        width  := Gdip_GetImageWidth(pBitmap)
        height := Gdip_GetImageHeight(pBitmap)
        if (width <= 0 || height <= 0)
            throw Error("Invalid bitmap dimensions")

        sizeKey   := ctrlW . "x" . ctrlH
        seriesNum := ""
        imaNum    := ""
        success   := false

        ; =====================================================
        ; 1) Try stored bounding box for this control size
        ; =====================================================
        if ControlCropMap.Has(sizeKey) {
            cropInfo   := ControlCropMap[sizeKey]
            cropWidth  := cropInfo.w
            cropHeight := cropInfo.h

			success := TryOcrOnBitmapRegion(pBitmap, width, height, cropWidth, cropHeight,
											annotationColor, &seriesNum, &imaNum)

			if success {
				verify := VerifyOcrWithLargerCrop(pBitmap, width, height, cropWidth, cropHeight,
												  annotationColor, &seriesNum, &imaNum)
				if (verify.ok) {
					; Update stored crop to the bigger verified size (more future-proof)
					ControlCropMap[sizeKey] := { w: verify.w, h: verify.h }
					cropWidth  := verify.w
					cropHeight := verify.h
				}
			}

        }

		; =====================================================
		; 2) If no success, do learning step with progress popup
		; =====================================================
		
		; Figure out monitor area at the mouse position (where the control is)
		monIdx := GetMonitorAtPoint(mouseX, mouseY)
		MonitorGet(monIdx, &mLeft, &mTop, &mRight, &mBottom)
		monW := mRight - mLeft
		monH := mBottom - mTop
		maxCropArea := monW * monH * MAX_OCR_MONITOR_FRACTION


		if !success {
			; Create / reset the search progress GUI
			if IsObject(g_SearchGui) {
				try g_SearchGui.Destroy()
			}

			g_SearchGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Border")
			g_SearchGui.BackColor := "Black"
			g_SearchGui.SetFont("s10 cWhite")

			percentText := "Searching 10% of control..."
			g_SearchText := g_SearchGui.Add("Text", "w220 Center", percentText)

			g_SearchGui.Show("AutoSize X" . (mouseX + 30) . " Y" . (mouseY + 30))

			for stepPercent in [10, 15, 20, 30, 50, 70, 100] {
				if !IsObject(g_SearchGui)
					break

				g_SearchText.Text := "Searching " . stepPercent . "% of control..."
				g_SearchGui.Show()

				; --- raw percent-based sizes ---
				rawW := width  * stepPercent // 100
				rawH := height * stepPercent // 100

				; --- clamp to absolute pixel caps ---
				cropWidth  := Min(rawW,  MAX_OCR_WIDTH,  width)
				cropHeight := Min(rawH,  MAX_OCR_HEIGHT, height)

				; --- enforce monitor-area fraction cap ---
				cropArea := cropWidth * cropHeight
				if (maxCropArea > 0 && cropArea > maxCropArea) {
					scale := Sqrt(maxCropArea / cropArea)
					cropWidth  := Max(10, Round(cropWidth  * scale))
					cropHeight := Max(10, Round(cropHeight * scale))
					cropArea   := cropWidth * cropHeight
				}

				if TryOcrOnBitmapRegion(pBitmap, width, height, cropWidth, cropHeight,
						annotationColor, &seriesNum, &imaNum) {

					verify := VerifyOcrWithLargerCrop(pBitmap, width, height, cropWidth, cropHeight,
													  annotationColor, &seriesNum, &imaNum)

					if (verify.ok) {
						; also clamp verified size
						cropWidth  := Min(verify.w, MAX_OCR_WIDTH,  width)
						cropHeight := Min(verify.h, MAX_OCR_HEIGHT, height)
					}

					success := true
					ControlCropMap[sizeKey] := { w: cropWidth, h: cropHeight }
					break
				}
			}

			if IsObject(g_SearchGui) {
				try g_SearchGui.Destroy()
				g_SearchGui := ""
				g_SearchText := ""
			}
		}


        ; =====================================================
        ; 3) If still no success, show one red error message
        ; =====================================================
        if !success {
            Gdip_DisposeImage(pBitmap)
            Gdip_Shutdown(pToken)
            ShowFlashMessage("No series/image captured. Series IMA pair not found in screen grab.", mouseX, mouseY, "Red")
            return
        }

        ; Done with bitmap
        Gdip_DisposeImage(pBitmap)
        Gdip_Shutdown(pToken)

        ; =====================================================
        ; 4) SUCCESS: send to PowerScribe
        ; =====================================================
        textToSend := " (" . seriesNum . ":" . imaNum . ")"

        flashHandles := ShowFlashMessage("Series " . seriesNum . " IMA " . imaNum,
                                         mouseX, mouseY, "Green")

        currentWinID := WinGetID("A")

		try {
			if WinExist("ahk_exe Nuance.PowerScribeOne.exe") {
				BlockInput("Mouse")
				BlockInput("MouseMove")
				safetyTimer := () => (BlockInput("Off"), BlockInput("MouseMoveOff"))
				SetTimer(safetyTimer, -3000)
				escUnblock := (*) => (BlockInput("Off"), BlockInput("MouseMoveOff"), SetTimer(safetyTimer, 0), Hotkey("Escape", escUnblock, "Off"))
				Hotkey("Escape", escUnblock, "On")

				WinActivate("ahk_exe Nuance.PowerScribeOne.exe")
				WinWaitActive("ahk_exe Nuance.PowerScribeOne.exe", , 1)
				Sleep(20)

				oldClipboard := A_Clipboard
				A_Clipboard := textToSend
				Sleep(20)
				SendInput("^v")

				SetTimer(safetyTimer, 0)
				Hotkey("Escape", escUnblock, "Off")
				BlockInput("Off")
				BlockInput("MouseMoveOff")

				SetTimer(RestoreClipboard.Bind(oldClipboard, textToSend), -500)

				for flashGui in flashHandles {
					try flashGui.Destroy()
				}

				Sleep(20)
				WinActivate("ahk_id " . currentWinID)
			} else {
				ShowFlashMessage("PowerScribe window not found.", mouseX, mouseY, "Red")
			}
		} catch as err {
			BlockInput("Off")
			BlockInput("MouseMoveOff")
			ShowFlashMessage("Error activating PowerScribe: " . err.Message, mouseX, mouseY, "Red")
		}
    } catch as err {
        if (pBitmap)
            Gdip_DisposeImage(pBitmap)
        if (pToken)
            Gdip_Shutdown(pToken)
        if IsObject(g_SearchGui) {
            try g_SearchGui.Destroy()
            g_SearchGui := ""
            g_SearchText := ""
        }
        ShowFlashMessage("Error capturing: " . err.Message, mouseX, mouseY, "Red")
        return
    }
}



TryFallbackOCR(screenX, screenY, ctrlW, ctrlH, cropX, cropY, cropWidth, cropHeight, annotationColor) {

    tempPath := "C:\temp"
    fallbackFile := tempPath . "\dicom_fallback_ocr.png"

    try {
        pToken := Gdip_Startup()
        pBitmap := Gdip_BitmapFromScreen(screenX . "|" . screenY . "|" . ctrlW . "|" . ctrlH)

        pCroppedBitmap := Gdip_CloneBitmapArea(pBitmap, cropX, cropY, cropWidth, cropHeight)
        pFilteredBitmap := CreateWhiteOnBlackBitmap(pCroppedBitmap, cropWidth, cropHeight, annotationColor)

        Gdip_SaveBitmapToFile(pFilteredBitmap, fallbackFile)

        Gdip_DisposeImage(pFilteredBitmap)
        Gdip_DisposeImage(pCroppedBitmap)
        Gdip_DisposeImage(pBitmap)
        Gdip_Shutdown(pToken)
    } catch {
        return false
    }

    ocrText := PerformOCRTesseract(fallbackFile)
    try FileDelete(fallbackFile)

    if RegExMatch(ocrText, "i)Series[\s:]*(\d++)[^\d\r\n]*IMA?g?e?[\s:]*(\d++)(?=\D|$)", &match) {
        seriesNum := match[1]
        imaNum := match[2]
        ShowFlashMessage("Series " . seriesNum . " IMA " . imaNum, screenX + 100, screenY + 100, "Green")
        return true
    }

    return false
}


CompareImages(file1, file2) {
    ; Compare two images pixel by pixel
    ; Returns true if identical, false if different
    
    if !FileExist(file1) || !FileExist(file2) {
        return false
    }
    
    pToken := Gdip_Startup()
    pBitmap1 := Gdip_CreateBitmapFromFile(file1)
    pBitmap2 := Gdip_CreateBitmapFromFile(file2)
    
    if !pBitmap1 || !pBitmap2 {
        if pBitmap1
            Gdip_DisposeImage(pBitmap1)
        if pBitmap2
            Gdip_DisposeImage(pBitmap2)
        Gdip_Shutdown(pToken)
        return false
    }
    
    width1 := Gdip_GetImageWidth(pBitmap1)
    height1 := Gdip_GetImageHeight(pBitmap1)
    width2 := Gdip_GetImageWidth(pBitmap2)
    height2 := Gdip_GetImageHeight(pBitmap2)
    
    ; Different dimensions = different images
    if (width1 != width2 || height1 != height2) {
        Gdip_DisposeImage(pBitmap1)
        Gdip_DisposeImage(pBitmap2)
        Gdip_Shutdown(pToken)
        return false
    }
    
    ; Compare pixels
    identical := true
    loop height1 {
        y := A_Index - 1
        loop width1 {
            x := A_Index - 1
            pixel1 := Gdip_GetPixel(pBitmap1, x, y)
            pixel2 := Gdip_GetPixel(pBitmap2, x, y)
            
            if (pixel1 != pixel2) {
                identical := false
                break 2  ; Break out of both loops
            }
        }
    }
    
    Gdip_DisposeImage(pBitmap1)
    Gdip_DisposeImage(pBitmap2)
    Gdip_Shutdown(pToken)
    
    return identical
}

ReformatSeriesImage(*) {
    ; Check if we're in PowerScribe
    if !WinActive("ahk_exe Nuance.PowerScribeOne.exe") {
        return
    }
    
    ; Save old clipboard
    oldClipboard := A_Clipboard
    A_Clipboard := ""
    
    ; Copy selected text
    SendInput("^c")
    Sleep(100)
    
    ; Get copied text
    copiedText := A_Clipboard
    
    ; Parse and reformat
    if (copiedText != "") {
        reformatted := ReformatPairs(copiedText)
        
        ; Put reformatted text in clipboard
        A_Clipboard := reformatted
        Sleep(50)
        
        ; Paste it back
        SendInput("^v")
        
        ; Restore clipboard after a delay (one-shot)
        SetTimer(RestoreClipboard.Bind(oldClipboard, reformatted), -500)

    } else {
        ; Nothing was selected, restore clipboard
        A_Clipboard := oldClipboard
    }
}

ReformatPairs(text) {
    ; -----------------------------
    ; Helpers
    ; -----------------------------
    IsValidCallout(inner) {
        ; Valid formats:
        ; (3:25)
        ; (3:88, 92)
        ; (3:88, 92; 602:91)
        return RegExMatch(inner, "^\d+:[\d,;\s:]+$")
    }

    SummarizeCalloutGroup(calloutsArr) {
        ; Build temp text from this group only, reuse old extraction logic
        tempText := ""
        for _, c in calloutsArr
            tempText .= c . " "

        pairs := Map()

        ; Pattern 1: Simple pairs like (12:447)
        pos2 := 1
        while (pos2 := RegExMatch(tempText, "\((\d+):(\d+)\)", &m2, pos2)) {
            series := m2[1], image := m2[2]
            if !pairs.Has(series)
                pairs[series] := []
            pairs[series].Push(image)
            pos2 += m2.Len
        }

        ; Pattern 2: Summarized format like (3:88, 92; 602:91)
        pos2 := 1
        while (pos2 := RegExMatch(tempText, "\(([^)]+)\)", &m2, pos2)) {
            inner := m2[1]
            seriesGroups := StrSplit(inner, ";")
            for sg in seriesGroups {
                sg := Trim(sg)
                if RegExMatch(sg, "^(\d+):(.+)$", &sm) {
                    series := sm[1]
                    imagesStr := sm[2]
                    imageList := StrSplit(imagesStr, ",")
                    for imgStr in imageList {
                        img := Trim(imgStr)
                        if RegExMatch(img, "^\d+$") {
                            if !pairs.Has(series)
                                pairs[series] := []
                            pairs[series].Push(img)
                        }
                    }
                }
            }
            pos2 += m2.Len
        }

        ; Sort series
        seriesArray := []
        for s, _ in pairs
            seriesArray.Push(s)
        seriesArray := SortNumericArray(seriesArray)

        result := ""
        for _, s in seriesArray {
            if (result != "")
                result .= "; "
            result .= s . ":"
            imgs := SortNumericArray(RemoveDuplicates(pairs[s]))
            for i, img in imgs {
                if (i > 1)
                    result .= ", "
                result .= img
            }
        }

        return (result != "") ? "(" . result . ")" : ""
    }


    ; -----------------------------
    ; 1) Tokenize text into pieces:
    ;    plain text vs valid callout blocks
    ; -----------------------------
    pieces := []
    pos := 1
    while (posMatch := RegExMatch(text, "\(([^)]+)\)", &m, pos)) {
        start := m.Pos
        full  := m[0]
        inner := m[1]

        ; text before this paren
        if (start > pos) {
            pieces.Push({ type: "text", value: SubStr(text, pos, start - pos) })
        }

        if IsValidCallout(inner) {
            pieces.Push({ type: "callout", value: full })
        } else {
            ; not a callout, keep as text
            pieces.Push({ type: "text", value: full })
        }

        pos := start + StrLen(full)
    }
    if (pos <= StrLen(text))
        pieces.Push({ type: "text", value: SubStr(text, pos) })


    ; -----------------------------
    ; 2) Condense only adjacent callouts
    ;    separated by zero or more whitespace
    ; -----------------------------
    out := ""
    i := 1
    while (i <= pieces.Length) {
        p := pieces[i]

        if (p.type = "callout") {
            group := [p.value]
            j := i + 1

            ; absorb sequences of adjacent callouts, with optional whitespace-only text between them
            while (j <= pieces.Length) {
                if (pieces[j].type = "callout") {
                    ; directly adjacent callout (zero spaces)
                    group.Push(pieces[j].value)
                    j++
                } else if (pieces[j].type = "text"
                    && RegExMatch(pieces[j].value, "^\s*$")
                    && j + 1 <= pieces.Length
                    && pieces[j + 1].type = "callout") {
                    ; whitespace-only gap then callout
                    group.Push(pieces[j + 1].value)
                    j += 2
                } else {
                    break
                }
            }

            out .= SummarizeCalloutGroup(group)
            i := j
            continue
        }

        ; plain text passes through
        out .= p.value
        i++
    }


    ; -----------------------------
    ; 3) Punctuation + spacing fixes
    ; -----------------------------

    ; (2) move period before a valid callout to after it
    ; e.g. "word. (3:69)" -> "word(3:69)."
    out := RegExReplace(out, "\.\s*(\(\d+:[\d,;\s:]+\))", "$1.")

    ; (3a) at most one space before a callout when preceded by non-space
    ; e.g. "word(3:69)" or "word     (3:69)" -> "word (3:69)"
    out := RegExReplace(out, "(\S)\s*\((\d+:[\d,;\s:]+)\)", "$1 ($2)")

    ; (3b) ensure a space after a callout if followed by a word character
    ; e.g. "(3:69)word" -> "(3:69) word"
    out := RegExReplace(out, "\)(?=\w)", ") ")

    ; cleanup multiple spaces
    while InStr(out, "  ")
        out := StrReplace(out, "  ", " ")

    return out
}


SortNumericArray(arr) {
    ; Convert strings to numbers for proper sorting
    n := arr.Length
    loop n - 1 {
        i := A_Index
        loop n - i {
            j := A_Index
            if (Integer(arr[j]) > Integer(arr[j + 1])) {
                temp := arr[j]
                arr[j] := arr[j + 1]
                arr[j + 1] := temp
            }
        }
    }
    return arr
}

RemoveDuplicates(arr) {
    ; Remove duplicate values from array
    seen := Map()
    result := []
    
    for index, value in arr {
        if !seen.Has(value) {
            seen[value] := true
            result.Push(value)
        }
    }
    
    return result
}


ShowFlashMessage(message, centerX, centerY, color := "Red") {
    ; Get monitor info for mouse position
    mouseMonitor := GetMonitorAtPoint(centerX, centerY)
    
    ; Get monitor info for PowerScribe (if exists)
    psMonitor := 0
    try {
        if WinExist("ahk_exe Nuance.PowerScribeOne.exe") {
            WinGetPos(&psX, &psY, , , "ahk_exe Nuance.PowerScribeOne.exe")
            psMonitor := GetMonitorAtPoint(psX + 100, psY + 100)
        }
    }
    
    ; FIX #4: Collect GUI handles to return them
    flashHandles := []
    
    ; Show message on mouse monitor
    flashGui1 := ShowFlashOnMonitor(message, mouseMonitor, color)
    if flashGui1
        flashHandles.Push(flashGui1)
    
    ; Show message on PowerScribe monitor if different
    if (psMonitor > 0 && psMonitor != mouseMonitor) {
        flashGui2 := ShowFlashOnMonitor(message, psMonitor, color)
        if flashGui2
            flashHandles.Push(flashGui2)
    }
    
    return flashHandles
}

ShowFlashOnMonitor(message, monitorNum, color := "Red") {
    ; Get monitor dimensions
    MonitorGet(monitorNum, &mLeft, &mTop, &mRight, &mBottom)
    
    ; Calculate center
    centerX := mLeft + (mRight - mLeft) // 2
    centerY := mTop + (mBottom - mTop) // 2
    
    ; Create flash GUI
    flashGui := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20")
    flashGui.BackColor := color
    flashGui.SetFont("s16 bold cWhite")
    
    txtCtrl := flashGui.Add("Text", "w600 Center BackgroundTrans", message)
    
    ; Show at center of monitor
    flashGui.Show("NA X" . (centerX - 300) . " Y" . (centerY - 50) . " W600 H100")
    
    ; Set transparency
    WinSetTransparent(230, flashGui)
    
    ; FIX #4: Only auto-hide for red (error) messages, not green (success) messages
    if (color = "Red") {
        SetTimer(() => flashGui.Destroy(), -2000)
    }
    
    return flashGui
}

TryOcrOnBitmapRegion(pBitmap, width, height, cropWidth, cropHeight, annotationColor, &seriesNum, &imaNum) {
    ; Ensure temp folder exists
    tempPath := "C:\temp"
    if !DirExist(tempPath)
        DirCreate(tempPath)
    captureFile := tempPath . "\dicom_capture_ocr.png"

    ; Clamp crop to valid size
    cropWidth  := Min(cropWidth,  width)
    cropHeight := Min(cropHeight, height)
    if (cropWidth < 5 || cropHeight < 5)
        return false

    ; Bottom portion, left aligned
    cropX := 0
    cropY := height - cropHeight

    ; Crop and filter
    pCroppedBitmap := Gdip_CloneBitmapArea(pBitmap, cropX, cropY, cropWidth, cropHeight)
    pFilteredBitmap := CreateWhiteOnBlackBitmap(pCroppedBitmap, cropWidth, cropHeight, annotationColor)
    Gdip_SaveBitmapToFile(pFilteredBitmap, captureFile)

    Gdip_DisposeImage(pFilteredBitmap)
    Gdip_DisposeImage(pCroppedBitmap)

    ; OCR
    ocrText := PerformOCRTesseract(captureFile)
    try FileDelete(captureFile)

    if RegExMatch(ocrText, "i)Series[\s:]*(\d++)[^\d\r\n]*IMA?g?e?[\s:]*(\d++)(?=\D|$)", &match) {
        seriesNum := match[1]
        imaNum    := match[2]
        return true
    }
    return false
}


CreateWhiteOnBlackBitmap(pSourceBitmap, width, height, targetColor) {
    ; Create new bitmap (full size for quality)
    pNewBitmap := Gdip_CreateBitmap(width, height)
    pGraphics := Gdip_GraphicsFromImage(pNewBitmap)
    
    ; Fill with WHITE background
    Gdip_GraphicsClear(pGraphics, 0xFFFFFFFF)
    
    ; Extract target RGB
    targetR := (targetColor >> 16) & 0xFF
    targetG := (targetColor >> 8) & 0xFF
    targetB := targetColor & 0xFF
    
    tolerance := 15
    
    ; Copy matching pixels as BLACK
    loop height {
        y := A_Index - 1
        loop width {
            x := A_Index - 1
            pixelColor := Gdip_GetPixel(pSourceBitmap, x, y) & 0xFFFFFF
            
            pixelR := (pixelColor >> 16) & 0xFF
            pixelG := (pixelColor >> 8) & 0xFF
            pixelB := pixelColor & 0xFF
            
            if (Abs(pixelR - targetR) <= tolerance 
                && Abs(pixelG - targetG) <= tolerance 
                && Abs(pixelB - targetB) <= tolerance) {
                ; Set as BLACK
                Gdip_SetPixel(pNewBitmap, x, y, 0xFF000000)
            }
        }
    }
    
    Gdip_DeleteGraphics(pGraphics)
    return pNewBitmap
}

PerformOCRPowerShell(imagePath) {
    ; Use PowerShell with Windows.Media.Ocr
    psScript := "
    (
    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    [void][Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime]
    [void][Windows.Media.Ocr.OcrEngine,Windows.Foundation,ContentType=WindowsRuntime]
    [void][Windows.Foundation.IAsyncOperation``1,Windows.Foundation,ContentType=WindowsRuntime]
    [void][Windows.Graphics.Imaging.BitmapDecoder,Windows.Graphics,ContentType=WindowsRuntime]
    
    function Await(`$WinRtTask) {
        `$asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { `$_.Name -eq 'AsTask' -and `$_.GetParameters().Count -eq 1 -and `$_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation``1' })[0]
        `$netTask = `$asTask.MakeGenericMethod(`$WinRtTask.GetType().GenericTypeArguments)
        `$netTask.Invoke(`$null, @(`$WinRtTask))
        `$netTask.Result
    }
    
    `$file = Await([Windows.Storage.StorageFile]::GetFileFromPathAsync('" . imagePath . "'))
    `$stream = Await(`$file.OpenAsync([Windows.Storage.FileAccessMode]::Read))
    `$decoder = Await([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync(`$stream))
    `$bitmap = Await(`$decoder.GetSoftwareBitmapAsync())
    `$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
    `$result = Await(`$engine.RecognizeAsync(`$bitmap))
    `$result.Text
    )"
    
    try {
        ; Run PowerShell and capture output
        shell := ComObject("WScript.Shell")
        exec := shell.Exec("powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " . Chr(34) . psScript . Chr(34))
        result := exec.StdOut.ReadAll()
        return Trim(result)
    } catch {
        return ""
    }
}

PerformOCR(imagePath) {
    ; Use Windows OCR (requires Windows 10+)
    try {
        ; Create OCR engine
        ocrEngine := ComObject("Windows.Media.Ocr.OcrEngine")
        
        ; Load image
        imageFile := ComObject("Windows.Storage.StorageFile")
        asyncOp := imageFile.GetFileFromPathAsync(imagePath)
        
        ; Wait for async operation
        while !asyncOp.Status
            Sleep(10)
        
        file := asyncOp.GetResults()
        
        ; Create stream
        asyncStream := file.OpenAsync(1)
        while !asyncStream.Status
            Sleep(10)
        stream := asyncStream.GetResults()
        
        ; Decode image
        decoder := ComObject("Windows.Graphics.Imaging.BitmapDecoder")
        asyncDecoder := decoder.CreateAsync(stream)
        while !asyncDecoder.Status
            Sleep(10)
        bitmapDecoder := asyncDecoder.GetResults()
        
        ; Get software bitmap
        asyncBitmap := bitmapDecoder.GetSoftwareBitmapAsync()
        while !asyncBitmap.Status
            Sleep(10)
        softwareBitmap := asyncBitmap.GetResults()
        
        ; Perform OCR
        asyncResult := ocrEngine.RecognizeAsync(softwareBitmap)
        while !asyncResult.Status
            Sleep(10)
        result := asyncResult.GetResults()
        
        return result.Text
    } catch {
        return ""
    }
}

PerformOCRTesseract(imagePath) {
    ; Get script directory
    scriptDir := A_ScriptDir
    
    ; Check if running as module (included from RAHK_Main)
    if (A_LineFile != A_ScriptFullPath) {
        ; Running as module - Tesseract is in Modules directory
        tesseractPath := scriptDir . "\Modules\Tesseract-OCR\tesseract.exe"
    } else {
        ; Running standalone - Tesseract is in same directory as script
        tesseractPath := scriptDir . "\Tesseract-OCR\tesseract.exe"
    }
    
    ; Check if Tesseract exists
    if !FileExist(tesseractPath) {
        MsgBox("Tesseract not found at: " . tesseractPath . "`n`nPlease place tesseract.exe in a 'Tesseract-OCR' subfolder.", "Error", 16)
        return ""
    }
    
    outputFile := "C:\temp\ocr_output"
    
    ; Delete old output if exists
    if FileExist(outputFile . ".txt")
        FileDelete(outputFile . ".txt")
    
    ; Optimized for speed:
    ; --psm 6 = uniform block of text
    ; --oem 0 = legacy engine (often faster than LSTM)
    ; -c tessedit_char_whitelist limits to expected characters only
    RunWait('"' . tesseractPath . '" "' . imagePath . '" "' . outputFile . '" --psm 6 --oem 1 -c tessedit_char_whitelist="0123456789SeriesIMAImage: "', , "Hide")
    
    ; Read result
    if FileExist(outputFile . ".txt") {
        result := FileRead(outputFile . ".txt")
        FileDelete(outputFile . ".txt")
        return result
    }
    
    return ""
}




GetMonitorAtPoint(x, y) {
    ; Get monitor number that contains the point
    monCount := MonitorGetCount()
    loop monCount {
        MonitorGet(A_Index, &mLeft, &mTop, &mRight, &mBottom)
        if (x >= mLeft && x < mRight && y >= mTop && y < mBottom) {
            return A_Index
        }
    }
    return 1
}




; GDI+ Functions
Gdip_Startup() {
    if !DllCall("GetModuleHandle", "str", "gdiplus", "Ptr")
        DllCall("LoadLibrary", "str", "gdiplus")
    si := Buffer(A_PtrSize = 8 ? 24 : 16, 0)
    NumPut("UInt", 1, si)
    DllCall("gdiplus\GdiplusStartup", "PtrP", &pToken := 0, "Ptr", si, "Ptr", 0)
    return pToken
}

Gdip_Shutdown(pToken) {
    DllCall("gdiplus\GdiplusShutdown", "Ptr", pToken)
    if hModule := DllCall("GetModuleHandle", "str", "gdiplus", "Ptr")
        DllCall("FreeLibrary", "Ptr", hModule)
}

Gdip_CreateBitmapFromFile(sFile) {
    DllCall("gdiplus\GdipCreateBitmapFromFile", "WStr", sFile, "PtrP", &pBitmap := 0)
    return pBitmap
}

Gdip_CreateBitmap(w, h) {
    DllCall("gdiplus\GdipCreateBitmapFromScan0", "Int", w, "Int", h, "Int", 0, "Int", 0x26200A, "Ptr", 0, "PtrP", &pBitmap := 0)
    return pBitmap
}

Gdip_GetImageWidth(pBitmap) {
    DllCall("gdiplus\GdipGetImageWidth", "Ptr", pBitmap, "UIntP", &width := 0)
    return width
}

Gdip_GetImageHeight(pBitmap) {
    DllCall("gdiplus\GdipGetImageHeight", "Ptr", pBitmap, "UIntP", &height := 0)
    return height
}

Gdip_GetPixel(pBitmap, x, y) {
    DllCall("gdiplus\GdipBitmapGetPixel", "Ptr", pBitmap, "Int", x, "Int", y, "UIntP", &color := 0)
    return color
}

Gdip_SetPixel(pBitmap, x, y, color) {
    return DllCall("gdiplus\GdipBitmapSetPixel", "Ptr", pBitmap, "Int", x, "Int", y, "UInt", color)
}

Gdip_GraphicsFromImage(pBitmap) {
    DllCall("gdiplus\GdipGetImageGraphicsContext", "Ptr", pBitmap, "PtrP", &pGraphics := 0)
    return pGraphics
}

Gdip_GraphicsClear(pGraphics, color) {
    return DllCall("gdiplus\GdipGraphicsClear", "Ptr", pGraphics, "UInt", color)
}

Gdip_DeleteGraphics(pGraphics) {
    return DllCall("gdiplus\GdipDeleteGraphics", "Ptr", pGraphics)
}

Gdip_BitmapFromScreen(Screen := 0) {
    if (Screen = 0) {
        sX := SysGet(76)
        sY := SysGet(77)
        sW := SysGet(78)
        sH := SysGet(79)
    } else {
        Screen := StrSplit(Screen, "|")
        sX := Screen[1]
        sY := Screen[2]
        sW := Screen[3]
        sH := Screen[4]

        ; --- NEW: adjust for per-monitor DPI scaling ---
        scale := GetDpiScaleAtPoint(sX + sW // 2, sY + sH // 2)
        if (scale != 1.0) {
            sX := Round(sX * scale)
            sY := Round(sY * scale)
            sW := Round(sW * scale)
            sH := Round(sH * scale)
        }
	}
    
    hDC := DllCall("GetDC", "Ptr", 0, "Ptr")
    hBM := CreateDIBSection(sW, sH)
    hDC2 := CreateCompatibleDC()
    obm := SelectObject(hDC2, hBM)
    DllCall("gdi32\BitBlt", "Ptr", hDC2, "Int", 0, "Int", 0, "Int", sW, "Int", sH, "Ptr", hDC, "Int", sX, "Int", sY, "UInt", 0x00CC0020)
    DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", "Ptr", hBM, "Ptr", 0, "PtrP", &pBitmap := 0)
    SelectObject(hDC2, obm)
    DeleteObject(hBM)
    DeleteDC(hDC2)
    DllCall("ReleaseDC", "Ptr", 0, "Ptr", hDC)
    return pBitmap
}

CreateDIBSection(w, h, bpp := 32) {
    hDC := DllCall("GetDC", "Ptr", 0, "Ptr")
    bi := Buffer(40, 0)
    NumPut("UInt", 40, "UInt", w, "UInt", h, "UShort", 1, "UShort", bpp, "UInt", 0, bi)
    hBM := DllCall("CreateDIBSection", "Ptr", hDC, "Ptr", bi, "UInt", 0, "PtrP", 0, "Ptr", 0, "UInt", 0, "Ptr")
    DllCall("ReleaseDC", "Ptr", 0, "Ptr", hDC)
    return hBM
}

CreateCompatibleDC(hDC := 0) {
    return DllCall("CreateCompatibleDC", "Ptr", hDC, "Ptr")
}

SelectObject(hDC, hGDIObj) {
    return DllCall("SelectObject", "Ptr", hDC, "Ptr", hGDIObj, "Ptr")
}

DeleteObject(hGDIObj) {
    return DllCall("DeleteObject", "Ptr", hGDIObj)
}

DeleteDC(hDC) {
    return DllCall("DeleteDC", "Ptr", hDC)
}

Gdip_SaveBitmapToFile(pBitmap, sOutput) {
    SplitPath(sOutput, , , &Extension)
    if (Extension = "")
        Extension := "png"
    Extension := "." Extension
    
    DllCall("gdiplus\GdipGetImageEncodersSize", "UIntP", &nCount := 0, "UIntP", &nSize := 0)
    ci := Buffer(nSize)
    DllCall("gdiplus\GdipGetImageEncoders", "UInt", nCount, "UInt", nSize, "Ptr", ci)
    
    loop nCount {
        sString := StrGet(NumGet(ci, (idx := (48+7*A_PtrSize)*(A_Index-1))+32+3*A_PtrSize, "Ptr"), "UTF-16")
        if InStr(sString, Extension) {
            pCodec := ci.Ptr + idx
            break
        }
    }
    
    if !pCodec
        return -1
    
    DllCall("gdiplus\GdipSaveImageToFile", "Ptr", pBitmap, "WStr", sOutput, "Ptr", pCodec, "UInt", 0)
    return 0
}

Gdip_DisposeImage(pBitmap) {
    return DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap)
}

GetDpiScaleAtPoint(x, y) {
    scale := 1.0
    try {
        pt := (x & 0xFFFFFFFF) | (y << 32)
        ; MONITOR_DEFAULTTONEAREST = 2
        hMon := DllCall("MonitorFromPoint", "Int64", pt, "UInt", 2, "Ptr")
        if (hMon) {
            ; MDT_EFFECTIVE_DPI = 0
            if (DllCall("Shcore\GetDpiForMonitor"
                , "Ptr", hMon
                , "Int", 0
                , "UIntP", &dpiX := 0
                , "UIntP", &dpiY := 0) = 0) {
                scale := dpiX / 96.0
            }
        }
    }
    return scale
}

RestoreClipboard(savedClip, expectedText) {
    ; Only restore if clipboard still equals what *we* set.
    if (A_Clipboard = expectedText) {
        A_Clipboard := savedClip
    }
}
