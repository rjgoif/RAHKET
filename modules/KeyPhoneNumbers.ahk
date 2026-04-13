; ============================================================================
; MODULE: Key Phone Numbers
; Written by Reece J. Goiffon, MD, PhD
; AutoHotkey v2
; Can be run standalone or included in RAHK_Main.ahk
; ============================================================================

#Requires AutoHotkey v2.0

; Only set these if running standalone
if (A_LineFile = A_ScriptFullPath) {
    #SingleInstance Force
    Persistent
    
    ; Include Links module for standalone mode
    #Include Links.ahk
    
    ; Create standalone tray menu with links
    A_TrayMenu.Delete()
    A_TrayMenu.Add("Show Key Phone Numbers", (*) => ShowPhoneNumbers())
    A_TrayMenu.Add("Links", CreateLinksMenu())
    A_TrayMenu.Add()
    A_TrayMenu.Add("Exit", (*) => ExitApp())
    
    ; Show GUI on startup when standalone
    ShowPhoneNumbers()
}


; ============================================================================
; MODULE CODE
; ============================================================================

; Function to expand short extensions to full phone numbers
ExpandPhoneNumber(shortNum) {
    ; Remove any spaces and hyphens
    shortNum := StrReplace(StrReplace(shortNum, " ", ""), "-", "")
    
    ; Check if it's a 5-digit format (e.g., "3-2613" becomes "32613")
    if (StrLen(shortNum) = 5) {
        prefix := SubStr(shortNum, 1, 1)
        extension := SubStr(shortNum, 2, 4)
        
        switch prefix {
            case "3":
                return "(617) 643-" extension
            case "4":
                return "(617) 724-" extension
            case "6":
                return "(617) 726-" extension
            case "8":
                return "(857) 238-" extension
            default:
                return shortNum
        }
    }
    
    return shortNum
}

; Function to create phone number GUI
ShowPhoneNumbers() {
    static phoneGui := ""
    
    ; Close existing GUI if it exists
    if (phoneGui != "") {
        try phoneGui.Destroy()
    }
    
    ; Create new GUI
    phoneGui := Gui("+AlwaysOnTop", "Phones")
    phoneGui.SetFont("s9", "Segoe UI")
    phoneGui.MarginX := 15
    phoneGui.MarginY := 15
    
    ; Define phone number data with categories
    phoneData := [
        {category: "Abd resident", numbers: ["6-5162"]},
        {category: "Abd consult", numbers: ["4-2162"]},
        {category: "Abd fluoro", numbers: ["3-2613", "3-2605"]},
        {category: "Abd Y6", numbers: ["3-5446", "4-2481"]},
        {spacer: true},
        {category: "Chest CT", numbers: ["3-3899"]},
        {category: "Chest IP", numbers: ["4-2051"]},
        {category: "Chest OP", numbers: ["6-2197"]},
        {spacer: true},
        {category: "Breast", numbers: ["4-0228"]},
        {spacer: true},
        {category: "Cardiac CT", numbers: ["4-7132"]},
        {category: "Cardiac MR", numbers: ["3-4457"]},
        {category: "Vascular", numbers: ["4-7115"]},
        {spacer: true},
        {category: "ED", numbers: ["4-1533", "4-3458"]},
        {spacer: true},
        {category: "Neuro Consult", numbers: ["4-1931"]},
        {category: "Neuro ED", numbers: ["6-8188"]},
        {spacer: true},
        {category: "Pedi", numbers: ["4-2119", "6-6005"]},
        {spacer: true},
        {category: "MSK", numbers: ["6-5339"]},
        {spacer: true},
        {category: "NM PET", numbers: ["6-6737"]},
        {category: "NM Clinic", numbers: ["6-1404"]},
        {spacer: true},
        {category: "Telerads", numbers: ["4-4270"]}
    ]
    
    ; Add phone numbers to GUI
    currentY := 15
    for item in phoneData {
        ; Handle spacer
        if (item.HasOwnProp("spacer") && item.spacer) {
            currentY += 10
            continue
        }
        
        ; Add category label
        phoneGui.Add("Text", "x15 y" currentY " w110", item.category)
        
        ; Add numbers in columns
        numX := 130
        for index, num in item.numbers {
            fullNum := ExpandPhoneNumber(num)
            
            ; Create a text control that shows tooltip
            numCtrl := phoneGui.Add("Text", "x" numX " y" currentY " w55 Right", num)
            
            ; Use a custom method to show tooltips on hover
            numCtrl.OnEvent("Click", MakePhoneClickHandler(fullNum))
            
            ; Store the full number for later use
            numCtrl.fullNumber := fullNum
            
            ; Move to second column for next number
            if (index = 1) {
                numX := 190
            }
        }
        
        currentY += 22
    }
    
    ; Add some spacing before translation guide
    currentY += 10
    
    ; Add separator line
    phoneGui.Add("Text", "x15 y" currentY " w230 h1 0x10")
    currentY += 10
    
    ; Add translation guide
    phoneGui.SetFont("s8", "Segoe UI")
    phoneGui.Add("Text", "x15 y" currentY " w230", "3-xxxx → (617) 643-xxxx")
    currentY += 16
    phoneGui.Add("Text", "x15 y" currentY " w230", "4-xxxx → (617) 724-xxxx")
    currentY += 16
    phoneGui.Add("Text", "x15 y" currentY " w230", "6-xxxx → (617) 726-xxxx")
    currentY += 16
    phoneGui.Add("Text", "x15 y" currentY " w230", "8-xxxx → (857) 238-xxxx")
    currentY += 20
    
    ; Add separator line
    phoneGui.Add("Text", "x15 y" currentY " w230 h1 0x10")
    currentY += 10
    
    ; Reset font for link
    phoneGui.SetFont("s9", "Segoe UI")
    
    ; Add link to full phone list
    phoneGui.Add("Link", "x15 y" currentY " w230 Center", '<a href="https://radhub.massgeneral.org/reshub/phone-list/">Full Phone List</a>')
    currentY += 30
    
    ; Add close button
    closeBtn := phoneGui.Add("Button", "x80 y" currentY " w100", "Close")
    closeBtn.OnEvent("Click", (*) => phoneGui.Hide())
    
    ; Handle GUI close
    phoneGui.OnEvent("Close", (*) => phoneGui.Hide())
    phoneGui.OnEvent("Escape", (*) => phoneGui.Hide())
    
    ; Set up tooltip timer
    SetTimer(UpdateTooltip, 100)
    
    ; Show the GUI
    phoneGui.Show("w260")
}

; Function to update tooltip based on mouse position
UpdateTooltip() {
    static lastControl := 0
    
    try {
        ; Get control under mouse
        MouseGetPos(,, &winID, &ctrlID, 2)
        
        ; Only process if we're over the phone GUI
        if (!winID)
            return
            
        try {
            ctrl := GuiCtrlFromHwnd(ctrlID)
            
            ; If we moved to a different control
            if (ctrl != lastControl) {
                lastControl := ctrl
                
                ; Clear tooltip first
                ToolTip()
                
                ; Show tooltip if this control has a full number
                if (ctrl && ctrl.HasOwnProp("fullNumber")) {
                    ToolTip(ctrl.fullNumber)
                }
            }
        }
    } catch {
        ; If any error, just clear tooltip
        ToolTip()
        lastControl := 0
    }
}

; Handler for clicking phone numbers (optional - could copy to clipboard)
MakePhoneClickHandler(fullNum) {
    return (*) => (
        A_Clipboard := fullNum,
        ToolTip("Copied: " fullNum),
        SetTimer(() => ToolTip(), -2000)
    )
}

; Global hotkey to show phone numbers (can be customized)
; ^!p::ShowPhoneNumbers()  ; Ctrl+Alt+P (uncomment to enable)