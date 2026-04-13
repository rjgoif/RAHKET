; ============================================================================
; MODULE: VESSEL MEASUREMENTS
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
    A_TrayMenu.Add("Open Vessel Measurements", (*) => Show_VesselMeasurements())
    A_TrayMenu.Add()
    A_TrayMenu.Add("Exit", (*) => ExitApp())
    
    ; Auto-show the GUI when running standalone
    Show_VesselMeasurements()
}

Show_VesselMeasurements() {
    VM := Gui(, "Vessel Measurements")
    VM.OnEvent("Close", (*) => VM.Destroy())

    ; --- Header row (bold) ---
    VM.SetFont("Bold")
    VM.Add("Text", "x10  y10 w130 Center", "Vessel")
    VM.Add("Text", "x150 y10 w110 Center", "Normal (cm)")
    VM.Add("Text", "x270 y10 w140 Center", "Dilated/Ectatic (cm)")
    VM.Add("Text", "x420 y10 w120 Center", "Aneurysmal (cm)")
    VM.SetFont("Norm") ; back to normal for data rows

    ; --- Table data ---
    data := [
        ["Sinus of Valsalva", "<3.9", "4.0–4.4", "≥4.5"],
        ["Ascending Aorta",   "<3.9", "4.0–4.4", "≥4.5"],
        ["Descending Aorta",  "<2.4", "2.5–3.9", "≥4.0"],
        ["Abdominal Aorta",   "<2.0", "2.0–2.9", "≥3.0"],
        ["Iliac Arteries",    "<1.0", "1.0–1.4", "≥1.5"]
    ]

    y := 40
    rowH := 26
    For row in data {
        VM.Add("Text", "x10  y" y " w130", row[1])
        VM.Add("Text", "x150 y" y " w110 Center", row[2])
        VM.Add("Text", "x270 y" y " w140 Center", row[3])
        VM.Add("Text", "x420 y" y " w120 Center", row[4])
        y += rowH
    }

    ; --- Close button ---
    closeBtn := VM.Add("Button", "x260 y" (y + 10) " w100", "Close")
    closeBtn.OnEvent("Click", (*) => VM.Destroy())

    ; Adjust window size based on content
    totalH := y + 60
    VM.Show("w560 h" totalH)
}
