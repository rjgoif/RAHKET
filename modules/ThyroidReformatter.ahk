; ============================================================================
; MODULE: THYROID REFORMATTER
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
    A_TrayMenu.Add("Open Thyroid Reformatter", (*) => Show_ThyroidReformatter())
    A_TrayMenu.Add()
    A_TrayMenu.Add("Exit", (*) => ExitApp())
    
    ; Auto-show the GUI when running standalone
    Show_ThyroidReformatter()
}

Show_ThyroidReformatter() {
    TR := Gui(, "Thyroid Reformatter")
    TR.OnEvent("Close", (*) => TR.Destroy())
    
    TR.Add("Text", "x10 y10", "Paste text from Clinical Guidance tool for reformatting")
    TR_InputBox := TR.Add("Edit", "x10 y30 w300 h300")
    TR_InputBox.OnEvent("Change", TR_ProcessText)
    
    TR.Add("Text", "x320 y10", "Output Text:")
    TR_OutputBox := TR.Add("Edit", "x320 y30 w300 h300 ReadOnly")
    
    TR.Show("w630 h350")
    
    TR_ProcessText(*) {
        input := TR_InputBox.Value
        output := ProcessThyroidText(input)
        TR_OutputBox.Value := output
    }
}

ProcessThyroidText(inputText) {
    s := inputText
    
    ; Apply all find and replace operations
    s := StrReplace(s, "FINDINGS: ", "")
    s := StrReplace(s, " ---", "")
    s := StrReplace(s, "ACR TI-RADS", "TI-RADS")
    s := StrReplace(s, "( ", "(")
    s := StrReplace(s, "Prior FNA/Biopsy", "Prior biopsy")
    s := StrReplace(s, " NODULE", "NODULE")
    s := StrReplace(s, "pts", "points")
    s := StrReplace(s, " pt", " point")
    s := StrReplace(s, "wider than tall", "`"wider than tall`"")
    s := StrReplace(s, "taller than wide", "`"taller than wide`"")
    
    ; Add newline at the end
    s := s . "`n"
    
    return s
}