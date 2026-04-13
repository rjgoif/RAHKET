; ============================================================================
; MODULE: 
; Written by Reece J. Goiffon, MD, PhD
; AutoHotkey v2
; Can be run standalone or included in RAHKET_Main.ahk
; ============================================================================

#Requires AutoHotkey v2.0

; Only set these if running standalone
if (A_LineFile = A_ScriptFullPath) {
    #SingleInstance Force
    Persistent
    
    ; Create standalone tray menu with links
    A_TrayMenu.Delete()
    A_TrayMenu.Add("Exit", (*) => ExitApp())
}


; ============================================================================
; MODULE CODE
; ============================================================================

