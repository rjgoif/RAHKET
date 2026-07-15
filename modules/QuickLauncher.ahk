; ============================================================================
; QUICK LAUNCHER MODULE
; Win + ` opens a text search bar at the mouse position.
; Type a module name, Up/Down to navigate matches, Enter to launch, Esc to cancel.
; ============================================================================

; --- Manifest: display name -> launch function reference ---
; NOTE: keep this in sync with the tray menu list in RAHKET_Main.ahk.
; There is no automatic way to derive this from the tray menu, so if you add
; a module there, add it here too (and vice versa).
QuickLauncher_Modules := [
    { name: "Lung Nodule Hunter",  fn: Show_NoduleHunter },
    { name: "Lines and Tubes",     fn: Show_LinesAndTubes },
    { name: "Thyroid Nodules",     fn: Show_ThyroidNodules },
    { name: "Vessel Measurements", fn: Show_VesselMeasurements },
    { name: "Key Phone Numbers",   fn: ShowPhoneNumbers }
]

QuickLauncher_Gui      := ""
QuickLauncher_Edit     := ""
QuickLauncher_List     := ""
QuickLauncher_Filtered := []

; ============================================================================
; SHOW / OPEN
; ============================================================================

Show_QuickLauncher() {
    global QuickLauncher_Gui, QuickLauncher_Edit, QuickLauncher_List
    global QuickLauncher_Modules, QuickLauncher_Filtered

    if IsObject(QuickLauncher_Gui) {
        QuickLauncher_Close()
    }

    QuickLauncher_Gui := Gui("+AlwaysOnTop -Caption +ToolWindow +Border", "RAHKET Quick Launch")
    QuickLauncher_Gui.BackColor := "1e1e1e"
    QuickLauncher_Gui.MarginX := 8
    QuickLauncher_Gui.MarginY := 8

    QuickLauncher_Gui.SetFont("s11 cWhite", "Segoe UI")
    QuickLauncher_Edit := QuickLauncher_Gui.AddEdit("w320 -E0x200 Background1e1e1e")

    QuickLauncher_Gui.SetFont("s10 cWhite", "Segoe UI")
    QuickLauncher_List := QuickLauncher_Gui.AddListBox("w320 r8 y+4 Background1e1e1e")

    QuickLauncher_Edit.OnEvent("Change", QuickLauncher_OnType)
    QuickLauncher_Gui.OnEvent("Escape", (*) => QuickLauncher_Close())
    QuickLauncher_List.OnEvent("DoubleClick", (*) => QuickLauncher_Launch())
    QuickLauncher_List.OnEvent("Change", QuickLauncher_OnListClick)

    MouseGetPos(&mx, &my)
    QuickLauncher_Gui.Show("x" (mx - 160) " y" (my - 10) " NoActivate AutoSize")
    WinActivate("ahk_id " QuickLauncher_Gui.Hwnd)
    QuickLauncher_Edit.Focus()

    ; populate AFTER the window is shown — populating before Show() can leave
    ; the ListBox painted with only its first row until the next redraw
    QuickLauncher_Filtered := QuickLauncher_Modules.Clone()
    QuickLauncher_RefreshList()

    ; poll for focus loss and close the launcher if the user clicks away
    SetTimer(QuickLauncher_CheckFocus, 100)
}

QuickLauncher_CheckFocus() {
    global QuickLauncher_Gui
    if !IsObject(QuickLauncher_Gui) {
        SetTimer(QuickLauncher_CheckFocus, 0)
        return
    }
    if !WinActive("ahk_id " QuickLauncher_Gui.Hwnd)
        QuickLauncher_Close()
}

; ============================================================================
; CONTEXT-SENSITIVE HOTKEYS (active only while the launcher window is focused)
; ============================================================================

QuickLauncher_IsActive(*) {
    global QuickLauncher_Gui
    if !IsObject(QuickLauncher_Gui)
        return false
    return WinActive("ahk_id " QuickLauncher_Gui.Hwnd) ? true : false
}

#HotIf QuickLauncher_IsActive()
Up::QuickLauncher_MoveSelection(-1)
Down::QuickLauncher_MoveSelection(1)
Enter::QuickLauncher_Launch()
Escape::QuickLauncher_Close()
#HotIf

; ============================================================================
; FILTERING / LIST MANAGEMENT
; ============================================================================

QuickLauncher_OnType(ctrl, *) {
    global QuickLauncher_Modules, QuickLauncher_Filtered

    searchTerm := Trim(ctrl.Text)
    QuickLauncher_Filtered := []

    if (searchTerm = "") {
        QuickLauncher_Filtered := QuickLauncher_Modules.Clone()
    } else {
        ; starts-with matches first, then contains matches
        startsWith      := []
        containsMatches := []
        for mod in QuickLauncher_Modules {
            if (SubStr(mod.name, 1, StrLen(searchTerm)) = searchTerm)
                startsWith.Push(mod)
            else if InStr(mod.name, searchTerm)
                containsMatches.Push(mod)
        }
        for mod in startsWith
            QuickLauncher_Filtered.Push(mod)
        for mod in containsMatches
            QuickLauncher_Filtered.Push(mod)
    }

    QuickLauncher_RefreshList()
}

QuickLauncher_RefreshList() {
    global QuickLauncher_List, QuickLauncher_Filtered

    names := []
    for mod in QuickLauncher_Filtered
        names.Push(mod.name)

    QuickLauncher_List.Delete()
    if (names.Length > 0) {
        QuickLauncher_List.Add(names)
        QuickLauncher_List.Choose(1)
    }
}

QuickLauncher_OnListClick(ctrl, *) {
    global QuickLauncher_Edit
    ; native click already updates ctrl.Value; just return focus to the Edit
    ; so the Enter hotkey behaves identically to the arrow-key navigation path
    QuickLauncher_Edit.Focus()
}

QuickLauncher_MoveSelection(direction) {
    global QuickLauncher_List, QuickLauncher_Filtered

    count := QuickLauncher_Filtered.Length
    if (count = 0)
        return

    current := QuickLauncher_List.Value
    if (current = 0)
        current := 1

    newIndex := current + direction
    if (newIndex < 1)
        newIndex := 1
    if (newIndex > count)
        newIndex := count

    QuickLauncher_List.Choose(newIndex)
}

; ============================================================================
; LAUNCH / CLOSE
; ============================================================================

QuickLauncher_Launch(*) {
    global QuickLauncher_List, QuickLauncher_Filtered

    idx := QuickLauncher_List.Value
    if (idx = 0 && QuickLauncher_Filtered.Length = 1)
        idx := 1

    if (idx = 0 || idx > QuickLauncher_Filtered.Length) {
        QuickLauncher_Close()
        return
    }

    selectedFn := QuickLauncher_Filtered[idx].fn
    QuickLauncher_Close()
    selectedFn.Call()
}

QuickLauncher_Close(*) {
    global QuickLauncher_Gui
    SetTimer(QuickLauncher_CheckFocus, 0)
    if IsObject(QuickLauncher_Gui) {
        try QuickLauncher_Gui.Destroy()
        QuickLauncher_Gui := ""
    }
}
