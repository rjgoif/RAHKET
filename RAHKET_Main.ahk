; ============================================================================
; RADIOLOGY AUTOHOTKEY ENHANCEMENT TOOLS
; Written by Reece J. Goiffon, MD, PhD
; ============================================================================

#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent




; ============================================================================
; ZIP EXTRACT HELPER (uses Windows Shell)
; ============================================================================

ExtractZip(zipPath, destDir) {
    if !FileExist(zipPath)
        return false

    if !DirExist(destDir)
        DirCreate(destDir)

    shell := ComObject("Shell.Application")
    src   := shell.Namespace(zipPath)
    if !src
        return false

    dest := shell.Namespace(destDir)
    if !dest
        return false

    ; 16 = FOF_NOCONFIRMMKDIR (no "confirm new folder" dialogs)
    dest.CopyHere(src.Items, 16)
    return true
}



; ============================================================================
; INSTALLER / BOOTSTRAP: ensure Modules, tray icon, and Start Menu shortcut
; ============================================================================

if A_IsCompiled {

    ; ============================================================
    ;  NETWORK-LAUNCHED PROTECTION / SELF-COPY INSTALLER
    ; ============================================================

    exePath := A_ScriptFullPath
    exeDir  := A_ScriptDir

    ; Detect if app is launched from \\mghpacs\jpg or Z: drive
    isNetworkLaunch :=
        InStr(exeDir, "\\mghpacs\jpg", false)      ; UNC network
        || SubStr(exeDir, 1, 2) = "Z:"             ; Z: drive

    if (isNetworkLaunch) {

        ; --------------------------------------------------------
        ; Build the EXACT target install directory:
        ;   C:\Users\<user>\utils\RAHKET\
        ; --------------------------------------------------------

        userRoot := SubStr(A_AppData, 1, InStr(A_AppData, "\AppData") - 1)
        ; Example:
        ; A_AppData = C:\Users\JohnDoe\AppData\Roaming
        ; userRoot = C:\Users\JohnDoe

        localInstallDir := userRoot "\utils\RAHKET"

        if !DirExist(localInstallDir)
            DirCreate(localInstallDir)

        ; --- CLEAN OUT OLD MODULES + ZIP IN TARGET DIR ---
        localModulesDir := localInstallDir "\Modules"
        localZipPath    := localInstallDir "\modules.zip"

        if DirExist(localModulesDir)
            DirDelete(localModulesDir, true)   ; true = recursive delete

        if FileExist(localZipPath)
            FileDelete(localZipPath)

        ; Destination EXE
        localExe := localInstallDir "\RAHKET.exe"

        ; Copy this EXE locally (overwrite to ensure latest)
        FileCopy(exePath, localExe, true)      ; true = force overwrite

        ; Notify user
        MsgBox(
            "RAHKET should be run from your local device.`n"
          . "A new installation was sent to:`n"
          . localInstallDir,
            "Local Installation Recommended", 64
        )

        ; Launch local instance
        Run(localExe)

        ; Kill network-launched instance
        ExitApp()
    }



    ; --- Ensure modules.zip exists beside the EXE (embedded via FileInstall) ---
    zipPath    := A_ScriptDir "\modules.zip"
    modulesDir := A_ScriptDir "\Modules"

    if !FileExist(zipPath) {
        ; Embed modules.zip into the EXE at compile time and extract on first run
        FileInstall("modules.zip", zipPath, false)   ; false = don't overwrite existing
    }

    ; --- If Modules folder is missing, extract modules.zip into the EXE folder ---
    if !DirExist(modulesDir) && FileExist(zipPath) {
        ExtractZip(zipPath, A_ScriptDir)  ; zip should contain top-level "Modules" folder
    }

    ; --- Ensure tray icon .ico exists beside the EXE (also via FileInstall) ---
    icoPath := A_ScriptDir "\RAHKET_rocket_icon.ico"
    if !FileExist(icoPath) {
        ; Embed the icon and extract it; true = overwrite to keep icon in sync with builds
        FileInstall("assets\RAHKET_rocket_icon.ico", icoPath, true)
    }

    ; --- (Re)create a Start Menu shortcut to this EXE ---
    try {
        shortcutDir  := A_Programs                 ; Start Menu\Programs
        shortcutPath := shortcutDir "\RAHKET.lnk"

        if !DirExist(shortcutDir)
            DirCreate(shortcutDir)

        exePath := A_ScriptFullPath

        if FileExist(shortcutPath)
            FileDelete(shortcutPath)

        FileCreateShortcut(
            exePath,        ; Target
            shortcutPath,   ; Shortcut file path
            A_ScriptDir,    ; Working directory
            ,               ; Args (none)
            "RAHKET",       ; Description
            exePath,        ; Icon = EXE's icon
            ,               ; ShortcutKey (none)
            1,              ; IconNumber
            1               ; Run normal
        )
    }
}

; --- Set tray icon to the .ico in the same folder (now guaranteed to exist when compiled) ---
TraySetIcon(A_ScriptDir "\RAHKET_rocket_icon.ico")

; --- Include modules ---
#Include Modules\NoduleHunter.ahk
#Include Modules\ThyroidNodules.ahk
;#Include Modules\ThyroidReformatter.ahk
#Include Modules\VesselMeasurements.ahk
#Include Modules\KeyPhoneNumbers.ahk
#Include Modules\Links.ahk
#Include Modules\ImageInserter.ahk
#Include Modules\PS_TextFormatter.ahk

; Create system tray icon with custom menu
A_TrayMenu.Delete()
A_TrayMenu.Add("Lung Nodule Hunter", MenuItem_NoduleHunter)
A_TrayMenu.Add("Thyroid Nodules", MenuItem_ThyroidNodules)
;A_TrayMenu.Add("Thyroid Reformatter", MenuItem_ThyroidReformatter)
A_TrayMenu.Add("Vessel Measurements", MenuItem_VesselMeasurements)
A_TrayMenu.Add("Key Phone Numbers", MenuItem_KeyPhoneNumbers)

; Add Links submenu from module
A_TrayMenu.Add("Links", CreateLinksMenu())

; Add Image Inserter submenu
A_TrayMenu.Add("Image Inserter", CreateImageInserterMenu())

; Add Report Edits submenu
A_TrayMenu.Add("Report Edits", CreateReportEditsMenu())

A_TrayMenu.Add()
A_TrayMenu.Add("About", MenuItem_About)
A_TrayMenu.Add("Exit", MenuItem_Exit)

; ============================================================================
; MENU ITEM HANDLERS
; ============================================================================

MenuItem_NoduleHunter(*) {
    Show_NoduleHunter()
}

MenuItem_ThyroidNodules(*) {
    Show_ThyroidNodules()
}

;MenuItem_ThyroidReformatter(*) {
;    Show_ThyroidReformatter()
;}

MenuItem_VesselMeasurements(*) {
    Show_VesselMeasurements()
}

MenuItem_KeyPhoneNumbers(*) {
    ShowPhoneNumbers()
}						  

MenuItem_About(*) {
    MsgBox("RAHKET`nRadiology AutoHotKey Enhancement Tools`nWritten by Reece J. Goiffon, MD, PhD", "About", 64)
}

MenuItem_Exit(*) {
    ExitApp()
}

; ============================================================================
; IMAGE INSERTER SUBMENU CREATION
; ============================================================================

CreateImageInserterMenu() {
    imageInserterMenu := Menu()
    imageInserterMenu.Add("(Re)enable", MenuItem_ImageInserterEnable)
    imageInserterMenu.Add("Disable", MenuItem_ImageInserterDisable)
    return imageInserterMenu
}

MenuItem_ImageInserterEnable(*) {
    ImageInserter_Enable()
}

MenuItem_ImageInserterDisable(*) {
    ImageInserter_Disable()
}

; ============================================================================
; REPORT EDITS SUBMENU CREATION
; ============================================================================

CreateReportEditsMenu() {
    reportEditsMenu := Menu()
    reportEditsMenu.Add("Enable Report Editor", MenuItem_LineJoinerEnable)
    reportEditsMenu.Add("Disable Report Editor", MenuItem_LineJoinerDisable)
    return reportEditsMenu
}

MenuItem_LineJoinerEnable(*) {
    LineJoiner_Enable()
}

MenuItem_LineJoinerDisable(*) {
    LineJoiner_Disable()
}
