; ============================================================================
; RADIOLOGY AUTOHOTKEY ENHANCEMENT TOOLS
; Written by Reece J. Goiffon, MD, PhD
; ============================================================================

#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent

RAHKET_VERSION := "1.0.1.0"
; @Ahk2Exe-SetVersion 1.0.1.0
; @Ahk2Exe-SetDescription RAHKET - Radiology AutoHotKey Enhancement Tools
; @Ahk2Exe-SetProductName RAHKET
; @Ahk2Exe-SetCompanyName Reece J. Goiffon MD PhD

GITHUB_VERSION_URL  := "https://raw.githubusercontent.com/rjgoif/RAHKET/main/version"
GITHUB_RELEASES_URL := "https://github.com/rjgoif/RAHKET/releases/latest"



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
    if FileExist(zipPath)
        FileDelete(zipPath)
    return true
}



; ============================================================================
; CONFIG READER
; Returns value for a given key from rahket_config.ini, or "" if not found
; ============================================================================

ReadConfig(key) {
    configPath := A_ScriptDir "\rahket_config.ini"
    if !FileExist(configPath)
        return ""

    loop read, configPath {
        line := Trim(A_LoopReadLine)
        if (line = "" || SubStr(line, 1, 1) = ";")
            continue
        colonPos := InStr(line, "=")
        if !colonPos
            continue
        lineKey := Trim(SubStr(line, 1, colonPos - 1))
        lineVal := Trim(SubStr(line, colonPos + 1))
        if (lineKey = key)
            return lineVal
    }
    return ""
}

configPath := A_ScriptDir "\rahket_config.ini"

; ============================================================================
; INSTALLER / BOOTSTRAP
; ============================================================================

if A_IsCompiled {

    exePath := A_ScriptFullPath
    exeDir  := A_ScriptDir

    ; Generic network launch detection — no hardcoded paths in source
    isNetworkLaunch :=
        (SubStr(exeDir, 1, 2) = "\\")     ; any UNC path
        || (SubStr(exeDir, 1, 2) = "Z:")  ; Z: drive mapped network share

    if (isNetworkLaunch) {

        ; Build local install directory: C:\Users\<user>\utils\RAHKET\
        userRoot        := SubStr(A_AppData, 1, InStr(A_AppData, "\AppData") - 1)
        localInstallDir := userRoot "\utils\RAHKET"

        if !DirExist(localInstallDir)
            DirCreate(localInstallDir)

        ; Clean out old modules
        localModulesDir := localInstallDir "\Modules"
        localZipPath    := localInstallDir "\modules.zip"

        if DirExist(localModulesDir)
            DirDelete(localModulesDir, true)

        if FileExist(localZipPath)
            FileDelete(localZipPath)

        ; Copy installer EXE as RAHKET.exe
        localExe := localInstallDir "\RAHKET.exe"
        FileCopy(exePath, localExe, true)

        ; Copy rahket_config.ini if present alongside installer
        localConfig  := localInstallDir "\rahket_config.ini"
        sourceConfig := exeDir "\rahket_config.ini"
        if FileExist(sourceConfig)
            FileCopy(sourceConfig, localConfig, true)

        MsgBox(
            "RAHKET has been installed to:`n" localInstallDir
            "`n`nRAHKET will now launch from your local installation.",
            "RAHKET Installed", 64
        )

        Run(localExe)
        ExitApp()
    }


    ; --- Ensure modules.zip exists (embedded via FileInstall) ---
    zipPath    := A_ScriptDir "\modules.zip"
    modulesDir := A_ScriptDir "\Modules"

    if !FileExist(zipPath)
        FileInstall("modules.zip", zipPath, false)

    ; --- Extract Modules folder if missing ---
    if !DirExist(modulesDir) && FileExist(zipPath)
        ExtractZip(zipPath, A_ScriptDir)

    ; --- Ensure tray icon exists ---
    icoPath := A_ScriptDir "\RAHKET_rocket_icon.ico"
    if !FileExist(icoPath)
        FileInstall("assets\RAHKET_rocket_icon.ico", icoPath, true)

    ; --- Create Start Menu shortcut ---
    try {
        shortcutDir  := A_Programs
        shortcutPath := shortcutDir "\RAHKET.lnk"

        if !DirExist(shortcutDir)
            DirCreate(shortcutDir)

        if FileExist(shortcutPath)
            FileDelete(shortcutPath)

        FileCreateShortcut(
            A_ScriptFullPath,
            shortcutPath,
            A_ScriptDir,
            ,
            "RAHKET",
            A_ScriptFullPath,
            ,
            1,
            1
        )
    }
}



; ============================================================================
; TRAY ICON
; ============================================================================

if FileExist(A_ScriptDir "\RAHKET_rocket_icon.ico")
    TraySetIcon(A_ScriptDir "\RAHKET_rocket_icon.ico")
else if FileExist(A_ScriptDir "\assets\RAHKET_rocket_icon.ico")
    TraySetIcon(A_ScriptDir "\assets\RAHKET_rocket_icon.ico")



; ============================================================================
; INCLUDE MODULES
; ============================================================================

#Include Modules\NoduleHunter.ahk
#Include Modules\ThyroidNodules.ahk
;#Include Modules\ThyroidReformatter.ahk
#Include Modules\VesselMeasurements.ahk
#Include Modules\KeyPhoneNumbers.ahk
#Include Modules\Links.ahk
#Include Modules\ImageInserter.ahk
#Include Modules\PS_TextFormatter.ahk



; ============================================================================
; TRAY MENU
; ============================================================================

A_TrayMenu.Delete()
A_TrayMenu.Add("Lung Nodule Hunter",   MenuItem_NoduleHunter)
A_TrayMenu.Add("Thyroid Nodules",      MenuItem_ThyroidNodules)
;A_TrayMenu.Add("Thyroid Reformatter", MenuItem_ThyroidReformatter)
A_TrayMenu.Add("Vessel Measurements",  MenuItem_VesselMeasurements)
A_TrayMenu.Add("Key Phone Numbers",    MenuItem_KeyPhoneNumbers)
A_TrayMenu.Add("Links",                CreateLinksMenu())
A_TrayMenu.Add("Image Inserter",       CreateImageInserterMenu())
A_TrayMenu.Add("Report Edits",         CreateReportEditsMenu())
A_TrayMenu.Add()
A_TrayMenu.Add("Check for Updates",    MenuItem_CheckForUpdates)
A_TrayMenu.Add("About",                MenuItem_About)
A_TrayMenu.Add("Exit",                 MenuItem_Exit)



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
    MsgBox(
        "RAHKET v" RAHKET_VERSION "`n"
        "Radiology AutoHotKey Enhancement Tools`n"
        "Written by Reece J. Goiffon, MD, PhD",
        "About", 64
    )
}

MenuItem_Exit(*) {
    ExitApp()
}



; ============================================================================
; UPDATE CHECKER
; ============================================================================

MenuItem_CheckForUpdates(*) {
    CheckForUpdates()
}

CheckForUpdates() {
    global RAHKET_VERSION, GITHUB_VERSION_URL, GITHUB_RELEASES_URL

    networkPath := GetNetworkUpdatePath()

    if (networkPath != "") {
        networkVersion := ReadNetworkVersion(networkPath)

        if (networkVersion = "") {
            MsgBox("Could not read version information from the network location.", "Update Check Failed", 48)
            return
        }

        if (networkVersion = RAHKET_VERSION) {
            result := MsgBox(
                "RAHKET is up to date (v" RAHKET_VERSION ").`n`n"
                "Force reinstall from network anyway?",
                "No Update Available", 4 + 32
            )
            if (result = "Yes")
                DoNetworkUpdate(networkPath)
            return
        }

        result := MsgBox(
            "A new version of RAHKET is available.`n`n"
            "Current version:  v" RAHKET_VERSION "`n"
            "Available version: v" networkVersion "`n`n"
            "Install now?",
            "Update Available", 4 + 32
        )
        if (result = "Yes")
            DoNetworkUpdate(networkPath)
        return
    }

    ; --- No network path — fall back to GitHub ---
    tempVersionFile := A_Temp "\rahket_version_check.txt"
    try {
        Download(GITHUB_VERSION_URL, tempVersionFile)
    } catch {
        MsgBox("Could not reach the update server.`nCheck your internet connection and try again.", "Update Check Failed", 48)
        return
    }
    if !FileExist(tempVersionFile) {
        MsgBox("Update check failed — version file not downloaded.", "Update Check Failed", 48)
        return
    }
    githubVersion := Trim(FileRead(tempVersionFile))
    FileDelete(tempVersionFile)
    if (githubVersion = RAHKET_VERSION) {
        MsgBox("RAHKET is up to date (v" RAHKET_VERSION ").", "No Update Available", 64)
        return
    }
    result := MsgBox(
        "A new version of RAHKET is available.`n`n"
        "Current version:  v" RAHKET_VERSION "`n"
        "Available version: v" githubVersion "`n`n"
        "Open the downloads page?",
        "Update Available", 4 + 32
    )
    if (result = "Yes")
        Run(GITHUB_RELEASES_URL)
}


GetNetworkUpdatePath() {
    ; Read both paths from config
    zPath   := ReadConfig("main_update")
    uncPath := ReadConfig("alt_update")

    ; Try Z: path first
    if (zPath != "" && DirExist(zPath))
        return zPath

    ; Fall back to UNC path
    if (uncPath != "" && DirExist(uncPath))
        return uncPath

    return ""
}


ReadNetworkVersion(networkPath) {
    versionFile := networkPath "\version"
    if !FileExist(versionFile)
        return ""
    return Trim(FileRead(versionFile))
}


DoNetworkUpdate(networkPath) {
    newExe   := networkPath "\RAHKET Installer.exe"
    localExe := A_ScriptFullPath
    tempExe  := A_Temp "\RAHKET_update.exe"
    batFile  := A_Temp "\rahket_updater.bat"

    if !FileExist(newExe) {
        MsgBox("Update file not found at the network location.", "Update Failed", 48)
        return
    }

    ; Copy new EXE to temp location
    try {
        FileCopy(newExe, tempExe, true)
    } catch {
        MsgBox("Failed to copy update file. Check that you have access to the network location.", "Update Failed", 48)
        return
    }

    ; Also copy updated config if present
    newConfig   := networkPath "\rahket_config.ini"
    localConfig := A_ScriptDir "\rahket_config.ini"
    tempConfig  := A_Temp "\rahket_config_update.ini"
    hasConfig   := FileExist(newConfig)
    if hasConfig
        FileCopy(newConfig, tempConfig, true)

    ; Write batch file to wait for RAHKET to exit, replace EXE, restart
	q := Chr(34)
    batContent :=
        "@echo off`r`n"
        . "timeout /t 2 /nobreak >nul`r`n"
        . ":waitloop`r`n"
        . "tasklist | find /i " q "RAHKET.exe" q " >nul 2>&1`r`n"
        . "if not errorlevel 1 (`r`n"
        . "    timeout /t 1 /nobreak >nul`r`n"
        . "    goto waitloop`r`n"
        . ")`r`n"
        . "copy /y " q tempExe q " " q localExe q "`r`n"

    if hasConfig
        batContent .= "copy /y " q tempConfig q " " q localConfig q "`r`n"

    localModulesDir := A_ScriptDir "\Modules"
    localZipPath    := A_ScriptDir "\modules.zip"

    batContent .=
        "if exist " q localModulesDir q " rmdir /s /q " q localModulesDir q "`r`n"
        . "if exist " q localZipPath q " del /f /q " q localZipPath q "`r`n"
        . "start " q q " " q localExe q "`r`n"
        . "del " q batFile q "`r`n"
	
	if FileExist(batFile)
		FileDelete(batFile)
	FileAppend(batContent, batFile)

    MsgBox(
        "RAHKET will now close and update to v" ReadNetworkVersion(networkPath) ".`n"
        "It will restart automatically when complete.",
        "Updating RAHKET", 64
    )

    Run(batFile, , "Hide")
    ExitApp()
}



; ============================================================================
; IMAGE INSERTER SUBMENU
; ============================================================================

CreateImageInserterMenu() {
    imageInserterMenu := Menu()
    imageInserterMenu.Add("(Re)enable", MenuItem_ImageInserterEnable)
    imageInserterMenu.Add("Disable",    MenuItem_ImageInserterDisable)
    return imageInserterMenu
}

MenuItem_ImageInserterEnable(*) {
    ImageInserter_Enable()
}

MenuItem_ImageInserterDisable(*) {
    ImageInserter_Disable()
}



; ============================================================================
; REPORT EDITS SUBMENU
; ============================================================================

CreateReportEditsMenu() {
    reportEditsMenu := Menu()
    reportEditsMenu.Add("Enable Report Editor",  MenuItem_LineJoinerEnable)
    reportEditsMenu.Add("Disable Report Editor", MenuItem_LineJoinerDisable)
    return reportEditsMenu
}

MenuItem_LineJoinerEnable(*) {
    LineJoiner_Enable()
}

MenuItem_LineJoinerDisable(*) {
    LineJoiner_Disable()
}
