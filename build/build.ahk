; ============================================================================
; RAHKET Build Script
; Zips the Modules folder, then compiles RAHKET_Main.ahk into RAHKET installer.exe
; Run this from the repo root, or from the build\ subfolder.
; ============================================================================

#Requires AutoHotkey v2.0
#SingleInstance Force

; ── Resolve repo root (works whether script lives in build\ or repo root) ──
scriptDir := A_ScriptDir
repoRoot  := (FileExist(scriptDir "\..\RAHKET_Main.ahk")) ? scriptDir "\.." : scriptDir
repoRoot  := RTrim(repoRoot, "\")

mainScript  := repoRoot "\RAHKET_Main.ahk"
modulesDir  := repoRoot "\modules"
zipDest     := repoRoot "\modules.zip"
iconFile    := repoRoot "\assets\RAHKET_rocket_icon.ico"
versionFile := repoRoot "\version"
outputExe   := repoRoot "\build\RAHKET installer.exe"
buildDir    := repoRoot "\build"

; ── Verify repo root looks right ──
if !FileExist(mainScript) {
    MsgBox("Could not find RAHKET_Main.ahk.`nExpected it at:`n" mainScript, "Build Error", 16)
    ExitApp()
}

; ── Read version from version file ──
if !FileExist(versionFile) {
    MsgBox("Could not find version file.`nExpected it at:`n" versionFile, "Build Error", 16)
    ExitApp()
}
buildVersion := Trim(FileRead(versionFile))
if (buildVersion = "") {
    MsgBox("version file is empty.", "Build Error", 16)
    ExitApp()
}

; ── Update version strings in RAHKET_Main.ahk ──
mainContent := FileRead(mainScript)

; Replace RAHKET_VERSION := "x.x.x.x"
mainContent := RegExReplace(mainContent, 'RAHKET_VERSION := "[^"]*"', 'RAHKET_VERSION := "' buildVersion '"')

; Replace ;@Ahk2Exe-SetVersion x.x.x.x
mainContent := RegExReplace(mainContent, ';@Ahk2Exe-SetVersion [^\r\n]*', ';@Ahk2Exe-SetVersion ' buildVersion)

FileDelete(mainScript)
FileAppend(mainContent, mainScript)

; ── Locate Ahk2Exe ──
defaultCompiler := EnvGet("LOCALAPPDATA") "\Programs\AutoHotkey\Compiler\Ahk2Exe.exe"
compilerPath    := ""

if FileExist(defaultCompiler) {
    compilerPath := defaultCompiler
} else {
    MsgBox(
        "Ahk2Exe.exe not found at the default location:`n" defaultCompiler
        "`n`nClick OK to browse for Ahk2Exe.exe.",
        "Compiler Not Found", 64
    )
    compilerPath := FileSelect(1, , "Locate Ahk2Exe.exe", "Executable (*.exe)")
    if (compilerPath = "") {
        MsgBox("No compiler selected. Build cancelled.", "Build Cancelled", 16)
        ExitApp()
    }
    if !FileExist(compilerPath) {
        MsgBox("Selected file does not exist. Build cancelled.", "Build Error", 16)
        ExitApp()
    }
}

; ── Ensure build\ output folder exists ──
if !DirExist(buildDir)
    DirCreate(buildDir)

; ── Step 1: Delete old modules.zip if present ──
if FileExist(zipDest)
    FileDelete(zipDest)

; ── Step 2: Zip the modules folder using PowerShell ──
psCmd := 'powershell.exe -NoProfile -Command "Compress-Archive -Path \"' . modulesDir . '\" -DestinationPath \"' . zipDest . '\" -Force"'

RunWait(psCmd, , "Hide")

if !FileExist(zipDest) {
    MsgBox("Failed to create modules.zip.`nCheck that the modules\ folder exists at:`n" modulesDir, "Build Error", 16)
    ExitApp()
}

; ── Step 3: Compile ──
compileCmd := '"' compilerPath '"'
    . ' /in "'   mainScript '"'
    . ' /out "'  outputExe '"'
    . ' /icon "' iconFile '"'

RunWait(compileCmd, , "Hide")

if !FileExist(outputExe) {
    MsgBox("Compilation failed.`nAhk2Exe did not produce an output file.`nCheck that RAHKET_Main.ahk has no syntax errors.", "Build Error", 16)
    ExitApp()
}

; ── Step 4: Write version file to build\ ──
buildVersionFile := buildDir "\version"
if FileExist(buildVersionFile)
    FileDelete(buildVersionFile)
FileAppend(buildVersion, buildVersionFile)

; ── Step 5: Clean up modules.zip from repo root ──
if FileExist(zipDest)
    FileDelete(zipDest)

; ── Done ──
MsgBox(
    "Build complete.`n`n"
    "Version: " buildVersion "`n"
    "EXE: " outputExe "`n"
    "Version file: " buildVersionFile,
    "Build Successful", 64
)
Run('explorer.exe "' buildDir '"')
ExitApp()
