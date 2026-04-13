; ============================================================================
; MODULE: LINKS MENU
; Written by Reece J. Goiffon, MD, PhD
; AutoHotkey v2
; Can be run standalone or included in RAHK_Main.ahk
; ============================================================================

#Requires AutoHotkey v2.0

; Only set these if running standalone
if (A_LineFile = A_ScriptFullPath) {
    #SingleInstance Force
    Persistent
    
    ; Create standalone tray menu with links
    A_TrayMenu.Delete()
    A_TrayMenu.Add("Links", CreateLinksMenu())
    A_TrayMenu.Add()
    A_TrayMenu.Add("Exit", (*) => ExitApp())
}

; Function to create and return the Links submenu
CreateLinksMenu() {
    LinksMenu := Menu()
    LinksMenu.Add("Google", Link_Google)
    LinksMenu.Add("ChatGPT", Link_ChatGPT)
    LinksMenu.Add("PubMed", Link_PubMed)
    LinksMenu.Add()
    LinksMenu.Add("Edit Links...", Link_EditLinks)
    return LinksMenu
}

; ============================================================================
; LINK HANDLERS
; ============================================================================

Link_Google(*) {
    Run("https://www.google.com/")
}

Link_ChatGPT(*) {
    Run("https://chatgpt.com/")
}

Link_PubMed(*) {
    Run("https://pubmed.ncbi.nlm.nih.gov/")
}

Link_EditLinks(*) {
    MsgBox("To add or edit links, modify the Links.ahk file in the Modules folder.`n`nAdd new links using:`nLinksMenu.Add(`"Name`", Link_FunctionName)", "Edit Links", 64)
}