; ============================================================================
; MODULE: LINES AND TUBES
; Written by Reece J. Goiffon, MD, PhD
; AutoHotkey v2
; Can be run standalone or included in RAHKET_Main.ahk
;
; PROTOTYPE SCOPE: this first pass implements the shared architecture (add
; device -> configure -> dynamic output -> grouped removal sentence -> copy)
; plus five devices to prove out every field pattern needed by the rest of
; the eventual device list:
;   - Endotracheal tube      (mutually exclusive sub-forms, blank fill-in)
;   - Enteric tube           (radio grid w/ dependent sub-field)
;   - Feeding tube           (enteric pattern + toggle modifiers)
;   - Internal jugular catheter (buttons + toggle + cycling anatomic groups)
;   - Subclavian vein catheter  (same as IJ, extended cycling groups)
; Remaining devices from the full list can be added as additional entries
; in LT_BuildDeviceDefs() using the same field-definition pattern.
;
; UI NOTES:
;   - Mouse only. No text entry anywhere. "Other" and any unfilled blank
;     always renders as "_____" for later dictation in PowerScribe.
;   - The Gui window is created ONCE and never destroyed/recreated. Clicks
;     only hide/show or update controls in the middle column, so a manual
;     resize/reposition of the window persists across every click.
;   - Anatomic locations with sub-variants (e.g. right atrium: generic /
;     upper / mid / lower) are grouped into a family button. Clicking the
;     family selects a default variant and reveals a second row of buttons
;     for the specific sub-variants -- the same "click reveals more
;     buttons" pattern as the ET tube's mainstem-bronchus side picker.
; ============================================================================

#Requires AutoHotkey v2.0

; Path to the click-region diagrams (same pattern as TN_ImgDir in
; ThyroidNodules.ahk): default assumes included from RAHKET_Main, overridden
; below for the standalone case inside the existing "only if running
; standalone" block.
LT_ImageDir := A_ScriptDir "\modules\LinesAndTubesImages"

; ---- global state ----
LT_Instances       := Map()   ; instanceId -> {deviceKey, fields: Map}
LT_InstanceOrder   := []      ; instanceIds in add-order
LT_InstanceCounter := 0
LT_GuiObj          := ""
LT_DeviceDefs      := Map()
LT_DeviceDefsBuilt := false

; Layout coordinates, recomputed on resize (see LT_FlushResize)
LT_MidX   := 0
LT_MidW   := 0
LT_RightX := 0
LT_RightW := 0

; Static controls that need to move/resize when the window is resized
LT_ListBoxCtrl       := ""
LT_AddBtnCtrl        := ""
LT_ClearBtnCtrl      := ""
LT_NewPatientBtnCtrl := ""
LT_MidLabelCtrl      := ""
LT_RightLabelCtrl    := ""
LT_CopyBtnCtrl       := ""
LT_PendingResizeW    := 0
LT_PendingResizeH    := 0

; The middle column is a genuine child window, reparented into LT_GuiObj
; exactly ONCE (see LT_CreateMidPanel) and never destroyed/reparented again
; -- an earlier attempt destroyed and re-reparented it on every click, which
; is almost certainly what corrupted AHK's internal window bookkeeping and
; crashed the process. Content within it is still just hidden/replaced the
; same safe way the rest of this module already works, same as before.
; Being a real child window with a real WS_VSCROLL scrollbar means Windows
; handles clipping and repositioning natively -- no more hand-rolled
; per-control Move/Visible math for ordinary scrolling.
LT_MidPanel        := ""
LT_MidPanelX       := 0
LT_MidPanelY       := 0
LT_MidPanelW       := 0
LT_MidPanelH       := 0
LT_MidPanelClientW := 0
LT_MidPanelClientH := 0

; The click-region diagram popup: a plain, independent top-level window
; (not reparented into anything -- it just floats next to the main window),
; created once and only ever hidden/shown after that, same as everything
; else in this module. Position/size are remembered in memory for the
; running session (not persisted to disk) once the person moves/resizes it.
LT_ImageGui         := ""
LT_ImagePicture     := ""
LT_ImageWinX        := 0
LT_ImageWinY        := 0
LT_ImageWinW        := 0
LT_ImageWinH        := 0
LT_ImageWinKnown    := false  ; becomes true once we have a real remembered position/size
LT_ImageWinVisible  := false  ; avoids redundant Show() calls while already visible
LT_ImageCurrentKey  := ""     ; which image is currently loaded ("" = none shown)
LT_ImageCurrentInst := ""     ; which instance's fields a click should update

; GDI+ state for drawing a translucent highlight over the currently-
; selected region on top of the base diagram. Started once, never shut
; down (the process reclaims it on exit like everything else here).
LT_GdipToken          := 0
LT_GdipImageBitmaps    := Map()  ; imageKey -> loaded GDI+ Bitmap pointer
LT_ImageCurrentHBitmap := 0      ; the HBITMAP currently shown in the Picture control, so it can be freed before being replaced

; Dynamic middle-column controls, tracked so a rebuild can hide the old set
; instead of destroying anything
LT_MiddleControls := []

; instId|fieldId -> array of {ctrl, meta}, used to update button text in
; place without a full middle-column rebuild
LT_FieldControls := Map()

LT_ScrollOffset     := 0
LT_MidContentBottom := 10

; One clickable Text control per report line on the right; clicking a line
; jumps to and expands the matching device in the middle column
LT_RightControls := []

; Parallel to the device picker ListBox's items: "" for a category header
; row, a device key otherwise. Rebuilt by LT_DeviceLabelList.
LT_PickListDeviceKeys := []

; Canonical output order (also the order devices are listed on the left)
; Categories drive both the report's device order and the left-side picker's
; grouping/dividers -- one source of truth so they can't drift apart.
LT_DeviceCategories := [
    Map("name", "Airway", "keys", ["ETT", "TRACH"]),
    Map("name", "GI tubes", "keys", ["ENTERIC", "FEEDING", "GGJ"]),
    Map("name", "Chest tubes & drains", "keys", ["PLEURAL", "MEDIASTINAL"]),
    Map("name", "Vascular lines", "keys", ["IJ", "SCV", "PICC", "PORT", "VECMO", "IABP"]),
    Map("name", "Cardiac devices", "keys", ["IMPELLA", "LVAD", "EPIWIRE", "GENERATOR", "RETAINEDLEADS", "PADS", "LOOPZIO", "LEADLESS"]),
    Map("name", "Other", "keys", ["EPIDURAL", "ABDOMINAL", "NEPHROSTOMY", "OTHER"])
]

LT_DeviceOrder := []
for LT_Cat_ in LT_DeviceCategories {
    for LT_Key_ in LT_Cat_["keys"]
        LT_DeviceOrder.Push(LT_Key_)
}

OnMessage(0x20A, LT_OnMouseWheel)  ; WM_MOUSEWHEEL
OnMessage(0x115, LT_OnVScroll)     ; WM_VSCROLL
OnMessage(0x0003, LT_OnImageMove)  ; WM_MOVE (filtered to the image popup window; tracks its position for the session)
OnMessage(0x0214, LT_OnImageSizing)  ; WM_SIZING (filtered to the image popup window; enforces its aspect ratio live during drag)

; ============================================================================
; SHARED ANATOMIC LOCATION GROUPS
; Each group is a family of related locations. If a family has only one
; state it acts as a plain single-click button. If it has multiple states,
; clicking the family button selects a default state and reveals a second
; row of buttons for the specific variants (same pattern as the ET tube's
; mainstem-bronchus side picker) rather than cycling through them.
; ============================================================================

LT_CentralTipGroups := [
    Map("groupLabel", "Brachiocephalic vein", "states", [
        Map("label", "Right brachiocephalic vein", "short", "Right", "phrase", "over the right brachiocephalic vein"),
        Map("label", "Left brachiocephalic vein", "short", "Left", "phrase", "over the left brachiocephalic vein")
    ]),
    Map("groupLabel", "Confluence of brachiocephalic veins", "states", [
        Map("label", "Confluence of the brachiocephalic veins", "phrase", "at the confluence of the brachiocephalic veins")
    ]),
    Map("groupLabel", "SVC", "states", [
        Map("label", "SVC (unspecified)", "short", "(unspecified)", "phrase", "in the superior vena cava"),
        Map("label", "Proximal SVC", "short", "Proximal", "phrase", "in the proximal superior vena cava"),
        Map("label", "Mid SVC", "short", "Mid", "phrase", "in the mid superior vena cava"),
        Map("label", "Distal SVC", "short", "Distal", "phrase", "in the distal superior vena cava")
    ]),
    Map("groupLabel", "Cavoatrial junction", "states", [
        Map("label", "Superior cavoatrial junction", "short", "Superior", "phrase", "at the superior cavoatrial junction"),
        Map("label", "Inferior cavoatrial junction", "short", "Inferior", "phrase", "at the inferior cavoatrial junction")
    ]),
    Map("groupLabel", "Right atrium", "states", [
        Map("label", "Right atrium (unspecified)", "short", "(unspecified)", "phrase", "in the right atrium"),
        Map("label", "Upper right atrium", "short", "Upper", "phrase", "in the upper right atrium"),
        Map("label", "Mid right atrium", "short", "Mid", "phrase", "in the mid right atrium"),
        Map("label", "Lower right atrium", "short", "Lower", "phrase", "in the lower right atrium")
    ]),
    Map("groupLabel", "Right ventricle", "states", [
        Map("label", "Right ventricle", "short", "Right", "phrase", "in the right ventricle")
    ]),
    Map("groupLabel", "Pulmonary artery", "states", [
        Map("label", "Pulmonary artery (unspecified)", "short", "(unspecified)", "phrase", "in the pulmonary artery"),
        Map("label", "Proximal pulmonary artery", "short", "Proximal", "phrase", "in the proximal pulmonary artery"),
        Map("label", "Mid pulmonary artery", "short", "Mid", "phrase", "in the mid pulmonary artery"),
        Map("label", "Distal pulmonary artery", "short", "Distal", "phrase", "in the distal pulmonary artery")
    ]),
    Map("groupLabel", "Right pulmonary artery", "states", [
        Map("label", "Right pulmonary artery (unspecified)", "short", "(unspecified)", "phrase", "in the right pulmonary artery"),
        Map("label", "Proximal right pulmonary artery", "short", "Proximal", "phrase", "in the proximal right pulmonary artery"),
        Map("label", "Mid right pulmonary artery", "short", "Mid", "phrase", "in the mid right pulmonary artery"),
        Map("label", "Distal right pulmonary artery", "short", "Distal", "phrase", "in the distal right pulmonary artery")
    ]),
    Map("groupLabel", "Left pulmonary artery", "states", [
        Map("label", "Left pulmonary artery (unspecified)", "short", "(unspecified)", "phrase", "in the left pulmonary artery"),
        Map("label", "Proximal left pulmonary artery", "short", "Proximal", "phrase", "in the proximal left pulmonary artery"),
        Map("label", "Mid left pulmonary artery", "short", "Mid", "phrase", "in the mid left pulmonary artery"),
        Map("label", "Distal left pulmonary artery", "short", "Distal", "phrase", "in the distal left pulmonary artery")
    ]),
    Map("groupLabel", "Interlobar artery", "states", [
        Map("label", "Right interlobar artery", "short", "Right", "phrase", "in the right interlobar artery"),
        Map("label", "Left interlobar artery", "short", "Left", "phrase", "in the left interlobar artery")
    ]),
    Map("groupLabel", "Other", "states", [
        Map("label", "Other", "phrase", "__OTHER__")
    ])
]

; Same families/labels as LT_CentralTipGroups, but every phrase uses
; "projecting over X" uniformly. IJ and SCV catheter tips are described
; this way; PICC keeps the original in/at/over mix above.
LT_ProjectingTipGroups := [
    Map("groupLabel", "Brachiocephalic vein", "states", [
        Map("label", "Right brachiocephalic vein", "short", "Right", "phrase", "projecting over the right brachiocephalic vein"),
        Map("label", "Left brachiocephalic vein", "short", "Left", "phrase", "projecting over the left brachiocephalic vein")
    ]),
    Map("groupLabel", "Confluence of brachiocephalic veins", "states", [
        Map("label", "Confluence of the brachiocephalic veins", "phrase", "projecting over the confluence of the brachiocephalic veins")
    ]),
    Map("groupLabel", "SVC", "states", [
        Map("label", "SVC (unspecified)", "short", "(unspecified)", "phrase", "projecting over the superior vena cava"),
        Map("label", "Proximal SVC", "short", "Proximal", "phrase", "projecting over the proximal superior vena cava"),
        Map("label", "Mid SVC", "short", "Mid", "phrase", "projecting over the mid superior vena cava"),
        Map("label", "Distal SVC", "short", "Distal", "phrase", "projecting over the distal superior vena cava")
    ]),
    Map("groupLabel", "Cavoatrial junction", "states", [
        Map("label", "Superior cavoatrial junction", "short", "Superior", "phrase", "projecting over the superior cavoatrial junction"),
        Map("label", "Inferior cavoatrial junction", "short", "Inferior", "phrase", "projecting over the inferior cavoatrial junction")
    ]),
    Map("groupLabel", "Right atrium", "states", [
        Map("label", "Right atrium (unspecified)", "short", "(unspecified)", "phrase", "projecting over the right atrium"),
        Map("label", "Upper right atrium", "short", "Upper", "phrase", "projecting over the upper right atrium"),
        Map("label", "Mid right atrium", "short", "Mid", "phrase", "projecting over the mid right atrium"),
        Map("label", "Lower right atrium", "short", "Lower", "phrase", "projecting over the lower right atrium")
    ]),
    Map("groupLabel", "Right ventricle", "states", [
        Map("label", "Right ventricle", "short", "Right", "phrase", "projecting over the right ventricle")
    ]),
    Map("groupLabel", "Pulmonary artery", "states", [
        Map("label", "Pulmonary artery (unspecified)", "short", "(unspecified)", "phrase", "projecting over the pulmonary artery"),
        Map("label", "Proximal pulmonary artery", "short", "Proximal", "phrase", "projecting over the proximal pulmonary artery"),
        Map("label", "Mid pulmonary artery", "short", "Mid", "phrase", "projecting over the mid pulmonary artery"),
        Map("label", "Distal pulmonary artery", "short", "Distal", "phrase", "projecting over the distal pulmonary artery")
    ]),
    Map("groupLabel", "Right pulmonary artery", "states", [
        Map("label", "Right pulmonary artery (unspecified)", "short", "(unspecified)", "phrase", "projecting over the right pulmonary artery"),
        Map("label", "Proximal right pulmonary artery", "short", "Proximal", "phrase", "projecting over the proximal right pulmonary artery"),
        Map("label", "Mid right pulmonary artery", "short", "Mid", "phrase", "projecting over the mid right pulmonary artery"),
        Map("label", "Distal right pulmonary artery", "short", "Distal", "phrase", "projecting over the distal right pulmonary artery")
    ]),
    Map("groupLabel", "Left pulmonary artery", "states", [
        Map("label", "Left pulmonary artery (unspecified)", "short", "(unspecified)", "phrase", "projecting over the left pulmonary artery"),
        Map("label", "Proximal left pulmonary artery", "short", "Proximal", "phrase", "projecting over the proximal left pulmonary artery"),
        Map("label", "Mid left pulmonary artery", "short", "Mid", "phrase", "projecting over the mid left pulmonary artery"),
        Map("label", "Distal left pulmonary artery", "short", "Distal", "phrase", "projecting over the distal left pulmonary artery")
    ]),
    Map("groupLabel", "Interlobar artery", "states", [
        Map("label", "Right interlobar artery", "short", "Right", "phrase", "projecting over the right interlobar artery"),
        Map("label", "Left interlobar artery", "short", "Left", "phrase", "projecting over the left interlobar artery")
    ]),
    Map("groupLabel", "Other", "states", [
        Map("label", "Other", "phrase", "__OTHER__")
    ])
]

; PICC-specific tip locations. Unlike the shared central-line set, this is
; capped at the right ventricle (a PICC has no business projecting into the
; pulmonary arteries) and every phrase uniformly reads "projects to the X" --
; true regardless of location, and specific to PICC (IJ/SCV use "projecting
; over", Port uses the same "projecting over" set minus PA/interlobar).
LT_PICCTipGroups := [
    Map("groupLabel", "Brachiocephalic vein", "states", [
        Map("label", "Right brachiocephalic vein", "short", "Right", "phrase", "projects to the right brachiocephalic vein"),
        Map("label", "Left brachiocephalic vein", "short", "Left", "phrase", "projects to the left brachiocephalic vein")
    ]),
    Map("groupLabel", "Confluence of brachiocephalic veins", "states", [
        Map("label", "Confluence of the brachiocephalic veins", "phrase", "projects to the confluence of the brachiocephalic veins")
    ]),
    Map("groupLabel", "SVC", "states", [
        Map("label", "SVC (unspecified)", "short", "(unspecified)", "phrase", "projects to the superior vena cava"),
        Map("label", "Proximal SVC", "short", "Proximal", "phrase", "projects to the proximal superior vena cava"),
        Map("label", "Mid SVC", "short", "Mid", "phrase", "projects to the mid superior vena cava"),
        Map("label", "Distal SVC", "short", "Distal", "phrase", "projects to the distal superior vena cava")
    ]),
    Map("groupLabel", "Cavoatrial junction", "states", [
        Map("label", "Superior cavoatrial junction", "short", "Superior", "phrase", "projects to the superior cavoatrial junction"),
        Map("label", "Inferior cavoatrial junction", "short", "Inferior", "phrase", "projects to the inferior cavoatrial junction")
    ]),
    Map("groupLabel", "Right atrium", "states", [
        Map("label", "Right atrium (unspecified)", "short", "(unspecified)", "phrase", "projects to the right atrium"),
        Map("label", "Upper right atrium", "short", "Upper", "phrase", "projects to the upper right atrium"),
        Map("label", "Mid right atrium", "short", "Mid", "phrase", "projects to the mid right atrium"),
        Map("label", "Lower right atrium", "short", "Lower", "phrase", "projects to the lower right atrium")
    ]),
    Map("groupLabel", "Right ventricle", "states", [
        Map("label", "Right ventricle", "phrase", "projects to the right ventricle")
    ]),
    Map("groupLabel", "Clavicle", "states", [
        Map("label", "Right clavicle", "short", "Right", "phrase", "projects to the right clavicle"),
        Map("label", "Left clavicle", "short", "Left", "phrase", "projects to the left clavicle")
    ]),
    Map("groupLabel", "Other", "states", [
        Map("label", "Other", "phrase", "__OTHER__")
    ])
]

; Consolidated vein/central-line location set shared by IJ, SCV, PICC, and
; Port -- same families, same order, same labels across all four, so one
; picture/region-map can eventually drive all of them. Order matches
; top-to-bottom anatomic position for that future picture, not alphabetical.
; Families marked "requiresSubSelection" have no default state: picking the
; family reveals its sub-choices but leaves the field genuinely incomplete
; (red) until one is clicked, rather than silently assuming a side.
LT_VeinTipGroups_Projecting := [
    Map("groupLabel", "Axillary v.", "requiresSubSelection", true, "states", [
        Map("label", "Right axillary vein", "short", "Right", "phrase", "projecting over the right axillary vein"),
        Map("label", "Left axillary vein", "short", "Left", "phrase", "projecting over the left axillary vein")
    ]),
    Map("groupLabel", "Subclavian v.", "requiresSubSelection", true, "states", [
        Map("label", "Right subclavian vein", "short", "Right", "phrase", "projecting over the right subclavian vein"),
        Map("label", "Left subclavian vein", "short", "Left", "phrase", "projecting over the left subclavian vein")
    ]),
    Map("groupLabel", "Brachiocephalic v.", "requiresSubSelection", true, "states", [
        Map("label", "Right brachiocephalic vein", "short", "Right", "phrase", "projecting over the right brachiocephalic vein"),
        Map("label", "Left brachiocephalic vein", "short", "Left", "phrase", "projecting over the left brachiocephalic vein"),
        Map("label", "Confluence of the brachiocephalic veins", "short", "Confluence", "phrase", "projecting over the confluence of the brachiocephalic veins")
    ]),
    Map("groupLabel", "SVC", "states", [
        Map("label", "SVC (unspecified)", "short", "(unspecified)", "phrase", "projecting over the superior vena cava"),
        Map("label", "Proximal SVC", "short", "Proximal", "phrase", "projecting over the proximal superior vena cava"),
        Map("label", "Mid SVC", "short", "Mid", "phrase", "projecting over the mid superior vena cava"),
        Map("label", "Distal SVC", "short", "Distal", "phrase", "projecting over the distal superior vena cava")
    ]),
    Map("groupLabel", "Sup. cavoatrial jct.", "states", [
        Map("label", "Superior cavoatrial junction", "phrase", "projecting over the superior cavoatrial junction")
    ]),
    Map("groupLabel", "R. atrium", "states", [
        Map("label", "Right atrium (unspecified)", "short", "(unspecified)", "phrase", "projecting over the right atrium"),
        Map("label", "Upper right atrium", "short", "Upper", "phrase", "projecting over the upper right atrium"),
        Map("label", "Mid right atrium", "short", "Mid", "phrase", "projecting over the mid right atrium"),
        Map("label", "Lower right atrium", "short", "Lower", "phrase", "projecting over the lower right atrium")
    ]),
    Map("groupLabel", "Inf. cavoatrial jct.", "states", [
        Map("label", "Inferior cavoatrial junction", "phrase", "projecting over the inferior cavoatrial junction")
    ]),
    Map("groupLabel", "IVC", "states", [
        Map("label", "Inferior vena cava", "phrase", "projecting over the inferior vena cava")
    ]),
    Map("groupLabel", "R. ventricle", "states", [
        Map("label", "Right ventricle", "phrase", "projecting over the right ventricle")
    ]),
    Map("groupLabel", "Pulmonary art.", "states", [
        Map("label", "Pulmonary artery (unspecified)", "short", "(unspecified)", "phrase", "projecting over the pulmonary artery"),
        Map("label", "Proximal pulmonary artery", "short", "Proximal", "phrase", "projecting over the proximal pulmonary artery"),
        Map("label", "Mid pulmonary artery", "short", "Mid", "phrase", "projecting over the mid pulmonary artery"),
        Map("label", "Distal pulmonary artery", "short", "Distal", "phrase", "projecting over the distal pulmonary artery")
    ]),
    Map("groupLabel", "R. pulmonary art.", "states", [
        Map("label", "Right pulmonary artery (unspecified)", "short", "(unspecified)", "phrase", "projecting over the right pulmonary artery"),
        Map("label", "Proximal right pulmonary artery", "short", "Proximal", "phrase", "projecting over the proximal right pulmonary artery"),
        Map("label", "Mid right pulmonary artery", "short", "Mid", "phrase", "projecting over the mid right pulmonary artery"),
        Map("label", "Distal right pulmonary artery", "short", "Distal", "phrase", "projecting over the distal right pulmonary artery")
    ]),
    Map("groupLabel", "L. pulmonary art.", "states", [
        Map("label", "Left pulmonary artery", "phrase", "projecting over the left pulmonary artery")
    ]),
    Map("groupLabel", "Interlobar art.", "requiresSubSelection", true, "states", [
        Map("label", "Right interlobar artery", "short", "Right", "phrase", "projecting over the right interlobar artery"),
        Map("label", "Left interlobar artery", "short", "Left", "phrase", "projecting over the left interlobar artery")
    ]),
    Map("groupLabel", "Other", "states", [
        Map("label", "Other", "phrase", "__OTHER__")
    ])
]

; Same structure, PICC's "projects to" phrasing.
LT_VeinTipGroups_PICC := [
    Map("groupLabel", "Axillary v.", "requiresSubSelection", true, "states", [
        Map("label", "Right axillary vein", "short", "Right", "phrase", "projects to the right axillary vein"),
        Map("label", "Left axillary vein", "short", "Left", "phrase", "projects to the left axillary vein")
    ]),
    Map("groupLabel", "Subclavian v.", "requiresSubSelection", true, "states", [
        Map("label", "Right subclavian vein", "short", "Right", "phrase", "projects to the right subclavian vein"),
        Map("label", "Left subclavian vein", "short", "Left", "phrase", "projects to the left subclavian vein")
    ]),
    Map("groupLabel", "Brachiocephalic v.", "requiresSubSelection", true, "states", [
        Map("label", "Right brachiocephalic vein", "short", "Right", "phrase", "projects to the right brachiocephalic vein"),
        Map("label", "Left brachiocephalic vein", "short", "Left", "phrase", "projects to the left brachiocephalic vein"),
        Map("label", "Confluence of the brachiocephalic veins", "short", "Confluence", "phrase", "projects to the confluence of the brachiocephalic veins")
    ]),
    Map("groupLabel", "SVC", "states", [
        Map("label", "SVC (unspecified)", "short", "(unspecified)", "phrase", "projects to the superior vena cava"),
        Map("label", "Proximal SVC", "short", "Proximal", "phrase", "projects to the proximal superior vena cava"),
        Map("label", "Mid SVC", "short", "Mid", "phrase", "projects to the mid superior vena cava"),
        Map("label", "Distal SVC", "short", "Distal", "phrase", "projects to the distal superior vena cava")
    ]),
    Map("groupLabel", "Sup. cavoatrial jct.", "states", [
        Map("label", "Superior cavoatrial junction", "phrase", "projects to the superior cavoatrial junction")
    ]),
    Map("groupLabel", "R. atrium", "states", [
        Map("label", "Right atrium (unspecified)", "short", "(unspecified)", "phrase", "projects to the right atrium"),
        Map("label", "Upper right atrium", "short", "Upper", "phrase", "projects to the upper right atrium"),
        Map("label", "Mid right atrium", "short", "Mid", "phrase", "projects to the mid right atrium"),
        Map("label", "Lower right atrium", "short", "Lower", "phrase", "projects to the lower right atrium")
    ]),
    Map("groupLabel", "Inf. cavoatrial jct.", "states", [
        Map("label", "Inferior cavoatrial junction", "phrase", "projects to the inferior cavoatrial junction")
    ]),
    Map("groupLabel", "IVC", "states", [
        Map("label", "Inferior vena cava", "phrase", "projects to the inferior vena cava")
    ]),
    Map("groupLabel", "R. ventricle", "states", [
        Map("label", "Right ventricle", "phrase", "projects to the right ventricle")
    ]),
    Map("groupLabel", "Pulmonary art.", "states", [
        Map("label", "Pulmonary artery (unspecified)", "short", "(unspecified)", "phrase", "projects to the pulmonary artery"),
        Map("label", "Proximal pulmonary artery", "short", "Proximal", "phrase", "projects to the proximal pulmonary artery"),
        Map("label", "Mid pulmonary artery", "short", "Mid", "phrase", "projects to the mid pulmonary artery"),
        Map("label", "Distal pulmonary artery", "short", "Distal", "phrase", "projects to the distal pulmonary artery")
    ]),
    Map("groupLabel", "R. pulmonary art.", "states", [
        Map("label", "Right pulmonary artery (unspecified)", "short", "(unspecified)", "phrase", "projects to the right pulmonary artery"),
        Map("label", "Proximal right pulmonary artery", "short", "Proximal", "phrase", "projects to the proximal right pulmonary artery"),
        Map("label", "Mid right pulmonary artery", "short", "Mid", "phrase", "projects to the mid right pulmonary artery"),
        Map("label", "Distal right pulmonary artery", "short", "Distal", "phrase", "projects to the distal right pulmonary artery")
    ]),
    Map("groupLabel", "L. pulmonary art.", "states", [
        Map("label", "Left pulmonary artery", "phrase", "projects to the left pulmonary artery")
    ]),
    Map("groupLabel", "Interlobar art.", "requiresSubSelection", true, "states", [
        Map("label", "Right interlobar artery", "short", "Right", "phrase", "projects to the right interlobar artery"),
        Map("label", "Left interlobar artery", "short", "Left", "phrase", "projects to the left interlobar artery")
    ]),
    Map("groupLabel", "Other", "states", [
        Map("label", "Other", "phrase", "__OTHER__")
    ])
]

LT_EntericTipGroups := [
    Map("groupLabel", "Esophagus", "states", [
        Map("label", "Esophagus (unspecified)", "short", "(unspecified)", "phrase", "in the esophagus"),
        Map("label", "Upper esophagus", "short", "Upper", "phrase", "in the upper esophagus"),
        Map("label", "Mid esophagus", "short", "Mid", "phrase", "in the mid esophagus"),
        Map("label", "Distal esophagus", "short", "Distal", "phrase", "in the distal esophagus")
    ]),
    Map("groupLabel", "Esophagogastric junction", "states", [
        Map("label", "Near esophagogastric junction", "short", "Near", "phrase", "near the esophagogastric junction"),
        Map("label", "Above esophagogastric junction", "short", "Above", "phrase", "above the esophagogastric junction"),
        Map("label", "Below esophagogastric junction", "short", "Below", "phrase", "below the esophagogastric junction")
    ]),
    Map("groupLabel", "Gastric conduit", "states", [
        Map("label", "Gastric conduit", "phrase", "in the gastric conduit")
    ]),
    Map("groupLabel", "Stomach (tip and side port)", "states", [
        Map("label", "Stomach (tip and side port)", "short", "(unspecified)", "phrase", "__STOMACH_GEN__"),
        Map("label", "Proximal stomach", "short", "Proximal", "phrase", "__STOMACH_PROX__"),
        Map("label", "Mid stomach", "short", "Mid", "phrase", "__STOMACH_MID__"),
        Map("label", "Distal stomach", "short", "Distal", "phrase", "__STOMACH_DIST__")
    ]),
    Map("groupLabel", "Off-image below diaphragm", "states", [
        Map("label", "Off-image below diaphragm", "phrase", "__OFFIMAGE__")
    ]),
    Map("groupLabel", "Malpositioned bronchus", "states", [
        Map("label", "Right lower lobe bronchus", "short", "Right", "phrase", "__MALPOS_R__"),
        Map("label", "Left lower lobe bronchus", "short", "Left", "phrase", "__MALPOS_L__")
    ]),
    Map("groupLabel", "Other", "states", [
        Map("label", "Other", "phrase", "__OTHER__")
    ])
]

LT_FeedingTipGroups := [
    Map("groupLabel", "Esophagus", "states", [
        Map("label", "Esophagus (unspecified)", "short", "(unspecified)", "phrase", "in the esophagus"),
        Map("label", "Upper esophagus", "short", "Upper", "phrase", "in the upper esophagus"),
        Map("label", "Mid esophagus", "short", "Mid", "phrase", "in the mid esophagus"),
        Map("label", "Distal esophagus", "short", "Distal", "phrase", "in the distal esophagus")
    ]),
    Map("groupLabel", "Esophagogastric junction", "states", [
        Map("label", "Near the esophagogastric junction", "short", "Near", "phrase", "near the esophagogastric junction"),
        Map("label", "Before or at the esophagogastric junction", "short", "Before/at", "phrase", "before or at the esophagogastric junction"),
        Map("label", "At the esophagogastric junction", "short", "At", "phrase", "at the esophagogastric junction"),
        Map("label", "At or beyond the esophagogastric junction", "short", "At/beyond", "phrase", "at or beyond the esophagogastric junction")
    ]),
    Map("groupLabel", "Gastric conduit", "states", [
        Map("label", "Gastric conduit", "phrase", "in the gastric conduit")
    ]),
    Map("groupLabel", "Stomach", "states", [
        Map("label", "Stomach (unspecified)", "short", "(unspecified)", "phrase", "in the stomach"),
        Map("label", "Proximal stomach", "short", "Proximal", "phrase", "in the proximal stomach"),
        Map("label", "Fundus", "phrase", "in the gastric fundus"),
        Map("label", "Mid stomach", "short", "Mid", "phrase", "in the mid stomach"),
        Map("label", "Distal stomach", "short", "Distal", "phrase", "in the distal stomach")
    ]),
    Map("groupLabel", "Duodenum", "states", [
        Map("label", "Duodenum", "phrase", "postpyloric and into the duodenum")
    ]),
    Map("groupLabel", "Pylorus", "states", [
        Map("label", "Near the pylorus", "short", "Near", "phrase", "near the pylorus"),
        Map("label", "Before or at the pylorus", "short", "Before/at", "phrase", "before or at the pylorus"),
        Map("label", "At the pylorus", "short", "At", "phrase", "at the pylorus"),
        Map("label", "At or beyond the pylorus", "short", "At/beyond", "phrase", "at or beyond the pylorus")
    ]),
    Map("groupLabel", "Duodenojejunal junction", "states", [
        Map("label", "At duodenojejunal junction", "short", "At", "phrase", "at the duodenojejunal junction"),
        Map("label", "Beyond duodenojejunal junction", "short", "Beyond", "phrase", "beyond the duodenojejunal junction")
    ]),
    Map("groupLabel", "Below diaphragm", "states", [
        Map("label", "Below diaphragm", "phrase", "below the diaphragm"),
        Map("label", "Off-image below diaphragm", "short", "Off-image", "phrase", "off-image below the diaphragm")
    ]),
    Map("groupLabel", "Endobronchial", "requiresSubSelection", true, "states", [
        Map("label", "Right lung", "short", "Right", "phrase", "endobronchial, in the right lung"),
        Map("label", "Left lung", "short", "Left", "phrase", "endobronchial, in the left lung")
    ]),
    Map("groupLabel", "Other", "states", [
        Map("label", "Other", "phrase", "__OTHER__")
    ])
]

LT_EpiduralGroups := [
    Map("groupLabel", "Thoracic spine", "states", [
        Map("label", "Thoracic spine (unspecified)", "short", "(unspecified)", "phrase", "over the thoracic spine"),
        Map("label", "Upper thoracic spine", "short", "Upper", "phrase", "over the upper thoracic spine"),
        Map("label", "Mid thoracic spine", "short", "Mid", "phrase", "over the mid thoracic spine"),
        Map("label", "Lower thoracic spine", "short", "Lower", "phrase", "over the lower thoracic spine")
    ]),
    Map("groupLabel", "Lumbar spine", "states", [
        Map("label", "Lumbar spine (unspecified)", "short", "(unspecified)", "phrase", "over the lumbar spine"),
        Map("label", "Upper lumbar spine", "short", "Upper", "phrase", "over the upper lumbar spine"),
        Map("label", "Mid lumbar spine", "short", "Mid", "phrase", "over the mid lumbar spine")
    ]),
    Map("groupLabel", "Other", "states", [
        Map("label", "Other", "phrase", "__OTHER__")
    ])
]

; ============================================================================
; IMAGE-MAP DATA (click-region pickers)
; ============================================================================

; Traced regions for vein.png, in native image pixel coordinates (polygon
; vertices). Purely positional data -- LT_VeinRegionMap below is what maps
; a region's name to an actual field/family/state.
LT_VeinImageRegions := [
    Map("name", "R axillary", "points", [[258, 155], [276, 114], [175, 133], [82, 163], [95, 196], [184, 170], [223, 161]]),
    Map("name", "R subclavian", "points", [[276, 112], [258, 156], [308, 151], [352, 154], [374, 116], [325, 109]]),
    Map("name", "R side", "points", [[289, 80], [366, 80], [364, 11], [290, 10]]),
    Map("name", "L side", "points", [[562, 76], [623, 77], [625, 12], [562, 9]]),
    Map("name", "R brachiocephalic", "points", [[374, 115], [396, 133], [404, 147], [412, 172], [415, 197], [416, 208], [392, 210], [375, 221], [371, 184], [362, 158], [352, 154]]),
    Map("name", "L brachiocephalic", "points", [[418, 208], [459, 173], [521, 134], [517, 183], [488, 202], [456, 230], [446, 217], [430, 210]]),
    Map("name", "L subclavian", "points", [[521, 133], [557, 118], [614, 102], [643, 97], [662, 132], [624, 139], [557, 162], [519, 182]]),
    Map("name", "L axillary", "points", [[663, 132], [702, 130], [737, 138], [767, 155], [785, 171], [809, 146], [769, 114], [730, 97], [685, 92], [660, 94], [644, 97]]),
    Map("name", "bracioceph confluence", "points", [[456, 230], [439, 214], [416, 208], [394, 210], [380, 217], [376, 220], [377, 251], [389, 257], [406, 261], [427, 261], [430, 259], [440, 246], [450, 236]]),
    Map("name", "prox SVC", "points", [[376, 252], [387, 258], [400, 261], [399, 303], [377, 301]]),
    Map("name", "mid SVC", "points", [[376, 302], [398, 303], [397, 362], [375, 363], [376, 330]]),
    Map("name", "distal SVC", "points", [[374, 365], [373, 415], [395, 413], [397, 362]]),
    Map("name", "SVC", "points", [[400, 262], [414, 263], [428, 261], [420, 284], [419, 308], [419, 416], [396, 414], [398, 341]]),
    Map("name", "S CAJ", "points", [[429, 441], [418, 432], [418, 417], [396, 414], [373, 415], [373, 426], [371, 436], [368, 443], [390, 448], [413, 446]]),
    Map("name", "upper RA", "points", [[368, 445], [361, 468], [357, 495], [397, 494], [395, 447], [380, 446]]),
    Map("name", "mid RA", "points", [[356, 495], [355, 524], [358, 556], [399, 558], [398, 494]]),
    Map("name", "lower RA", "points", [[358, 558], [366, 576], [374, 593], [378, 606], [400, 611], [398, 558]]),
    Map("name", "RA", "points", [[395, 448], [401, 610], [420, 618], [439, 629], [450, 602], [448, 564], [447, 529], [449, 493], [444, 463], [431, 442], [412, 447]]),
    Map("name", "Inf CAJ", "points", [[439, 629], [433, 633], [422, 632], [414, 631], [413, 648], [377, 638], [377, 606], [409, 614], [431, 623]]),
    Map("name", "IVC", "points", [[377, 637], [395, 644], [412, 648], [412, 708], [377, 708]]),
    Map("name", "RV", "points", [[462, 455], [455, 473], [449, 482], [450, 510], [448, 534], [450, 562], [450, 598], [443, 626], [476, 642], [531, 648], [579, 645], [624, 638], [579, 574], [536, 512], [532, 485], [537, 463], [524, 463], [485, 459]]),
    Map("name", "PA", "points", [[533, 311], [553, 363], [538, 462], [527, 464], [499, 460]]),
    Map("name", "dist PA", "points", [[491, 370], [490, 381], [515, 382], [532, 311]]),
    Map("name", "mid PA", "points", [[489, 382], [483, 401], [474, 424], [505, 429], [516, 383]]),
    Map("name", "prox PA", "points", [[474, 425], [463, 454], [498, 459], [505, 429]]),
    Map("name", "R PA", "points", [[531, 311], [509, 314], [462, 337], [420, 357], [420, 383], [492, 349], [511, 341]]),
    Map("name", "prox RPA", "points", [[508, 343], [452, 368], [454, 386], [489, 370]]),
    Map("name", "mid RPA", "points", [[420, 383], [420, 400], [452, 385], [451, 369]]),
    Map("name", "dist RPA", "points", [[373, 383], [372, 423], [355, 438], [352, 422], [351, 397], [372, 382]]),
    Map("name", "R interlobar", "points", [[350, 398], [335, 412], [318, 442], [308, 479], [298, 526], [318, 530], [337, 473], [348, 447], [353, 441]]),
    Map("name", "L PA", "points", [[553, 362], [570, 362], [586, 376], [603, 332], [576, 320], [552, 312], [532, 311], [544, 339]]),
    Map("name", "L interlobar", "points", [[603, 333], [618, 348], [628, 370], [637, 396], [647, 414], [681, 474], [661, 486], [631, 445], [611, 419], [594, 388], [586, 377]])
]

; Native pixel size of vein.png, needed to convert a click's on-screen
; (possibly resized) coordinates back to the same space LT_VeinImageRegions
; was traced in.
LT_VeinImageNativeW := 951
LT_VeinImageNativeH := 937

; Maps a region's name to what clicking it should do: either
; Map("side", "Right"/"Left") to set the device's own laterality field, or
; Map("family", <groupLabel>, "state", <state label>) to resolve the tip
; field directly to that family+state -- looked up against whichever
; groups array (Projecting or PICC phrasing) the clicked device's own "tip"
; field actually uses, so this same map drives IJ/SCV/PICC/Port uniformly.
LT_VeinRegionMap := Map(
    "R side", Map("plainField", "laterality", "plainValue", "Right"),
    "L side", Map("plainField", "laterality", "plainValue", "Left"),
    "R axillary", Map("family", "Axillary v.", "state", "Right axillary vein"),
    "L axillary", Map("family", "Axillary v.", "state", "Left axillary vein"),
    "R subclavian", Map("family", "Subclavian v.", "state", "Right subclavian vein"),
    "L subclavian", Map("family", "Subclavian v.", "state", "Left subclavian vein"),
    "R brachiocephalic", Map("family", "Brachiocephalic v.", "state", "Right brachiocephalic vein"),
    "L brachiocephalic", Map("family", "Brachiocephalic v.", "state", "Left brachiocephalic vein"),
    "bracioceph confluence", Map("family", "Brachiocephalic v.", "state", "Confluence of the brachiocephalic veins"),
    "SVC", Map("family", "SVC", "state", "SVC (unspecified)"),
    "prox SVC", Map("family", "SVC", "state", "Proximal SVC"),
    "mid SVC", Map("family", "SVC", "state", "Mid SVC"),
    "distal SVC", Map("family", "SVC", "state", "Distal SVC"),
    "S CAJ", Map("family", "Sup. cavoatrial jct.", "state", "Superior cavoatrial junction"),
    "RA", Map("family", "R. atrium", "state", "Right atrium (unspecified)"),
    "upper RA", Map("family", "R. atrium", "state", "Upper right atrium"),
    "mid RA", Map("family", "R. atrium", "state", "Mid right atrium"),
    "lower RA", Map("family", "R. atrium", "state", "Lower right atrium"),
    "Inf CAJ", Map("family", "Inf. cavoatrial jct.", "state", "Inferior cavoatrial junction"),
    "IVC", Map("family", "IVC", "state", "Inferior vena cava"),
    "RV", Map("family", "R. ventricle", "state", "Right ventricle"),
    "PA", Map("family", "Pulmonary art.", "state", "Pulmonary artery (unspecified)"),
    "prox PA", Map("family", "Pulmonary art.", "state", "Proximal pulmonary artery"),
    "mid PA", Map("family", "Pulmonary art.", "state", "Mid pulmonary artery"),
    "dist PA", Map("family", "Pulmonary art.", "state", "Distal pulmonary artery"),
    "R PA", Map("family", "R. pulmonary art.", "state", "Right pulmonary artery (unspecified)"),
    "prox RPA", Map("family", "R. pulmonary art.", "state", "Proximal right pulmonary artery"),
    "mid RPA", Map("family", "R. pulmonary art.", "state", "Mid right pulmonary artery"),
    "dist RPA", Map("family", "R. pulmonary art.", "state", "Distal right pulmonary artery"),
    "L PA", Map("family", "L. pulmonary art.", "state", "Left pulmonary artery"),
    "R interlobar", Map("family", "Interlobar art.", "state", "Right interlobar artery"),
    "L interlobar", Map("family", "Interlobar art.", "state", "Left interlobar artery")
)

; Traced regions for enteric_feeding.png (951x939 native), same idea as the
; vein data above -- shared by Enteric and Feeding tube, not GGJ.
LT_EntericFeedingImageRegions := [
    Map("name", "R lung", "points", [[371, 311], [345, 324], [328, 351], [320, 393], [323, 421], [336, 458], [356, 476], [374, 480], [398, 470], [420, 431], [425, 395], [420, 361], [412, 340], [397, 322], [386, 316]]),
    Map("name", "L lung", "points", [[588, 311], [559, 328], [542, 364], [538, 400], [543, 432], [560, 466], [588, 481], [610, 473], [624, 457], [635, 434], [641, 401], [639, 370], [629, 341], [613, 321]]),
    Map("name", "esophagus", "points", [[449, 1], [452, 89], [448, 237], [448, 370], [450, 504], [462, 572], [469, 605], [482, 595], [495, 588], [496, 585], [486, 554], [477, 504], [469, 427], [463, 324], [464, 244], [465, 128], [467, 55], [468, 0]]),
    Map("name", "distal esophagus", "points", [[471, 462], [498, 462], [506, 502], [515, 540], [527, 579], [512, 579], [499, 584], [496, 586], [487, 558], [479, 521], [475, 489]]),
    Map("name", "mid esophagus", "points", [[465, 223], [463, 273], [465, 369], [468, 429], [471, 464], [497, 463], [491, 421], [485, 345], [485, 268], [484, 223]]),
    Map("name", "prox esophagus", "points", [[491, 1], [488, 93], [485, 187], [483, 224], [464, 223], [464, 143], [466, 48], [469, 0]]),
    Map("name", "EG junction", "points", [[497, 588], [504, 604], [518, 623], [535, 642], [543, 647], [580, 653], [586, 654], [575, 662], [558, 669], [542, 672], [526, 659], [507, 648], [481, 629], [471, 605], [483, 594]]),
    Map("name", "at beyond EG", "points", [[550, 629], [534, 639], [543, 647], [571, 653], [586, 655], [590, 639], [587, 630], [572, 630], [560, 631], [550, 629]]),
    Map("name", "at EG", "points", [[509, 611], [521, 605], [534, 599], [537, 607], [536, 615], [540, 624], [544, 629], [550, 628], [534, 638], [515, 620], [512, 616]]),
    Map("name", "above at EG", "points", [[534, 598], [527, 580], [513, 580], [497, 587], [501, 596], [509, 610]]),
    Map("name", "stomach", "points", [[589, 654], [647, 661], [672, 672], [682, 690], [677, 712], [664, 725], [611, 767], [572, 799], [529, 829], [507, 841], [479, 851], [478, 828], [469, 800], [463, 790], [516, 754], [552, 733], [579, 717], [589, 708], [593, 702], [590, 697], [585, 692], [578, 689], [564, 683], [555, 679], [548, 673], [546, 671], [566, 667], [580, 661], [584, 657]]),
    Map("name", "fundus", "points", [[647, 661], [669, 642], [691, 623], [672, 622], [622, 626], [595, 629], [587, 629], [590, 644], [588, 654], [618, 657]]),
    Map("name", "prox stomach", "points", [[692, 623], [721, 631], [746, 648], [762, 668], [772, 691], [774, 720], [769, 746], [764, 758], [661, 728], [678, 711], [683, 691], [678, 676], [661, 666], [647, 661]]),
    Map("name", "mid stomach", "points", [[661, 728], [631, 754], [593, 783], [560, 808], [606, 864], [664, 842], [689, 828], [723, 806], [742, 788], [763, 758]]),
    Map("name", "distal stomach", "points", [[479, 851], [499, 845], [514, 837], [536, 823], [561, 807], [606, 864], [561, 876], [525, 884], [498, 888], [475, 891], [479, 871], [479, 862]]),
    Map("name", "pylorus", "points", [[479, 851], [439, 859], [416, 866], [398, 873], [385, 881], [378, 886], [378, 863], [383, 838], [419, 817], [442, 805], [463, 790]]),
    Map("name", "above at pyloris", "points", [[449, 858], [456, 892], [475, 891], [478, 872], [478, 851]]),
    Map("name", "at pylorus", "points", [[412, 867], [414, 899], [456, 893], [449, 857]]),
    Map("name", "at beyond pylorus", "points", [[412, 868], [396, 873], [379, 885], [381, 899], [387, 917], [395, 907], [398, 904], [406, 902], [414, 900]]),
    Map("name", "postpyloric", "points", [[386, 918], [379, 892], [377, 874], [382, 837], [362, 854], [347, 875], [335, 905], [331, 928], [332, 936], [373, 938], [377, 927]]),
    Map("name", "tip", "points", [[849, 110], [925, 110], [924, 48], [849, 48]]),
    Map("name", "port", "points", [[807, 130], [806, 182], [924, 184], [924, 122], [807, 124]]),
    Map("name", "tip and port", "points", [[925, 242], [923, 191], [603, 189], [603, 247]]),
    Map("name", "below diaphragm", "points", [[780, 637], [782, 716], [835, 715], [835, 633]]),
    Map("name", "off image", "points", [[783, 732], [781, 909], [835, 911], [836, 729]])
]

LT_EntericFeedingImageNativeW := 951
LT_EntericFeedingImageNativeH := 939

LT_EntericFeedingRegionMap := Map(
    "R lung", Map("family", "Endobronchial", "state", "Right lung"),
    "L lung", Map("family", "Endobronchial", "state", "Left lung"),
    "esophagus", Map("family", "Esophagus", "state", "Esophagus (unspecified)"),
    "prox esophagus", Map("family", "Esophagus", "state", "Upper esophagus"),
    "mid esophagus", Map("family", "Esophagus", "state", "Mid esophagus"),
    "distal esophagus", Map("family", "Esophagus", "state", "Distal esophagus"),
    "EG junction", Map("family", "Esophagogastric junction", "state", "Near the esophagogastric junction"),
    "above at EG", Map("family", "Esophagogastric junction", "state", "Before or at the esophagogastric junction"),
    "at EG", Map("family", "Esophagogastric junction", "state", "At the esophagogastric junction"),
    "at beyond EG", Map("family", "Esophagogastric junction", "state", "At or beyond the esophagogastric junction"),
    "stomach", Map("family", "Stomach", "state", "Stomach (unspecified)"),
    "fundus", Map("family", "Stomach", "state", "Fundus"),
    "prox stomach", Map("family", "Stomach", "state", "Proximal stomach"),
    "mid stomach", Map("family", "Stomach", "state", "Mid stomach"),
    "distal stomach", Map("family", "Stomach", "state", "Distal stomach"),
    "pylorus", Map("family", "Pylorus", "state", "Near the pylorus"),
    "above at pyloris", Map("family", "Pylorus", "state", "Before or at the pylorus"),
    "at pylorus", Map("family", "Pylorus", "state", "At the pylorus"),
    "at beyond pylorus", Map("family", "Pylorus", "state", "At or beyond the pylorus"),
    "postpyloric", Map("family", "Duodenum", "state", "Duodenum"),
    "below diaphragm", Map("family", "Below diaphragm", "state", "Below diaphragm"),
    "off image", Map("family", "Below diaphragm", "state", "Off-image below diaphragm"),
    "tip", Map("plainField", "part", "plainValue", "Tip"),
    "port", Map("plainField", "part", "plainValue", "Side port"),
    "tip and port", Map("plainField", "part", "plainValue", "Tip and side port")
)
; Only set these if running standalone
if (A_LineFile = A_ScriptFullPath) {
    ; Override image dir when launched standalone (or compiled EXE on its own)
    LT_ImageDir := A_ScriptDir "\LinesAndTubesImages"

    #SingleInstance Force
    Persistent

    A_TrayMenu.Delete()
    A_TrayMenu.Add("Open Lines and Tubes", (*) => Show_LinesAndTubes())
    A_TrayMenu.Add()
    A_TrayMenu.Add("Exit", (*) => ExitApp())

    Show_LinesAndTubes()
}


; ============================================================================
; ENTRY POINT
; ============================================================================

Show_LinesAndTubes(*) {
    global LT_DeviceDefs, LT_DeviceDefsBuilt

    if (!LT_DeviceDefsBuilt) {
        LT_DeviceDefs := LT_BuildDeviceDefs()
        LT_DeviceDefsBuilt := true
    }

    LT_EnsureGui()
    LT_GuiObj.Show()
    LT_RebuildMiddleColumn()
}


; ============================================================================
; DEVICE DEFINITIONS
; ============================================================================

LT_BuildDeviceDefs() {
    global LT_CentralTipGroups, LT_ProjectingTipGroups, LT_PICCTipGroups, LT_VeinTipGroups_Projecting, LT_VeinTipGroups_PICC
    global LT_EntericTipGroups, LT_FeedingTipGroups, LT_EpiduralGroups

    defs := Map()

    defs["ETT"] := Map(
        "label", "Endotracheal tube",
        "shortLabel", "ETT",
        "fields", [
            Map("id", "form", "type", "buttons", "label", "Position",
                "options", ["Above carina", "Mainstem bronchus"]),
            Map("id", "side", "type", "buttons", "label", "Side",
                "options", ["Right", "Left"],
                "showIf", Map("field", "form", "value", "Mainstem bronchus"))
        ],
        "sentenceFn", LT_Sentence_ETT,
        "removalNoun", (fields) => Map("text", "endotracheal tube", "plural", false)
    )

    defs["ENTERIC"] := Map(
        "label", "Enteric tube",
        "fields", [
            Map("id", "location", "type", "grouped", "label", "Tip location",
                "groups", LT_FeedingTipGroups),
            Map("id", "part", "type", "buttons", "label", "Part",
                "options", ["Tip", "Side port", "Tip and side port"])
        ],
        "sentenceFn", (fields) => LT_Sentence_Enteric(fields, "Enteric tube"),
        "removalNoun", (fields) => Map("text", "enteric tube", "plural", false),
        "imageKey", "enteric_feeding"
    )

    defs["FEEDING"] := Map(
        "label", "Feeding tube",
        "fields", [
            Map("id", "location", "type", "grouped", "label", "Tip location",
                "groups", LT_FeedingTipGroups),
            Map("id", "part", "type", "buttons", "label", "Part",
                "options", ["Tip", "Side port", "Tip and side port"]),
            Map("id", "weighted", "type", "toggle", "label", "Weighted tip"),
            Map("id", "stylet", "type", "toggle", "label", "With stylet")
        ],
        "sentenceFn", LT_Sentence_Feeding,
        "removalNoun", (fields) => Map("text", fields.Get("weighted", false) ? "weighted feeding tube" : "feeding tube", "plural", false),
        "imageKey", "enteric_feeding"
    )

    defs["GGJ"] := Map(
        "label", "G/GJ/J tube",
        "fields", [
            Map("id", "tubeType", "type", "buttons", "label", "Type",
                "options", ["G tube", "GJ tube", "J tube"]),
            Map("id", "location", "type", "grouped", "label", "Tip location",
                "groups", LT_BuildGGJTipGroups())
        ],
        "sentenceFn", LT_Sentence_GGJ,
        "removalNoun", LT_RemovalNoun_GGJ
    )

    defs["IJ"] := Map(
        "label", "Internal jugular catheter",
        "shortLabel", "IJ catheter",
        "fields", [
            Map("id", "laterality", "type", "buttons", "label", "Side",
                "options", ["Right", "Left"]),
            Map("id", "boreSize", "type", "buttons", "label", "Bore",
                "options", ["", "Large bore", "Small bore"]),
            Map("id", "modRepositioned", "type", "toggle", "label", "Repositioned"),
            Map("id", "modTunneled", "type", "toggle", "label", "Tunneled"),
            Map("id", "modSheathed", "type", "toggle", "label", "Sheathed"),
            Map("id", "sheathOnly", "type", "toggle", "label", "Sheath only (empty)"),
            Map("id", "tip", "type", "grouped", "label", "Tip location",
                "groups", LT_VeinTipGroups_Projecting)
        ],
        "sentenceFn", (fields) => LT_Sentence_CentralLine(fields, "internal jugular catheter"),
        "removalNoun", (fields) => LT_RemovalNoun_CentralLine(fields, "internal jugular catheter"),
        "imageKey", "vein"
    )

    defs["SCV"] := Map(
        "label", "Subclavian vein catheter",
        "shortLabel", "SCV catheter",
        "fields", [
            Map("id", "laterality", "type", "buttons", "label", "Side",
                "options", ["Right", "Left"]),
            Map("id", "boreSize", "type", "buttons", "label", "Bore",
                "options", ["", "Large bore", "Small bore"]),
            Map("id", "modRepositioned", "type", "toggle", "label", "Repositioned"),
            Map("id", "modTunneled", "type", "toggle", "label", "Tunneled"),
            Map("id", "modSheathed", "type", "toggle", "label", "Sheathed"),
            Map("id", "sheathOnly", "type", "toggle", "label", "Sheath only (empty)"),
            Map("id", "tip", "type", "grouped", "label", "Tip location",
                "groups", LT_VeinTipGroups_Projecting)
        ],
        "sentenceFn", (fields) => LT_Sentence_CentralLine(fields, "subclavian vein catheter"),
        "removalNoun", (fields) => LT_RemovalNoun_CentralLine(fields, "subclavian vein catheter"),
        "imageKey", "vein"
    )

    defs["PICC"] := Map(
        "label", "Peripherally inserted central catheter (PICC)",
        "shortLabel", "PICC",
        "fields", [
            Map("id", "laterality", "type", "buttons", "label", "Side",
                "options", ["Right", "Left"]),
            Map("id", "tip", "type", "grouped", "label", "Tip location",
                "groups", LT_VeinTipGroups_PICC)
        ],
        "sentenceFn", (fields) => LT_Sentence_CentralLine(fields, "peripherally inserted central catheter (PICC)"),
        "removalNoun", (fields) => LT_RemovalNoun_CentralLine(fields, "peripherally inserted central catheter (PICC)"),
        "imageKey", "vein"
    )

    defs["PORT"] := Map(
        "label", "Port",
        "fields", [
            Map("id", "laterality", "type", "buttons", "label", "Side",
                "options", ["Right", "Left"]),
            Map("id", "position", "type", "buttons", "label", "Position",
                "options", ["Anterior chest wall", "Lower chest wall"]),
            Map("id", "accessed", "type", "buttons", "label", "Access",
                "options", ["Accessed", "Unaccessed"]),
            Map("id", "portType", "type", "buttons", "label", "Type",
                "options", ["Single", "Dual"]),
            Map("id", "tip", "type", "grouped", "label", "Tip location",
                "groups", LT_VeinTipGroups_Projecting)
        ],
        "sentenceFn", LT_Sentence_Port,
        "removalNoun", LT_RemovalNoun_Port,
        "imageKey", "vein"
    )

    defs["VECMO"] := Map(
        "label", "Venous extracorporeal membrane oxygenation (ECMO) cannula",
        "shortLabel", "vECMO",
        "fields", [
            Map("id", "approach", "type", "buttons", "label", "Approach",
                "options", ["Superior", "Inferior"]),
            Map("id", "tip", "type", "radio", "label", "Tip location",
                "options", [
                    Map("label", "Infrahepatic IVC", "phrase", "in the infrahepatic inferior vena cava"),
                    Map("label", "Hepatic IVC", "phrase", "in the hepatic inferior vena cava"),
                    Map("label", "Inferior cavoatrial junction", "short", "Inferior", "phrase", "at the inferior cavoatrial junction"),
                    Map("label", "Right atrium", "short", "Right", "phrase", "in the right atrium"),
                    Map("label", "Superior cavoatrial junction", "short", "Superior", "phrase", "at the superior cavoatrial junction"),
                    Map("label", "SVC", "phrase", "in the superior vena cava"),
                    Map("label", "Other", "phrase", "__OTHER__")
                ])
        ],
        "sentenceFn", LT_Sentence_VECMO,
        "removalNoun", LT_RemovalNoun_VECMO
    )

    defs["IABP"] := Map(
        "label", "Intra-aortic balloon pump",
        "shortLabel", "IABP",
        "fields", [],
        "sentenceFn", (fields) => "Intra-aortic balloon pump, tip _____ below the left subclavian artery.",
        "removalNoun", (fields) => Map("text", "intra-aortic balloon pump", "plural", false)
    )

    defs["TRACH"] := Map(
        "label", "Tracheostomy tube",
        "fields", [
            Map("id", "form", "type", "buttons", "label", "Position",
                "options", ["At thoracic inlet", "Above carina"])
        ],
        "sentenceFn", LT_Sentence_Trach,
        "removalNoun", (fields) => Map("text", "tracheostomy tube", "plural", false)
    )

    defs["PLEURAL"] := Map(
        "label", "Pleural tube",
        "fields", [
            Map("id", "laterality", "type", "buttons", "label", "Side",
                "options", ["Right", "Left"]),
            Map("id", "deviceType", "type", "buttons", "label", "Type",
                "options", ["Chest tube", "Pleural catheter", "Pigtail catheter"]),
            Map("id", "location", "type", "buttons", "label", "Location",
                "options", ["Apical", "Basal", "Mid", "Chest wall", "Other"]),
            Map("id", "count", "type", "counter", "label", "Count", "min", 1, "max", 9)
        ],
        "sentenceFn", LT_Sentence_Pleural,
        "removalNoun", LT_RemovalNoun_Pleural,
        "aggregate", true,
        "aggregateFn", LT_AggregatePleural
    )

    defs["MEDIASTINAL"] := Map(
        "label", "Mediastinal drain",
        "fields", [
            Map("id", "inferiorApproach", "type", "toggle", "label", "Inferior approach"),
            Map("id", "position", "type", "buttons", "label", "Position",
                "options", ["Anterior", "Posterior"]),
            Map("id", "count", "type", "counter", "label", "Count", "min", 1, "max", 9)
        ],
        "sentenceFn", LT_Sentence_Mediastinal,
        "removalNoun", LT_RemovalNoun_Mediastinal,
        "aggregate", true,
        "aggregateFn", LT_AggregateMediastinal
    )

    defs["EPIDURAL"] := Map(
        "label", "Epidural catheter",
        "fields", [
            Map("id", "tip", "type", "grouped", "label", "Tip location",
                "groups", LT_EpiduralGroups)
        ],
        "sentenceFn", LT_Sentence_Epidural,
        "removalNoun", (fields) => Map("text", "epidural catheter", "plural", false)
    )

    defs["EPIWIRE"] := Map(
        "label", "Epicardial pacing wires",
        "shortLabel", "Epicardial wires",
        "fields", [
            Map("id", "retained", "type", "toggle", "label", "Retained")
        ],
        "sentenceFn", (fields) => (fields.Get("retained", false)
            ? "Retained epicardial pacing wires."
            : "Epicardial pacing wires."),
        "removalNoun", (fields) => Map("text", fields.Get("retained", false)
            ? "retained epicardial pacing wires"
            : "epicardial pacing wires", "plural", true)
    )

    defs["IMPELLA"] := Map(
        "label", "Impella left ventricular assist device",
        "shortLabel", "Impella LVAD",
        "fields", [
            Map("id", "approach", "type", "buttons", "label", "Approach",
                "options", ["Right axillary", "Left axillary", "Right radial", "Left radial", "Aortofemoral"]),
            Map("id", "crossingValve", "type", "toggle", "label", "Crossing the aortic valve plane")
        ],
        "sentenceFn", LT_Sentence_Impella,
        "removalNoun", (fields) => Map("text", "Impella left ventricular assist device (LVAD)", "plural", false)
    )

    defs["LVAD"] := Map(
        "label", "Left ventricular assist device (LVAD)",
        "shortLabel", "LVAD",
        "fields", [
            Map("id", "model", "type", "buttons", "label", "Model",
                "options", ["HeartMate 3", "HVAD"])
        ],
        "sentenceFn", LT_Sentence_LVAD,
        "removalNoun", (fields) => Map("text", "left ventricular assist device (LVAD)", "plural", false)
    )

    defs["PADS"] := Map(
        "label", "External pacing/defibrillation pads",
        "shortLabel", "ACD pads",
        "fields", [],
        "sentenceFn", (fields) => "External pacing/defibrillation pads.",
        "removalNoun", (fields) => Map("text", "external pacing/defibrillation pads", "plural", true)
    )

    defs["LOOPZIO"] := Map(
        "label", "Superficial cardiac monitor",
        "shortLabel", "Loop/Zio",
        "fields", [],
        "sentenceFn", (fields) => "Superficial cardiac monitor.",
        "removalNoun", (fields) => Map("text", "superficial cardiac monitor", "plural", false)
    )

    defs["LEADLESS"] := Map(
        "label", "Leadless pacemaker",
        "fields", [],
        "sentenceFn", (fields) => "Leadless pacemaker over the right ventricle.",
        "removalNoun", (fields) => Map("text", "leadless pacemaker", "plural", false)
    )

    defs["GENERATOR"] := Map(
        "label", "Chest wall generator",
        "fields", [
            Map("id", "laterality", "type", "buttons", "label", "Side",
                "options", ["Right", "Left"]),
            Map("id", "position", "type", "buttons", "label", "Position",
                "options", ["Anterior", "Lower"]),
            Map("id", "intact", "type", "toggle", "label", "Intact"),
            Map("id", "leadType", "type", "buttons", "label", "Lead type",
                "options", ["Pacing", "Defibrillation"]),
            Map("id", "locRA", "type", "toggle", "label", "Right atrium"),
            Map("id", "locRV", "type", "toggle", "label", "Right ventricle"),
            Map("id", "locCS", "type", "toggle", "label", "Coronary sinus"),
            Map("id", "locCV", "type", "toggle", "label", "Cardiac veins"),
            Map("id", "locPS", "type", "toggle", "label", "Presternal")
        ],
        "sentenceFn", LT_Sentence_Generator,
        "removalNoun", LT_RemovalNoun_Generator
    )

    defs["RETAINEDLEADS"] := Map(
        "label", "Retained pacing/defibrillation leads",
        "shortLabel", "Retained leads",
        "fields", [
            Map("id", "leadType", "type", "buttons", "label", "Lead type",
                "options", ["Pacing", "Defibrillation"]),
            Map("id", "locRA", "type", "toggle", "label", "Right atrium"),
            Map("id", "locRV", "type", "toggle", "label", "Right ventricle"),
            Map("id", "locCS", "type", "toggle", "label", "Coronary sinus"),
            Map("id", "locCV", "type", "toggle", "label", "Cardiac veins"),
            Map("id", "locPS", "type", "toggle", "label", "Presternal")
        ],
        "sentenceFn", LT_Sentence_RetainedLeads,
        "removalNoun", (fields) => Map("text", "retained pacing/defibrillation leads", "plural", true)
    )

    defs["ABDOMINAL"] := Map(
        "label", "Abdominal drain",
        "fields", [
            Map("id", "location", "type", "buttons", "label", "Location",
                "options", ["Left upper quadrant", "Epigastric", "Right upper quadrant", "Subhepatic", "Other"]),
            Map("id", "deviceType", "type", "buttons", "label", "Type",
                "options", ["Pigtail catheter", "Drain"]),
            Map("id", "count", "type", "counter", "label", "Count", "min", 1, "max", 9)
        ],
        "sentenceFn", LT_Sentence_Abdominal,
        "removalNoun", LT_RemovalNoun_Abdominal
    )

    defs["NEPHROSTOMY"] := Map(
        "label", "Percutaneous nephrostomy catheter",
        "shortLabel", "Nephrostomy",
        "fields", [
            Map("id", "laterality", "type", "buttons", "label", "Side",
                "options", ["Left", "Right", "Bilateral"])
        ],
        "sentenceFn", LT_Sentence_Nephrostomy,
        "removalNoun", LT_RemovalNoun_Nephrostomy
    )

    defs["OTHER"] := Map(
        "label", "Other",
        "fields", [],
        "sentenceFn", (fields) => "_____",
        "removalNoun", (fields) => Map("text", "_____", "plural", false)
    )

    return defs
}

LT_BuildSubclavianTipGroups() {
    global LT_CentralTipGroups
    arr := []
    for grp in LT_CentralTipGroups {
        if (grp["groupLabel"] = "Other")
            continue
        arr.Push(grp)
    }
    arr.Push(Map("groupLabel", "Clavicle", "states", [
        Map("label", "Right clavicle", "short", "Right", "phrase", "projecting over the right clavicle"),
        Map("label", "Left clavicle", "short", "Left", "phrase", "projecting over the left clavicle")
    ]))
    arr.Push(Map("groupLabel", "Other", "states", [
        Map("label", "Other", "phrase", "__OTHER__")
    ]))
    return arr
}

; Same idea, but built on the "projecting over" phrasing set -- used by SCV.
LT_BuildProjectingSubclavianGroups() {
    global LT_ProjectingTipGroups
    arr := []
    for grp in LT_ProjectingTipGroups {
        if (grp["groupLabel"] = "Other")
            continue
        arr.Push(grp)
    }
    arr.Push(Map("groupLabel", "Clavicle", "states", [
        Map("label", "Right clavicle", "short", "Right", "phrase", "projecting over the right clavicle"),
        Map("label", "Left clavicle", "short", "Left", "phrase", "projecting over the left clavicle")
    ]))
    arr.Push(Map("groupLabel", "Other", "states", [
        Map("label", "Other", "phrase", "__OTHER__")
    ]))
    return arr
}

; G/GJ/J tube shares the feeding tube's stomach-and-beyond locations, minus
; esophagus, the esophagogastric junction, gastric conduit, and the
; malpositioned-bronchus option -- none of those apply to a tube that
; starts in the stomach.
LT_BuildGGJTipGroups() {
    global LT_FeedingTipGroups
    excluded := ["Esophagus", "Esophagogastric junction", "Gastric conduit", "Endobronchial"]
    arr := []
    for grp in LT_FeedingTipGroups {
        skip := false
        for ex in excluded {
            if (grp["groupLabel"] = ex) {
                skip := true
                break
            }
        }
        if (!skip)
            arr.Push(grp)
    }
    return arr
}

; Port catheters get the same central-line families minus the pulmonary
; artery ones (a port catheter tip doesn't sit in the PA).
LT_BuildPortTipGroups() {
    global LT_CentralTipGroups
    excluded := ["Pulmonary artery", "Right pulmonary artery", "Left pulmonary artery", "Interlobar artery"]
    arr := []
    for grp in LT_CentralTipGroups {
        skip := false
        for ex in excluded {
            if (grp["groupLabel"] = ex) {
                skip := true
                break
            }
        }
        if (!skip)
            arr.Push(grp)
    }
    return arr
}

; Port catheters aren't placed under real-time US guidance the way IJ/SCV
; lines are, so their tip position is described the same "projecting over"
; way, not "in"/"at".
LT_BuildProjectingPortTipGroups() {
    global LT_ProjectingTipGroups
    excluded := ["Pulmonary artery", "Right pulmonary artery", "Left pulmonary artery", "Interlobar artery"]
    arr := []
    for grp in LT_ProjectingTipGroups {
        skip := false
        for ex in excluded {
            if (grp["groupLabel"] = ex) {
                skip := true
                break
            }
        }
        if (!skip)
            arr.Push(grp)
    }
    return arr
}


LT_FindFieldDef(def, fieldId) {
    for fd in def["fields"] {
        if (fd["id"] = fieldId)
            return fd
    }
    return ""
}

LT_FieldHasDependents(def, fieldId) {
    for fd in def["fields"] {
        if (fd.Has("showIf") && fd["showIf"]["field"] = fieldId)
            return true
    }
    return false
}

; Devices with a "laterality" field (currently the vascular catheters, and
; any future device that follows the same pattern e.g. chest tubes) require
; a side to be selected before the report is considered complete.
LT_DeviceHasLateralityField(def) {
    for fd in def["fields"] {
        if (fd["id"] = "laterality")
            return true
    }
    return false
}

; True if any grouped field's currently-active family requires an explicit
; sub-selection (e.g. Axillary v. -- Right/Left, no default) but hasn't
; gotten one yet.
LT_FieldsHaveIncompleteGrouped(fields, def) {
    for fd in def["fields"] {
        if (fd["type"] != "grouped")
            continue
        fieldId := fd["id"]
        activeFamily := fields.Get(fieldId "_activeFamily", "")
        if (activeFamily = "")
            continue
        for grp in fd["groups"] {
            if (grp["groupLabel"] = activeFamily && grp.Get("requiresSubSelection", false) && fields.Get(fieldId, "") = "")
                return true
        }
    }
    return false
}

LT_FieldsNeedSideWarning(fields, def) {
    if (fields.Get("removed", false))
        return false
    if (LT_DeviceHasLateralityField(def) && fields.Get("laterality", "") = "")
        return true
    if (LT_FieldsHaveIncompleteGrouped(fields, def))
        return true
    return false
}

LT_InstanceNeedsSideWarning(instId) {
    global LT_Instances, LT_DeviceDefs
    inst := LT_Instances[instId]
    def := LT_DeviceDefs[inst["deviceKey"]]
    return LT_FieldsNeedSideWarning(inst["fields"], def)
}

LT_AnyInstanceNeedsSideWarning() {
    global LT_InstanceOrder
    for instId in LT_InstanceOrder {
        if (LT_InstanceNeedsSideWarning(instId))
            return true
    }
    return false
}


; ============================================================================
; SENTENCE BUILDERS
; ============================================================================

LT_Capitalize(s) {
    if (s = "")
        return s
    return StrUpper(SubStr(s, 1, 1)) SubStr(s, 2)
}

LT_LowercaseFirst(s) {
    if (s = "")
        return s
    return StrLower(SubStr(s, 1, 1)) SubStr(s, 2)
}

; "A" / "A and B" / "A, B, and C"
LT_JoinWithAnd(arr) {
    n := arr.Length
    if (n = 0)
        return ""
    if (n = 1)
        return arr[1]
    if (n = 2)
        return arr[1] " and " arr[2]
    out := ""
    Loop n - 1
        out .= arr[A_Index] ", "
    out .= "and " arr[n]
    return out
}

LT_Sentence_CentralLine(fields, nounBase) {
    side       := fields.Get("laterality", "")
    sheathOnly := fields.Get("sheathOnly", false)
    tipPhrase  := fields.Get("tip_phrase", "")

    if (tipPhrase = "__OTHER__")
        tipPhrase := "_____"
    if (tipPhrase = "")
        tipPhrase := "_____"

    ; Attributes are independently toggleable (a line can be tunneled AND
    ; large bore AND sheathed at once); bore size stays mutually exclusive
    ; with itself since a catheter can't be both large and small bore.
    attrParts := []
    if (fields.Get("modRepositioned", false))
        attrParts.Push("repositioned")
    if (fields.Get("modTunneled", false))
        attrParts.Push("tunneled")
    if (fields.Get("modSheathed", false))
        attrParts.Push("sheathed")

    boreSize := fields.Get("boreSize", "")
    if (boreSize != "")
        attrParts.Push(StrLower(boreSize))

    modText := ""
    for i, p in attrParts
        modText .= (i > 1 ? " " : "") p
    if (modText != "")
        modText .= " "

    sideText := (side != "") ? StrLower(side) " " : ""
    noun     := sheathOnly ? (nounBase " sheath") : nounBase
    subject  := (sheathOnly ? "Empty " : "") modText sideText noun
    subject  := LT_Capitalize(subject)

    return subject " tip " tipPhrase "."
}

LT_RemovalNoun_CentralLine(fields, nounBase) {
    side := fields.Get("laterality", "")
    sheathOnly := fields.Get("sheathOnly", false)
    sideText := (side != "") ? StrLower(side) " " : ""
    noun := sheathOnly ? (nounBase " sheath") : nounBase
    return Map("text", sideText noun, "plural", false)
}

LT_Sentence_ETT(fields) {
    form := fields.Get("form", "")

    if (form = "Mainstem bronchus") {
        side := fields.Get("side", "")
        sideText := (side != "") ? StrLower(side) : "_____"
        return "Endotracheal tube tip in the " sideText " mainstem bronchus."
    }

    return "Endotracheal tube tip _____ above the carina."
}

LT_Sentence_Enteric(fields, deviceLabel) {
    loc := fields.Get("location_phrase", "")
    part := fields.Get("part", "")
    partText := (part = "Side port") ? "side port" : (part = "Tip and side port") ? "tip and side port" : "tip"

    if (loc = "")
        return deviceLabel "."
    if (loc = "__OTHER__")
        return deviceLabel " " partText " _____."

    return deviceLabel " " partText " " loc "."
}

LT_Sentence_Feeding(fields) {
    weighted := fields.Get("weighted", false)
    stylet   := fields.Get("stylet", false)
    label    := (weighted ? "Weighted feeding" : "Feeding") " tube"

    base := LT_Sentence_Enteric(fields, label)

    if (stylet && SubStr(base, -1) = ".")
        base := SubStr(base, 1, StrLen(base) - 1) ", with stylet in place."

    return base
}

LT_GGJLabel(fields) {
    tubeType := fields.Get("tubeType", "")
    if (tubeType = "GJ tube")
        return "Gastrojejunostomy tube"
    if (tubeType = "J tube")
        return "Jejunostomy tube"
    return "Gastrostomy tube"
}

LT_Sentence_GGJ(fields) {
    return LT_Sentence_Enteric(fields, LT_GGJLabel(fields))
}

LT_RemovalNoun_GGJ(fields) {
    return Map("text", StrLower(LT_GGJLabel(fields)), "plural", false)
}

LT_Sentence_Trach(fields) {
    form := fields.Get("form", "")
    if (form = "At thoracic inlet")
        return "Tracheostomy tube tip at the thoracic inlet."
    if (form = "Above carina")
        return "Tracheostomy tube tip _____ above the carina."
    return "Tracheostomy tube."
}

LT_Sentence_Pleural(fields) {
    side := fields.Get("laterality", "")
    location := fields.Get("location", "")
    deviceType := fields.Get("deviceType", "")
    count := fields.Get("count", 1)

    sideText := (side != "") ? StrLower(side) " " : ""
    locText := (location = "Other") ? "_____ " : (location != "") ? StrLower(location) " " : ""
    typeNoun := (deviceType != "") ? StrLower(deviceType) : "chest tube"
    if (count > 1)
        typeNoun .= "s"

    base := LT_Capitalize(sideText locText typeNoun)
    if (count > 1)
        base .= " [" count "]"
    base .= "."
    return base
}

LT_RemovalNoun_Pleural(fields) {
    side := fields.Get("laterality", "")
    deviceType := fields.Get("deviceType", "")
    count := fields.Get("count", 1)
    sideText := (side != "") ? StrLower(side) " " : ""
    typeNoun := (deviceType != "") ? StrLower(deviceType) : "chest tube"
    if (count > 1)
        typeNoun .= "s"
    return Map("text", sideText typeNoun, "plural", count > 1)
}

; Combines every active pleural tube/pigtail/catheter instance onto one
; shared line, e.g. "Right apical chest tubes [2] and left basal pigtail
; catheters [3]" instead of one line per instance. Each instance already
; produces a complete self-contained descriptor via LT_Sentence_Pleural, so
; this just joins those fragments together (lowercasing every fragment
; after the first, since only the start of the whole line should be
; capitalized).
LT_AggregatePleural(instancesFields) {
    frags := []
    for i, f in instancesFields {
        s := LT_Sentence_Pleural(f)
        s := LT_AppendOtherNote(s, f)
        s := LT_PrependNewModifier(s, f)
        s := LT_StripTrailingPeriod(s)
        if (i > 1)
            s := LT_LowercaseFirst(s)
        frags.Push(s)
    }
    if (frags.Length = 0)
        return ""
    return LT_JoinWithAnd(frags)
}

LT_Sentence_VECMO(fields) {
    approach := fields.Get("approach", "")
    tipPhrase := fields.Get("tip_phrase", "")
    if (tipPhrase = "" || tipPhrase = "__OTHER__")
        tipPhrase := "_____"

    approachText := (approach != "") ? StrLower(approach) " approach " : ""
    base := approachText "venous extracorporeal membrane oxygenation (ECMO) cannula"
    return LT_Capitalize(base) ", tip " tipPhrase "."
}

LT_RemovalNoun_VECMO(fields) {
    approach := fields.Get("approach", "")
    approachText := (approach != "") ? StrLower(approach) " approach " : ""
    return Map("text", approachText "venous extracorporeal membrane oxygenation (ECMO) cannula", "plural", false)
}

LT_Sentence_Port(fields) {
    accessed := fields.Get("accessed", "")
    side := fields.Get("laterality", "")
    position := fields.Get("position", "")
    portType := fields.Get("portType", "")
    tipPhrase := fields.Get("tip_phrase", "")
    if (tipPhrase = "__OTHER__" || tipPhrase = "")
        tipPhrase := "_____"

    parts := []
    if (accessed != "")
        parts.Push(StrLower(accessed))
    if (portType != "")
        parts.Push(StrLower(portType))
    if (side != "")
        parts.Push(StrLower(side))
    if (position != "")
        parts.Push(StrLower(position))

    prefix := ""
    for i, p in parts
        prefix .= (i > 1 ? " " : "") p
    if (prefix != "")
        prefix .= " "

    subject := LT_Capitalize(prefix "port catheter")
    return subject " tip " tipPhrase "."
}

LT_RemovalNoun_Port(fields) {
    side := fields.Get("laterality", "")
    position := fields.Get("position", "")
    parts := []
    if (side != "")
        parts.Push(StrLower(side))
    if (position != "")
        parts.Push(StrLower(position))
    parts.Push("port")
    out := ""
    for i, p in parts
        out .= (i > 1 ? " " : "") p
    return Map("text", out, "plural", false)
}

LT_Sentence_Impella(fields) {
    approach := fields.Get("approach", "")
    crossing := fields.Get("crossingValve", false)

    approachText := (approach != "") ? StrLower(approach) : "_____"
    base := "Impella left ventricular assist device (LVAD) via " approachText " approach"
    if (crossing)
        base .= ", crossing the aortic valve plane"
    base .= "."
    return base
}

LT_Sentence_Epidural(fields) {
    phrase := fields.Get("tip_phrase", "")
    if (phrase = "" || phrase = "__OTHER__")
        phrase := "_____"
    return "Epidural catheter tip " phrase "."
}

LT_Sentence_Mediastinal(fields) {
    inferior := fields.Get("inferiorApproach", false)
    position := fields.Get("position", "")
    count := fields.Get("count", 1)

    parts := []
    if (inferior)
        parts.Push("inferior approach")
    if (position != "")
        parts.Push(StrLower(position))

    desc := ""
    for i, p in parts
        desc .= (i > 1 ? ", " : "") p

    noun := "mediastinal drain"
    if (count > 1)
        noun .= "s"

    base := LT_Capitalize(noun)
    if (count > 1)
        base .= " [" count "]"
    if (desc != "")
        base .= " (" desc ")"
    base .= "."
    return base
}

LT_RemovalNoun_Mediastinal(fields) {
    count := fields.Get("count", 1)
    noun := "mediastinal drain"
    if (count > 1)
        noun .= "s"
    return Map("text", noun, "plural", count > 1)
}

; Combines every active mediastinal drain instance into one shared line,
; e.g. "Anterior [2] and posterior [2] inferior approach mediastinal
; drains" -- grouped by position, with inferior approach and the info note
; applied once if any instance has them checked, rather than repeating
; "mediastinal drain(s)" and the modifiers on a separate line per instance.
LT_AggregateMediastinal(instancesFields) {
    groups := Map()
    order := []
    anyInferior := false
    anyOtherNote := false
    anyIsNew := false

    for f in instancesFields {
        pos := f.Get("position", "")
        cnt := f.Get("count", 1)
        if (!groups.Has(pos)) {
            groups[pos] := 0
            order.Push(pos)
        }
        groups[pos] += cnt
        if (f.Get("inferiorApproach", false))
            anyInferior := true
        if (f.Get("otherNote", false))
            anyOtherNote := true
        if (f.Get("isNew", false))
            anyIsNew := true
    }

    canonical := ["Anterior", "Posterior", ""]
    keys := []
    for k in canonical
        if (groups.Has(k))
            keys.Push(k)
    for k in order {
        found := false
        for kk in keys {
            if (kk = k) {
                found := true
                break
            }
        }
        if (!found)
            keys.Push(k)
    }

    groupTexts := []
    totalCount := 0
    for k in keys {
        c := groups[k]
        totalCount += c
        if (k != "") {
            t := StrLower(k)
            if (c > 1)
                t .= " [" c "]"
            groupTexts.Push(t)
        } else if (c > 1) {
            groupTexts.Push("[" c "]")
        }
    }

    modifier := anyInferior ? "inferior approach " : ""
    noun := (totalCount > 1) ? "mediastinal drains" : "mediastinal drain"
    newPrefix := anyIsNew ? "New " : ""

    line := LT_Capitalize(newPrefix modifier noun)
    if (groupTexts.Length > 0)
        line .= ": " LT_JoinWithAnd(groupTexts)
    if (anyOtherNote)
        line .= " _____"
    return line
}

LT_GeneratorLeadLocations(fields) {
    locs := []
    if (fields.Get("locRA", false))
        locs.Push("right atrium")
    if (fields.Get("locRV", false))
        locs.Push("right ventricle")
    if (fields.Get("locCS", false))
        locs.Push("coronary sinus")
    if (fields.Get("locCV", false))
        locs.Push("cardiac veins")
    if (fields.Get("locPS", false))
        locs.Push("presternal")
    return locs
}

LT_Sentence_Generator(fields) {
    side := fields.Get("laterality", "")
    position := fields.Get("position", "")
    intact := fields.Get("intact", false)
    leadType := fields.Get("leadType", "")

    sideText := (side != "") ? StrLower(side) " " : ""
    posText := (position != "") ? StrLower(position) " " : ""
    subject := LT_Capitalize(sideText posText "chest wall generator")

    descParts := []
    if (intact)
        descParts.Push("intact")
    if (leadType != "")
        descParts.Push(StrLower(leadType))
    descText := ""
    for i, p in descParts
        descText .= (i > 1 ? " " : "") p
    if (descText != "")
        descText .= " "

    locs := LT_GeneratorLeadLocations(fields)
    leadWord := (locs.Length = 1) ? "lead" : "leads"
    locsText := (locs.Length = 0) ? "_____" : LT_JoinWithAnd(locs)

    return subject " projecting " descText leadWord " to the " locsText "."
}

LT_RemovalNoun_Generator(fields) {
    side := fields.Get("laterality", "")
    position := fields.Get("position", "")
    sideText := (side != "") ? StrLower(side) " " : ""
    posText := (position != "") ? StrLower(position) " " : ""
    return Map("text", sideText posText "chest wall generator", "plural", false)
}

LT_Sentence_Abdominal(fields) {
    location := fields.Get("location", "")
    deviceType := fields.Get("deviceType", "")
    count := fields.Get("count", 1)

    if (location = "Other")
        locText := "_____ "
    else if (location != "")
        locText := StrLower(location) " "
    else
        locText := "upper abdominal "

    typeNoun := (deviceType != "") ? StrLower(deviceType) : "drain"
    if (count > 1)
        typeNoun .= "s"

    base := LT_Capitalize(locText typeNoun)
    if (count > 1)
        base .= " [" count "]"
    base .= "."
    return base
}

LT_RemovalNoun_Abdominal(fields) {
    location := fields.Get("location", "")
    deviceType := fields.Get("deviceType", "")
    count := fields.Get("count", 1)
    if (location = "Other")
        locText := "_____ "
    else if (location != "")
        locText := StrLower(location) " "
    else
        locText := "upper abdominal "
    typeNoun := (deviceType != "") ? StrLower(deviceType) : "drain"
    if (count > 1)
        typeNoun .= "s"
    return Map("text", locText typeNoun, "plural", count > 1)
}

LT_Sentence_Nephrostomy(fields) {
    side := fields.Get("laterality", "")
    isPlural := (side = "Bilateral")
    sideText := (side != "") ? StrLower(side) " " : ""
    noun := "percutaneous nephrostomy tube"
    if (isPlural)
        noun .= "s"
    return LT_Capitalize(sideText noun) "."
}

LT_RemovalNoun_Nephrostomy(fields) {
    side := fields.Get("laterality", "")
    isPlural := (side = "Bilateral")
    sideText := (side != "") ? StrLower(side) " " : ""
    noun := "percutaneous nephrostomy tube"
    if (isPlural)
        noun .= "s"
    return Map("text", sideText noun, "plural", isPlural)
}

LT_Sentence_RetainedLeads(fields) {
    leadType := fields.Get("leadType", "")
    locs := LT_GeneratorLeadLocations(fields)
    leadWord := (locs.Length = 1) ? "lead" : "leads"
    locsText := (locs.Length = 0) ? "_____" : LT_JoinWithAnd(locs)
    typeText := (leadType != "") ? StrLower(leadType) " " : ""
    return "Retained " typeText leadWord " to the " locsText "."
}

LT_Sentence_LVAD(fields) {
    model := fields.Get("model", "")
    if (model = "")
        return "Left ventricular assist device (LVAD)."
    return model " left ventricular assist device (LVAD)."
}

; Each entry is Map("text", "...", "plural", true/false) -- plural reflects
; the noun's own grammatical number (e.g. "pads" is always plural, a
; multi-count chest tube entry is plural when count > 1), independent of
; how many devices are being grouped together.
LT_JoinRemovalSentence(nouns) {
    n := nouns.Length
    texts := []
    anyPlural := false
    for item in nouns {
        texts.Push(item["text"])
        if (item["plural"])
            anyPlural := true
    }
    joined := LT_JoinWithAnd(texts)
    verb := (n > 1 || anyPlural) ? "are" : "is"
    return "Prior " joined " " verb " now absent."
}

; Every device instance has an "Add other/unlisted finding" toggle,
; regardless of what fields it otherwise has (even zero-field devices like
; the loop/Zio monitor). When checked, tack on a blank placeholder for
; dictating whatever else needs to be said about it.
LT_AppendOtherNote(s, fields) {
    if (!fields.Get("otherNote", false))
        return s
    if (SubStr(s, -1) = ".")
        return SubStr(s, 1, StrLen(s) - 1) " _____."
    return s " _____"
}

; Every device instance also has a universal "New" toggle, same idea as
; Now-absent/Add-info -- applies to any device, not just IJ/SCV (which
; used to have their own dedicated New checkbox; removed in favor of this,
; see LT_Sentence_CentralLine).
LT_PrependNewModifier(s, fields) {
    if (!fields.Get("isNew", false) || s = "")
        return s
    return "New " LT_LowercaseFirst(s)
}

; No terminal periods on any output line -- strip one if a sentence
; function (or the removal-grouping sentence) happens to end with one.
LT_StripTrailingPeriod(s) {
    if (SubStr(s, -1) = ".")
        return SubStr(s, 1, StrLen(s) - 1)
    return s
}

LT_BuildActiveLines() {
    global LT_DeviceOrder, LT_InstanceOrder, LT_Instances, LT_DeviceDefs

    lines := []

    for deviceKey in LT_DeviceOrder {
        def := LT_DeviceDefs[deviceKey]
        isAggregate := def.Get("aggregate", false)
        activeFieldsList := []

        for instId in LT_InstanceOrder {
            inst := LT_Instances[instId]
            if (inst["deviceKey"] != deviceKey)
                continue
            fields := inst["fields"]

            if (fields.Get("removed", false)) {
                continue
            } else if (isAggregate) {
                activeFieldsList.Push(fields)
            } else {
                sentenceFn := def["sentenceFn"]
                s := sentenceFn(fields)
                if (s != "") {
                    s := LT_AppendOtherNote(s, fields)
                    s := LT_PrependNewModifier(s, fields)
                    s := LT_StripTrailingPeriod(s)
                    lines.Push(s)
                }
            }
        }

        if (isAggregate && activeFieldsList.Length > 0) {
            aggFn := def["aggregateFn"]
            s := aggFn(activeFieldsList)
            if (s != "")
                lines.Push(LT_StripTrailingPeriod(s))
        }
    }

    return lines
}

; "" if nothing is marked removed, otherwise the one grouped "X, Y are now
; absent" sentence.
LT_BuildRemovalLine() {
    global LT_DeviceOrder, LT_InstanceOrder, LT_Instances, LT_DeviceDefs

    removalNouns := []
    for deviceKey in LT_DeviceOrder {
        def := LT_DeviceDefs[deviceKey]
        for instId in LT_InstanceOrder {
            inst := LT_Instances[instId]
            if (inst["deviceKey"] != deviceKey)
                continue
            fields := inst["fields"]
            if (fields.Get("removed", false))
                removalNouns.Push(def["removalNoun"](fields))
        }
    }

    if (removalNouns.Length = 0)
        return ""
    return LT_StripTrailingPeriod(LT_JoinRemovalSentence(removalNouns))
}

; Any full term + parenthetical acronym pattern that appears in actual
; report text (not just device-picker labels) -- add here to get the
; "spell it out once, acronym after that" treatment automatically anywhere
; it's used.
LT_AcronymPairs := [
    ["peripherally inserted central catheter (PICC)", "PICC"],
    ["venous extracorporeal membrane oxygenation (ECMO) cannula", "ECMO cannula"],
    ["left ventricular assist device (LVAD)", "LVAD"]
]

; Scans a set of lines (in output order) for each known full-term/acronym
; pair; every occurrence after the first gets collapsed to just the
; acronym, preserving whether that particular occurrence was capitalized
; (sentence-initial) or not. Matching is case-insensitive (AHK's default),
; so this catches both "Peripherally..." and "peripherally..." uniformly.
LT_CollapseRepeatedAcronyms(lines) {
    for pair in LT_AcronymPairs {
        full := pair[1]
        short := pair[2]
        seenCount := 0
        for i, line in lines {
            pos := InStr(line, full)
            if (!pos)
                continue
            seenCount += 1
            if (seenCount = 1)
                continue  ; first mention keeps the full spelled-out form
            matchedText := SubStr(line, pos, StrLen(full))
            firstChar := SubStr(matchedText, 1, 1)
            replacement := (firstChar = StrUpper(firstChar)) ? LT_Capitalize(short) : short
            lines[i] := SubStr(line, 1, pos - 1) replacement SubStr(line, pos + StrLen(full))
        }
    }
    return lines
}

LT_BuildLines() {
    lines := LT_BuildActiveLines()
    removalLine := LT_BuildRemovalLine()
    if (removalLine != "")
        lines.Push(removalLine)
    return LT_CollapseRepeatedAcronyms(lines)
}

; Plain-text clipboard format: newline-joined, with a leading blank line so
; pasting somewhere that only takes plain text can still start its own list.
LT_JoinLinesPlain(lines) {
    out := ""
    for line in lines
        out .= line "`r`n"
    out := RTrim(out, "`r`n")
    return "`r`n" out
}

LT_BuildOutput() {
    return LT_JoinLinesPlain(LT_BuildLines())
}


; ============================================================================
; INSTANCE MANAGEMENT
; ============================================================================

LT_AddInstance(deviceKey) {
    global LT_Instances, LT_InstanceOrder, LT_InstanceCounter, LT_GuiObj

    ; From the moment a device is added, keep the window pinned above other
    ; windows (e.g. the PACS viewer) so it's usable while looking at images.
    ; Cleared on Close or New Patient -- see those handlers.
    if (IsObject(LT_GuiObj))
        try LT_GuiObj.Opt("+AlwaysOnTop")

    ; collapse everything currently in the list so the newly added device
    ; is the only one expanded
    for id in LT_InstanceOrder
        LT_Instances[id]["collapsed"] := true

    LT_InstanceCounter += 1
    instId := deviceKey "_" LT_InstanceCounter
    LT_Instances[instId] := Map("deviceKey", deviceKey, "fields", Map(), "collapsed", false)
    LT_InstanceOrder.Push(instId)
    LT_RebuildMiddleColumn(instId)
}

LT_RemoveInstance(instId, *) {
    global LT_Instances, LT_InstanceOrder
    LT_Instances.Delete(instId)
    for i, id in LT_InstanceOrder {
        if (id = instId) {
            LT_InstanceOrder.RemoveAt(i)
            break
        }
    }
    LT_RebuildMiddleColumn()
}

; Shift+Delete: removes whichever instance is currently expanded (the same
; notion of "what you're working on" the image popup already tracks).
LT_HotkeyRemoveSelected(*) {
    instId := LT_FindExpandedInstance()
    if (instId != "")
        LT_RemoveInstance(instId)
}

LT_ClearAll(*) {
    global LT_Instances, LT_InstanceOrder, LT_InstanceCounter
    LT_Instances := Map()
    LT_InstanceOrder := []
    LT_InstanceCounter := 0
    LT_RebuildMiddleColumn()
}

; Clears all state AND tears down/rebuilds the window itself, preserving its
; current size and position. Use this between patients — it's also the
; cleanest way to reclaim the hidden controls that pile up over a long
; session, since a fresh window has none.
LT_NewPatient(*) {
    global LT_GuiObj, LT_MidPanel, LT_Instances, LT_InstanceOrder, LT_InstanceCounter
    global LT_MiddleControls, LT_FieldControls, LT_ScrollOffset

    gx := 0, gy := 0, gw := 0, gh := 0
    if (IsObject(LT_GuiObj)) {
        try LT_GuiObj.GetPos(&gx, &gy, &gw, &gh)
        try LT_GuiObj.Destroy()  ; also destroys LT_MidPanel, since it's a child of this window
    }
    LT_GuiObj := ""
    LT_MidPanel := ""
    LT_MiddleControls := []
    LT_FieldControls := Map()
    LT_ScrollOffset := 0

    LT_Instances := Map()
    LT_InstanceOrder := []
    LT_InstanceCounter := 0

    LT_EnsureGui()

    if (gw > 0)
        LT_GuiObj.Move(gx, gy, gw, gh)

    LT_RebuildMiddleColumn()
}

LT_InstanceDisplayLabel(instId) {
    global LT_Instances, LT_InstanceOrder, LT_DeviceDefs
    inst := LT_Instances[instId]
    deviceKey := inst["deviceKey"]
    def := LT_DeviceDefs[deviceKey]

    sameKeyIds := []
    for id in LT_InstanceOrder {
        if (LT_Instances[id]["deviceKey"] = deviceKey)
            sameKeyIds.Push(id)
    }
    if (sameKeyIds.Length <= 1)
        return def.Get("shortLabel", def["label"])

    for i, id in sameKeyIds {
        if (id = instId)
            return def.Get("shortLabel", def["label"]) " #" i
    }
    return def.Get("shortLabel", def["label"])
}


; ============================================================================
; FIELD EVENT HANDLERS
; ============================================================================

LT_FieldSet(instId, fieldId, val, *) {
    global LT_Instances, LT_DeviceDefs
    inst := LT_Instances[instId]
    def := LT_DeviceDefs[inst["deviceKey"]]
    fields := inst["fields"]
    cur := fields.Get(fieldId, "")
    fields[fieldId] := (cur = val) ? "" : val

    if (LT_FieldHasDependents(def, fieldId) || fieldId = "laterality") {
        LT_RebuildMiddleColumn(instId)
    } else {
        fieldDef := LT_FindFieldDef(def, fieldId)
        LT_RefreshFieldControls(instId, fieldId, fieldDef, fields)
        LT_UpdateOutputBoxOnly()
    }
}

LT_FieldToggle(instId, fieldId, *) {
    global LT_Instances
    inst := LT_Instances[instId]
    fields := inst["fields"]
    cur := fields.Get(fieldId, false)
    fields[fieldId] := !cur

    if (fieldId = "removed")
        LT_RebuildMiddleColumn(instId)
    else
        LT_UpdateOutputBoxOnly()
}

LT_CounterAdjust(instId, fieldId, delta, minV, maxV, *) {
    global LT_Instances, LT_FieldControls
    inst := LT_Instances[instId]
    fields := inst["fields"]
    cur := fields.Get(fieldId, 1)
    newVal := cur + delta
    if (newVal < minV)
        newVal := minV
    if (newVal > maxV)
        newVal := maxV
    fields[fieldId] := newVal

    key := instId "|" fieldId
    if (LT_FieldControls.Has(key)) {
        for entry in LT_FieldControls[key] {
            if (entry["meta"]["kind"] = "counterDisplay")
                try entry["ctrl"].Text := String(newVal)
        }
    }
    LT_UpdateOutputBoxOnly()
}

LT_RadioSet(instId, fieldId, optLabel, optPhrase, *) {
    global LT_Instances, LT_DeviceDefs
    inst := LT_Instances[instId]
    def := LT_DeviceDefs[inst["deviceKey"]]
    fields := inst["fields"]
    cur := fields.Get(fieldId, "")

    if (cur = optLabel) {
        fields[fieldId] := ""
        fields[fieldId "_phrase"] := ""
    } else {
        fields[fieldId] := optLabel
        fields[fieldId "_phrase"] := optPhrase
    }

    if (LT_FieldHasDependents(def, fieldId)) {
        LT_RebuildMiddleColumn(instId)
    } else {
        fieldDef := LT_FindFieldDef(def, fieldId)
        LT_RefreshFieldControls(instId, fieldId, fieldDef, fields)
        LT_UpdateOutputBoxOnly()
    }
}

; Clicking a family button either activates that family (defaulting to its
; first state) or, if it's already active, clears the field. Since this
; changes which sub-state buttons are visible, it always does a full
; middle-column refresh.
LT_GroupFamilyClick(instId, fieldId, groupIndex, *) {
    global LT_Instances, LT_DeviceDefs
    inst := LT_Instances[instId]
    def := LT_DeviceDefs[inst["deviceKey"]]
    fields := inst["fields"]
    fieldDef := LT_FindFieldDef(def, fieldId)
    grp := fieldDef["groups"][groupIndex]

    activeFamily := fields.Get(fieldId "_activeFamily", "")
    isActive := (activeFamily = grp["groupLabel"])

    if (isActive) {
        fields[fieldId] := ""
        fields[fieldId "_phrase"] := ""
        fields[fieldId "_activeFamily"] := ""
    } else if (grp.Get("requiresSubSelection", false)) {
        ; Mark this family as active (so its sub-state row reveals) without
        ; picking a default -- stays incomplete/red until a specific state
        ; is clicked.
        fields[fieldId] := ""
        fields[fieldId "_phrase"] := ""
        fields[fieldId "_activeFamily"] := grp["groupLabel"]
    } else {
        st1 := grp["states"][1]
        fields[fieldId] := st1["label"]
        fields[fieldId "_phrase"] := st1["phrase"]
        fields[fieldId "_activeFamily"] := grp["groupLabel"]
    }

    LT_RebuildMiddleColumn(instId)
}

; Clicking a sub-state button within the already-active family just swaps
; which state is selected; the button layout doesn't change, so this can be
; refreshed in place.
LT_GroupStateClick(instId, fieldId, stateLabel, statePhrase, *) {
    global LT_Instances, LT_DeviceDefs
    inst := LT_Instances[instId]
    def := LT_DeviceDefs[inst["deviceKey"]]
    fields := inst["fields"]
    fields[fieldId] := stateLabel
    fields[fieldId "_phrase"] := statePhrase

    fieldDef := LT_FindFieldDef(def, fieldId)
    LT_RefreshFieldControls(instId, fieldId, fieldDef, fields)
    LT_UpdateOutputBoxOnly()
}

LT_ToggleCollapse(instId, *) {
    global LT_Instances
    inst := LT_Instances[instId]
    inst["collapsed"] := !inst.Get("collapsed", false)
    LT_RebuildMiddleColumn(instId)
}

; Escapes a plain-text line for safe inclusion in RTF: backslash and braces
; are RTF-significant, and anything outside ASCII needs \uNNNN? escaping.
LT_RtfEscape(s) {
    out := ""
    Loop Parse, s {
        c := A_LoopField
        code := Ord(c)
        if (c = "\")
            out .= "\\"
        else if (c = "{")
            out .= "\{"
        else if (c = "}")
            out .= "\}"
        else if (code > 127)
            out .= "\u" code "?"
        else
            out .= c
    }
    return out
}

; Builds each line as its own bulleted paragraph using RTF's \bullet control
; word -- a semantic instruction telling the reader to draw its own bullet
; glyph, not a literal typed character -- with a hanging indent so wrapped
; continuation lines still align under the text rather than the bullet.
; A single line doesn't need list machinery -- no leading blank paragraph
; (that's only there to let PowerScribe's own list continue from it), no
; \pn bullet definition, just the plain text.
LT_BuildPlainRTF(lines) {
    header := "
    (LTrim
    {\rtf1\ansi\ansicpg1252\deff0\deflang1033
    {\fonttbl{\f0\fswiss\fcharset0 Calibri;}}
    {\*\generator LinesAndTubes;}
    \viewkind4\uc1\pard\f0\fs22
    )"
    rtf := header
    for i, line in lines {
        if (i > 1)
            rtf .= "\par "
        rtf .= LT_RtfEscape(line)
    }
    rtf .= "}"
    return rtf
}

LT_BuildBulletRTF(lines) {
    ; Matched directly against real RTF captured from PowerScribe's own
    ; editor (Riched20), not the Word-style {\listtable}/{\listoverridetable}
    ; mechanism used earlier -- RichEdit uses the older, simpler \pn
    ; destination group instead. It's defined once, on the first bullet
    ; paragraph; every later bullet -- including the last one -- just
    ; repeats the {\pntext...} fallback marker and gets its own trailing
    ; \par, exactly like PowerScribe's own captured reference does. (An
    ; earlier version dropped every line's own \par to fix a trailing
    ; blank-line bug; that bug was actually the separate \pard\par reset
    ; that used to follow the loop, not the per-line \par itself.)
    header := "
    (LTrim
    {\rtf1\ansi\ansicpg1252\deff0\deflang1033
    {\fonttbl{\f0\fswiss\fcharset0 Calibri;}{\f1\fnil\fcharset2 Symbol;}}
    {\*\generator LinesAndTubes;}
    \viewkind4\uc1\pard\f0\fs22
    \pard\li0\fi0\par
    )"

    rtf := header
    for i, line in lines {
        if (i = 1)
            rtf .= "\pard{\pntext\f1\'B7\tab}{\*\pn\pnlvlblt\pnf1\pnindent0{\pntxtb\'B7}}\fi-360\li360 " LT_RtfEscape(line) "\par"
        else
            rtf .= "{\pntext\f1\'B7\tab}" LT_RtfEscape(line) "\par"
    }
    rtf .= "}"
    return rtf
}

; Same technique as TN_SetClipboardTextAndRTF / CTAP_SetClipboardTextAndRTF /
; RTM_SetClipboardTextAndRTF elsewhere in RAHKET: writes both CF_UNICODETEXT
; and the registered "Rich Text Format" clipboard format in one pass, so a
; paste target uses whichever it understands (falls back to plain text if it
; doesn't know RTF).
LT_SetClipboardTextAndRTF(plain, rtf) {
    if !DllCall("OpenClipboard", "ptr", 0, "int") {
        MsgBox "Could not open clipboard."
        return
    }

    DllCall("EmptyClipboard")

    lenW := (StrLen(plain) + 1) * 2
    hText := DllCall("GlobalAlloc", "uint", 0x2, "uptr", lenW, "ptr")
    if (hText) {
        pText := DllCall("GlobalLock", "ptr", hText, "ptr")
        StrPut(plain, pText, "UTF-16")
        DllCall("GlobalUnlock", "ptr", hText)
        DllCall("SetClipboardData", "uint", 13, "ptr", hText)  ; CF_UNICODETEXT
    }

    cfRtf := DllCall("RegisterClipboardFormat", "str", "Rich Text Format", "uint")
    lenA := StrLen(rtf) + 1
    hRtf := DllCall("GlobalAlloc", "uint", 0x2, "uptr", lenA, "ptr")
    if (hRtf) {
        pRtf := DllCall("GlobalLock", "ptr", hRtf, "ptr")
        StrPut(rtf, pRtf, "CP0")
        DllCall("GlobalUnlock", "ptr", hRtf)
        DllCall("SetClipboardData", "uint", cfRtf, "ptr", hRtf)
    }

    DllCall("CloseClipboard")
}

LT_CopyOutput(*) {
    if (LT_AnyInstanceNeedsSideWarning()) {
        MsgBox("Please complete the selection for all red devices.", "Incomplete selection", 48)
        return
    }
    activeLines := LT_BuildActiveLines()
    removalLine := LT_BuildRemovalLine()

    if (activeLines.Length <= 1) {
        lines := activeLines.Clone()
        if (removalLine != "")
            lines.Push(removalLine)
        lines := LT_CollapseRepeatedAcronyms(lines)
        plain := ""
        for i, line in lines
            plain .= (i > 1 ? "`r`n" : "") line
        rtf := LT_BuildPlainRTF(lines)
    } else {
        lines := LT_BuildLines()
        plain := LT_JoinLinesPlain(lines)
        rtf := LT_BuildBulletRTF(lines)
    }
    LT_SetClipboardTextAndRTF(plain, rtf)
    ToolTip("Copied to clipboard")
    SetTimer(() => ToolTip(), -1000)
}

LT_AddDeviceFromListbox(*) {
    global LT_GuiObj, LT_PickListDeviceKeys
    lb := LT_GuiObj["LT_DeviceListBox"]
    idx := lb.Value
    if (idx = 0)
        return
    key := (idx >= 1 && idx <= LT_PickListDeviceKeys.Length) ? LT_PickListDeviceKeys[idx] : ""
    if (key = "")
        return  ; a category header row, not a real device
    LT_AddInstance(key)
}


; ============================================================================
; GUI RENDERING
; ============================================================================

LT_UpdateOutputBoxOnly() {
    LT_RenderOutputPane()
}

; Rebuilds the right pane as one Text control per report line so each line
; can be clicked to jump back to its device in the middle column. Hides the
; previous set of line controls rather than destroying the window.
LT_RenderOutputPane() {
    global LT_GuiObj, LT_RightControls, LT_RightX, LT_RightW
    global LT_DeviceOrder, LT_InstanceOrder, LT_Instances, LT_DeviceDefs

    for ctrl in LT_RightControls {
        try ctrl.Visible := false
    }
    LT_RightControls := []

    y := 35
    removalNouns := []

    for deviceKey in LT_DeviceOrder {
        def := LT_DeviceDefs[deviceKey]
        isAggregate := def.Get("aggregate", false)
        activeFieldsList := []
        anyNeedsSide := false

        for instId in LT_InstanceOrder {
            inst := LT_Instances[instId]
            if (inst["deviceKey"] != deviceKey)
                continue
            fields := inst["fields"]

            if (fields.Get("removed", false)) {
                removalNouns.Push(def["removalNoun"](fields))
                continue
            }

            if (isAggregate) {
                activeFieldsList.Push(fields)
                if (LT_FieldsNeedSideWarning(fields, def))
                    anyNeedsSide := true
                continue
            }

            sentenceFn := def["sentenceFn"]
            s := sentenceFn(fields)
            if (s = "")
                continue
            s := LT_AppendOtherNote(s, fields)
            s := LT_PrependNewModifier(s, fields)
            s := LT_StripTrailingPeriod(s)

            approxRows := Ceil(StrLen(s) / 50)
            h := Max(20, approxRows * 16 + 6)
            lineColor := LT_InstanceNeedsSideWarning(instId) ? "cRed" : "cBlue"
            txt := LT_GuiObj.Add("Text", "x" LT_RightX " y" y " w" LT_RightW " h" h " " lineColor, s)
            txt.OnEvent("Click", LT_OpenInstanceFromOutput.Bind(instId))
            LT_RightControls.Push(txt)
            y += h + 8
        }

        if (isAggregate && activeFieldsList.Length > 0) {
            aggFn := def["aggregateFn"]
            s := LT_StripTrailingPeriod(aggFn(activeFieldsList))
            if (s != "") {
                approxRows := Ceil(StrLen(s) / 50)
                h := Max(20, approxRows * 16 + 6)
                lineColor := anyNeedsSide ? "cRed" : "cBlue"
                txt := LT_GuiObj.Add("Text", "x" LT_RightX " y" y " w" LT_RightW " h" h " " lineColor, s)
                txt.OnEvent("Click", LT_OpenDeviceFromOutput.Bind(deviceKey))
                LT_RightControls.Push(txt)
                y += h + 8
            }
        }
    }

    if (removalNouns.Length > 0) {
        s := LT_StripTrailingPeriod(LT_JoinRemovalSentence(removalNouns))
        approxRows := Ceil(StrLen(s) / 50)
        h := Max(20, approxRows * 16 + 6)
        txt := LT_GuiObj.Add("Text", "x" LT_RightX " y" y " w" LT_RightW " h" h, s)
        LT_RightControls.Push(txt)
    }
}

; Clicking a line in the report pane expands that device (collapsing the
; rest) back in the middle column.
LT_OpenInstanceFromOutput(instId, *) {
    global LT_Instances, LT_InstanceOrder
    for id in LT_InstanceOrder
        LT_Instances[id]["collapsed"] := (id != instId)
    LT_RebuildMiddleColumn(instId)
}

; Clicking an aggregated line (mediastinal drains, chest tubes) expands
; every instance of that device, since the line represents all of them.
LT_OpenDeviceFromOutput(deviceKey, *) {
    global LT_Instances, LT_InstanceOrder
    firstId := ""
    for id in LT_InstanceOrder {
        inst := LT_Instances[id]
        belongs := (inst["deviceKey"] = deviceKey)
        inst["collapsed"] := !belongs
        if (belongs && firstId = "")
            firstId := id
    }
    LT_RebuildMiddleColumn(firstId)
}

; Builds the picker's display list with a header row before each category
; and tracks, per row, which device key (if any) it corresponds to -- see
; LT_PickListDeviceKeys, used by LT_AddDeviceFromListbox to skip headers.
LT_DeviceLabelList() {
    global LT_DeviceCategories, LT_DeviceDefs, LT_PickListDeviceKeys
    labels := []
    LT_PickListDeviceKeys := []
    for cat in LT_DeviceCategories {
        labels.Push("── " cat["name"] " ──")
        LT_PickListDeviceKeys.Push("")
        for key in cat["keys"] {
            def := LT_DeviceDefs[key]
            labels.Push(def.Get("shortLabel", def["label"]))
            LT_PickListDeviceKeys.Push(key)
        }
    }
    return labels
}

LT_FieldHidden(fieldDef, fields) {
    if (!fieldDef.Has("showIf"))
        return false
    cond := fieldDef["showIf"]
    curVal := fields.Get(cond["field"], "")
    return curVal != cond["value"]
}

; Creates the middle-column child window ONCE: a plain top-level popup that
; gets surgically converted to a WS_CHILD of the main window and given a
; native WS_VSCROLL scrollbar. Never destroyed/reparented again after this
; -- only its contents get hidden/replaced (see LT_RebuildMiddleColumn),
; the same safe pattern the rest of this module already uses. Being a real
; child window means Windows clips its content and repositions its children
; on scroll natively; we don't hand-roll any of that.
LT_CreateMidPanel() {
    global LT_GuiObj, LT_MidPanelX, LT_MidPanelY, LT_MidPanelW, LT_MidPanelH
    global LT_MidPanelClientW, LT_MidPanelClientH

    panel := Gui("-Caption +ToolWindow -DPIScale", "")
    panel.BackColor := "F0F0F0"
    panel.Show("Hide x0 y0 w" LT_MidPanelW " h" LT_MidPanelH)

    hwnd := panel.Hwnd
    GWL_STYLE := -16
    WS_POPUP := 0x80000000
    WS_CAPTION := 0x00C00000
    WS_CHILD := 0x40000000
    WS_VSCROLL := 0x00200000
    WS_CLIPSIBLINGS := 0x04000000

    style := DllCall("GetWindowLong", "ptr", hwnd, "int", GWL_STYLE, "int")
    style := (style & ~WS_POPUP & ~WS_CAPTION) | WS_CHILD | WS_VSCROLL | WS_CLIPSIBLINGS
    DllCall("SetWindowLong", "ptr", hwnd, "int", GWL_STYLE, "int", style)
    DllCall("SetParent", "ptr", hwnd, "ptr", LT_GuiObj.Hwnd)
    DllCall("MoveWindow", "ptr", hwnd, "int", LT_MidPanelX, "int", LT_MidPanelY, "int", LT_MidPanelW, "int", LT_MidPanelH, "int", 1)
    DllCall("ShowWindow", "ptr", hwnd, "int", 5)  ; SW_SHOW

    rect := Buffer(16, 0)
    DllCall("GetClientRect", "ptr", hwnd, "ptr", rect)
    LT_MidPanelClientW := NumGet(rect, 8, "int")
    LT_MidPanelClientH := NumGet(rect, 12, "int")

    return panel
}

; Moves/resizes the panel (e.g. on a window resize) without touching its
; WS_CHILD/reparented state, and re-queries its usable client size.
LT_MoveMidPanel(x, y, w, h) {
    global LT_MidPanel, LT_MidPanelX, LT_MidPanelY, LT_MidPanelW, LT_MidPanelH
    global LT_MidPanelClientW, LT_MidPanelClientH

    LT_MidPanelX := x
    LT_MidPanelY := y
    LT_MidPanelW := w
    LT_MidPanelH := h

    if (!IsObject(LT_MidPanel))
        return

    try DllCall("MoveWindow", "ptr", LT_MidPanel.Hwnd, "int", x, "int", y, "int", w, "int", h, "int", 1)

    rect := Buffer(16, 0)
    DllCall("GetClientRect", "ptr", LT_MidPanel.Hwnd, "ptr", rect)
    LT_MidPanelClientW := NumGet(rect, 8, "int")
    LT_MidPanelClientH := NumGet(rect, 12, "int")
}

; Adds a control into the middle-column panel and tracks it so a later
; rebuild can hide it instead of destroying anything.
LT_AddMid(g, ctrlType, opts, text) {
    global LT_MiddleControls
    ctrl := g.Add(ctrlType, opts, text)
    LT_MiddleControls.Push(ctrl)
    return ctrl
}

; Pushes the scrollbar's range/page/position to match current content.
LT_UpdateScrollInfo() {
    global LT_MidPanel, LT_MidContentBottom, LT_MidPanelClientH, LT_ScrollOffset
    if (!IsObject(LT_MidPanel))
        return
    SB_VERT := 1
    SIF_ALL := 0x17
    si := Buffer(28, 0)
    NumPut("uint", 28, si, 0)
    NumPut("uint", SIF_ALL, si, 4)
    NumPut("int", 0, si, 8)                                     ; nMin
    NumPut("int", Max(LT_MidContentBottom, LT_MidPanelClientH), si, 12) ; nMax
    NumPut("uint", LT_MidPanelClientH, si, 16)                   ; nPage
    NumPut("int", LT_ScrollOffset, si, 20)                       ; nPos
    DllCall("SetScrollInfo", "ptr", LT_MidPanel.Hwnd, "int", SB_VERT, "ptr", si, "int", 1)
}

; Scrolls the panel to an absolute position. Because this panel's whole
; client area IS the scrollable content (no restricted rect needed, unlike
; the earlier attempt that had to share one window with two other columns),
; this is the simple, fully-supported form of ScrollWindow: passing NULL
; rects moves every child control automatically and is the faster path per
; the Win32 docs.
LT_ScrollTo(newPos) {
    global LT_MidPanel, LT_ScrollOffset, LT_MidContentBottom, LT_MidPanelClientH
    if (!IsObject(LT_MidPanel))
        return
    maxScroll := Max(0, LT_MidContentBottom - LT_MidPanelClientH)
    if (newPos < 0)
        newPos := 0
    if (newPos > maxScroll)
        newPos := maxScroll
    if (newPos = LT_ScrollOffset)
        return
    dy := LT_ScrollOffset - newPos
    try DllCall("ScrollWindow", "ptr", LT_MidPanel.Hwnd, "int", 0, "int", dy, "ptr", 0, "ptr", 0)
    LT_ScrollOffset := newPos
    LT_UpdateScrollInfo()
}

; Native scrollbar drag/click handling for the middle-column panel.
LT_OnVScroll(wParam, lParam, msg, hwnd) {
    global LT_MidPanel, LT_ScrollOffset

    if (!IsObject(LT_MidPanel) || hwnd != LT_MidPanel.Hwnd)
        return

    SB_VERT := 1
    SIF_ALL := 0x17
    si := Buffer(28, 0)
    NumPut("uint", 28, si, 0)
    NumPut("uint", SIF_ALL, si, 4)
    DllCall("GetScrollInfo", "ptr", hwnd, "int", SB_VERT, "ptr", si)
    minP := NumGet(si, 8, "int")
    maxP := NumGet(si, 12, "int")
    pageSz := NumGet(si, 16, "uint")
    trackPos := NumGet(si, 24, "int")

    action := wParam & 0xFFFF
    newPos := LT_ScrollOffset
    if (action = 0)        ; SB_LINEUP
        newPos -= 30
    else if (action = 1)   ; SB_LINEDOWN
        newPos += 30
    else if (action = 2)   ; SB_PAGEUP
        newPos -= pageSz
    else if (action = 3)   ; SB_PAGEDOWN
        newPos += pageSz
    else if (action = 4 || action = 5)  ; SB_THUMBTRACK / SB_THUMBPOSITION
        newPos := trackPos
    else if (action = 6)   ; SB_TOP
        newPos := minP
    else if (action = 7)   ; SB_BOTTOM
        newPos := maxP
    else
        return

    LT_ScrollTo(newPos)
}

; Mouse wheel scrolling: hit-test in screen coordinates against the panel's
; actual window rect.
LT_OnMouseWheel(wParam, lParam, msg, hwnd) {
    global LT_MidPanel, LT_ScrollOffset

    if (!IsObject(LT_MidPanel))
        return

    sx := lParam & 0xFFFF
    sy := (lParam >> 16) & 0xFFFF

    prect := Buffer(16, 0)
    DllCall("GetWindowRect", "ptr", LT_MidPanel.Hwnd, "ptr", prect)
    left := NumGet(prect, 0, "int")
    top := NumGet(prect, 4, "int")
    right := NumGet(prect, 8, "int")
    bottom := NumGet(prect, 12, "int")
    if (sx < left || sx > right || sy < top || sy > bottom)
        return

    raw := wParam & 0xFFFFFFFF
    deltaRaw := (raw >> 16) & 0xFFFF
    delta := (deltaRaw > 32767) ? (deltaRaw - 65536) : deltaRaw

    LT_ScrollTo(LT_ScrollOffset - (delta / 120) * 40)
}

LT_RegisterFieldControl(instId, fieldId, ctrl, meta) {
    global LT_FieldControls
    key := instId "|" fieldId
    if (!LT_FieldControls.Has(key))
        LT_FieldControls[key] := []
    LT_FieldControls[key].Push(Map("ctrl", ctrl, "meta", meta))
}

; Updates already-rendered button text for one field without touching layout.
LT_RefreshFieldControls(instId, fieldId, fieldDef, fields) {
    global LT_FieldControls
    key := instId "|" fieldId
    if (!LT_FieldControls.Has(key))
        return

    curVal := fields.Get(fieldId, "")
    type := fieldDef["type"]

    for entry in LT_FieldControls[key] {
        ctrl := entry["ctrl"]

        if (type = "grouped") {
            kind := entry["meta"]["kind"]
            if (kind = "family") {
                grp := fieldDef["groups"][entry["meta"]["groupIndex"]]
                famLabel := (grp["states"].Length = 1) ? grp["states"][1]["label"] : grp["groupLabel"]
                isActive := (fields.Get(fieldId "_activeFamily", "") = grp["groupLabel"])
                try ctrl.Text := (isActive ? "> " : "") famLabel
            } else if (kind = "state") {
                lbl := entry["meta"]["label"]
                short := entry["meta"].Get("short", lbl)
                selected := (curVal = lbl)
                try ctrl.Text := (selected ? "v " : "") short
            }

        } else if (type = "buttons" || type = "radio") {
            optVal := entry["meta"]
            optLabel := (optVal = "") ? "(none)" : optVal
            selected := (curVal = optVal)
            try ctrl.Text := (selected ? "> " : "") optLabel
        }
    }
}

; Buttons are sized once at creation for the longest text they could ever
; show (including the "> " selected-prefix), since a later click only
; updates .Text in place and never resizes the control.
LT_BTN_CHAR_W := 7
LT_BTN_PAD    := 26

LT_TextButtonWidth(text) {
    global LT_BTN_CHAR_W, LT_BTN_PAD
    return LT_BTN_PAD + StrLen(text) * LT_BTN_CHAR_W
}

LT_RenderField(g, x, y, w, instId, fieldDef, fields) {
    type := fieldDef["type"]
    id := fieldDef["id"]
    curVal := fields.Get(id, "")

    if (type = "buttons") {
        labelColor := (id = "laterality" && curVal = "") ? " cRed" : ""
        LT_AddMid(g, "Text", "x" x " y" (y + 3) " w90" labelColor, fieldDef["label"] ":")
        bx := x + 95
        for opt in fieldDef["options"] {
            optLabel := (opt = "") ? "(none)" : opt
            btnW := LT_TextButtonWidth("> " optLabel)
            selected := (curVal = opt)
            if (bx + btnW > x + w) {
                bx := x + 95
                y += 26
            }
            btn := LT_AddMid(g, "Button", "x" bx " y" y " w" btnW " h22", (selected ? "> " : "") optLabel)
            btn.OnEvent("Click", LT_FieldSet.Bind(instId, id, opt))
            LT_RegisterFieldControl(instId, id, btn, opt)
            bx += btnW + 4
        }
        return y + 28

    } else if (type = "radio") {
        LT_AddMid(g, "Text", "x" x " y" y " w" (w - 20), fieldDef["label"] ":")
        y += 20
        bx := x
        for opt in fieldDef["options"] {
            optLabel := opt["label"]
            optPhrase := opt["phrase"]
            btnW := LT_TextButtonWidth("> " optLabel)
            if (btnW > w)
                btnW := w
            selected := (curVal = optLabel)
            if (bx + btnW > x + w) {
                bx := x
                y += 24
            }
            btn := LT_AddMid(g, "Button", "x" bx " y" y " w" btnW " h22", (selected ? "> " : "") optLabel)
            btn.OnEvent("Click", LT_RadioSet.Bind(instId, id, optLabel, optPhrase))
            LT_RegisterFieldControl(instId, id, btn, optLabel)
            bx += btnW + 4
        }
        return y + 26

    } else if (type = "grouped") {
        LT_AddMid(g, "Text", "x" x " y" y " w" (w - 20), fieldDef["label"] ":")
        y += 20

        activeFamily := fields.Get(id "_activeFamily", "")

        bx := x
        justRevealed := false
        for gi, grp in fieldDef["groups"] {
            isSelected := (activeFamily = grp["groupLabel"])
            incomplete := (isSelected && grp.Get("requiresSubSelection", false) && curVal = "")

            famLabel := (grp["states"].Length = 1) ? grp["states"][1]["label"] : grp["groupLabel"]
            btnW := LT_TextButtonWidth("> " famLabel)
            if (bx + btnW > x + w) {
                bx := x
                y += 24
            }
            famOpts := "x" bx " y" y " w" btnW " h22" (incomplete ? " cRed" : "")
            btn := LT_AddMid(g, "Button", famOpts, (isSelected ? "> " : "") famLabel)
            btn.OnEvent("Click", LT_GroupFamilyClick.Bind(instId, id, gi))
            LT_RegisterFieldControl(instId, id, btn, Map("kind", "family", "groupIndex", gi))
            bx += btnW + 4

            ; reveal this family's specific states directly beneath its own
            ; button, not after the whole family list
            if (isSelected && grp["states"].Length > 1) {
                y += 26
                subIndent := 20
                subX := x + subIndent
                subW := w - subIndent
                sbx := subX
                for st in grp["states"] {
                    stSelected := (st["label"] = curVal)
                    stShort := st.Get("short", st["label"])
                    sBtnW := LT_TextButtonWidth("v " stShort)
                    if (sbx + sBtnW > subX + subW) {
                        sbx := subX
                        y += 22
                    }
                    sbtn := LT_AddMid(g, "Button", "x" sbx " y" y " w" sBtnW " h20", (stSelected ? "v " : "") stShort)
                    sbtn.SetFont("s8")
                    sbtn.OnEvent("Click", LT_GroupStateClick.Bind(instId, id, st["label"], st["phrase"]))
                    LT_RegisterFieldControl(instId, id, sbtn, Map("kind", "state", "label", st["label"], "short", stShort))
                    sbx += sBtnW + 4
                }
                y += 24
                bx := x  ; resume remaining family buttons on a fresh row
                justRevealed := true
            } else {
                justRevealed := false
            }
        }

        ; y tracks the top of whatever row was placed last. If that row was
        ; a plain (non-revealed) family row, its own height hasn't been
        ; accounted for yet -- add it now, same as every other field type
        ; does. If the loop ended mid-reveal, y already moved past that
        ; row's bottom, so only a small buffer is needed.
        if (!justRevealed)
            y += 26

        return y + 2

    } else if (type = "toggle") {
        cb := LT_AddMid(g, "CheckBox", "x" x " y" y " w" (w - 20), fieldDef["label"])
        cb.Value := fields.Get(id, false) ? 1 : 0
        cb.OnEvent("Click", LT_FieldToggle.Bind(instId, id))
        return y + 26

    } else if (type = "counter") {
        minV := fieldDef.Get("min", 1)
        maxV := fieldDef.Get("max", 9)
        curCount := fields.Get(id, 1)

        LT_AddMid(g, "Text", "x" x " y" (y + 3) " w90", fieldDef["label"] ":")
        minusBtn := LT_AddMid(g, "Button", "x" (x + 95) " y" y " w24 h22", "-")
        minusBtn.OnEvent("Click", LT_CounterAdjust.Bind(instId, id, -1, minV, maxV))

        countText := LT_AddMid(g, "Text", "x" (x + 123) " y" (y + 3) " w28 Center", String(curCount))
        LT_RegisterFieldControl(instId, id, countText, Map("kind", "counterDisplay"))

        plusBtn := LT_AddMid(g, "Button", "x" (x + 155) " y" y " w24 h22", "+")
        plusBtn.OnEvent("Click", LT_CounterAdjust.Bind(instId, id, 1, minV, maxV))
        return y + 28
    }

    return y
}

LT_RenderInstancePanel(g, x, y, w, instId) {
    global LT_Instances, LT_DeviceDefs

    inst := LT_Instances[instId]
    def := LT_DeviceDefs[inst["deviceKey"]]
    fields := inst["fields"]
    label := LT_InstanceDisplayLabel(instId)

    innerX := x + 10
    innerW := w - 20
    curY := y

    removeBtn := LT_AddMid(g, "Button", "x" (x + w - 65) " y" y " w65 h20", "Remove")
    removeBtn.OnEvent("Click", LT_RemoveInstance.Bind(instId))

    isRemoved := fields.Get("removed", false)
    isCollapsed := inst.Get("collapsed", false)
    arrow := isRemoved ? "" : (isCollapsed ? "▶ " : "▼ ")
    labelColor := LT_InstanceNeedsSideWarning(instId) ? "cRed" : "cBlue"

    labelCtrl := LT_AddMid(g, "Text", "x" innerX " y" (y + 3) " w" (innerW - 75) " " labelColor, arrow label)
    labelCtrl.OnEvent("Click", LT_ToggleCollapse.Bind(instId))
    curY += 24

    cb := LT_AddMid(g, "CheckBox", "x" innerX " y" curY " w200", "Now absent")
    cb.Value := isRemoved ? 1 : 0
    cb.OnEvent("Click", LT_FieldToggle.Bind(instId, "removed"))
    curY += 26

    newCb := LT_AddMid(g, "CheckBox", "x" innerX " y" curY " w200", "New")
    newCb.Value := fields.Get("isNew", false) ? 1 : 0
    newCb.OnEvent("Click", LT_FieldToggle.Bind(instId, "isNew"))
    curY += 26

    otherCb := LT_AddMid(g, "CheckBox", "x" innerX " y" curY " w200", "Add info")
    otherCb.Value := fields.Get("otherNote", false) ? 1 : 0
    otherCb.OnEvent("Click", LT_FieldToggle.Bind(instId, "otherNote"))
    curY += 26

    if (isRemoved) {
        LT_AddMid(g, "Text", "x" innerX " y" curY " w" (innerW - 20), "(marked as now absent)")
        curY += 22
    } else if (isCollapsed) {
        LT_AddMid(g, "Text", "x" innerX " y" curY " w" (innerW - 20), "(click the label above to expand)")
        curY += 22
    } else {
        for fieldDef in def["fields"] {
            if (LT_FieldHidden(fieldDef, fields))
                continue
            curY := LT_RenderField(g, innerX, curY, innerW, instId, fieldDef, fields)
        }
    }

    LT_AddMid(g, "Text", "x" x " y" curY " w" w " h2 +0x10", "")  ; SS_ETCHEDHORZ separator
    curY += 12

    return curY
}

; ============================================================================
; IMAGE-MAP POPUP (click-region diagram)
; ============================================================================

LT_GdipStartup() {
    global LT_GdipToken
    if (LT_GdipToken)
        return
    si := Buffer(24, 0)
    NumPut("UInt", 1, si, 0)  ; GdiplusVersion = 1
    DllCall("gdiplus\GdiplusStartup", "ptr*", &LT_GdipToken, "ptr", si, "ptr", 0)
}

; Loads (once per image, cached) the base diagram as a GDI+ Bitmap so it
; can be redrawn with a highlight composited on top every time the
; selection or window size changes.
LT_GdipLoadImage(imageKey, path) {
    global LT_GdipImageBitmaps
    if (LT_GdipImageBitmaps.Has(imageKey))
        return LT_GdipImageBitmaps[imageKey]
    LT_GdipStartup()
    pBitmap := 0
    DllCall("gdiplus\GdipCreateBitmapFromFile", "wstr", path, "ptr*", &pBitmap)
    LT_GdipImageBitmaps[imageKey] := pBitmap
    return pBitmap
}

; Which region name(s), if any, correspond to the current instance's
; already-made selections -- the reverse of LT_VeinRegionMap, used to
; decide what to highlight. Both a resolved tip state and a chosen side
; can be highlighted at once, since they're independent fields.
LT_FindSelectedRegionNames() {
    global LT_Instances, LT_ImageCurrentInst, LT_ImageCurrentKey
    names := []
    if (LT_ImageCurrentInst = "" || !LT_Instances.Has(LT_ImageCurrentInst))
        return names

    cfg := LT_GetImageConfig(LT_ImageCurrentKey)
    if (cfg = "")
        return names

    fields := LT_Instances[LT_ImageCurrentInst]["fields"]
    curLocState := fields.Get(cfg["locField"], "")
    regionMap := cfg["regionMap"]

    for name, target in regionMap {
        if (target.Has("state") && curLocState != "" && target["state"] = curLocState)
            names.Push(name)
        else if (target.Has("plainField")) {
            curVal := fields.Get(target["plainField"], "")
            if (curVal != "" && curVal = target["plainValue"])
                names.Push(name)
        }
    }
    return names
}

; Composites the base image plus a translucent fill over every currently-
; selected region's actual traced polygon (scaled to the current display
; size) into a fresh bitmap, and shows that on the Picture control.
LT_RenderImageWithHighlight(dispW, dispH) {
    global LT_ImagePicture, LT_ImageCurrentKey, LT_ImageCurrentHBitmap
    global LT_GdipImageBitmaps

    if (LT_ImageCurrentKey = "" || !LT_GdipImageBitmaps.Has(LT_ImageCurrentKey))
        return
    pBitmap := LT_GdipImageBitmaps[LT_ImageCurrentKey]
    if (!pBitmap || dispW <= 0 || dispH <= 0)
        return

    cfg := LT_GetImageConfig(LT_ImageCurrentKey)

    PixelFormat32bppPARGB := 0x26200A
    pCanvas := 0
    DllCall("gdiplus\GdipCreateBitmapFromScan0", "int", dispW, "int", dispH, "int", 0, "int", PixelFormat32bppPARGB, "ptr", 0, "ptr*", &pCanvas)
    if (!pCanvas)
        return

    pGraphics := 0
    DllCall("gdiplus\GdipGetImageGraphicsContext", "ptr", pCanvas, "ptr*", &pGraphics)
    DllCall("gdiplus\GdipSetInterpolationMode", "ptr", pGraphics, "int", 7)  ; HighQualityBicubic
    DllCall("gdiplus\GdipDrawImageRectI", "ptr", pGraphics, "ptr", pBitmap, "int", 0, "int", 0, "int", dispW, "int", dispH)

    regionNames := (cfg != "") ? LT_FindSelectedRegionNames() : []
    if (regionNames.Length > 0 && cfg != "") {
        scaleX := dispW / cfg["nativeW"]
        scaleY := dispH / cfg["nativeH"]
        pBrush := 0
        DllCall("gdiplus\GdipCreateSolidFill", "uint", 0x6E1E90FF, "ptr*", &pBrush)  ; translucent dodger-blue
        for region in cfg["regions"] {
            matched := false
            for n in regionNames {
                if (n = region["name"]) {
                    matched := true
                    break
                }
            }
            if (!matched)
                continue
            pts := region["points"]
            n := pts.Length
            buf := Buffer(8 * n, 0)
            Loop n {
                NumPut("Int", Round(pts[A_Index][1] * scaleX), buf, (A_Index - 1) * 8)
                NumPut("Int", Round(pts[A_Index][2] * scaleY), buf, (A_Index - 1) * 8 + 4)
            }
            DllCall("gdiplus\GdipFillPolygonI", "ptr", pGraphics, "ptr", pBrush, "ptr", buf, "int", n, "int", 0)
        }
        DllCall("gdiplus\GdipDeleteBrush", "ptr", pBrush)
    }

    DllCall("gdiplus\GdipDeleteGraphics", "ptr", pGraphics)

    hBitmap := 0
    DllCall("gdiplus\GdipCreateHBITMAPFromBitmap", "ptr", pCanvas, "ptr*", &hBitmap, "uint", 0xFFFFFFFF)
    DllCall("gdiplus\GdipDisposeImage", "ptr", pCanvas)
    if (!hBitmap)
        return

    try LT_ImagePicture.Value := "HBITMAP:" hBitmap
    if (LT_ImageCurrentHBitmap)
        try DllCall("DeleteObject", "ptr", LT_ImageCurrentHBitmap)
    LT_ImageCurrentHBitmap := hBitmap
}

; Re-renders at whatever size the popup is currently displayed at --
; called after showing, resizing, or any click that changes a selection.
LT_RefreshImageDisplay() {
    global LT_ImageGui, LT_ImageCurrentKey
    if (!IsObject(LT_ImageGui) || LT_ImageCurrentKey = "")
        return
    rect := Buffer(16, 0)
    try DllCall("GetClientRect", "ptr", LT_ImageGui.Hwnd, "ptr", rect)
    dispW := NumGet(rect, 8, "int")
    dispH := NumGet(rect, 12, "int")
    if (dispW > 0 && dispH > 0)
        LT_RenderImageWithHighlight(dispW, dispH)
}

; Point-in-polygon test (standard even-odd ray-casting rule).
LT_PointInPolygon(px, py, points) {
    inside := false
    n := points.Length
    j := n
    Loop n {
        i := A_Index
        xi := points[i][1], yi := points[i][2]
        xj := points[j][1], yj := points[j][2]
        if (((yi > py) != (yj > py)) && (px < (xj - xi) * (py - yi) / (yj - yi) + xi))
            inside := !inside
        j := i
    }
    return inside
}

; Which instance (if any) is the natural target for the popup right now:
; the one instance that's expanded and not marked removed. Mirrors the
; existing "only one thing expanded at a time" collapse model.
LT_FindExpandedInstance() {
    global LT_Instances, LT_InstanceOrder
    for id in LT_InstanceOrder {
        inst := LT_Instances[id]
        if (!inst.Get("collapsed", false) && !inst["fields"].Get("removed", false))
            return id
    }
    return ""
}

; Creates the popup window exactly once. Never destroyed/recreated after
; this -- only hidden/shown -- same lesson learned from the main window and
; the scrollable panel.
LT_EnsureImageGui() {
    global LT_ImageGui, LT_ImagePicture

    if (IsObject(LT_ImageGui))
        return

    LT_ImageGui := Gui("+Resize", "Lines and Tubes - Diagram")
    LT_ImageGui.SetFont("s9", "Segoe UI")
    LT_ImagePicture := LT_ImageGui.Add("Picture", "x0 y0 w480 h480", "")
    LT_ImagePicture.OnEvent("Click", LT_OnImagePictureClick)
    LT_ImageGui.OnEvent("Close", (*) => LT_HideImagePopup())
    LT_ImageGui.OnEvent("Size", LT_OnImageResize)
}

; Records the popup's current position/size (called on both move and
; resize) so it can be restored the next time it's shown this session.
LT_SaveImageWinGeometry() {
    global LT_ImageGui, LT_ImageWinX, LT_ImageWinY, LT_ImageWinW, LT_ImageWinH, LT_ImageWinKnown
    if (!IsObject(LT_ImageGui))
        return
    try {
        LT_ImageGui.GetPos(&LT_ImageWinX, &LT_ImageWinY, &LT_ImageWinW, &LT_ImageWinH)
        LT_ImageWinKnown := true
    }
}

; Resizes the Picture control to exactly match the window's actual client
; area, queried via GetClientRect (the same reliable technique already used
; for the scrollable middle-column panel) rather than trusting that
; whatever we last told Show()/the control matches reality.
LT_SyncImagePictureSize() {
    global LT_ImageGui, LT_ImagePicture
    if (!IsObject(LT_ImageGui) || !IsObject(LT_ImagePicture))
        return
    rect := Buffer(16, 0)
    try DllCall("GetClientRect", "ptr", LT_ImageGui.Hwnd, "ptr", rect)
    cw := NumGet(rect, 8, "int")
    ch := NumGet(rect, 12, "int")
    if (cw > 0 && ch > 0)
        try LT_ImagePicture.Move(0, 0, cw, ch)
}

LT_OnImageResize(GuiObj, MinMax, W, H) {
    global LT_ImageWinVisible
    if (MinMax = -1)  ; minimized
        return
    ; Ignore resize events that fire while the popup isn't supposed to be
    ; visible (e.g. during a Hide() transition) -- Windows can report
    ; transient/wrong geometry at that moment, and caching it here would
    ; corrupt the size used for every subsequent reshow this session.
    if (!LT_ImageWinVisible)
        return
    LT_SyncImagePictureSize()
    LT_RefreshImageDisplay()
    LT_SaveImageWinGeometry()
}

LT_OnImageMove(wParam, lParam, msg, hwnd) {
    global LT_ImageGui
    if (!IsObject(LT_ImageGui) || hwnd != LT_ImageGui.Hwnd)
        return
    LT_SaveImageWinGeometry()
}

; Live-drag aspect ratio lock, matching whichever image is currently
; loaded (so this stays correct even if a future diagram isn't square).
; wParam identifies which edge/corner is being dragged (WMSZ_* constants);
; lParam points at the proposed RECT, which we adjust in place.
; Outer window size minus client area size -- title bar + border overhead.
; Needed because the image fills the client area, not the outer window
; frame, and those two sizes differ (mainly by title bar height).
; Outer window size minus client area size -- title bar + border overhead.
; Needed because the image fills the client area, not the outer window
; frame, and those two sizes differ (mainly by title bar height).
;
; Computed via AdjustWindowRectEx from the window's style bits, not by
; measuring actual on-screen geometry (GetWindowRect/GetClientRect) --
; measuring right after Show() can read stale layout before Windows has
; actually finished sizing the window, which is what made the first version
; of this unreliable. Style-based calculation has no such timing dependency
; and works even on a window that's never been shown yet.
LT_GetWindowChrome(hwnd, &deltaW, &deltaH) {
    GWL_STYLE := -16
    GWL_EXSTYLE := -20
    style := DllCall("GetWindowLong", "ptr", hwnd, "int", GWL_STYLE, "int")
    exStyle := DllCall("GetWindowLong", "ptr", hwnd, "int", GWL_EXSTYLE, "int")

    ; arbitrary reference client size -- the delta is the same regardless
    ; of what size is used here, since chrome overhead doesn't scale
    rect := Buffer(16, 0)
    NumPut("Int", 0, rect, 0)
    NumPut("Int", 0, rect, 4)
    NumPut("Int", 100, rect, 8)
    NumPut("Int", 100, rect, 12)
    DllCall("AdjustWindowRectEx", "ptr", rect, "int", style, "int", 0, "int", exStyle)

    outerW := NumGet(rect, 8, "Int") - NumGet(rect, 0, "Int")
    outerH := NumGet(rect, 12, "Int") - NumGet(rect, 4, "Int")

    deltaW := outerW - 100
    deltaH := outerH - 100
}

LT_OnImageSizing(wParam, lParam, msg, hwnd) {
    global LT_ImageGui, LT_ImageCurrentKey
    if (!IsObject(LT_ImageGui) || hwnd != LT_ImageGui.Hwnd)
        return

    cfg := LT_GetImageConfig(LT_ImageCurrentKey)
    if (cfg = "")
        return
    ratio := cfg["nativeW"] / cfg["nativeH"]

    LT_GetWindowChrome(hwnd, &dw, &dh)

    left := NumGet(lParam, 0, "Int")
    top := NumGet(lParam, 4, "Int")
    right := NumGet(lParam, 8, "Int")
    bottom := NumGet(lParam, 12, "Int")

    ; work in client-area-equivalent dimensions, since that's what actually
    ; needs to match the image's ratio, not the outer window frame
    w := (right - left) - dw
    h := (bottom - top) - dh

    WMSZ_TOP := 3, WMSZ_TOPLEFT := 4, WMSZ_TOPRIGHT := 5, WMSZ_BOTTOM := 6

    if (wParam = WMSZ_TOP || wParam = WMSZ_BOTTOM) {
        ; dragging a horizontal edge only -- height changed, derive width
        right := left + Round(h * ratio) + dw
    } else {
        ; dragging a vertical edge or any corner -- width changed, derive
        ; height; anchor the bottom instead of the top when dragging from
        ; a top-anchored corner, so the edge actually being dragged moves
        newH := Round(w / ratio)
        if (wParam = WMSZ_TOPLEFT || wParam = WMSZ_TOPRIGHT)
            top := bottom - newH - dh
        else
            bottom := top + newH + dh
    }

    NumPut("Int", left, lParam, 0)
    NumPut("Int", top, lParam, 4)
    NumPut("Int", right, lParam, 8)
    NumPut("Int", bottom, lParam, 12)
    return true
}

; Shows (creating if needed) the popup for the given image key, positioned
; to the left of the main window the first time this session, or wherever
; it was last left afterward. Only actually calls Show() when transitioning
; from hidden to visible -- calling it again while already visible (which
; would otherwise happen on every single click while a vein device is
; expanded) isn't needed and risks exactly the kind of redraw glitch that
; showed up as a duplicated-looking image.
; Work-area bounds (left/top/right/bottom, in screen coordinates) of
; whichever monitor the given window is currently on -- used to check
; whether the popup's default left-side position would land off-screen.
LT_GetMonitorWorkArea(hwnd, &left, &top, &right, &bottom) {
    MONITOR_DEFAULTTONEAREST := 2
    hMonitor := DllCall("MonitorFromWindow", "ptr", hwnd, "uint", MONITOR_DEFAULTTONEAREST, "ptr")
    mi := Buffer(40, 0)
    NumPut("UInt", 40, mi, 0)  ; cbSize
    DllCall("GetMonitorInfo", "ptr", hMonitor, "ptr", mi)
    ; MONITORINFO: cbSize(4), rcMonitor(4x Int32), rcWork(4x Int32), dwFlags(4)
    left := NumGet(mi, 20, "Int")
    top := NumGet(mi, 24, "Int")
    right := NumGet(mi, 28, "Int")
    bottom := NumGet(mi, 32, "Int")
}

; True if the given window currently has the WS_EX_TOPMOST style -- used
; to make the image popup mirror the main window's always-on-top state
; exactly, rather than trying to duplicate every place that toggles it.
LT_IsAlwaysOnTop(hwnd) {
    GWL_EXSTYLE := -20
    WS_EX_TOPMOST := 0x8
    exStyle := DllCall("GetWindowLong", "ptr", hwnd, "int", GWL_EXSTYLE, "int")
    return (exStyle & WS_EX_TOPMOST) != 0
}

LT_ShowImagePopup(imageKey, instId) {
    global LT_ImageGui, LT_ImagePicture, LT_GuiObj, LT_ImageDir
    global LT_ImageCurrentKey, LT_ImageCurrentInst, LT_ImageWinVisible
    global LT_ImageWinX, LT_ImageWinY, LT_ImageWinW, LT_ImageWinH, LT_ImageWinKnown

    LT_EnsureImageGui()

    ; Mirror the main window's current always-on-top state -- checked live
    ; every time, so it's correct regardless of whether the image window
    ; was created before or after the main window last toggled it.
    if (IsObject(LT_GuiObj)) {
        try {
            if (LT_IsAlwaysOnTop(LT_GuiObj.Hwnd))
                LT_ImageGui.Opt("+AlwaysOnTop")
            else
                LT_ImageGui.Opt("-AlwaysOnTop")
        }
    }
    LT_ImageCurrentInst := instId

    if (imageKey != LT_ImageCurrentKey) {
        imgPath := LT_ImageDir "\" imageKey ".png"
        if (!FileExist(imgPath)) {
            MsgBox("Diagram image not found:`n" imgPath, "Lines and Tubes", 48)
            return
        }
        LT_GdipLoadImage(imageKey, imgPath)
        LT_ImageCurrentKey := imageKey
    }

    if (!LT_ImageWinVisible) {
        if (!LT_ImageWinKnown) {
            mx := 0, my := 0, mw := 0, mh := 0
            if (IsObject(LT_GuiObj))
                try LT_GuiObj.GetPos(&mx, &my, &mw, &mh)

            ; Size from the real image dimensions -- fixing the width at
            ; 480 and deriving height from the actual ratio, rather than a
            ; hardcoded square that only self-corrected once the person
            ; manually resized (WM_SIZING only enforces the ratio live
            ; during a drag, not at creation time).
            cfg := LT_GetImageConfig(imageKey)
            LT_ImageWinW := 480
            LT_ImageWinH := (cfg != "") ? Round(480 * cfg["nativeH"] / cfg["nativeW"]) : 480

            leftX := mx - LT_ImageWinW - 15
            rightX := mx + mw + 15

            gotMonitor := false
            if (IsObject(LT_GuiObj)) {
                try {
                    LT_GetMonitorWorkArea(LT_GuiObj.Hwnd, &monLeft, &monTop, &monRight, &monBottom)
                    gotMonitor := true
                }
            }

            ; Only override to the right side if the left position would
            ; genuinely land off that monitor's left edge -- otherwise keep
            ; the normal left placement.
            LT_ImageWinX := (gotMonitor && leftX < monLeft) ? rightX : leftX
            LT_ImageWinY := my
            LT_ImageWinKnown := true

            ; LT_ImageWinW/H above are the *target client area* size --
            ; convert to the outer window size the image actually needs
            ; Show() to use, via the style-based chrome calculation (works
            ; before the window's ever been shown, unlike measuring actual
            ; on-screen geometry).
            try {
                LT_GetWindowChrome(LT_ImageGui.Hwnd, &dw, &dh)
                LT_ImageWinW += dw
                LT_ImageWinH += dh
            }
        }
        try LT_ImageGui.Show("x" LT_ImageWinX " y" LT_ImageWinY " w" LT_ImageWinW " h" LT_ImageWinH)
        LT_ImageWinVisible := true
        LT_SyncImagePictureSize()
    }

    LT_RefreshImageDisplay()
}

LT_HideImagePopup() {
    global LT_ImageGui, LT_ImageCurrentInst, LT_ImageWinVisible
    LT_ImageWinVisible := false
    if (IsObject(LT_ImageGui))
        try LT_ImageGui.Hide()
    LT_ImageCurrentInst := ""
}

; Called after every middle-column rebuild: shows the popup if the
; currently-expanded instance's device has an image, hides it otherwise --
; this is what keeps it from persisting or stacking when you're not
; actively working on an image-capable device.
LT_UpdateImagePopup() {
    global LT_DeviceDefs, LT_Instances

    instId := LT_FindExpandedInstance()
    if (instId = "") {
        LT_HideImagePopup()
        return
    }

    inst := LT_Instances[instId]
    def := LT_DeviceDefs[inst["deviceKey"]]
    imageKey := def.Get("imageKey", "")
    if (imageKey = "") {
        LT_HideImagePopup()
        return
    }

    LT_ShowImagePopup(imageKey, instId)
}

; Click on the diagram, via the Picture control's own Click event (same
; proven mechanism as the report-pane's clickable lines) rather than a raw
; WM_LBUTTONDOWN hook of uncertain reliability for this control type.
; MouseGetPos with CoordMode "Client" gives coordinates relative to the
; image window's client area -- since the Picture control always exactly
; fills that client area (see LT_SyncImagePictureSize), those are already
; the control-relative coordinates we need.
; Per-image config: which regions/native size/routing map to use, and which
; field on the device the location regions resolve against (vein uses
; "tip", enteric/feeding uses "location"). Centralizing this is what lets
; the click handler, highlighter, and reverse-lookup all stay image-agnostic
; instead of hardcoding "vein" specifically.
LT_GetImageConfig(imageKey) {
    global LT_VeinImageRegions, LT_VeinImageNativeW, LT_VeinImageNativeH, LT_VeinRegionMap
    global LT_EntericFeedingImageRegions, LT_EntericFeedingImageNativeW, LT_EntericFeedingImageNativeH, LT_EntericFeedingRegionMap

    if (imageKey = "vein")
        return Map("regions", LT_VeinImageRegions, "nativeW", LT_VeinImageNativeW,
            "nativeH", LT_VeinImageNativeH, "regionMap", LT_VeinRegionMap, "locField", "tip")
    if (imageKey = "enteric_feeding")
        return Map("regions", LT_EntericFeedingImageRegions, "nativeW", LT_EntericFeedingImageNativeW,
            "nativeH", LT_EntericFeedingImageNativeH, "regionMap", LT_EntericFeedingRegionMap, "locField", "location")
    return ""
}

LT_OnImagePictureClick(ctrl, info) {
    global LT_ImageGui, LT_ImageCurrentKey, LT_ImageCurrentInst
    global LT_Instances, LT_DeviceDefs

    if (LT_ImageCurrentInst = "" || !LT_Instances.Has(LT_ImageCurrentInst))
        return

    cfg := LT_GetImageConfig(LT_ImageCurrentKey)
    if (cfg = "")
        return  ; no region data for this image yet

    prevMode := A_CoordModeMouse
    CoordMode("Mouse", "Client")
    MouseGetPos(&dispX, &dispY)
    CoordMode("Mouse", prevMode)

    rect := Buffer(16, 0)
    try DllCall("GetClientRect", "ptr", LT_ImageGui.Hwnd, "ptr", rect)
    dispW := NumGet(rect, 8, "int")
    dispH := NumGet(rect, 12, "int")
    if (dispW = 0 || dispH = 0)
        return

    nx := dispX * (cfg["nativeW"] / dispW)
    ny := dispY * (cfg["nativeH"] / dispH)

    matchName := ""
    for region in cfg["regions"] {
        if (LT_PointInPolygon(nx, ny, region["points"])) {
            matchName := region["name"]
            break
        }
    }
    regionMap := cfg["regionMap"]
    if (matchName = "" || !regionMap.Has(matchName))
        return

    target := regionMap[matchName]
    instId := LT_ImageCurrentInst
    inst := LT_Instances[instId]
    def := LT_DeviceDefs[inst["deviceKey"]]
    fields := inst["fields"]

    if (target.Has("plainField")) {
        fields[target["plainField"]] := target["plainValue"]
        LT_RebuildMiddleColumn(instId)
        return
    }

    ; family+state: resolve directly against whichever groups array this
    ; device's own location field actually uses (e.g. vein's Projecting or
    ; PICC phrasing variant), so the same region map drives every device
    ; that shares this image.
    locField := cfg["locField"]
    fieldDef := LT_FindFieldDef(def, locField)
    if (fieldDef = "")
        return
    for grp in fieldDef["groups"] {
        if (grp["groupLabel"] != target["family"])
            continue
        for st in grp["states"] {
            if (st["label"] = target["state"]) {
                fields[locField] := st["label"]
                fields[locField "_phrase"] := st["phrase"]
                fields[locField "_activeFamily"] := grp["groupLabel"]
                LT_RebuildMiddleColumn(instId)
                return
            }
        }
    }
}

; Builds the window shell exactly once. Left/right columns are static;
; only the middle column's contents change after this.
; Closing the window is the other way (besides New Patient) to end the
; "pinned above everything" period.
LT_OnGuiClose(*) {
    global LT_GuiObj
    if (IsObject(LT_GuiObj)) {
        try LT_GuiObj.Opt("-AlwaysOnTop")
        LT_GuiObj.Hide()
    }
    LT_HideImagePopup()
}

LT_EnsureGui() {
    global LT_GuiObj, LT_MidX, LT_MidW, LT_RightX, LT_RightW
    global LT_ListBoxCtrl, LT_AddBtnCtrl, LT_ClearBtnCtrl, LT_NewPatientBtnCtrl
    global LT_MidLabelCtrl, LT_RightLabelCtrl, LT_CopyBtnCtrl
    global LT_MidPanel, LT_MidPanelX, LT_MidPanelY, LT_MidPanelW, LT_MidPanelH

    if (IsObject(LT_GuiObj))
        return

    LT_GuiObj := Gui("+Resize", "Lines and Tubes")
    LT_GuiObj.SetFont("s9", "Segoe UI")

    HotIfWinActive("ahk_id " LT_GuiObj.Hwnd)
    Hotkey("^c", LT_CopyOutput)
    Hotkey("^n", LT_NewPatient)
    Hotkey("^Delete", LT_ClearAll)
    Hotkey("+Delete", LT_HotkeyRemoveSelected)
    HotIfWinActive()

    leftW := 180
    LT_GuiObj.Add("Text", "x10 y10 w" leftW, "Add device:")
    LT_ListBoxCtrl := LT_GuiObj.Add("ListBox", "x10 y30 w" leftW " h405 vLT_DeviceListBox", LT_DeviceLabelList())
    LT_ListBoxCtrl.OnEvent("DoubleClick", LT_AddDeviceFromListbox)
    LT_AddBtnCtrl := LT_GuiObj.Add("Button", "x10 y440 w" leftW " h28", "Add Selected Device")
    LT_AddBtnCtrl.OnEvent("Click", LT_AddDeviceFromListbox)
    LT_ClearBtnCtrl := LT_GuiObj.Add("Button", "x10 y473 w" leftW " h28", "Clear All")
    LT_ClearBtnCtrl.OnEvent("Click", LT_ClearAll)
    LT_NewPatientBtnCtrl := LT_GuiObj.Add("Button", "x10 y506 w" leftW " h28", "New Patient")
    LT_NewPatientBtnCtrl.OnEvent("Click", LT_NewPatient)

    LT_MidX := 10 + leftW + 15
    LT_MidW := 330
    LT_MidLabelCtrl := LT_GuiObj.Add("Text", "x" LT_MidX " y10 w" LT_MidW, "Active devices:")

    ; The middle column is a real child window with its own native
    ; scrollbar -- created once here and never destroyed again.
    LT_MidPanelX := LT_MidX
    LT_MidPanelY := 35
    LT_MidPanelW := LT_MidW
    LT_MidPanelH := 525
    LT_MidPanel := LT_CreateMidPanel()

    LT_RightX := LT_MidX + LT_MidW + 15
    LT_RightW := 330
    LT_RightLabelCtrl := LT_GuiObj.Add("Text", "x" LT_RightX " y10 w" LT_RightW, "Report text (click a line to edit it):")
    LT_CopyBtnCtrl := LT_GuiObj.Add("Button", "x" LT_RightX " y480 w" LT_RightW " h30", "Copy to Clipboard")
    LT_CopyBtnCtrl.OnEvent("Click", LT_CopyOutput)

    LT_GuiObj.OnEvent("Close", LT_OnGuiClose)
    LT_GuiObj.OnEvent("Size", LT_OnResize)

    defaultW := LT_RightX + LT_RightW + 20
    defaultH := Round(580 * 1.25)  ; 25% taller, so the pick list fits without scrolling
    LT_GuiObj.Show("w" defaultW " h" defaultH)

    ; The controls above were laid out for the old default height; reflow
    ; immediately (reusing the same logic a manual resize triggers) so the
    ; extra height is actually used rather than left as dead space.
    global LT_PendingResizeW, LT_PendingResizeH
    LT_PendingResizeW := defaultW
    LT_PendingResizeH := defaultH
    LT_FlushResize()
}

; A resize drag fires the Size event continuously -- debounce it the same
; way scrolling is debounced, so dragging the border doesn't hammer a full
; rebuild dozens of times a second.
LT_OnResize(GuiObj, MinMax, W, H) {
    global LT_PendingResizeW, LT_PendingResizeH
    if (MinMax = -1)  ; minimized -- nothing meaningful to lay out
        return
    LT_PendingResizeW := W
    LT_PendingResizeH := H
    SetTimer(LT_FlushResize, -60)
}

; Recomputes the whole three-column layout for the new window size and
; repositions every static control, then rebuilds the middle column and
; report pane at their new coordinates.
LT_FlushResize() {
    global LT_GuiObj, LT_PendingResizeW, LT_PendingResizeH
    global LT_MidX, LT_MidW, LT_RightX, LT_RightW
    global LT_ListBoxCtrl, LT_AddBtnCtrl, LT_ClearBtnCtrl, LT_NewPatientBtnCtrl
    global LT_MidLabelCtrl, LT_RightLabelCtrl, LT_CopyBtnCtrl

    if (!IsObject(LT_GuiObj))
        return

    W := LT_PendingResizeW
    H := LT_PendingResizeH
    if (W < 300 || H < 200)  ; ignore degenerate sizes mid-drag
        return

    leftX := 10
    leftW := 180
    gap := 15
    minColW := 220
    bottomMargin := 20

    remaining := W - leftX - leftW - gap * 2 - 20
    if (remaining < minColW * 2)
        remaining := minColW * 2
    midW := Floor(remaining / 2)
    rightW := remaining - midW
    midX := leftX + leftW + gap
    rightX := midX + midW + gap

    LT_MidX := midX
    LT_MidW := midW
    LT_RightX := rightX
    LT_RightW := rightW

    ; left column: list plus three stacked buttons, filling available height
    listY := 30
    buttonH := 28
    buttonGap := 5
    buttonsBlockH := 3 * buttonH + 2 * buttonGap
    listH := Max(100, H - listY - buttonsBlockH - buttonGap - bottomMargin)
    btn1Y := listY + listH + buttonGap
    btn2Y := btn1Y + buttonH + buttonGap
    btn3Y := btn2Y + buttonH + buttonGap

    try LT_ListBoxCtrl.Move(leftX, listY, leftW, listH)
    try LT_AddBtnCtrl.Move(leftX, btn1Y, leftW, buttonH)
    try LT_ClearBtnCtrl.Move(leftX, btn2Y, leftW, buttonH)
    try LT_NewPatientBtnCtrl.Move(leftX, btn3Y, leftW, buttonH)

    ; middle column: header label + the scrollable panel itself
    try LT_MidLabelCtrl.Move(midX, 10, midW)
    midPanelH := Max(100, H - 35 - bottomMargin)
    LT_MoveMidPanel(midX, 35, midW, midPanelH)

    ; right column: header label + copy button
    try LT_RightLabelCtrl.Move(rightX, 10, rightW)
    copyBtnY := H - bottomMargin - buttonH
    try LT_CopyBtnCtrl.Move(rightX, copyBtnY, rightW, buttonH)

    LT_RebuildMiddleColumn()
    LT_RenderOutputPane()
}

; Refreshes only the middle column: hides the previously rendered instance
; panels (rather than destroying the window) and draws the current state.
; If focusInstId is given, the scroll position is adjusted (only if needed)
; so that instance's panel starts within the visible viewport -- this is
; what stops an expand/collapse from silently scrolling the thing you just
; clicked out of view.
LT_RebuildMiddleColumn(focusInstId := "") {
    global LT_MidPanel, LT_MiddleControls, LT_FieldControls, LT_InstanceOrder
    global LT_ScrollOffset, LT_MidContentBottom, LT_MidPanelClientW, LT_MidPanelClientH

    if (!IsObject(LT_MidPanel))
        return

    for ctrl in LT_MiddleControls {
        try ctrl.Visible := false
    }
    LT_MiddleControls := []
    LT_FieldControls := Map()

    margin := 10
    contentW := LT_MidPanelClientW - margin

    instanceStartY := Map()
    instanceEndY := Map()
    curY := margin
    for instId in LT_InstanceOrder {
        instanceStartY[instId] := curY
        curY := LT_RenderInstancePanel(LT_MidPanel, 0, curY, contentW, instId)
        instanceEndY[instId] := curY
    }
    LT_MidContentBottom := curY

    desiredOffset := LT_ScrollOffset
    if (focusInstId != "" && instanceStartY.Has(focusInstId)) {
        startY := instanceStartY[focusInstId]
        endY := instanceEndY[focusInstId]
        viewportH := LT_MidPanelClientH
        viewTop := desiredOffset
        viewBottom := desiredOffset + viewportH

        if (startY < viewTop || endY > viewBottom) {
            if (endY - startY <= viewportH) {
                ; the whole expanded block fits in the viewport -- prefer
                ; showing all of it rather than just snapping to the top,
                ; so a newly revealed row at the bottom isn't left hidden
                desiredOffset := Max(0, endY - viewportH)
                if (startY - desiredOffset < 0)
                    desiredOffset := Max(0, startY)
            } else {
                ; doesn't fit regardless -- anchor to the top and let the
                ; person scroll the rest, same as before
                desiredOffset := Max(0, startY)
            }
        }
    }

    maxOff := Max(0, LT_MidContentBottom - LT_MidPanelClientH)
    if (desiredOffset > maxOff)
        desiredOffset := maxOff
    if (desiredOffset < 0)
        desiredOffset := 0

    ; Newly (re)created controls sit at their raw/unscrolled positions --
    ; one ScrollWindow call brings the whole panel to the resolved offset.
    if (desiredOffset > 0)
        try DllCall("ScrollWindow", "ptr", LT_MidPanel.Hwnd, "int", 0, "int", -desiredOffset, "ptr", 0, "ptr", 0)
    LT_ScrollOffset := desiredOffset
    LT_UpdateScrollInfo()

    LT_UpdateOutputBoxOnly()
    LT_UpdateImagePopup()
}
