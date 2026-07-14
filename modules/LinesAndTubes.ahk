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

; Canonical output order (also the order devices are listed on the left)
LT_DeviceOrder := ["ETT", "TRACH", "ENTERIC", "FEEDING", "IJ", "SCV", "PICC", "PORT",
    "IMPELLA", "LVAD", "EPIWIRE", "GENERATOR", "RETAINEDLEADS", "PADS", "PLEURAL", "MEDIASTINAL", "EPIDURAL"]

OnMessage(0x20A, LT_OnMouseWheel)  ; WM_MOUSEWHEEL
OnMessage(0x115, LT_OnVScroll)     ; WM_VSCROLL

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
        Map("label", "Right brachiocephalic vein", "phrase", "over the right brachiocephalic vein"),
        Map("label", "Left brachiocephalic vein", "phrase", "over the left brachiocephalic vein")
    ]),
    Map("groupLabel", "Confluence of brachiocephalic veins", "states", [
        Map("label", "Confluence of the brachiocephalic veins", "phrase", "at the confluence of the brachiocephalic veins")
    ]),
    Map("groupLabel", "SVC", "states", [
        Map("label", "SVC (unspecified)", "phrase", "in the superior vena cava"),
        Map("label", "Proximal SVC", "phrase", "in the proximal superior vena cava"),
        Map("label", "Mid SVC", "phrase", "in the mid superior vena cava"),
        Map("label", "Distal SVC", "phrase", "in the distal superior vena cava")
    ]),
    Map("groupLabel", "Cavoatrial junction", "states", [
        Map("label", "Superior cavoatrial junction", "phrase", "at the superior cavoatrial junction"),
        Map("label", "Inferior cavoatrial junction", "phrase", "at the inferior cavoatrial junction")
    ]),
    Map("groupLabel", "Right atrium", "states", [
        Map("label", "Right atrium (unspecified)", "phrase", "in the right atrium"),
        Map("label", "Upper right atrium", "phrase", "in the upper right atrium"),
        Map("label", "Mid right atrium", "phrase", "in the mid right atrium"),
        Map("label", "Lower right atrium", "phrase", "in the lower right atrium")
    ]),
    Map("groupLabel", "Right ventricle", "states", [
        Map("label", "Right ventricle", "phrase", "in the right ventricle")
    ]),
    Map("groupLabel", "Pulmonary artery", "states", [
        Map("label", "Pulmonary artery (unspecified)", "phrase", "in the pulmonary artery"),
        Map("label", "Proximal pulmonary artery", "phrase", "in the proximal pulmonary artery"),
        Map("label", "Mid pulmonary artery", "phrase", "in the mid pulmonary artery"),
        Map("label", "Distal pulmonary artery", "phrase", "in the distal pulmonary artery")
    ]),
    Map("groupLabel", "Right pulmonary artery", "states", [
        Map("label", "Right pulmonary artery (unspecified)", "phrase", "in the right pulmonary artery"),
        Map("label", "Proximal right pulmonary artery", "phrase", "in the proximal right pulmonary artery"),
        Map("label", "Mid right pulmonary artery", "phrase", "in the mid right pulmonary artery"),
        Map("label", "Distal right pulmonary artery", "phrase", "in the distal right pulmonary artery")
    ]),
    Map("groupLabel", "Left pulmonary artery", "states", [
        Map("label", "Left pulmonary artery (unspecified)", "phrase", "in the left pulmonary artery"),
        Map("label", "Proximal left pulmonary artery", "phrase", "in the proximal left pulmonary artery"),
        Map("label", "Mid left pulmonary artery", "phrase", "in the mid left pulmonary artery"),
        Map("label", "Distal left pulmonary artery", "phrase", "in the distal left pulmonary artery")
    ]),
    Map("groupLabel", "Interlobar artery", "states", [
        Map("label", "Right interlobar artery", "phrase", "in the right interlobar artery"),
        Map("label", "Left interlobar artery", "phrase", "in the left interlobar artery")
    ]),
    Map("groupLabel", "Other", "states", [
        Map("label", "Other", "phrase", "__OTHER__")
    ])
]

LT_EntericTipGroups := [
    Map("groupLabel", "Esophagus", "states", [
        Map("label", "Esophagus (unspecified)", "phrase", "in the esophagus"),
        Map("label", "Upper esophagus", "phrase", "in the upper esophagus"),
        Map("label", "Mid esophagus", "phrase", "in the mid esophagus"),
        Map("label", "Distal esophagus", "phrase", "in the distal esophagus")
    ]),
    Map("groupLabel", "Esophagogastric junction", "states", [
        Map("label", "Near esophagogastric junction", "phrase", "near the esophagogastric junction"),
        Map("label", "Above esophagogastric junction", "phrase", "above the esophagogastric junction"),
        Map("label", "Below esophagogastric junction", "phrase", "below the esophagogastric junction")
    ]),
    Map("groupLabel", "Gastric conduit", "states", [
        Map("label", "Gastric conduit", "phrase", "in the gastric conduit")
    ]),
    Map("groupLabel", "Stomach (tip and side port)", "states", [
        Map("label", "Stomach (tip and side port)", "phrase", "__STOMACH_GEN__"),
        Map("label", "Proximal stomach", "phrase", "__STOMACH_PROX__"),
        Map("label", "Mid stomach", "phrase", "__STOMACH_MID__"),
        Map("label", "Distal stomach", "phrase", "__STOMACH_DIST__")
    ]),
    Map("groupLabel", "Off-image below diaphragm", "states", [
        Map("label", "Off-image below diaphragm", "phrase", "__OFFIMAGE__")
    ]),
    Map("groupLabel", "Malpositioned bronchus", "states", [
        Map("label", "Right lower lobe bronchus", "phrase", "__MALPOS_R__"),
        Map("label", "Left lower lobe bronchus", "phrase", "__MALPOS_L__")
    ]),
    Map("groupLabel", "Other", "states", [
        Map("label", "Other", "phrase", "__OTHER__")
    ])
]

LT_FeedingTipGroups := [
    Map("groupLabel", "Esophagus", "states", [
        Map("label", "Esophagus (unspecified)", "phrase", "in the esophagus"),
        Map("label", "Upper esophagus", "phrase", "in the upper esophagus"),
        Map("label", "Mid esophagus", "phrase", "in the mid esophagus"),
        Map("label", "Distal esophagus", "phrase", "in the distal esophagus")
    ]),
    Map("groupLabel", "Esophagogastric junction", "states", [
        Map("label", "Near esophagogastric junction", "phrase", "near the esophagogastric junction"),
        Map("label", "Above esophagogastric junction", "phrase", "above the esophagogastric junction"),
        Map("label", "Below esophagogastric junction", "phrase", "below the esophagogastric junction")
    ]),
    Map("groupLabel", "Gastric conduit", "states", [
        Map("label", "Gastric conduit", "phrase", "in the gastric conduit")
    ]),
    Map("groupLabel", "Stomach (tip and side port)", "states", [
        Map("label", "Stomach (tip and side port)", "phrase", "__STOMACH_GEN__"),
        Map("label", "Proximal stomach", "phrase", "__STOMACH_PROX__"),
        Map("label", "Mid stomach", "phrase", "__STOMACH_MID__"),
        Map("label", "Distal stomach", "phrase", "__STOMACH_DIST__")
    ]),
    Map("groupLabel", "Duodenum", "states", [
        Map("label", "Duodenum (unspecified)", "phrase", "in the duodenum"),
        Map("label", "Proximal duodenum", "phrase", "in the proximal duodenum"),
        Map("label", "Mid duodenum", "phrase", "in the mid duodenum"),
        Map("label", "Distal duodenum", "phrase", "in the distal duodenum")
    ]),
    Map("groupLabel", "Near the pylorus", "states", [
        Map("label", "Near the pylorus", "phrase", "near the pylorus")
    ]),
    Map("groupLabel", "Duodenojejunal junction", "states", [
        Map("label", "At duodenojejunal junction", "phrase", "at the duodenojejunal junction"),
        Map("label", "Beyond duodenojejunal junction", "phrase", "beyond the duodenojejunal junction")
    ]),
    Map("groupLabel", "Off-image below diaphragm", "states", [
        Map("label", "Off-image below diaphragm", "phrase", "__OFFIMAGE__")
    ]),
    Map("groupLabel", "Malpositioned bronchus", "states", [
        Map("label", "Right lower lobe bronchus", "phrase", "__MALPOS_R__"),
        Map("label", "Left lower lobe bronchus", "phrase", "__MALPOS_L__")
    ]),
    Map("groupLabel", "Other", "states", [
        Map("label", "Other", "phrase", "__OTHER__")
    ])
]

LT_EpiduralGroups := [
    Map("groupLabel", "Thoracic spine", "states", [
        Map("label", "Thoracic spine (unspecified)", "phrase", "over the thoracic spine"),
        Map("label", "Upper thoracic spine", "phrase", "over the upper thoracic spine"),
        Map("label", "Mid thoracic spine", "phrase", "over the mid thoracic spine"),
        Map("label", "Lower thoracic spine", "phrase", "over the lower thoracic spine")
    ]),
    Map("groupLabel", "Lumbar spine", "states", [
        Map("label", "Lumbar spine (unspecified)", "phrase", "over the lumbar spine"),
        Map("label", "Upper lumbar spine", "phrase", "over the upper lumbar spine"),
        Map("label", "Mid lumbar spine", "phrase", "over the mid lumbar spine")
    ])
]

; Only set these if running standalone
if (A_LineFile = A_ScriptFullPath) {
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
    global LT_CentralTipGroups, LT_EntericTipGroups, LT_FeedingTipGroups, LT_EpiduralGroups

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
        "removalNoun", (fields) => "endotracheal tube"
    )

    defs["ENTERIC"] := Map(
        "label", "Enteric tube",
        "fields", [
            Map("id", "location", "type", "grouped", "label", "Tip location",
                "groups", LT_EntericTipGroups)
        ],
        "sentenceFn", (fields) => LT_Sentence_Enteric(fields, "Enteric tube"),
        "removalNoun", (fields) => "enteric tube"
    )

    defs["FEEDING"] := Map(
        "label", "Feeding tube",
        "fields", [
            Map("id", "location", "type", "grouped", "label", "Tip location",
                "groups", LT_FeedingTipGroups),
            Map("id", "weighted", "type", "toggle", "label", "Weighted tip"),
            Map("id", "stylet", "type", "toggle", "label", "With stylet")
        ],
        "sentenceFn", LT_Sentence_Feeding,
        "removalNoun", (fields) => fields.Get("weighted", false) ? "weighted feeding tube" : "feeding tube"
    )

    defs["IJ"] := Map(
        "label", "Internal jugular catheter",
        "shortLabel", "IJ catheter",
        "fields", [
            Map("id", "laterality", "type", "buttons", "label", "Side",
                "options", ["Right", "Left"]),
            Map("id", "boreSize", "type", "buttons", "label", "Bore",
                "options", ["", "Large bore", "Small bore"]),
            Map("id", "modNew", "type", "toggle", "label", "New"),
            Map("id", "modRepositioned", "type", "toggle", "label", "Repositioned"),
            Map("id", "modTunneled", "type", "toggle", "label", "Tunneled"),
            Map("id", "modSheathed", "type", "toggle", "label", "Sheathed"),
            Map("id", "sheathOnly", "type", "toggle", "label", "Sheath only (empty)"),
            Map("id", "tip", "type", "grouped", "label", "Tip location",
                "groups", LT_CentralTipGroups)
        ],
        "sentenceFn", (fields) => LT_Sentence_CentralLine(fields, "internal jugular catheter"),
        "removalNoun", (fields) => LT_RemovalNoun_CentralLine(fields, "internal jugular catheter")
    )

    defs["SCV"] := Map(
        "label", "Subclavian vein catheter",
        "shortLabel", "SCV catheter",
        "fields", [
            Map("id", "laterality", "type", "buttons", "label", "Side",
                "options", ["Right", "Left"]),
            Map("id", "boreSize", "type", "buttons", "label", "Bore",
                "options", ["", "Large bore", "Small bore"]),
            Map("id", "modNew", "type", "toggle", "label", "New"),
            Map("id", "modRepositioned", "type", "toggle", "label", "Repositioned"),
            Map("id", "modTunneled", "type", "toggle", "label", "Tunneled"),
            Map("id", "modSheathed", "type", "toggle", "label", "Sheathed"),
            Map("id", "sheathOnly", "type", "toggle", "label", "Sheath only (empty)"),
            Map("id", "tip", "type", "grouped", "label", "Tip location",
                "groups", LT_BuildSubclavianTipGroups())
        ],
        "sentenceFn", (fields) => LT_Sentence_CentralLine(fields, "subclavian vein catheter"),
        "removalNoun", (fields) => LT_RemovalNoun_CentralLine(fields, "subclavian vein catheter")
    )

    defs["PICC"] := Map(
        "label", "Peripherally inserted central catheter (PICC)",
        "shortLabel", "PICC",
        "fields", [
            Map("id", "laterality", "type", "buttons", "label", "Side",
                "options", ["Right", "Left"]),
            Map("id", "tip", "type", "grouped", "label", "Tip location",
                "groups", LT_BuildSubclavianTipGroups())
        ],
        "sentenceFn", (fields) => LT_Sentence_CentralLine(fields, "peripherally inserted central catheter"),
        "removalNoun", (fields) => LT_RemovalNoun_CentralLine(fields, "peripherally inserted central catheter")
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
                "groups", LT_BuildPortTipGroups())
        ],
        "sentenceFn", LT_Sentence_Port,
        "removalNoun", LT_RemovalNoun_Port
    )

    defs["TRACH"] := Map(
        "label", "Tracheostomy tube",
        "fields", [
            Map("id", "form", "type", "buttons", "label", "Position",
                "options", ["At thoracic inlet", "Above carina"])
        ],
        "sentenceFn", LT_Sentence_Trach,
        "removalNoun", (fields) => "tracheostomy tube"
    )

    defs["PLEURAL"] := Map(
        "label", "Pleural tube",
        "fields", [
            Map("id", "laterality", "type", "buttons", "label", "Side",
                "options", ["Right", "Left"]),
            Map("id", "deviceType", "type", "buttons", "label", "Type",
                "options", ["Chest tube", "Pleural catheter"]),
            Map("id", "location", "type", "buttons", "label", "Location",
                "options", ["Apical pleural", "Basal pleural", "Mid pleural", "In the chest wall"])
        ],
        "sentenceFn", LT_Sentence_Pleural,
        "removalNoun", LT_RemovalNoun_Pleural
    )

    defs["MEDIASTINAL"] := Map(
        "label", "Mediastinal drain",
        "fields", [
            Map("id", "inferiorApproach", "type", "toggle", "label", "Inferior approach"),
            Map("id", "position", "type", "buttons", "label", "Position",
                "options", ["Anterior", "Posterior"])
        ],
        "sentenceFn", LT_Sentence_Mediastinal,
        "removalNoun", (fields) => "mediastinal drain"
    )

    defs["EPIDURAL"] := Map(
        "label", "Epidural catheter",
        "fields", [
            Map("id", "tip", "type", "grouped", "label", "Tip location",
                "groups", LT_EpiduralGroups)
        ],
        "sentenceFn", LT_Sentence_Epidural,
        "removalNoun", (fields) => "epidural catheter"
    )

    defs["EPIWIRE"] := Map(
        "label", "Epicardial pacing wires",
        "shortLabel", "Epicardial wires",
        "fields", [
            Map("id", "retained", "type", "toggle", "label", "Retained")
        ],
        "sentenceFn", (fields) => (fields.Get("retained", false)
            ? "Retained epicardial pacing wires are present."
            : "Epicardial pacing wires are present."),
        "removalNoun", (fields) => (fields.Get("retained", false)
            ? "retained epicardial pacing wires"
            : "epicardial pacing wires")
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
        "removalNoun", (fields) => "Impella left ventricular assist device"
    )

    defs["LVAD"] := Map(
        "label", "Left ventricular assist device (LVAD)",
        "shortLabel", "LVAD",
        "fields", [
            Map("id", "model", "type", "buttons", "label", "Model",
                "options", ["HeartMate 3", "HVAD"])
        ],
        "sentenceFn", LT_Sentence_LVAD,
        "removalNoun", (fields) => "left ventricular assist device"
    )

    defs["PADS"] := Map(
        "label", "External pacing/defibrillation pads",
        "shortLabel", "ACD pads",
        "fields", [],
        "sentenceFn", (fields) => "External pacing/defibrillation pads are present.",
        "removalNoun", (fields) => "external pacing/defibrillation pads"
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
        "removalNoun", (fields) => "retained pacing/defibrillation leads"
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
        Map("label", "Right clavicle", "phrase", "projecting over the right clavicle"),
        Map("label", "Left clavicle", "phrase", "projecting over the left clavicle")
    ]))
    arr.Push(Map("groupLabel", "Other", "states", [
        Map("label", "Other", "phrase", "__OTHER__")
    ]))
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

LT_InstanceNeedsSideWarning(instId) {
    global LT_Instances, LT_DeviceDefs
    inst := LT_Instances[instId]
    def := LT_DeviceDefs[inst["deviceKey"]]
    fields := inst["fields"]
    if (fields.Get("removed", false))
        return false
    if (!LT_DeviceHasLateralityField(def))
        return false
    return (fields.Get("laterality", "") = "")
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
    if (fields.Get("modNew", false))
        attrParts.Push("new")
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
    return sideText noun
}

LT_Sentence_ETT(fields) {
    form := fields.Get("form", "")

    if (form = "Mainstem bronchus") {
        side := fields.Get("side", "")
        sideText := (side != "") ? StrLower(side) : "_____"
        return "Endotracheal tube tip is in the " sideText " mainstem bronchus."
    }

    return "Endotracheal tube tip is _____ above the carina."
}

LT_Sentence_Enteric(fields, deviceLabel) {
    loc := fields.Get("location_phrase", "")

    if (loc = "")
        return deviceLabel " tip location not specified."

    if (loc = "__OTHER__")
        return deviceLabel " tip is _____."
    if (loc = "__STOMACH_GEN__")
        return deviceLabel " tip and side port are in the stomach."
    if (loc = "__STOMACH_PROX__")
        return deviceLabel " tip and side port are in the proximal stomach."
    if (loc = "__STOMACH_MID__")
        return deviceLabel " tip and side port are in the mid stomach."
    if (loc = "__STOMACH_DIST__")
        return deviceLabel " tip and side port are in the distal stomach."
    if (loc = "__OFFIMAGE__")
        return deviceLabel " tip is off-image below the diaphragm."
    if (loc = "__MALPOS_R__")
        return deviceLabel " tip is malpositioned in the right lower lobe bronchus."
    if (loc = "__MALPOS_L__")
        return deviceLabel " tip is malpositioned in the left lower lobe bronchus."

    return deviceLabel " tip is " loc "."
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

LT_Sentence_Trach(fields) {
    form := fields.Get("form", "")
    if (form = "At thoracic inlet")
        return "Tracheostomy tube tip is at the thoracic inlet."
    if (form = "Above carina")
        return "Tracheostomy tube tip is _____ above the carina."
    return "Tracheostomy tube tip location not specified."
}

LT_Sentence_Pleural(fields) {
    side := fields.Get("laterality", "")
    location := fields.Get("location", "")
    deviceType := fields.Get("deviceType", "")

    sideText := (side != "") ? StrLower(side) " " : ""
    typeNoun := (deviceType != "") ? StrLower(deviceType) : "chest tube"
    subject := LT_Capitalize(sideText typeNoun)

    if (location = "Apical pleural")
        locPhrase := "in the apical pleural space"
    else if (location = "Basal pleural")
        locPhrase := "in the basal pleural space"
    else if (location = "Mid pleural")
        locPhrase := "in the mid pleural space"
    else if (location = "In the chest wall")
        locPhrase := "within the chest wall"
    else
        locPhrase := "_____"

    return subject " tip is " locPhrase "."
}

LT_RemovalNoun_Pleural(fields) {
    side := fields.Get("laterality", "")
    deviceType := fields.Get("deviceType", "")
    sideText := (side != "") ? StrLower(side) " " : ""
    typeNoun := (deviceType != "") ? StrLower(deviceType) : "chest tube"
    return sideText typeNoun
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
    return out
}

LT_Sentence_Impella(fields) {
    approach := fields.Get("approach", "")
    crossing := fields.Get("crossingValve", false)

    approachText := (approach != "") ? StrLower(approach) : "_____"
    base := "Impella left ventricular assist device via " approachText " approach"
    if (crossing)
        base .= ", crossing the aortic valve plane"
    base .= "."
    return base
}

LT_Sentence_Epidural(fields) {
    phrase := fields.Get("tip_phrase", "")
    if (phrase = "")
        phrase := "_____"
    return "Epidural catheter tip is " phrase "."
}

LT_Sentence_Mediastinal(fields) {
    inferior := fields.Get("inferiorApproach", false)
    position := fields.Get("position", "")

    parts := []
    if (inferior)
        parts.Push("inferior approach")
    if (position != "")
        parts.Push(StrLower(position))

    desc := ""
    for i, p in parts
        desc .= (i > 1 ? ", " : "") p

    base := "Mediastinal drain is present"
    if (desc != "")
        base .= " (" desc ")"
    base .= "."
    return base
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
    return sideText posText "chest wall generator"
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
        return "Left ventricular assist device is present."
    return model " left ventricular assist device is present."
}

LT_JoinRemovalSentence(nouns) {
    n := nouns.Length
    if (n = 1) {
        joined := nouns[1]
        verb := "is"
    } else {
        joined := ""
        Loop n - 1 {
            joined .= nouns[A_Index]
            if (A_Index < n - 1)
                joined .= ", "
        }
        joined .= (n > 2 ? "," : "") " and " nouns[n]
        verb := "are"
    }
    return "Prior " joined " " verb " now absent."
}

LT_BuildOutput() {
    global LT_DeviceOrder, LT_InstanceOrder, LT_Instances, LT_DeviceDefs

    lines := []
    removalNouns := []

    for deviceKey in LT_DeviceOrder {
        for instId in LT_InstanceOrder {
            inst := LT_Instances[instId]
            if (inst["deviceKey"] != deviceKey)
                continue
            def := LT_DeviceDefs[deviceKey]
            fields := inst["fields"]

            if (fields.Get("removed", false)) {
                removalFn := def["removalNoun"]
                removalNouns.Push(removalFn(fields))
            } else {
                sentenceFn := def["sentenceFn"]
                s := sentenceFn(fields)
                if (s != "")
                    lines.Push(s)
            }
        }
    }

    if (removalNouns.Length > 0)
        lines.Push(LT_JoinRemovalSentence(removalNouns))

    out := ""
    for line in lines
        out .= line "`n"
    out := RTrim(out, "`n")
    return "`n" out
}


; ============================================================================
; INSTANCE MANAGEMENT
; ============================================================================

LT_AddInstance(deviceKey) {
    global LT_Instances, LT_InstanceOrder, LT_InstanceCounter

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
        return def["label"]

    for i, id in sameKeyIds {
        if (id = instId)
            return def["label"] " #" i
    }
    return def["label"]
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
    curVal := fields.Get(fieldId, "")

    isActive := false
    for st in grp["states"] {
        if (st["label"] = curVal) {
            isActive := true
            break
        }
    }

    if (isActive) {
        fields[fieldId] := ""
        fields[fieldId "_phrase"] := ""
    } else {
        st1 := grp["states"][1]
        fields[fieldId] := st1["label"]
        fields[fieldId "_phrase"] := st1["phrase"]
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

LT_CopyOutput(*) {
    if (LT_AnyInstanceNeedsSideWarning()) {
        MsgBox("Please select a side for all red devices.", "Missing side", 48)
        return
    }
    A_Clipboard := LT_BuildOutput()
    ToolTip("Copied to clipboard")
    SetTimer(() => ToolTip(), -1000)
}

LT_AddDeviceFromListbox(*) {
    global LT_GuiObj, LT_DeviceOrder
    lb := LT_GuiObj["LT_DeviceListBox"]
    idx := lb.Value
    if (idx = 0)
        return
    LT_AddInstance(LT_DeviceOrder[idx])
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
        for instId in LT_InstanceOrder {
            inst := LT_Instances[instId]
            if (inst["deviceKey"] != deviceKey)
                continue
            def := LT_DeviceDefs[deviceKey]
            fields := inst["fields"]

            if (fields.Get("removed", false)) {
                removalNouns.Push(def["removalNoun"](fields))
                continue
            }

            sentenceFn := def["sentenceFn"]
            s := sentenceFn(fields)
            if (s = "")
                continue

            approxRows := Ceil(StrLen(s) / 50)
            h := Max(20, approxRows * 16 + 6)
            lineColor := LT_InstanceNeedsSideWarning(instId) ? "cRed" : "cBlue"
            txt := LT_GuiObj.Add("Text", "x" LT_RightX " y" y " w" LT_RightW " h" h " " lineColor, s)
            txt.OnEvent("Click", LT_OpenInstanceFromOutput.Bind(instId))
            LT_RightControls.Push(txt)
            y += h + 8
        }
    }

    if (removalNouns.Length > 0) {
        s := LT_JoinRemovalSentence(removalNouns)
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

LT_DeviceLabelList() {
    global LT_DeviceOrder, LT_DeviceDefs
    labels := []
    for key in LT_DeviceOrder {
        def := LT_DeviceDefs[key]
        labels.Push(def.Get("shortLabel", def["label"]))
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
                isActive := false
                for st in grp["states"] {
                    if (st["label"] = curVal) {
                        isActive := true
                        break
                    }
                }
                try ctrl.Text := (isActive ? "> " : "") famLabel
            } else if (kind = "state") {
                lbl := entry["meta"]["label"]
                selected := (curVal = lbl)
                try ctrl.Text := (selected ? "v " : "") lbl
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

        bx := x
        justRevealed := false
        for gi, grp in fieldDef["groups"] {
            isSelected := false
            for st in grp["states"] {
                if (st["label"] = curVal) {
                    isSelected := true
                    break
                }
            }

            famLabel := (grp["states"].Length = 1) ? grp["states"][1]["label"] : grp["groupLabel"]
            btnW := LT_TextButtonWidth("> " famLabel)
            if (bx + btnW > x + w) {
                bx := x
                y += 24
            }
            btn := LT_AddMid(g, "Button", "x" bx " y" y " w" btnW " h22", (isSelected ? "> " : "") famLabel)
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
                    sBtnW := LT_TextButtonWidth("v " st["label"])
                    if (sbx + sBtnW > subX + subW) {
                        sbx := subX
                        y += 22
                    }
                    sbtn := LT_AddMid(g, "Button", "x" sbx " y" y " w" sBtnW " h20", (stSelected ? "v " : "") st["label"])
                    sbtn.SetFont("s8")
                    sbtn.OnEvent("Click", LT_GroupStateClick.Bind(instId, id, st["label"], st["phrase"]))
                    LT_RegisterFieldControl(instId, id, sbtn, Map("kind", "state", "label", st["label"]))
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

; Builds the window shell exactly once. Left/right columns are static;
; only the middle column's contents change after this.
LT_EnsureGui() {
    global LT_GuiObj, LT_MidX, LT_MidW, LT_RightX, LT_RightW
    global LT_ListBoxCtrl, LT_AddBtnCtrl, LT_ClearBtnCtrl, LT_NewPatientBtnCtrl
    global LT_MidLabelCtrl, LT_RightLabelCtrl, LT_CopyBtnCtrl
    global LT_MidPanel, LT_MidPanelX, LT_MidPanelY, LT_MidPanelW, LT_MidPanelH

    if (IsObject(LT_GuiObj))
        return

    LT_GuiObj := Gui("+Resize", "Lines and Tubes")
    LT_GuiObj.SetFont("s9", "Segoe UI")

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

    LT_GuiObj.OnEvent("Close", (*) => LT_GuiObj.Hide())
    LT_GuiObj.OnEvent("Size", LT_OnResize)
    LT_GuiObj.Show("w" (LT_RightX + LT_RightW + 20) " h580")
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
}
