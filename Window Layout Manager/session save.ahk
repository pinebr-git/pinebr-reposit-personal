#Requires AutoHotkey v2.0
#SingleInstance Force
#Warn All
#Warn LocalSameAsGlobal, Off

if !A_IsAdmin {
    Run('*RunAs "' A_ScriptFullPath '"')
    ExitApp
}

scriptDir := A_ScriptDir
layoutDir := scriptDir "\layouts"
settingsDir := scriptDir "\settings"

if !DirExist(layoutDir)
    DirCreate(layoutDir)
if !DirExist(settingsDir)
    DirCreate(settingsDir)

quickFile := layoutDir "\quicksave.jsonl"
quick1 := layoutDir "\quicksave_1.jsonl"
quick2 := layoutDir "\quicksave_2.jsonl"
restoreLogFile := layoutDir "\restore_log.txt"
legacyRestoreLogFile := layoutDir "\logs\restore_log.txt"
if FileExist(legacyRestoreLogFile) && !FileExist(restoreLogFile) {
    try FileMove(legacyRestoreLogFile, restoreLogFile, 1)
}
dumpPath := settingsDir "\window_dump.txt"
uiFont := "Segoe UI"
uiFontSize := 11
menuHoverState := 0

settingsIniFile := settingsDir "\settings.ini"
SHORTCUT_RULES := []
if FileExist(settingsIniFile)
    SHORTCUT_RULES := LoadShortcutRules(settingsIniFile)

BROWSER_EXE := CsvToExeMap("brave.exe,brave-beta.exe,brave-dev.exe,msedge.exe,msedgebeta.exe,msedgedev.exe,msedgecanary.exe,whale.exe,firefox.exe,firefox-beta.exe,firefox-nightly.exe,librewolf.exe,waterfox.exe,tor.exe,chrome.exe,chrome-beta.exe,chrome-dev.exe,chromium.exe,vivaldi.exe,opera.exe,launcher.exe,arc.exe,zen.exe,floorp.exe")
SHORT_WAIT_EXE := CloneBoolMap(BROWSER_EXE)
SHORT_WAIT_EXE["notepad.exe"] := true
SHORT_WAIT_EXE["notepad++.exe"] := true
SHORT_WAIT_EXE["sumatrapdf.exe"] := true
LONG_WAIT_EXE := CsvToExeMap("winword.exe,excel.exe,powerpnt.exe,outlook.exe,code.exe,devenv.exe,idea64.exe,hwp.exe,hword.exe,hpdf.exe,photoshop.exe,illustrator.exe,afterfx.exe,premierepro.exe,resolve.exe")
TIMEOUT_BY_EXE := Map("hwp.exe", 9000, "hword.exe", 9000)
FILE_EXTENSIONS := DefaultFileExtensions()

EXCLUDE_EXE := Map(
    "nvidia overlay.exe", true,
    "applicationframehost.exe", true,
    "autohotkey.exe", true,
    "autohotkey64.exe", true
)
if FileExist(settingsIniFile) {
    loadedExcludes := LoadExcludeExe(settingsIniFile)
    if loadedExcludes.Count
        EXCLUDE_EXE := loadedExcludes
}

if FileExist(settingsIniFile) {
    FILE_EXTENSIONS := LoadFileExtensions(settingsIniFile)
    LoadWaitRules(settingsIniFile, &SHORT_WAIT_EXE, &LONG_WAIT_EXE, &TIMEOUT_BY_EXE)
} else {
    FILE_EXTENSIONS := DefaultFileExtensions()
    SaveWaitRules(settingsIniFile, SHORT_WAIT_EXE, LONG_WAIT_EXE, TIMEOUT_BY_EXE)
    SaveFileExtensions(settingsIniFile, FILE_EXTENSIONS)
    SaveExcludeExe(settingsIniFile, EXCLUDE_EXE)
    SaveShortcutRules(settingsIniFile, SHORTCUT_RULES)
}
FILE_PATH_REGEX := BuildFilePathRegex(FILE_EXTENSIONS)

EXCLUDE_CLASS := Map(
    "consolewindowclass", true,
    "cascadia_hosting_window_class", true,
    "windowsterminalwindow", true,
    "progman", true
)

TIMEOUT_SHORT := 2000
TIMEOUT_NORMAL := 3000
TIMEOUT_LONG := 5000
POLL_SHORT := 80
POLL_NORMAL := 100
POLL_LONG := 120

lastTime := "(No Save)"
if FileExist(quickFile) {
    lastTime := FormatDateEnglish(FileGetTime(quickFile, "M"))
}

while true {
    if FileExist(quickFile)
        lastTime := FormatDateEnglish(FileGetTime(quickFile, "M"))
    else
        lastTime := "(No Save)"

    sel := ShowMainMenu(lastTime)
    if !sel
        break

    switch sel {
        case "1":
            if !FileExist(quickFile) {
                MsgBox("Quicksave file not found.")
                continue
            }
            if Confirm("Restore quicksave?")
                RestoreLayout(quickFile)
            continue

        case "2":
            if Confirm("Save current layout to quicksave?") {
                RollBackup()
                SaveLayout(quickFile)
                MsgBox("Quicksave saved.")
            }
            continue

        case "3":
            name := ManualRestoreSelect(layoutDir)
            if !name
                continue
            layoutPath := layoutDir "\" name
            if FileExist(layoutPath) {
                t := FileGetTime(layoutPath, "M")
                if Confirm("File: " name "`nSaved at: " t "`nRestore this layout?")
                    RestoreLayout(layoutPath)
            } else {
                MsgBox("File not found.")
            }
            continue

        case "4":
            nameInfo := ManualSaveSelectName(layoutDir)
            if !IsObject(nameInfo)
                continue
            name := nameInfo["name"]
            overwriteApproved := nameInfo["overwriteApproved"]
            if !IsValidFileName(name) {
                MsgBox("Invalid characters in filename.`n\ / : * ? " Chr(34) " < > |")
                continue
            }
            if !RegExMatch(name, "i)\.jsonl$")
                name .= ".jsonl"
            layoutPath := layoutDir "\" name
            if FileExist(layoutPath) {
                if !overwriteApproved {
                    t := FileGetTime(layoutPath, "M")
                    if !Confirm("File already exists.`nModified: " t "`nOverwrite?")
                        continue
                }
            }
            SaveLayout(layoutPath)
            MsgBox("Saved successfully.")
            continue

        case "5":
            name := SelectLayoutFile(layoutDir)
            if !name
                continue
            layoutPath := layoutDir "\" name
            if !FileExist(layoutPath) {
                MsgBox("File not found.")
                continue
            }
            exeFilter := SelectExeFilterForLayout(layoutPath)
            if !IsObject(exeFilter)
                continue
            if exeFilter.Count = 0 {
                MsgBox("No processes selected.")
                continue
            }
            if Confirm("Restore selected processes only?")
                RestoreLayout(layoutPath, exeFilter)
            continue

        case "6":
            ShowSettingsMenu()
            continue
    }
}
ExitApp

RollBackup() {
    global quickFile, quick1, quick2
    try {
        if FileExist(quick2)
            FileDelete(quick2)
        if FileExist(quick1)
            FileMove(quick1, quick2, 1)
        if FileExist(quickFile)
            FileMove(quickFile, quick1, 1)
    }
}

Confirm(msg) {
    ans := MsgBox(msg, "Confirm", "YesNo")
    return ans = "Yes"
}

ShowMainMenu(lastTime) {
    global uiFont, uiFontSize, menuHoverState, layoutDir
    menu := Gui("-DPIScale +AlwaysOnTop", "Layout Manager")
    menu.BackColor := "F3F6FA"
    menu.SetFont("s" uiFontSize, uiFont)
    menu.Add("Text", "xm ym c1F2937 BackgroundTrans", "Select a Task")
    labels := [
        "1) Quick Restore (Last: " lastTime ")",
        "2) Quick Save (Overwrite Quicksave)",
        "3) Manual Restore",
        "4) Manual Save",
        "5) Selective Restore (By Process)",
        "6) Settings"
    ]
    state := Map("sel", 0)
    rows := []
    for i, line in labels {
        row := menu.Add("Text", "xm y+6 w760 h36 +0x100 +0x200 Border BackgroundFFFFFF c111827", "  " line)
        row.OnEvent("Click", MenuSelect.Bind(state, menu, i, row))
        rows.Push(row)
    }
    openBtn := menu.Add("Button", "xm y+12 w150", "Open Explorer")
    openBtn.OnEvent("Click", (*) => Run('explorer.exe "' layoutDir '"'))
    cancelBtn := menu.Add("Button", "x650 yp w120", "Exit")
    cancelBtn.OnEvent("Click", (*) => menu.Destroy())
    menu.OnEvent("Close", (*) => menu.Destroy())
    menu.OnEvent("Escape", (*) => menu.Destroy())
    menuHoverState := Map("menuHwnd", menu.Hwnd, "rows", rows, "hovered", 0)
    OnMessage(0x200, MenuMouseMove)
    menu.Show("AutoSize Center")
    WinWaitClose("ahk_id " menu.Hwnd)
    OnMessage(0x200, MenuMouseMove, 0)
    menuHoverState := 0
    return state["sel"] ? state["sel"] "" : ""
}

ShowSettingsMenu() {
    global settingsDir
    while true {
        idx := ShowSelectableList("Settings", "Select a setting", [
            "1) Shortcut Rules",
            "2) Wait Rules",
            "3) File Extensions",
            "4) Exclusion Rules",
            "5) Dump Windows Info"
        ], "Back", true)
        if !idx
            return
        switch idx {
            case 1:
                ManageShortcutRules()
            case 2:
                ManageWaitRules()
            case 3:
                ManageFileExtensions()
            case 4:
                ManageExcludeRules()
            case 5:
                dumpPath := settingsDir "\window_dump.txt"
                DumpWindows(dumpPath)
                try Run('"' dumpPath '"')
                MsgBox("Dump completed.")
        }
    }
}

SelectLayoutFile(baseDir) {
    layouts := GetManualLayouts(baseDir)
    labels := []
    for item in layouts {
        displayName := item["name"]
        if IsQuicksaveFile(item["name"])
            displayName := "* " displayName
        labels.Push(displayName " (" FormatDateEnglish(item["modified"]) ")")
    }
    if layouts.Length = 0 {
        MsgBox("No save files found.")
        return ""
    }
    idx := ShowSelectableList("Manual Restore", "Select a file to restore", labels, "Back")
    return idx ? layouts[idx]["name"] : ""
}

GetManualLayouts(baseDir) {
    layouts := []
    Loop Files baseDir "\*.jsonl", "F"
        layouts.Push(Map("name", A_LoopFileName, "modified", FileGetTime(A_LoopFileFullPath, "M")))
    SortLayoutsForManual(layouts)
    return layouts
}

SortLayoutsForManual(layouts) {
    n := layouts.Length
    if n < 2
        return
    i := 1
    while (i < n) {
        minIdx := i
        j := i + 1
        while (j <= n) {
            if LayoutEntryLess(layouts[j], layouts[minIdx])
                minIdx := j
            j += 1
        }
        if (minIdx != i) {
            tmp := layouts[i]
            layouts[i] := layouts[minIdx]
            layouts[minIdx] := tmp
        }
        i += 1
    }
}

LayoutEntryLess(a, b) {
    aq := IsQuicksaveFile(a["name"])
    bq := IsQuicksaveFile(b["name"])
    if aq && !bq
        return true
    if !aq && bq
        return false
    return StrCompare(StrLower(a["name"]), StrLower(b["name"])) < 0
}

IsQuicksaveFile(name) {
    return StrLower(name) = "quicksave.jsonl"
}

ManualRestoreSelect(baseDir) {
    while true {
        res := ShowLayoutActionDialog(baseDir, "restore")
        if !IsObject(res)
            return ""
        action := res["action"]
        if action = "back"
            return ""
        if action = "restore"
            return res["name"]
        if action = "rename" {
            RenameLayoutFile(baseDir, res["name"])
            continue
        }
        if action = "delete" {
            DeleteLayoutFile(baseDir, res["name"])
            continue
        }
    }
}

ManualSaveSelectName(baseDir) {
    while true {
        res := ShowLayoutActionDialog(baseDir, "save")
        if !IsObject(res)
            return 0
        action := res["action"]
        if action = "back"
            return 0
        if action = "save_new" {
            name := EnsureJsonlName(res["input"])
            if !IsValidFileName(name) {
                MsgBox("Invalid characters in filename.")
                continue
            }
            return Map("name", name, "overwriteApproved", false)
        }
        if action = "save_overwrite" {
            if Confirm("Overwrite selected file?`n" res["name"])
                return Map("name", res["name"], "overwriteApproved", true)
            continue
        }
        if action = "rename" {
            RenameLayoutFile(baseDir, res["name"])
            continue
        }
        if action = "delete" {
            DeleteLayoutFile(baseDir, res["name"])
            continue
        }
    }
}

ShowLayoutActionDialog(baseDir, mode) {
    global uiFont, uiFontSize
    layouts := GetManualLayouts(baseDir)
    if layouts.Length = 0 && mode = "restore" {
        MsgBox("No save files found.")
        return 0
    }
    dialogTitle := (mode = "restore") ? "Manual Restore" : "Manual Save"
    header := (mode = "restore") ? "Select a file to restore" : "Manage files or enter a new name"
    dialog := Gui("-DPIScale +AlwaysOnTop", dialogTitle)
    dialog.BackColor := "F3F6FA"
    dialog.SetFont("s" uiFontSize, uiFont)
    listW := 760
    renameW := 90
    deleteW := 72
    btnGap := 6
    dialog.Add("Text", "xm ym c1F2937 BackgroundTrans", header)
    state := Map("action", "back", "name", "", "input", "")
    rowToName := Map()
    if mode = "save" {
        lv := dialog.Add("ListView", "xm y+8 w" listW " r12", ["Filename", "Saved At"])
        lv.ModifyCol(1, 430), lv.ModifyCol(2, 300)
        for item in layouts {
            displayName := IsQuicksaveFile(item["name"]) ? "* " item["name"] : item["name"]
            row := lv.Add("", displayName, FormatDateEnglish(item["modified"]))
            rowToName[row] := item["name"]
        }
        lv.Modify(1, "Select Focus")
        renameBtn := dialog.Add("Button", "x" (listW - (renameW + btnGap + deleteW)) " y+8 w" renameW " h24", "Rename")
        deleteBtn := dialog.Add("Button", "x+" btnGap " yp w" deleteW " h24", "Delete")
        dialog.Add("Text", "xm y+10", "New Name:")
        input := dialog.Add("Edit", "x+8 yp-3 w540")
        dialog.Add("Text", "x+8 yp+3 c374151 BackgroundTrans", ".jsonl")
        backBtn := dialog.Add("Button", "xm y+12 w120", "Back")
        overwriteBtn := dialog.Add("Button", "x510 yp w120", "Overwrite")
        saveBtn := dialog.Add("Button", "x+8 yp w120 Default", "Save")
        renameBtn.OnEvent("Click", LayoutDialogPickAction.Bind(state, dialog, lv, rowToName, "rename"))
        deleteBtn.OnEvent("Click", LayoutDialogPickAction.Bind(state, dialog, lv, rowToName, "delete"))
        overwriteBtn.OnEvent("Click", LayoutDialogPickAction.Bind(state, dialog, lv, rowToName, "save_overwrite"))
        saveBtn.OnEvent("Click", LayoutDialogInputAction.Bind(state, dialog, input, "save_new"))
        backBtn.OnEvent("Click", (*) => dialog.Destroy())
    } else {
        lv := dialog.Add("ListView", "xm y+8 w" listW " r14", ["Filename", "Saved At"])
        lv.ModifyCol(1, 430), lv.ModifyCol(2, 300)
        for item in layouts {
            displayName := IsQuicksaveFile(item["name"]) ? "* " item["name"] : item["name"]
            row := lv.Add("", displayName, FormatDateEnglish(item["modified"]))
            rowToName[row] := item["name"]
        }
        lv.Modify(1, "Select Focus")
        renameBtn := dialog.Add("Button", "x" (listW - (renameW + btnGap + deleteW)) " y+8 w" renameW " h24", "Rename")
        deleteBtn := dialog.Add("Button", "x+" btnGap " yp w" deleteW " h24", "Delete")
        backBtn := dialog.Add("Button", "xm y+12 w120", "Back")
        actionBtn := dialog.Add("Button", "x560 yp w200 Default", "Restore")
        renameBtn.OnEvent("Click", LayoutDialogPickAction.Bind(state, dialog, lv, rowToName, "rename"))
        deleteBtn.OnEvent("Click", LayoutDialogPickAction.Bind(state, dialog, lv, rowToName, "delete"))
        backBtn.OnEvent("Click", (*) => dialog.Destroy())
        actionBtn.OnEvent("Click", LayoutDialogPickAction.Bind(state, dialog, lv, rowToName, "restore"))
    }
    dialog.OnEvent("Close", (*) => dialog.Destroy())
    dialog.OnEvent("Escape", (*) => dialog.Destroy())
    dialog.Show("AutoSize Center")
    WinWaitClose("ahk_id " dialog.Hwnd)
    return state
}

LayoutDialogPickAction(state, dialog, lv, rowToName, action, *) {
    row := lv.GetNext(0, "F")
    if !row
        row := lv.GetNext(0)
    if !row || !rowToName.Has(row) {
        MsgBox("Please select a file from the list.")
        return
    }
    state["action"] := action
    state["name"] := rowToName[row]
    dialog.Destroy()
}

LayoutDialogInputAction(state, dialog, input, action, *) {
    v := Trim(input.Value)
    if !v {
        MsgBox("Please enter a filename.")
        return
    }
    state["action"] := action
    state["input"] := v
    dialog.Destroy()
}

RenameLayoutFile(baseDir, oldName) {
    oldPath := baseDir "\" oldName
    if !FileExist(oldPath) {
        MsgBox("File not found.")
        return false
    }
    defaultName := RegExReplace(oldName, "i)\.jsonl$")
    newName := PromptInput("Rename", "Original: " oldName, defaultName, ".jsonl", "Cancel", "New Name:")
    if !newName
        return false
    newName := EnsureJsonlName(newName)
    if !IsValidFileName(newName) {
        MsgBox("Invalid characters in filename.")
        return false
    }
    if StrLower(newName) = StrLower(oldName)
        return false
    newPath := baseDir "\" newName
    if FileExist(newPath) && !Confirm("File already exists. Overwrite?")
        return false
    if !Confirm("Rename this file?`n" oldName " -> " newName)
        return false
    try {
        FileMove(oldPath, newPath, 1)
        return true
    } catch as err {
        MsgBox("Rename failed.`n" err.Message)
        return false
    }
}

DeleteLayoutFile(baseDir, name) {
    path := baseDir "\" name
    if !FileExist(path) {
        MsgBox("File not found.")
        return false
    }
    if !Confirm("Delete this file?`n" name)
        return false
    try {
        FileDelete(path)
        return true
    } catch as err {
        MsgBox("Delete failed.`n" err.Message)
        return false
    }
}

EnsureJsonlName(name) {
    name := Trim(name)
    if !RegExMatch(name, "i)\.jsonl$")
        name .= ".jsonl"
    return name
}

ShowSelectableList(title, header, items, cancelText := "Cancel", showExit := false) {
    global uiFont, uiFontSize, menuHoverState
    menu := Gui("-DPIScale +AlwaysOnTop", title)
    menu.BackColor := "F3F6FA"
    menu.SetFont("s" uiFontSize, uiFont)
    listW := 760
    btnW := 120
    menu.Add("Text", "xm ym c1F2937 BackgroundTrans", header)
    state := Map("sel", 0)
    rows := []
    starRows := Map()
    for i, line in items {
        baseOpt := "xm y+6 w" listW " h36 +0x100 +0x200 Border "
        if SubStr(line, 1, 2) = "* " {
            row := menu.Add("Text", baseOpt "BackgroundFFF9E8 c7A4B00", "  " line)
            starRows[i] := true
        } else
            row := menu.Add("Text", baseOpt "BackgroundFFFFFF c111827", "  " line)
        row.OnEvent("Click", MenuSelect.Bind(state, menu, i, row))
        rows.Push(row)
    }
    cancelBtn := menu.Add("Button", "xm y+12 w" btnW, cancelText)
    cancelBtn.OnEvent("Click", (*) => menu.Destroy())
    if showExit {
        exitBtn := menu.Add("Button", "x" (listW - btnW) " yp w" btnW, "Exit")
        exitBtn.OnEvent("Click", (*) => ExitApp())
    }
    menu.OnEvent("Close", (*) => menu.Destroy())
    menu.OnEvent("Escape", (*) => menu.Destroy())
    menuHoverState := Map("menuHwnd", menu.Hwnd, "rows", rows, "hovered", 0, "starRows", starRows)
    OnMessage(0x200, MenuMouseMove)
    menu.Show("AutoSize Center")
    WinWaitClose("ahk_id " menu.Hwnd)
    OnMessage(0x200, MenuMouseMove, 0)
    menuHoverState := 0
    return state["sel"]
}

SelectExeFilterForLayout(layoutPath) {
    global uiFont, uiFontSize
    items := ParseJsonl(layoutPath)
    if items.Length = 0 {
        MsgBox("Nothing to restore.")
        return 0
    }
    exeCount := Map()
    for item in items {
        exeName := item["exeName"] ? item["exeName"] : "(unknown)"
        exeCount[exeName] := exeCount.Has(exeName) ? exeCount[exeName] + 1 : 1
    }
    exes := []
    for exeName, cnt in exeCount
        exes.Push(Map("exeName", exeName, "count", cnt))
    SortExeEntriesByCountDesc(exes)
    dialog := Gui("-DPIScale +AlwaysOnTop", "Selective Restore")
    dialog.BackColor := "F3F6FA"
    dialog.SetFont("s" uiFontSize, uiFont)
    dialog.Add("Text", "xm ym c1F2937 BackgroundTrans", "Select processes to restore (Default: All)")
    dialog.Add("Text", "xm y+4 c6B7280 BackgroundTrans", "Total items: " items.Length " / Processes: " exes.Length)
    lv := dialog.Add("ListView", "xm y+8 w760 r14 Checked -Multi", ["Restore", "Process", "Count"])
    lv.ModifyCol(1, 60), lv.ModifyCol(2, 540), lv.ModifyCol(3, 120)
    rowExeMap := Map()
    for entry in exes {
        displayExe := GetExeDisplayName(entry["exeName"])
        row := lv.Add("Check", "Y", displayExe, entry["count"])
        rowExeMap[row] := entry["exeName"]
        lv.Modify(row, "Check")
    }
    state := Map("ok", false, "selected", Map())
    allBtn := dialog.Add("Button", "xm y+10 w110", "Select All")
    noneBtn := dialog.Add("Button", "x+8 yp w110", "Deselect All")
    cancelBtn := dialog.Add("Button", "x+320 yp w100", "Back")
    okBtn := dialog.Add("Button", "x+8 yp w120 Default", "Restore")
    allBtn.OnEvent("Click", (*) => SelectAllListViewRows(lv, true))
    noneBtn.OnEvent("Click", (*) => SelectAllListViewRows(lv, false))
    cancelBtn.OnEvent("Click", (*) => dialog.Destroy())
    okBtn.OnEvent("Click", (*) => (state["ok"] := true, state["selected"] := CollectCheckedExeNames(lv, rowExeMap), dialog.Destroy()))
    dialog.OnEvent("Close", (*) => dialog.Destroy())
    dialog.OnEvent("Escape", (*) => dialog.Destroy())
    dialog.Show("AutoSize Center")
    WinWaitClose("ahk_id " dialog.Hwnd)
    if !state["ok"]
        return 0
    return state["selected"]
}

SelectAllListViewRows(lv, checked := true) {
    option := checked ? "Check" : "-Check"
    loop lv.GetCount()
        lv.Modify(A_Index, option)
}

SortExeEntriesByCountDesc(entries) {
    n := entries.Length
    if n < 2
        return
    i := 1
    while (i < n) {
        maxIdx := i
        j := i + 1
        while (j <= n) {
            left := entries[j]["count"]
            right := entries[maxIdx]["count"]
            if (left > right) || (left = right && StrCompare(entries[j]["exeName"], entries[maxIdx]["exeName"]) < 0)
                maxIdx := j
            j += 1
        }
        if maxIdx != i {
            tmp := entries[i], entries[i] := entries[maxIdx], entries[maxIdx] := tmp
        }
        i += 1
    }
}

CollectCheckedExeNames(lv, rowExeMap) {
    selected := Map()
    row := 0
    while row := lv.GetNext(row, "C") {
        if !rowExeMap.Has(row)
            continue
        selected[rowExeMap[row]] := true
    }
    return selected
}

GetExeDisplayName(exeName) {
    n := StrLower(exeName)
    if n = "code.exe"
        return "VSCode"
    if n = "pdfxedit.exe"
        return "PDF-XChange Editor"
    return exeName
}

MenuSelect(state, menu, value, row, *) {
    row.Opt("BackgroundCBD5E1 c0B1220")
    Sleep(90)
    state["sel"] := value
    menu.Destroy()
}

MenuMouseMove(*) {
    global menuHoverState
    if !IsObject(menuHoverState)
        return
    MouseGetPos(, , &winHwnd, &ctrlHwnd, 2)
    newHovered := 0
    if (winHwnd = menuHoverState["menuHwnd"]) {
        for i, row in menuHoverState["rows"] {
            if (ctrlHwnd = row.Hwnd) {
                newHovered := i
                break
            }
        }
    }
    oldHovered := menuHoverState["hovered"]
    if (newHovered = oldHovered)
        return
    if oldHovered {
        oldRow := menuHoverState["rows"][oldHovered]
        if menuHoverState.Has("starRows") && menuHoverState["starRows"].Has(oldHovered)
            oldRow.Opt("BackgroundFFF9E8 c7A4B00")
        else
            oldRow.Opt("BackgroundFFFFFF c111827")
    }
    if newHovered {
        newRow := menuHoverState["rows"][newHovered]
        newRow.Opt("BackgroundDDE5F2 c0B1220")
    }
    menuHoverState["hovered"] := newHovered
}

FormatDateEnglish(ts) {
    if StrLen(ts) < 14
        return ts
    hour24 := Integer(SubStr(ts, 9, 2))
    ampm := hour24 < 12 ? "AM" : "PM"
    hour12 := Mod(hour24, 12)
    if hour12 = 0
        hour12 := 12
    return SubStr(ts, 1, 4) "-" SubStr(ts, 5, 2) "-" SubStr(ts, 7, 2) " " hour12 ":" SubStr(ts, 11, 2) " " ampm
}

PromptInput(title, message, defaultValue := "", suffix := "", cancelText := "Cancel", inputLabel := "Input:") {
    global uiFont, uiFontSize
    dialog := Gui("-DPIScale +AlwaysOnTop", title)
    dialog.SetFont("s" uiFontSize, uiFont)
    lineCount := StrSplit(message, "`n").Length + 1
    if lineCount < 6
        lineCount := 6
    dialog.Add("Edit", "xm ym w680 r" lineCount " ReadOnly -Wrap", message)
    dialog.Add("Text", "xm y+10", inputLabel)
    inputWidth := suffix ? 540 : 620
    input := dialog.Add("Edit", "x+8 yp-3 w" inputWidth, defaultValue)
    if suffix
        dialog.Add("Text", "x+8 yp+3 c374151 BackgroundTrans", suffix)
    cancelBtn := dialog.Add("Button", "xm y+12 w120", cancelText)
    okBtn := dialog.Add("Button", "x560 yp w120 Default", "OK")
    state := Map("ok", false, "value", "")
    okBtn.OnEvent("Click", (*) => (state["ok"] := true, state["value"] := Trim(input.Value), dialog.Destroy()))
    cancelBtn.OnEvent("Click", (*) => dialog.Destroy())
    dialog.OnEvent("Close", (*) => dialog.Destroy())
    dialog.OnEvent("Escape", (*) => dialog.Destroy())
    dialog.Show("AutoSize Center")
    input.Focus()
    WinWaitClose("ahk_id " dialog.Hwnd)
    return state["ok"] ? state["value"] : ""
}

IsValidFileName(name) {
    invalidChars := "[\\/:*?" Chr(34) "<>|]"
    return !RegExMatch(name, invalidChars)
}

SaveLayout(layoutPath) {
    global EXCLUDE_EXE, EXCLUDE_CLASS, SHORTCUT_RULES
    lines := ""
    for hwnd in WinGetList() {
        try {
            title := WinGetTitle("ahk_id " hwnd)
            if !title
                continue
            cls := WinGetClass("ahk_id " hwnd)
            if EXCLUDE_CLASS.Has(StrLower(cls))
                continue
            exePath := WinGetProcessPath("ahk_id " hwnd)
            exeName := ""
            if exePath
                SplitPath(exePath, &exeName)
            exeName := StrLower(exeName)
            if EXCLUDE_EXE.Has(exeName) || IsCliExe(exeName)
                continue
            if (WinGetExStyle("ahk_id " hwnd) & 0x80)
                continue
            mm := WinGetMinMax("ahk_id " hwnd)
            WinGetPos(&x, &y, &w, &h, "ahk_id " hwnd)
            mon := MonitorFromPoint(x + w // 2, y + h // 2)
            wa := GetWorkArea(mon)
            rx := (x - wa.L) / (wa.R - wa.L), ry := (y - wa.T) / (wa.B - wa.T)
            rw := w / (wa.R - wa.L), rh := h / (wa.B - wa.T)
            extraType := "", extraVal := ""
            if (cls = "CabinetWClass" || cls = "ExploreWClass") {
                p := GetExplorerPath(hwnd)
                if p
                    TrySetExtra(&extraType, &extraVal, "explorerPath", p)
            }
            if exeName = "winword.exe" {
                doc := GetWordDocFullNameByHwnd(hwnd, title)
                if doc
                    TrySetExtra(&extraType, &extraVal, "wordDoc", doc)
            }
            if exeName = "excel.exe" {
                doc := GetExcelDocFullNameByTitle(title)
                if doc
                    TrySetExtra(&extraType, &extraVal, "excelDoc", doc)
            }
            if exeName = "powerpnt.exe" {
                doc := GetPowerPointDocFullNameByTitle(title)
                if doc
                    TrySetExtra(&extraType, &extraVal, "pptDoc", doc)
            }
            if (exeName = "hwp.exe" || exeName = "hword.exe") {
                doc := ExtractHwpDocPath(title)
                if doc
                    TrySetExtra(&extraType, &extraVal, "hwpDoc", doc)
            }
            if exeName = "hpdf.exe" {
                pdfPath := ExtractHanpdfDocPath(title)
                if pdfPath
                    TrySetExtra(&extraType, &extraVal, "hpdfDoc", pdfPath)
            }
            genericPath := ExtractGenericFilePathFromTitle(title)
            if genericPath
                TrySetExtra(&extraType, &extraVal, "filePath", genericPath)
            for rule in SHORTCUT_RULES {
                if (exeName = rule["exe"]) && InStr(title, rule["key"]) {
                    if FileExist(rule["targetPath"])
                        TrySetExtra(&extraType, &extraVal, "shortcutPath", rule["targetPath"])
                }
            }
            lines .= Format('{{"exe":"{1}","exeName":"{2}","class":"{3}","title":"{4}","mon":{5},"rx":{6},"ry":{7},"rw":{8},"rh":{9},"state":{10},"extraType":"{11}","extraVal":"{12}"}}`n',
                Esc(exePath), exeName, Esc(cls), Esc(title), mon, rx, ry, rw, rh, mm, extraType, Esc(extraVal))
        }
    }
    try FileDelete(layoutPath)
    FileAppend(lines, layoutPath, "UTF-8")
}

RestoreLayout(layoutPath, exeFilter := 0) {
    global SHORT_WAIT_EXE, LONG_WAIT_EXE
    global TIMEOUT_SHORT, TIMEOUT_NORMAL, TIMEOUT_LONG, TIMEOUT_BY_EXE
    if !FileExist(layoutPath) {
        MsgBox("File not found.")
        return
    }
    items := ParseJsonl(layoutPath)
    if IsObject(exeFilter) {
        filtered := []
        for item in items {
            if exeFilter.Has(item["exeName"])
                filtered.Push(item)
        }
        items := filtered
    }
    if items.Length = 0 {
        MsgBox("Nothing to restore.")
        return
    }
    AppendRestoreLog("START | file=" layoutPath)
    restored := 0, notFound := 0, errorCount := 0
    usedHwnd := Map(), launchedBrowserExe := Map(), manualNotices := [], manualNoticeSet := Map()
    for item in items {
        beforeSet := SnapshotWindowsByExe(item["exeName"])
        isBrowser := IsBrowserExe(item["exeName"])
        notice := BuildManualNotice(item)
        if notice && !manualNoticeSet.Has(notice) {
            manualNoticeSet[notice] := true
            manualNotices.Push(notice)
        }
        pid := 0
        if item["extraType"] = "shortcutPath" || item["extraType"] = "wordDoc" || item["extraType"] = "excelDoc" || item["extraType"] = "pptDoc" || item["extraType"] = "hwpDoc" || item["extraType"] = "hpdfDoc" || item["extraType"] = "filePath" {
            pid := SafeRun('"' item["extraVal"] '"')
        } else if item["extraType"] = "explorerPath" {
            pid := SafeRun('explorer.exe "' item["extraVal"] '"')
        } else if item["exe"] {
            if isBrowser {
                if !launchedBrowserExe.Has(item["exeName"]) {
                    pid := SafeRun('"' item["exe"] '"'), launchedBrowserExe[item["exeName"]] := true
                }
            } else pid := SafeRun('"' item["exe"] '"')
        }
        timeout := SHORT_WAIT_EXE.Has(item["exeName"]) ? TIMEOUT_SHORT : (LONG_WAIT_EXE.Has(item["exeName"]) ? TIMEOUT_LONG : TIMEOUT_NORMAL)
        if TIMEOUT_BY_EXE.Has(item["exeName"])
            timeout := TIMEOUT_BY_EXE[item["exeName"]]
        hwnd := 0
        if isBrowser {
            hwnd := WaitNewWindowFromSnapshot(item["exeName"], beforeSet, timeout, usedHwnd)
            if !hwnd && pid
                hwnd := WaitWindowByPid(pid, timeout, usedHwnd)
        } else {
            if pid
                hwnd := WaitWindowByPid(pid, timeout, usedHwnd)
            if !hwnd
                hwnd := WaitNewWindowFromSnapshot(item["exeName"], beforeSet, timeout, usedHwnd)
        }
        if !hwnd
            hwnd := FindFallbackWindow(item, usedHwnd)
        if !hwnd {
            notFound += 1
            AppendRestoreLog("NOT_FOUND | exe=" item["exeName"] " | title=" item["title"])
            continue
        }
        usedHwnd[hwnd] := true
        wa := GetWorkArea(item["mon"])
        x := Round(wa.L + item["rx"] * (wa.R - wa.L)), y := Round(wa.T + item["ry"] * (wa.B - wa.T))
        w := Round(item["rw"] * (wa.R - wa.L)), h := Round(item["rh"] * (wa.B - wa.T))
        try {
            WinRestore("ahk_id " hwnd)
            WinMove(x, y, w, h, "ahk_id " hwnd)
            if item["state"] = -1
                WinMinimize("ahk_id " hwnd)
            else if item["state"] = 1
                WinMaximize("ahk_id " hwnd)
            restored += 1
        } catch as err {
            errorCount += 1
            AppendRestoreLog("ERROR | exe=" item["exeName"] " | " err.Message)
        }
    }
    summary := (notFound + errorCount = 0) ? "Restore Completed." : "Restore finished with issues.`nNotFound: " notFound " | Error: " errorCount
    if manualNotices.Length > 0
        summary .= "`n`nPlease check manually:`n" JoinLines(manualNotices, 10)
    MsgBox(summary)
}

BuildManualNotice(item) {
    if !NeedsManualContentNotice(item)
        return ""
    exeLabel := GetExeDisplayName(item["exeName"])
    return item["title"] ? exeLabel " (" item["title"] ")" : exeLabel
}

NeedsManualContentNotice(item) {
    n := StrLower(item["exeName"])
    return (n = "code.exe" || n = "pdfxedit.exe" || (IsBrowserExe(n) && item["extraType"] != "shortcutPath"))
}

JoinLines(lines, maxShow := 12) {
    out := ""
    for i, line in lines {
        if i > maxShow
            break
        out .= "- " line "`n"
    }
    return RTrim(out, "`n")
}

GetWordDocFullNameByHwnd(hwnd, title) {
    try wd := ComObjActive("Word.Application")
    catch
        return ""
    try {
        for w in wd.Windows
            if w.Hwnd = hwnd
                return w.Document.FullName
    }
    return ""
}

AppendRestoreLog(message, maxLines := 300) {
    global restoreLogFile
    ts := FormatTime(, "yyyy-MM-dd HH:mm:ss")
    line := ts " | " message
    lines := []
    if FileExist(restoreLogFile) {
        for old in StrSplit(FileRead(restoreLogFile, "UTF-8"), "`n", "`r") {
            if Trim(old)
                lines.Push(old)
        }
    }
    lines.Push(line)
    while lines.Length > maxLines
        lines.RemoveAt(1)
    out := ""
    for row in lines
        out .= row "`n"
    try FileDelete(restoreLogFile)
    FileAppend(out, restoreLogFile, "UTF-8")
}

GetExcelDocFullNameByTitle(title) {
    try xl := ComObjActive("Excel.Application")
    catch
        return ""
    try {
        for wb in xl.Workbooks
            if InStr(title, wb.Name)
                return wb.FullName
    }
    return ""
}

GetPowerPointDocFullNameByTitle(title) {
    try pp := ComObjActive("PowerPoint.Application")
    catch
        return ""
    try {
        for p in pp.Presentations
            if InStr(title, p.Name)
                return p.FullName
    }
    return ""
}

SafeRun(cmd) {
    pid := 0
    try Run(cmd, , , &pid)
    return pid
}

IsCliExe(n) {
    n := StrLower(n)
    return n = "cmd.exe" || n = "powershell.exe" || n = "wt.exe"
}

MonitorFromPoint(x, y) {
    Loop MonitorGetCount() {
        wa := GetWorkArea(A_Index)
        if x >= wa.L && x < wa.R && y >= wa.T && y < wa.B
            return A_Index
    }
    return 1
}

GetWorkArea(mon) {
    MonitorGetWorkArea(mon, &L, &T, &R, &B)
    return {L: L, T: T, R: R, B: B}
}

GetExplorerPath(hwnd) {
    try {
        shell := ComObject("Shell.Application")
        for w in shell.Windows
            if w.HWND = hwnd
                return w.Document.Folder.Self.Path
    }
    return ""
}

ExtractHanpdfDocPath(title) {
    return ExtractBracketDocPath(title, "HanPDF")
}

ExtractHwpDocPath(title) {
    return ExtractBracketDocPath(title, "Hanword")
}

ExtractGenericFilePathFromTitle(title) {
    global FILE_PATH_REGEX
    if !title
        return ""
    if RegExMatch(title, FILE_PATH_REGEX, &m) {
        candidate := Trim(m[1], " `t")
        if FileExist(candidate)
            return candidate
        if !RegExMatch(candidate, "i)^[A-Z]:\\") && FileExist(A_WorkingDir "\\" candidate)
            return A_WorkingDir "\\" candidate
        return ""
    }
    return ""
}

ExtractBracketDocPath(title, tailKeywords := "") {
    if IsObject(tailKeywords) {
        matched := false
        for keyword in tailKeywords {
            if InStr(title, keyword) {
                matched := true
                break
            }
        }
        if !matched
            return ""
    } else if tailKeywords && !InStr(title, tailKeywords)
        return ""
    closePos := FindLastPos(title, "] - ")
    if !closePos
        return ""
    depth := 1, i := closePos - 1, startPos := 0
    while i >= 1 {
        ch := SubStr(title, i, 1)
        if ch = "]"
            depth += 1
            else if ch = "[" {
                depth -= 1
                if depth = 0 {
                    startPos := i
                    break
                }
            }
            i -= 1
    }
    if !startPos
        return ""
    path := Trim(SubStr(title, startPos + 1, closePos - startPos - 1)) "\" Trim(SubStr(title, 1, startPos - 1))
    return FileExist(path) ? path : ""
}

LoadShortcutRules(filePath) {
    rules := []
    sectionName := "ShortcutRules"
    count := Integer(IniReadSafe(filePath, sectionName, "Count", "0"))
    if !FileExist(filePath) || count <= 0
        return rules
    Loop count {
        i := A_Index
        exe := NormalizeExeName(IniReadSafe(filePath, sectionName, "Rule" i "_Exe", ""))
        key := Trim(IniReadSafe(filePath, sectionName, "Rule" i "_Key", ""))
        targetPath := Trim(IniReadSafe(filePath, sectionName, "Rule" i "_TargetPath", ""))
        if !exe || !key || !targetPath
            continue
        rules.Push(Map("exe", exe, "key", key, "targetPath", targetPath))
    }
    SortShortcutRules(rules)
    return rules
}

SaveShortcutRules(filePath, rules) {
    if !EnsureIniParentDir(filePath)
        return false
    if !IsObject(rules)
        rules := []
    try IniDelete(filePath, "ShortcutRules")
    catch
    try {
        IniWrite(rules.Length, filePath, "ShortcutRules", "Count")
        idx := 1
        for rule in rules {
            IniWrite(rule["exe"], filePath, "ShortcutRules", "Rule" idx "_Exe")
            IniWrite(rule["key"], filePath, "ShortcutRules", "Rule" idx "_Key")
            IniWrite(rule["targetPath"], filePath, "ShortcutRules", "Rule" idx "_TargetPath")
            idx += 1
        }
        return true
    } catch as err {
        MsgBox("Failed to save shortcut rules.`n" err.Message, "Settings Error", "Icon!")
        return false
    }
}

SortShortcutRules(rules) {
    n := rules.Length
    if n < 2
        return
    i := 1
    while (i < n) {
        minIdx := i
        j := i + 1
        while (j <= n) {
            left := StrLower(rules[j]["key"])
            right := StrLower(rules[minIdx]["key"])
            if StrCompare(left, right) < 0
                minIdx := j
            j += 1
        }
        if minIdx != i {
            tmp := rules[i]
            rules[i] := rules[minIdx]
            rules[minIdx] := tmp
        }
        i += 1
    }
}

CloneShortcutRules(rules) {
    copy := []
    for rule in rules
        copy.Push(Map("exe", rule["exe"], "key", rule["key"], "targetPath", rule["targetPath"]))
    return copy
}

RefreshShortcutRuleListView(state) {
    lv := state["lv"]
    rules := state["rules"]
    rowToIndex := Map()
    lv.Delete()
    for idx, rule in rules {
        row := lv.Add("", rule["exe"], rule["key"], rule["targetPath"])
        rowToIndex[row] := idx
    }
    if rules.Length > 0
        lv.Modify(1, "Select Focus")
    state["rowToIndex"] := rowToIndex
}

NormalizeExeName(exe) {
    exe := StrLower(Trim(exe))
    if !exe
        return ""
    if !RegExMatch(exe, "i)\.exe$")
        exe .= ".exe"
    return exe
}

ShortcutRuleExists(rules, candidate) {
    for rule in rules {
        if StrLower(rule["exe"]) = StrLower(candidate["exe"])
            && rule["key"] = candidate["key"]
            && StrLower(rule["targetPath"]) = StrLower(candidate["targetPath"])
            return true
    }
    return false
}

ShortcutRuleAdd(state, *) {
    exe := NormalizeExeName(PromptInput("Add Shortcut Rule", "Enter executable name.", "chrome.exe", "", "Cancel", "Executable:"))
    if !exe
        return
    key := Trim(PromptInput("Add Shortcut Rule", "Enter title keyword to match.", "", "", "Cancel", "Title Keyword:"))
    if !key
        return
    targetPath := FileSelect(3, A_MyDocuments, "Select shortcut target", "Shortcut or File (*.*)")
    if !targetPath
        return
    state["rules"].Push(Map("exe", exe, "key", key, "targetPath", targetPath))
    state["dirty"] := true
    SortShortcutRules(state["rules"])
    RefreshShortcutRuleListView(state)
}

ShortcutRuleDelete(state, *) {
    lv := state["lv"]
    row := lv.GetNext(0, "F")
    if !row
        row := lv.GetNext(0)
    if !row || !state["rowToIndex"].Has(row) {
        MsgBox("Please select a rule to delete.", "Notice", "Iconi Owner" state["dialog"].Hwnd)
        return
    }
    idx := state["rowToIndex"][row]
    rule := state["rules"][idx]
    if !Confirm("Delete this rule?`n" rule["exe"] " | " rule["key"])
        return
    state["rules"].RemoveAt(idx)
    state["dirty"] := true
    RefreshShortcutRuleListView(state)
}

ShortcutRuleEdit(state, *) {
    lv := state["lv"]
    row := lv.GetNext(0, "F")
    if !row
        row := lv.GetNext(0)
    if !row || !state["rowToIndex"].Has(row) {
        MsgBox("Please select a rule to edit.", "Notice", "Iconi Owner" state["dialog"].Hwnd)
        return
    }
    idx := state["rowToIndex"][row]
    rule := state["rules"][idx]
    newExe := NormalizeExeName(PromptInput("Edit Shortcut Rule", "Edit executable name.", rule["exe"], "", "Cancel", "Executable:"))
    if !newExe
        return
    newKey := Trim(PromptInput("Edit Shortcut Rule", "Edit title keyword to match.", rule["key"], "", "Cancel", "Title Keyword:"))
    if !newKey
        return
    newTargetPath := rule["targetPath"]
    if Confirm("Change target path?`nCurrent: " newTargetPath) {
        picked := FileSelect(3, A_MyDocuments, "Select shortcut target", "Shortcut or File (*.*)")
        if !picked
            return
        newTargetPath := picked
    }
    state["rules"][idx] := Map("exe", newExe, "key", newKey, "targetPath", newTargetPath)
    state["dirty"] := true
    SortShortcutRules(state["rules"])
    RefreshShortcutRuleListView(state)
}

ShortcutRuleSaveAndClose(state, *) {
    global SHORTCUT_RULES, settingsIniFile
    SHORTCUT_RULES := CloneShortcutRules(state["rules"])
    if (SaveShortcutRules(settingsIniFile, SHORTCUT_RULES)) {
        state["dirty"] := false
        MsgBox("Shortcut rules saved.", "Success", "Iconi Owner" state["dialog"].Hwnd)
    }
    state["dialog"].Destroy()
}

HandleShortcutRulesClose(state, *) {
    if (!state["dirty"]) {
        state["dialog"].Destroy()
        return
    }
    if !Confirm("Save changes before leaving?")
        return state["dialog"].Destroy()
    ShortcutRuleSaveAndClose(state)
}

ManageShortcutRules() {
    global SHORTCUT_RULES, uiFont, uiFontSize
    rules := CloneShortcutRules(SHORTCUT_RULES)
    dialog := Gui("-DPIScale +AlwaysOnTop", "Shortcut Rules")
    dialog.BackColor := "F3F6FA"
    dialog.SetFont("s" uiFontSize, uiFont)
    dialog.Add("Text", "xm ym c1F2937 BackgroundTrans", "Configure shortcut launch rules used during save/restore.")
    dialog.Add("Text", "xm y+4 c6B7280 BackgroundTrans", "Match by executable + title keyword -> target path (shortcut/documents/executables).")
    lv := dialog.Add("ListView", "xm y+8 w980 r14", ["Executable", "Title Keyword", "Target Path"])
    lv.ModifyCol(1, 170), lv.ModifyCol(2, 210), lv.ModifyCol(3, 540)
    state := Map("rules", rules, "lv", lv, "dialog", dialog, "rowToIndex", Map(), "dirty", false)
    RefreshShortcutRuleListView(state)
    addBtn := dialog.Add("Button", "xm y+10 w90", "Add")
    editBtn := dialog.Add("Button", "x+8 yp w90", "Edit")
    delBtn := dialog.Add("Button", "x+8 yp w90", "Delete")
    cancelBtn := dialog.Add("Button", "x+332 yp w100", "Back")
    saveBtn := dialog.Add("Button", "x+8 yp w120 Default", "Save")
    addBtn.OnEvent("Click", ShortcutRuleAdd.Bind(state))
    editBtn.OnEvent("Click", ShortcutRuleEdit.Bind(state))
    delBtn.OnEvent("Click", ShortcutRuleDelete.Bind(state))
    saveBtn.OnEvent("Click", ShortcutRuleSaveAndClose.Bind(state))
    cancelBtn.OnEvent("Click", HandleShortcutRulesClose.Bind(state))
    dialog.OnEvent("Close", HandleShortcutRulesClose.Bind(state))
    dialog.OnEvent("Escape", HandleShortcutRulesClose.Bind(state))
    dialog.Show("AutoSize Center")
    WinWaitClose("ahk_id " dialog.Hwnd)
}

CsvToExeMap(csv) {
    m := Map()
    for token in StrSplit(csv, ",") {
        exe := NormalizeExeName(token)
        if exe
            m[exe] := true
    }
    return m
}

CloneBoolMap(src) {
    out := Map()
    for k, v in src
        out[k] := v
    return out
}

DefaultFileExtensions() {
    return ["txt", "md", "rtf", "log", "ini", "cfg", "json", "jsonl", "yaml", "yml", "xml", "csv", "pdf", "doc", "docx", "xls", "xlsx", "ppt", "pptx", "hwp", "hwpx", "psd", "ai", "png", "jpg", "jpeg", "gif", "bmp", "zip", "7z", "rar"]
}

NormalizeExtensionToken(token) {
    token := StrLower(Trim(token))
    if !token
        return ""
    if SubStr(token, 1, 1) = "."
        token := SubStr(token, 2)
    token := RegExReplace(token, "[^a-z0-9]", "")
    return token
}

LoadFileExtensions(filePath) {
    listText := IniReadSafe(filePath, "FileExtensions", "extensions", "")
    exts := ParseFileExtensionsFromText(listText)
    return exts.Length ? exts : DefaultFileExtensions()
}

LoadExcludeExe(filePath) {
    exesText := IniReadSafe(filePath, "Exclusions", "exe", "")
    exes := ParseExeMapFromText(exesText)
    return exes
}

SaveFileExtensions(filePath, exts) {
    try IniDelete(filePath, "FileExtensions")
    catch
    if !IsObject(exts)
        return false
    expected := FileExtensionsToText(exts)
    if !EnsureIniParentDir(filePath)
        return false
    try {
        IniWrite(expected, filePath, "FileExtensions", "extensions")
        return true
    } catch as err {
        MsgBox("Failed to save file extensions settings.`n" err.Message, "Settings Error", "Icon!")
        return false
    }
}

SaveExcludeExe(filePath, exes) {
    try IniDelete(filePath, "Exclusions")
    catch
    if !IsObject(exes)
        exes := Map()
    expected := ExeMapToText(exes)
    if !EnsureIniParentDir(filePath)
        return false
    try {
        IniWrite(expected, filePath, "Exclusions", "exe")
        return true
    } catch as err {
        MsgBox(
            "Failed to save exclusion rules.`n"
            "Message: " err.Message "`n"
            "Line: " err.Line
            , "Settings Error", "Icon!"
        )
        return false
    }
}

BuildFilePathRegex(exts) {
    safeParts := []
    for ext in exts {
        t := NormalizeExtensionToken(ext)
        if t
            safeParts.Push(t)
    }
    if safeParts.Length = 0
        safeParts := DefaultFileExtensions()
    extAlternatives := JoinRegexAlternatives(safeParts)
    pathPart := "[^" Chr(34) ":<>|?*\\\\]+?"
    extSuffix := "\.(" extAlternatives ")"
    absolutePath := "[A-Z]:(?:\\" pathPart ")+" extSuffix
    relativePath := pathPart "(?:\\" pathPart ")+" extSuffix
    return "i)(" absolutePath "|" relativePath ")"
}

JoinRegexAlternatives(parts) {
    out := ""
    for i, part in parts {
        if i > 1
            out .= "|"
        out .= part
    }
    return out
}

FileExtensionsToText(exts) {
    out := ""
    for ext in exts
        out .= ext "`n"
    return RTrim(out, "`n")
}

ParseFileExtensionsFromText(text) {
    exts := []
    seen := Map()
    for line in StrSplit(text, "`n", "`r") {
        line := StrReplace(line, ",", " ")
        for token in StrSplit(line, " `t") {
            ext := NormalizeExtensionToken(token)
            if ext && !seen.Has(ext) {
                seen[ext] := true
                exts.Push(ext)
            }
        }
    }
    return exts
}

MapKeysToSortedArray(m) {
    arr := []
    for k in m
        arr.Push(k)
    SortStringArray(arr)
    return arr
}

SortStringArray(arr) {
    n := arr.Length
    if n < 2
        return
    i := 1
    while i < n {
        minIdx := i
        j := i + 1
        while j <= n {
            if StrCompare(StrLower(arr[j]), StrLower(arr[minIdx])) < 0
                minIdx := j
            j += 1
        }
        if minIdx != i {
            tmp := arr[i]
            arr[i] := arr[minIdx]
            arr[minIdx] := tmp
        }
        i += 1
    }
}

ExeMapToText(m) {
    out := ""
    for exe in MapKeysToSortedArray(m)
        out .= exe "`n"
    return RTrim(out, "`n")
}

TimeoutMapToText(m) {
    keys := MapKeysToSortedArray(m)
    out := ""
    for exe in keys
        out .= exe "=" m[exe] "`n"
    return RTrim(out, "`n")
}

ParseExeMapFromText(text) {
    global BROWSER_EXE
    m := Map()
    for line in StrSplit(text, "`n", "`r") {
        for token in StrSplit(line, ",") {
            token := Trim(token)
            if !token
                continue
            if StrLower(token) = "@browser" || StrLower(token) = "@browsers" {
                for exeName in BROWSER_EXE
                    m[exeName] := true
                continue
            }
            exe := NormalizeExeName(token)
            if exe
                m[exe] := true
        }
    }
    return m
}

ParseTimeoutMapFromText(text) {
    m := Map()
    for line in StrSplit(text, "`n", "`r") {
        line := Trim(line)
        if !line
            continue
        if !RegExMatch(line, "i)^([^=,\s]+)\s*(?:=|,|\s)\s*(\d+)$", &cap)
            continue
        exe := NormalizeExeName(cap[1])
        ms := Integer(cap[2])
        if exe && ms > 0
            m[exe] := ms
    }
    return m
}

LoadWaitRules(filePath, &shortMap, &longMap, &timeoutMap) {
    if !FileExist(filePath) {
        shortMap := Map()
        longMap := Map()
        timeoutMap := Map()
        return
    }
    shortText := IniReadSafe(filePath, "WaitRules", "short", "")
    longText := IniReadSafe(filePath, "WaitRules", "long", "")
    timeoutText := IniReadSafe(filePath, "WaitTimeout", "items", "")
    shortMap := ParseExeMapFromText(shortText)
    longMap := ParseExeMapFromText(longText)
    timeoutMap := ParseTimeoutMapFromText(timeoutText)
}

SaveWaitRules(filePath, shortMap, longMap, timeoutMap) {
    try IniDelete(filePath, "WaitRules")
    catch
    try IniDelete(filePath, "WaitTimeout")
    catch
    if !IsObject(shortMap)
        shortMap := Map()
    if !IsObject(longMap)
        longMap := Map()
    if !IsObject(timeoutMap)
        timeoutMap := Map()
    expectedShort := ExeMapToText(shortMap)
    expectedLong := ExeMapToText(longMap)
    expectedTimeout := TimeoutMapToText(timeoutMap)
    if !EnsureIniParentDir(filePath)
        return false
    try {
        IniWrite(expectedShort, filePath, "WaitRules", "short")
        IniWrite(expectedLong, filePath, "WaitRules", "long")
        IniWrite(expectedTimeout, filePath, "WaitTimeout", "items")
        return true
    } catch as err {
        MsgBox("Failed to save wait rules.`n" err.Message, "Settings Error", "Icon!")
        return false
    }
}

EnsureIniParentDir(filePath) {
    if !IsSet(filePath) || !filePath
        return false
    dirSepPos := InStr(filePath, "\", , -1)
    if !dirSepPos
        return true
    dirPath := SubStr(filePath, 1, dirSepPos - 1)
    if !DirExist(dirPath) {
        try DirCreate(dirPath)
        catch as err {
            MsgBox("Failed to create settings folder.`n" err.Message, "Settings Error", "Icon!")
            return false
        }
    }
    return true
}

IniReadSafe(filePath, section, key, defaultValue := "") {
    try {
        return IniRead(filePath, section, key)
    } catch {
        return defaultValue
    }
}

WaitRulesSave(state, *) {
    global SHORT_WAIT_EXE, LONG_WAIT_EXE, TIMEOUT_BY_EXE, settingsIniFile
    shortMap := ParseExeMapFromText(state["shortEdit"].Value)
    longMap := ParseExeMapFromText(state["longEdit"].Value)
    timeoutMap := ParseTimeoutMapFromText(state["timeoutEdit"].Value)
    SHORT_WAIT_EXE := shortMap
    LONG_WAIT_EXE := longMap
    TIMEOUT_BY_EXE := timeoutMap
    if (SaveWaitRules(settingsIniFile, SHORT_WAIT_EXE, LONG_WAIT_EXE, TIMEOUT_BY_EXE)) {
        MsgBox("Wait rules saved.", "Success", "Iconi Owner" state["dialog"].Hwnd)
        state["dialog"].Destroy()
    }
}

ManageWaitRules() {
    global SHORT_WAIT_EXE, LONG_WAIT_EXE, TIMEOUT_BY_EXE, settingsIniFile, uiFont, uiFontSize
    dialog := Gui("-DPIScale +AlwaysOnTop", "Wait Rules")
    dialog.BackColor := "F3F6FA"
    dialog.SetFont("s" uiFontSize, uiFont)
    waitDialogW := 540
    colGap := 20
    colW := 250
    dialog.Add("Text", "xm ym section w" waitDialogW " c1F2937 BackgroundTrans", "Configure process launch wait behavior.")
    dialog.Add("Text", "xm y+4 w" waitDialogW " c6B7280 BackgroundTrans", "Use one EXE per line. You can use @browser. Timeout override format: exe=milliseconds")
    dialog.Add("Text", "xm y+10 section w" colW, "Short Wait EXEs")
    shortEdit := dialog.Add("Edit", "xs y+4 w" colW " r10", ExeMapToText(SHORT_WAIT_EXE))
    dialog.Add("Text", "x+" colGap " ys w" colW, "Long Wait EXEs")
    longEdit := dialog.Add("Edit", "xp y+4 w" colW " r10", ExeMapToText(LONG_WAIT_EXE))
    dialog.Add("Text", "xm y+10 w" waitDialogW, "Timeout Overrides")
    timeoutEdit := dialog.Add("Edit", "xm y+4 w" waitDialogW " r6", TimeoutMapToText(TIMEOUT_BY_EXE))
    backBtn := dialog.Add("Button", "xm y+12 w120", "Back")
    saveBtn := dialog.Add("Button", "x" (waitDialogW - 120) " yp w120 Default", "Save")
    state := Map("dialog", dialog, "shortEdit", shortEdit, "longEdit", longEdit, "timeoutEdit", timeoutEdit)
    saveBtn.OnEvent("Click", WaitRulesSave.Bind(state))
    backBtn.OnEvent("Click", (*) => dialog.Destroy())
    dialog.OnEvent("Close", (*) => dialog.Destroy())
    dialog.OnEvent("Escape", (*) => dialog.Destroy())
    dialog.Show("AutoSize Center")
    WinWaitClose("ahk_id " dialog.Hwnd)
}

ExcludeRulesSave(state, *) {
    global EXCLUDE_EXE, settingsIniFile
    excludeMap := ParseExeMapFromText(state["excludeEdit"].Value)
    EXCLUDE_EXE := excludeMap
    if (SaveExcludeExe(settingsIniFile, EXCLUDE_EXE)) {
        MsgBox("Exclusion rules saved.", "Success", "Iconi Owner" state["dialog"].Hwnd)
        state["dialog"].Destroy()
    }
}

ManageExcludeRules() {
    global EXCLUDE_EXE, settingsIniFile, uiFont, uiFontSize
    dialog := Gui("-DPIScale +AlwaysOnTop", "Exclusion Rules")
    dialog.BackColor := "F3F6FA"
    dialog.SetFont("s" uiFontSize, uiFont)
    settingDialogW := 560
    dialog.Add("Text", "xm ym section w" settingDialogW " c1F2937 BackgroundTrans", "Configure process names to skip during save.")
    dialog.Add("Text", "xm y+4 w" settingDialogW " c6B7280 BackgroundTrans", "Use one EXE per line, including .exe.")
    excludeEdit := dialog.Add("Edit", "xm y+10 w" settingDialogW " r16", ExeMapToText(EXCLUDE_EXE))
    backBtn := dialog.Add("Button", "xm y+12 w120", "Back")
    saveBtn := dialog.Add("Button", "x" (settingDialogW - 120) " yp w120 Default", "Save")
    state := Map("dialog", dialog, "excludeEdit", excludeEdit)
    saveBtn.OnEvent("Click", ExcludeRulesSave.Bind(state))
    backBtn.OnEvent("Click", (*) => dialog.Destroy())
    dialog.OnEvent("Close", (*) => dialog.Destroy())
    dialog.OnEvent("Escape", (*) => dialog.Destroy())
    dialog.Show("AutoSize Center")
    WinWaitClose("ahk_id " dialog.Hwnd)
}

FileExtensionsSave(state, *) {
    global FILE_EXTENSIONS, FILE_PATH_REGEX, settingsIniFile
    exts := ParseFileExtensionsFromText(state["extEdit"].Value)
    if exts.Length = 0 {
        MsgBox("Please enter at least one extension.")
        return
    }
    FILE_EXTENSIONS := exts
    FILE_PATH_REGEX := BuildFilePathRegex(FILE_EXTENSIONS)
    if (SaveFileExtensions(settingsIniFile, FILE_EXTENSIONS)) {
        MsgBox("File extensions saved.", "Success", "Iconi Owner" state["dialog"].Hwnd)
        state["dialog"].Destroy()
    }
}

ManageFileExtensions() {
    global FILE_EXTENSIONS, FILE_PATH_REGEX, settingsIniFile, uiFont, uiFontSize
    dialog := Gui("-DPIScale +AlwaysOnTop", "File Extensions")
    dialog.BackColor := "F3F6FA"
    dialog.SetFont("s" uiFontSize, uiFont)
    settingDialogW := 560
    dialog.Add("Text", "xm ym section w" settingDialogW " c1F2937 BackgroundTrans", "Configure extensions used for generic file-path detection.")
    dialog.Add("Text", "xm y+4 w" settingDialogW " c6B7280 BackgroundTrans", "Use one extension per line, without dot. Example: pdf")
    extEdit := dialog.Add("Edit", "xm y+10 w" settingDialogW " r16", FileExtensionsToText(FILE_EXTENSIONS))
    backBtn := dialog.Add("Button", "xm y+12 w120", "Back")
    saveBtn := dialog.Add("Button", "x" (settingDialogW - 120) " yp w120 Default", "Save")
    state := Map("dialog", dialog, "extEdit", extEdit)
    saveBtn.OnEvent("Click", FileExtensionsSave.Bind(state))
    backBtn.OnEvent("Click", (*) => dialog.Destroy())
    dialog.OnEvent("Close", (*) => dialog.Destroy())
    dialog.OnEvent("Escape", (*) => dialog.Destroy())
    dialog.Show("AutoSize Center")
    WinWaitClose("ahk_id " dialog.Hwnd)
}

FindLastPos(haystack, needle) {
    last := 0, pos := 1
    while found := InStr(haystack, needle, false, pos)
        last := found, pos := found + 1
    return last
}

TrySetExtra(&extraType, &extraVal, candidateType, candidateVal) {
    if !candidateVal
        return
    if ExtraPriority(candidateType) >= ExtraPriority(extraType)
        extraType := candidateType, extraVal := candidateVal
}

ExtraPriority(t) {
    return (t = "shortcutPath") ? 100 : ((t ~= "wordDoc|excelDoc|pptDoc|hwpDoc|hpdfDoc") ? 90 : ((t = "filePath") ? 85 : ((t = "explorerPath") ? 80 : 0)))
}

Esc(s) {
    s := StrReplace(StrReplace(s "", "`r", " "), "`n", " ")
    return StrReplace(StrReplace(s, "\", "\\"), Chr(34), "\" Chr(34))
}

ParseJsonl(file) {
    arr := [], txt := FileRead(file, "UTF-8")
    for line in StrSplit(txt, "`n", "`r") {
        if !Trim(line)
            continue
        o := Map()
        for k in ["exe", "exeName", "class", "title", "extraType", "extraVal"]
            o[k] := JGet(line, k)
        for k in ["mon", "state"]
            o[k] := JGetNum(line, k)
        for k in ["rx", "ry", "rw", "rh"]
            o[k] := JGetFloat(line, k)
        o["exeName"] := StrLower(o["exeName"])
        arr.Push(o)
    }
    return arr
}

IsBrowserExe(n) {
    global BROWSER_EXE
    exe := NormalizeExeName(n)
    return exe ? BROWSER_EXE.Has(exe) : false
}

SnapshotWindowsByExe(exeName) {
    snap := Map()
    if exeName
        for hwnd in WinGetList("ahk_exe " exeName)
            snap[hwnd] := true
    return snap
}

WaitWindowByPid(pid, timeout, usedHwnd) {
    start := A_TickCount
    while A_TickCount - start < timeout {
        for hwnd in WinGetList("ahk_pid " pid)
            if !usedHwnd.Has(hwnd)
                return hwnd
        Sleep(100)
    }
    return 0
}

WaitNewWindowFromSnapshot(exeName, beforeSet, timeout, usedHwnd) {
    start := A_TickCount
    while A_TickCount - start < timeout {
        for hwnd in WinGetList("ahk_exe " exeName)
            if !usedHwnd.Has(hwnd) && !beforeSet.Has(hwnd)
                return hwnd
        Sleep(100)
    }
    return 0
}

FindFallbackWindow(item, usedHwnd) {
    exeName := item["exeName"], cls := item["class"], title := item["title"], best := 0
    if exeName {
        for hwnd in WinGetList("ahk_exe " exeName) {
            if usedHwnd.Has(hwnd)
                continue
            try {
                if cls && WinGetClass("ahk_id " hwnd) != cls
                    continue
                if title && InStr(WinGetTitle("ahk_id " hwnd), title)
                    return hwnd
                if !best
                    best := hwnd
            }
        }
    }
    return best
}

JGet(line, key) {
    return RegExMatch(line, '"' key '":"((?:\\.|[^"\\])*)"', &m) ? StrReplace(StrReplace(m[1], "\" Chr(34), Chr(34)), "\\", "\") : ""
}

JGetNum(line, key) {
    return RegExMatch(line, '"' key '":(-?\d+)', &m) ? Integer(m[1]) : 0
}

JGetFloat(line, key) {
    return RegExMatch(line, '"' key '":(-?\d+(?:\.\d+)?)', &m) ? m[1] + 0.0 : 0.0
}

DumpWindows(outFile) {
    lines := ""
    for hwnd in WinGetList() {
        try {
            title := WinGetTitle("ahk_id " hwnd)
            if !title
                continue
            lines .= "EXE=" WinGetProcessName("ahk_id " hwnd) "`nTITLE=" title "`nCLASS=" WinGetClass("ahk_id " hwnd) "`n-----------------------------`n"
        }
    }
    try FileDelete(outFile)
    FileAppend(lines, outFile, "UTF-8")
}
