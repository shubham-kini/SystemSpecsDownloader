#Requires AutoHotkey v2.0
SetTitleMatchMode 2  ; Allow partial title matching
SendMode "Input"      ; More reliable keystroke sending

; --- Constants ---
startWindowTitle := "HWiNFO"
analysisWindowTitle := "Analyzing"
updateWindowTitle := "Update"
summaryWindowTitle := "System Summary"
mainWindowTitle := "HWiNFO"
createLogfileWindowTitle := "Create Logfile"

global startOrMainWin := false
global startWinCLosed := false
global analysisWinClosed := false
global updateWinCLosed := false
global summaryWinClosed := false
global mainWinCLosed := false

downloadsFolder := EnvGet("USERPROFILE") "\Downloads\"
reportFileName := "HWiNFOReport."
cleanedFileName := "HWiNFOReport-" EnvGet("USERNAME") "-Cleaned."
targetHtmlString := "<TABLE WIDTH=`"100%`"><TD CLASS=`"dt`" id=`"Network`">Network<TR><TD></TABLE>"

;Run the file in Admin mode
#SingleInstance Force
full_command_line := DllCall("GetCommandLine", "str")
if not (A_IsAdmin or RegExMatch(full_command_line, " /restart(?!\S)")) {
    try
    {
        if A_IsCompiled
            Run '*RunAs "' A_ScriptFullPath '" /restart'
        else
            Run '*RunAs "' A_AhkPath '" /restart "' A_ScriptFullPath '"'
    }
    ExitApp
}

; Function to handle ControlClick with error checking
ControlClickSafe(Control, WinTitle) {
    try {
        WinActivate(WinTitle)
        SetControlDelay -1
        ControlClick(Control, WinTitle, , 'Left', 1, 'NA')
        Sleep 400
        if (WinTitle = startWindowTitle && Control = "Start") {
            global startWinCLosed := true
        }
    }
    catch Error as e {
        if (WinTitle = mainWindowTitle && Control = "Start") {
            ;This mostly indicates that the window is Main window and not the start window
            global startWinCLosed := false
        } else {
            MsgBox "ControlClick Failed on " Control ": " e.Message
            ExitApp
        }
    }
    ;return True  ; Return True if successful
}

; Function to check if HWiNFO64 is installed and run it
RunHWiNFO() {
    ; Get the Windows installation drive (C:, D:, etc.)
    local windowsDrive := SubStr(EnvGet("SystemRoot"), 1, 2)  ; This gives you the drive letter, e.g., C: or D:

    ; Define possible install paths relative to the Windows installation drive
    local installPaths := [
        windowsDrive . "\Program Files\HWiNFO64\HWiNFO64.EXE",
        windowsDrive . "\Program Files (x86)\HWiNFO64\HWiNFO64.EXE",
        windowsDrive . "\HWiNFO64\HWiNFO64.EXE",
    ]

    ; Check if HWiNFO64 exists in any of the defined paths
    for path in installPaths {
        ; MsgBox path
        ; MsgBox FileExist(path)
        if FileExist(path) {
            try {
                Run path  ; Run HWiNFO64 from the found path
                WinWait startWindowTitle, , 60  ; Wait for the window to appear
                return True  ; Successfully launched the program
            }
            catch Error as e {
                MsgBox "Failed to launch HWiNFO64 from " path ": " e.Message
                return False  ; Return False if launching failed
            }
        }
    }

    ; ; If HWiNFO64 wasn't found in any of the expected paths, prompt for a custom path
    ; Loop until a valid path is provided or the user cancels/close the InputBox
    while True {
        ; Ask the user to input a custom path for HWiNFO64
        customPath := InputBox("Please enter the full path to HWiNFO64.EXE", "Enter Custom Path", "W720 H70")

        ; Check if the user clicked "Ok", provided a valid path, and the file exists
        if (customPath.Result = "Ok" && customPath.Value != "" && FileExist(customPath.Value)) {
            try {
                Run customPath.Value  ; Run HWiNFO64 from the custom path
                WinWait startWindowTitle, , 60  ; Wait for the window to appear
                return True  ; Successfully launched the program
            }
            catch Error as e {
                MsgBox "Failed to launch HWiNFO64 from the custom path: " e.Message
                return False  ; Return False if launching failed
            }
        }
        ; If the user canceled the input or closed the input box, exit the loop
        else if (customPath.Result = "Cancel") {
            MsgBox "User canceled the path input. Exiting script."
            return False  ; Return False if the user cancels the input
        }
        ; If the input is invalid, prompt again
        else {
            MsgBox "The specified path is invalid or empty. Please try again."
        }
    }

}

; Step 1: Launch HWiNFO64 (if not already running)
if !WinExist(startWindowTitle) {

    if !RunHWiNFO() {
        ExitApp  ; Exit if the program wasn't launched
    }
    else {
        ControlClickSafe("Start", startWindowTitle)
    }
}
else {
    ;This code will check if the window with startWindowTitle is one of the popups, if not, it will assume it is either the main window or the start window.
    winId := WinExist(startWindowTitle)
    popupTitles := [updateWindowTitle, summaryWindowTitle]
    isPopup := false
    for popupTitle in popupTitles {
        popupId := WinExist(popupTitle)
        if (popupId == winId) {
            isPopup := true
            if (closePopUpWindow(popupId, popupTitle)) {
                if (popupTitle == updateWindowTitle) {
                    global updateWinCLosed := true
                }
                if (popupTitle == summaryWindowTitle) {
                    global summaryWinClosed := true
                }
            }
        }
    }
    if (!isPopup) {
        ControlClickSafe("Start", startWindowTitle)
    }
}
Sleep 500

if (startWinCLosed) {
    ; MsgBox "analysisWinClosed := true"
    ; Step 2: Click the Start button to begin the analysis
    ;ControlClickSafe("Start", startWindowTitle)

    ; Step 3: Wait for the loading UI to finish
    if (!analysisWinClosed) {
        try {
            if WinWait(analysisWindowTitle, , 10) {
                WinWaitClose(analysisWindowTitle, , 30)
                global analysisWinClosed := true

            }
            else {
                throw Error("Analysis UI not Found!")
            }
        }
        catch Error as e {
            MsgBox "Error While Analyzing!: " e.Message
            ExitApp
        }
    }

    ; Step 4: Handle the Update and System Summary UIs (close them using Esc key)
    if (!updateWinCLosed) {
        if (closePopUpWindow(updateWindowTitle, updateWindowTitle)) {
            global updateWinCLosed := true
        }
    }
    if (!summaryWinClosed) {
        if (closePopUpWindow(summaryWindowTitle, summaryWindowTitle)) {
            global summaryWinClosed := true
        }
    }
}

;Save 2 files
;for fileType in ["XML", "HTML"] {
for fileType in ["HTML"] {
    ; Step 5: Activate the Main UI and press the necessary keys to create the XML log
    try {
        if WinWait(mainWindowTitle, , 10) {
            WinActivate(mainWindowTitle)
            Sleep 400
            Send("{Alt}")
            Sleep 100
            Send("R")
            Sleep 100
            Send("C")
            Sleep 400
        }
        else {
            throw Error("Main UI not Found!")
        }
    }
    catch Error as e {
        MsgBox "Error activating main UI: " e.Message
        ExitApp
    }

    ; Step 6: Wait for the "Create Logfile" popup to appear and interact with it
    try {
        ; if !WinWait(createLogfileWindowTitle, , 20) {
        ;     throw Error("Create Logfile window did not appear!")
        ; }
        WinWait(createLogfileWindowTitle, , 20)
        WinActivate(createLogfileWindowTitle)
        Sleep 400
    }
    catch Error as e {
        MsgBox "Failed to open Create Logfile: " e.Message
        ExitApp
    }

    ; Step 7: Click XML Button
    ControlClickSafe(fileType, createLogfileWindowTitle)

    ; Step 8: Set File Path
    try {
        ControlClickSafe("Edit1", createLogfileWindowTitle)
        fullFilePath := downloadsFolder reportFileName fileType
        ControlSetText(fullFilePath, "Edit1", createLogfileWindowTitle)
        Sleep 500
    }
    catch Error as e {
        MsgBox "Failed to set file path: " e.Message
    }

    ; Step 9: Click "Next >"
    ControlClickSafe("Next >", createLogfileWindowTitle)

    ; Step 10: Click "Finish"
    ControlClickSafe("Finish", createLogfileWindowTitle)
    Sleep 2000

    ; Step 11: Clean the HTML File and make a new HTML File
    if (fileType = "HTML") {
        inputFilePath := downloadsFolder reportFileName fileType
        outputFilePath := downloadsFolder cleanedFileName fileType
        if (removeNetworkAndPortTables(inputFilePath, outputFilePath)) {
            MsgBox("Cleaned HTML saved to: " outputFilePath)
        }
    }

}

; Step 12: Delete the original HTML and XML files
try {
    ;FileDelete(downloadsFolder reportFileName "XML")
    FileDelete(downloadsFolder reportFileName "HTML")
    ;MsgBox("Original XML and HTML files deleted.")
} catch Error as e {
    MsgBox("Failed to delete original files: " e.Message)
}

; Step 13: Open Downloads folder
try {
    Run(downloadsFolder)
    ;MsgBox("Downloads folder opened.")
} catch Error as e {
    MsgBox("Failed to open Downloads folder: " e.Message)
}

; Final Message
MsgBox "HWiNFO report creation process is complete!"

; --- Function to close PopUp Windows
closePopUpWindow(winTitle, winRealTitle) {
    try {
        if WinWait(winTitle, , 10) {
            WinActivate(winTitle)
            Sleep 500
            Send "{Esc}"
            return true
        }
        else {
            throw Error(winRealTitle "UI not Found!")
        }
    }
    catch Error as e {
        MsgBox "Failed to close " winRealTitle ": " e.Message
        ExitApp
    }
}

; --- Function to clean HTML ---
removeNetworkAndPortTables(inputFilePath, outputFilePath) {
    try {
        file := FileOpen(inputFilePath, "r")
        htmlString := file.Read()
        file.Close()
    } catch Error as e {
        MsgBox("Error reading input file: " e.Message)
        return false
    }

    position := InStr(htmlString, targetHtmlString)

    if (position > 0) {
        modifiedHtml := SubStr(htmlString, 1, position - 1) ; Keep text before match.
    } else {
        modifiedHtml := htmlString
    }

    try {
        file := FileOpen(outputFilePath, "w")
        file.Write(modifiedHtml)
        file.Close()
        return true
    } catch Error as e {
        MsgBox("Error writing output file: " e.Message)
        return false
    }
}
