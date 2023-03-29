package main

import (
    "fmt"
    "os"
    "ole"
    "ole/oleutil"
)

func main() {
    // get the path to the Start Menu folder
    startMenu, _ := os.UserConfigDir()
    startMenu += "\\Microsoft\\Windows\\Start Menu\\Programs"

    // get the list of all programs installed on the system
    programsFolder := "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs"
    programs, _ := os.ReadDir(programsFolder)

    // create a Windows Shell object
    ole.CoInitialize(0)
    defer ole.CoUninitialize()
    shell, _ := oleutil.CreateObject("WScript.Shell")

    // loop through each program and create a shortcut in the Start Menu folder
    for _, program := range programs {
        shortcut := oleutil.MustCallMethod(shell, "CreateShortcut", fmt.Sprintf("%s\\%s.lnk", startMenu, program.Name())).ToIDispatch()
        oleutil.MustPutProperty(shortcut, "TargetPath", fmt.Sprintf("%s\\%s", programsFolder, program.Name()))
        oleutil.MustPutProperty(shortcut, "WorkingDirectory", fmt.Sprintf("%s\\%s", programsFolder, program.Name()))
        oleutil.MustCallMethod(shortcut, "Save")
    }
}
