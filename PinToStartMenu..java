import java.io.File;
import java.io.IOException;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class ShortcutCreator {
    public static void main(String[] args) throws IOException {
        // get the path to the Start Menu folder
        String startMenu = System.getenv("APPDATA") + "\\Microsoft\\Windows\\Start Menu\\Programs";

        // get the list of all programs installed on the system
        File programsFolder = new File("C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs");
        String[] programs = programsFolder.list();

        // create a Windows Shell object
        ActiveXComponent shell = new ActiveXComponent("WScript.Shell");

        // loop through each program and create a shortcut in the Start Menu folder
        for (String program : programs) {
            Dispatch shortcut = Dispatch.call(shell, "CreateShortcut", startMenu + "\\" + program + ".lnk").toDispatch();
            Dispatch.put(shortcut, "TargetPath", program);
            Dispatch.put(shortcut, "WorkingDirectory", program);
            Dispatch.call(shortcut, "Save");
        }
    }
}
