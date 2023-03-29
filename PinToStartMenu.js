const winshell = require('winshell');
const shell = new ActiveXObject('WScript.Shell');

// get the path to the Start Menu folder
const startMenu = winshell.startup();

// get the list of all programs installed on the system
const programs = winshell.programs();

// loop through each program and create a shortcut in the Start Menu folder
for (let i = 0; i < programs.length; i++) {
  const program = programs[i];
  const shortcut = shell.CreateShortcut(`${startMenu}\\${program}.lnk`);
  shortcut.TargetPath = program;
  shortcut.WorkingDirectory = program;
  shortcut.Save();
}
