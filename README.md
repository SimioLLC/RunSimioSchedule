# RunSimioSchedule
This Github has two examples that show how to run Simio Projects in "Headless" mode using the Simio API.

In summary, there are two examples of using file-drop techniques to initiate the Simio Engine running a Simio project (a .spfx file).
The first example (RunSimioSchedule) has the Simio project loaded at startup, and it is run any time an updated file with scheduling "events" (such as downtime) is dropped into the IN folder.
The second example (RunSimioSchedule2) instead watches the IN folder for a Simio project file (anything with a .SPFX extension). This project is loaded and then an Experiment or Plan is run according to the Settings.

The Settings file is a standard VisualStudio artifact that can be seen in the code and is distributed with the executable file (for example, see \bin\Release\RunSimioSchedule.exe.config).

Look to the Documentation folder contains detailed documentation.

