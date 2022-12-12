# RunSimioSchedule
This Github has two examples that show how to run Simio Projects in "Headless" mode using the Simio API.

In summary, there are three examples of using file-drop techniques to initiate the Simio Engine running a Simio project (a .spfx file).
The first example (RunSimioSchedule) has the Simio project loaded at startup, and it is run any time an updated file with scheduling "events" (such as downtime) is dropped into the IN folder.
The second example (RunSimioSchedule2) instead watches the IN folder for a Simio project file (anything with a .SPFX extension). This project is loaded and then an Experiment or Plan is run according to the Settings.

The third example is a very bare example that uses example projects (from the Example folder of the desktop release) to run an Experiment and a Plan.

The Settings file is a standard VisualStudio artifact that can be seen in the code and is distributed with the executable file (for example, see \bin\Release\RunSimioSchedule.exe.config).

Under the Source folder are the three projects, and under each of these is a Configuration folder with a single folder (e.g. RunSimioSchedule2). You can take this and place it under the root folder defined in the project's Settings file (e.g. c:\temp).

Note that these code samples reference the Simio Engine from where it is installed for the desktop version: e.g. c:\program files\Simio LLC\simio. If you wish to run it elsewhere, you need to move the DLLs, including all of their dependencies and the licensing DLLs.

Look to the Documentation folder that contains detailed documentation.

