# RunSimioSchedule
This Github contains examples that show how to run Simio Projects in "Headless" mode using the SimioEngine API.

Simio releases v18.272 and up are now built on .NET Core. This project was updated to .NET Core v9 from .NET Framework v4.7.2. To upgrade your project from .NET Framework to .NET Core, Simio has found this VS 2022 extension useful -> https://marketplace.visualstudio.com/items?itemName=ms-dotnettools.upgradeassistant

<b>Note: Versions past 15.245 require additional references to Sqlite as Sqlite is now embedded within Simio. These are best done with NuGet references. For an example, examine the BareBones References.</b>

In summary, there are three examples, two of which using file-drop techniques to initiate the Simio Engine running a Simio project (a .spfx file).
The first example (RunSimioSchedule) has the Simio project loaded at startup, and it is run any time an updated file with scheduling "events" (such as downtime) is dropped into the IN folder.
The second example (RunSimioSchedule2) instead watches the IN folder for a Simio project file (anything with a .SPFX extension). This project is loaded and then an Experiment or Plan is run according to the Settings.

The third example is a very bare-bones example that uses example projects (from the Example folder of the desktop release) to run an Experiment and a Plan.

The Settings file is a standard VisualStudio artifact that can be seen in the code and is distributed with the executable file (for example, see \bin\Release\RunSimioSchedule.exe.config).

Under the Source folder are the three projects, and under each of these is a Configuration folder with a single folder (e.g. RunSimioSchedule2). You can take this and place it under the root folder defined in the project's Settings file (e.g. c:\temp).

Note that these code samples reference the Simio Engine from where it is installed for the desktop version: e.g. c:\program files\Simio LLC\simio. If you wish to run it elsewhere, you need to move the DLLs, including all of their dependencies and the licensing DLLs.

Look to the Documentation folder that contains more detailed information.