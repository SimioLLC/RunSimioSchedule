using SimioAPI;
using System;
using System.IO;

namespace RunSimioScheduleOrExperiment
{
    /// <summary>
    /// Everything needed for Running a Schedule headless.
    /// </summary>
    public class RunContext
    {
        public ISimioProject SimioProject { get; set; }

        /// <summary>
        /// The name of the project file.
        /// </summary>
        public string ProjectFilename { get; set; }

        /// <summary>
        /// This model name is looked for in each dropped project
        /// </summary>
        public string ModelName { get; set; } = null;

        /// <summary>
        /// This experiment name is looked for in each dropped project.
        /// If it is not null/empty and IsExperimentToBeRun is true, then the Experiment is run.
        /// </summary>
        public string ExperimentName { get; set; } = null;


        /// <summary>
        /// Full path to project (e.g. spfx) file while processing
        /// </summary>
        public string ProcessingFilepath
        {
            get
            {
                string fullpath = Path.Combine(ProcessingFolderpath, ProjectFilename);
                return fullpath;
            }
        }

        /// <summary>
        /// Full path to project (e.g. spfx) file while processing
        /// </summary>
        public string SuccessFilepath
        {
            get
            {
                string fullpath = Path.Combine(SuccessFolderpath, ProjectFilename);
                return fullpath;
            }
        }

        /// <summary>
        /// Full path to project (e.g. spfx) file while processing
        /// </summary>
        public string ErrorFilepath
        {
            get
            {
                string fullpath = Path.Combine(ErrorFolderpath, ProjectFilename);
                return fullpath;
            }
        }

        /// <summary>
        /// Folder monitored by System.FileWatcher
        /// </summary>
        public string WatcherFolderpath { get; set; }

        /// <summary>
        /// Folder holding the project file while it is being run.
        /// </summary>
        public string ProcessingFolderpath { get; set; }
        /// <summary>
        /// Folder where project files that have successfully run are placed.
        /// </summary>
        public string SuccessFolderpath { get; set; }
        /// <summary>
        /// Folder where projects that cannot process are placed
        /// </summary>
        public string ErrorFolderpath { get; set; }

        /// <summary>
        /// File filter for System.FileWatcher
        /// </summary>
        public string WatcherFilter { get; set; }


        /// <summary>
        /// The full path to the constructed status file for the project
        /// </summary>
        public string StatusFilepath { get; set; }

        /// <summary>
        /// Where the extensions path is set to during the constructor.
        /// </summary>
        public string ExtensionsFolderpath { get; private set; }

        public bool IsStatusToBeDeletedBeforeRun { get; set; }

        /// <summary>
        /// Save project after the run
        /// </summary>
        public bool IsProjectToBeSaved { get; set; }

        /// <summary>
        /// Should the plan be run
        /// </summary>
        public bool IsPlanToBeRun { get; set; } = true;

        /// <summary>
        /// Should the experiment named in ExperimentName be run?
        /// </summary>
        public bool IsExperimentToBeRun { get; set; } = true;

        /// <summary>
        /// Should risk analysis be run.
        /// </summary>
        public bool IsRiskAnalysisToBeRun { get; set; } = false;

        /// <summary>
        /// Constructor. Set up the ISimioProject and check directories, build paths, etc.
        /// Underneath the root, we have:
        /// /In - the folder being watched for .spfx files
        /// /Processing - where the project is placed upon receipt and run from
        /// /Saved - where the saved project goes after running
        /// /Error - where the project file goes if there are errors.
        /// A status file is written directly under the root folder
        /// The model name is the model that is run (from the settings file).
        /// In this scenario, all projects must have this model, or default to 'Model' if not found.
        /// </summary>
        /// <param name="modelName"></param>
        /// <param name="rootFolderpath"></param>
        /// <param name="statusFilename"></param>
        public RunContext(string modelName
            , string experimentName
            , string rootFolderpath
            , string statusFilepath)
        {
            string marker = "Begin.";
            try
            {
                ModelName = modelName;
                ExperimentName = experimentName;

                WatcherFolderpath = Path.Combine(rootFolderpath, "In");
                ProcessingFolderpath = Path.Combine(rootFolderpath, "Processing");
                SuccessFolderpath = Path.Combine(rootFolderpath, "Success");
                ErrorFolderpath = Path.Combine(rootFolderpath, "Error");
                StatusFilepath = statusFilepath;

                WatcherFilter = "*.SPFX";

                if (!Directory.Exists(rootFolderpath))
                    throw new ApplicationException($"Root Folder={rootFolderpath} not found.");

                string statusFolderpath = Path.GetDirectoryName(statusFilepath);
                if (!Directory.Exists(statusFolderpath))
                    throw new ApplicationException($"Status Folder={statusFolderpath} not found.");

                if (!Directory.Exists(WatcherFolderpath))
                    throw new ApplicationException($"Watcher folder={WatcherFolderpath} not found.");

                marker = $"RunContext Watching={WatcherFolderpath} with Filter-{WatcherFolderpath} Status to={StatusFilepath} Model={ModelName}";

                string programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
                this.ExtensionsFolderpath = Path.Combine(programFiles, "Simio LLC", "Simio", "UserExtensions");
                if (Directory.Exists(ExtensionsFolderpath) == false)
                {
                    throw new ApplicationException($"Could not locate ExtensionsPath={ExtensionsFolderpath}");
                }

                marker = $"Setting ExtensionPath to={ExtensionsFolderpath}";
                //SimioProjectFactory.SetExtensionsPath(ExtensionsFolderpath);

            }
            catch (Exception ex)
            {
                throw new ApplicationException($"RootFolder={rootFolderpath} Marker={marker} Err={ex}");
            }
        } // constructor


        /// <summary>
        /// Log to our status file.
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="append"></param>
        public void Logit(string msg, bool append)
        {
            HeadlessHelpers.LogIt(msg, append);
        }

        /// <summary>
        /// Log to our status file.
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="append"></param>
        public void Logit(string msg)
        {
            Logit(msg, true);
        }


    }  // RunContext class

}

