using System;
using System.IO;
using System.ServiceProcess;
using System.Timers;


namespace RunSimioScheduleOrExperiment
{
    class Program : ServiceBase
    {
        /// <summary>
        /// RunContext needed for running
        /// </summary>
        private static RunContext RContext { get; set; }

        /// <summary>
        /// Used for locking the code that looks for and processes the triggering file.
        /// </summary>
        private static readonly object _runLock = new object();

        static void Main(string[] args)
        {
            string rootFolder = Properties.Settings.Default.RootFolderpath;
            // Quick check
            if (!Directory.Exists(rootFolder))
            {
                throw new ApplicationException($"Fatal error: No Root Folder found={rootFolder}");
            }

            string modelName = Properties.Settings.Default.ModelName;
            string experimentName = Properties.Settings.Default.ExperimentName;
            string statusFilepath = Properties.Settings.Default.StatusFilepath;

            // Create the run context
            RContext = new RunContext(modelName, experimentName, rootFolder, statusFilepath);

            HeadlessHelpers.LogIt($"ExtensionsPath={RContext.ExtensionsFolderpath}");

            RContext.IsProjectToBeSaved = Properties.Settings.Default.IsProjectToBeSaved;
            RContext.IsPlanToBeRun = Properties.Settings.Default.IsPlanToBeRun;
            RContext.IsRiskAnalysisToBeRun = Properties.Settings.Default.IsRiskAnalysisToBeRun;
            RContext.IsExperimentToBeRun = Properties.Settings.Default.IsExperimentToBeRun;

            if (!Environment.UserInteractive)
            {
                // running as service
                using (var service = new Program())
                {
                    ServiceBase.Run(service);
                }
            }
            else
            {
                using (var service = new Program())
                {
                    // running as console app
                    service.OnStart(args);
                    Console.WriteLine("Hit the 'q' key to stop...");
                    do
                    {
                        System.Threading.Thread.Sleep(50);
                    } while (Console.ReadLine() != "q");

                    service.OnStop();
                }

            }
        }

        /// <summary>
        /// On Start
        /// </summary>
        /// <param name="args"></param>
        protected override void OnStart(string[] args)
        {
            try
            {
                ShutdownHost();

                // Belts and suspenders. Set up a timer to watch for event (i.e. never trust a FileWatcher)
                System.Timers.Timer CheckTimer = new System.Timers.Timer();
                CheckTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
                CheckTimer.Interval = 5000;
                CheckTimer.Enabled = true;

                LogIt($"Setting up FileWatcher");
                FileSystemWatcher watcher = new FileSystemWatcher(
                    RContext.WatcherFolderpath,
                    RContext.WatcherFilter);

                watcher.Created += Watcher_Created;
                watcher.Changed += Watcher_Changed;
                watcher.EnableRaisingEvents = true;

                base.OnStart(args);

            }
            catch (Exception ex)
            {
                LogIt(ex.Message);
                throw new ApplicationException($"Err={ex}");
            }

        }


        /// <summary>
        /// Currently does nothing, but could be used for reset/cleanup conditions.
        /// </summary>
        private void ShutdownHost()
        {
            LogIt($"ShutdownHost called.");
        }

        protected override void OnStop()
        {
            ShutdownHost();
            base.OnStop();
        }

        static void Watcher_Changed(object sender, FileSystemEventArgs e) => CheckAndRun("Changed", e.FullPath);

        static void Watcher_Created(object sender, FileSystemEventArgs e) => CheckAndRun("Created", e.FullPath);

        private void OnTimedEvent(object sender, ElapsedEventArgs e) => CheckAndRun("Timer", "");


        /// <summary>
        /// See if the event file is present and - if  so - run the model.
        /// We lock it to prevent the timer and fileWatcher from interfering with each other.
        /// </summary>
        /// <param name="caller"></param>
        private static void CheckAndRun(string caller, string filepath)
        {
            string incomingFilepath = "";

            lock (_runLock)
            {
                if ( caller == "Timer")
                {
                    // Find the filepath
                    foreach ( string fullpath in Directory.GetFiles(RContext.WatcherFolderpath))
                    {
                        string extension = Path.GetExtension(fullpath);
                        if (extension.ToLower() == ".spfx")
                            incomingFilepath= fullpath;
                        else
                        {
                            string filename = Path.GetFileName(fullpath);
                            HeadlessHelpers.MoveFileToNewFolder(fullpath, RContext.ErrorFolderpath);
                        }
                    }
                }
                else
                {
                    incomingFilepath = filepath;
                }

                if ( File.Exists(incomingFilepath))
                {
                    LogIt($"Calling RunSchedule From={caller}. Found file={incomingFilepath}");

                    try
                    {
                        // Run the schedule plan/experiment
                        HeadlessHelpers.RunScheduleAndSaveProject(RContext, incomingFilepath);
                    }
                    catch (Exception ex)
                    {
                        LogIt(ex.Message);
                    }
                }
            } // lock

        }

        private static void LogIt(string msg)
        {
            HeadlessHelpers.LogIt(msg);
        }

    }
}
