using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Configuration.Install;
using SimioAPI;
using System.Timers;

namespace RunSimioSchedule
{
    class Program : ServiceBase
    {
        /// <summary>
        /// Information needed for HeadlessHelpers
        /// </summary>
        private static RunContext RContext { get; set; }

        /// <summary>
        /// Used for locking the code that looks for and processes the triggering file.
        /// </summary>
        private static readonly object _runLock = new object();

        static void Main(string[] args)
        {
            string modelName = Properties.Settings.Default.ModelName;
            string projectPath = Properties.Settings.Default.SimioProjectFile;
            string eventPath = Properties.Settings.Default.EventFile;
            string exportSchedulePath = Properties.Settings.Default.ExportScheduleFile;

            // Create the run context
            RContext = new RunContext(modelName, projectPath, eventPath, exportSchedulePath);

            LogIt($"Project={projectPath}");
            LogIt($"ExtensionsPath={RContext.ExtensionPath}");

            RContext.DeleteStatusBeforeRun = Properties.Settings.Default.DeleteStatusBeforeEachRun;
            RContext.SaveProject = Properties.Settings.Default.SaveProject;
            RContext.RunPlan = Properties.Settings.Default.RunPlan;
            RContext.RunRiskAnalysis = Properties.Settings.Default.RunRiskAnalysis;
            RContext.ExportSchedule = Properties.Settings.Default.ExportSchedule;

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
                    Console.WriteLine("Hit any key to stop...");
                    do
                    {
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
            string marker = "Begin";

            try
            {
                ShutdownHost();

                // Belts and suspenders. Set up a timer to watch for event (never trust the FileWatcher)
                System.Timers.Timer CheckTimer = new System.Timers.Timer();
                CheckTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
                CheckTimer.Interval = 5000;
                CheckTimer.Enabled = true;

                LogIt($"Setting up FileWatcher");
                FileSystemWatcher watcher = new FileSystemWatcher(
                    Path.GetDirectoryName(RContext.EventFilepath),
                    Path.GetFileName(RContext.EventFilepath));

                watcher.Created += watcher_Created;
                watcher.Changed += watcher_Changed;
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

        static void watcher_Changed(object sender, FileSystemEventArgs e) => CheckAndRun("Changed");

        static void watcher_Created(object sender, FileSystemEventArgs e) => CheckAndRun("Created");

        private void OnTimedEvent(object sender, ElapsedEventArgs e) => CheckAndRun("Timer");


        /// <summary>
        /// See if the event file is presenty and - if  so - run the model.
        /// We lock it to prevent the timer and fileWatcher from interfering with each other.
        /// </summary>
        /// <param name="caller"></param>
        private static void CheckAndRun(string caller)
        {
            lock (_runLock)
            {
                if (System.IO.File.Exists(RContext.EventFilepath))
                {
                    LogIt($"Calling RunSchedule From={caller}. Found file={RContext.EventFilepath}");

                    try
                    {
                        HeadlessHelpers.RunScheduleExportResultsAndSaveProject(RContext);
                    }
                    catch (Exception ex)
                    {
                        LogIt(ex.Message);
                    }
                }
            } // lock

        }

        /// <summary>
        /// Log message to a text file specified in the Properties
        /// The message is prepended with a timetamp
        /// and the lines are truncated to 32000 characters.
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="append message"></param>
        private static void LogIt(string msg, bool append)
        {
            try
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(Properties.Settings.Default.StatusFile, append))
                {
                    string timestamp = System.DateTime.Now.ToString("ddMMMyy HH:mm:ss");
                    msg = $"{timestamp}: {msg}";
                    if (msg.Length > 32000) msg = msg.Substring(0, 32000);
                    file.WriteLine(msg);
                }
            }
            catch { }
        }

        private static void LogIt(string msg)
        {
            LogIt(msg, true);
        }

    }
}
