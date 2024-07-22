using RunBarebones.Properties;
using SimioAPI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace RunBarebones
{
    /// <summary>
    /// The simplest of runs with no fluff, to get you up and running with a known solution.
    /// Also, it deviates from good programming practices (such as modularity) to show all
    /// the logic in-line in a single method.
    /// 
    /// Install the desktop to your computer (e.g. c:\program files\Simio LLC\Simio
    /// Create a run root folder (say, c:\temp\SimioProjects\245 - 245 is the Simio version in this example).
    /// Get two sample projects from the desktop examples folder (e.g. c:\program files\Simio LLC\Simio):
    ///    1. HospitalEmergencyDepartment - for running an Experiment
    ///    2. SchedulingDiscretePartProduction - for running a Plan
    ///   
    /// The following code will run an Experiment from the first, and a Plan from the second
    /// The model "Model" is first located
    /// and then the experiment is Reset and then Run, or the Plan is Reset and then Run.
    /// The project is saved back to a file appended with "Saved"
    /// You can then open the project file from Simio desktop to see the results.
    /// </summary>
    internal class Program
    {
        static void Main(string[] args)
        {
            // This project is used to test various versions of Simio.
            // To do this, there is the "simioVersion" that corresponds to a folder on a temporary
            // path that holds our example projects included with that version.
            // Note that this "version" should correspond to the SimioAPI and SimioDLL References/Dependencies that
            // the project is built with (see the Solution Explorer)
            string programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

            // Determine whether to save Project with Source Control
            bool saveWithSourceControlSupport = Settings.Default.SaveWithSourceControlSupport;

            //==== Do some checks to make sure root folders are present ====
            string workPath = Settings.Default.WorkPath;
            if (!Directory.Exists(workPath))
            {
                AlertAndExit($"Cannot locate WorkPath={workPath}");
            }

            string simioPath = Settings.Default.SimioDesktopPath;
            if (!Directory.Exists(simioPath))
            {
                AlertAndExit($"Simio desktop path is not located. Path Setting={simioPath}");
            }


            Logit($"Info: Simio Installed at Root={simioPath}.");

            //========== Common items (e.g. setting up ExtensionsFolderpath) ===============
            string projectPath = Path.Combine(workPath, Settings.Default.ExperimentProjectFilename);
            if (!File.Exists(projectPath))
            {
                AlertAndExit($"Cannot locate Experiment Project Path={projectPath}. Model Name (from Settings)={Settings.Default.ExperimentModelName}");
            }

            string projectName = Path.GetFileNameWithoutExtension(projectPath);

            //========== Common items (e.g. setting up ExtensionsFolderpath) ===============
            string extensionsFolderpath = Path.Combine( simioPath, "UserExtensions");
            if (!Directory.Exists(extensionsFolderpath))
            {
                AlertAndExit($"Cannot location Extensions Path={extensionsFolderpath}");
            }

            Logit($"Info: Setting (User) Extensions path to={extensionsFolderpath}");
            SimioProjectFactory.SetExtensionsPath(extensionsFolderpath);

            //========== Do an Experiment ===============
            Logit($"Info: Project={projectPath} exists. Loading...");
            ISimioProject projectExperiment = SimioProjectFactory.LoadProject(projectPath, out string[] loadWarnings);
            if (loadWarnings?.Length > 0)
            {
                Logit($"Found {loadWarnings?.Length} Warnings while loading={projectPath}");

                int count = 0;
                foreach (var warning in loadWarnings)
                {
                    Logit($"Warning: Load Warning #{++count}={warning}");
                }
            }

            string modelName = Settings.Default.ExperimentModelName;
            Logit($"Info: From Settings: model={modelName}.");
            IModel model = projectExperiment.Models[modelName];

            if ( model != null )
            {
                if ( !model.Experiments.Any() ) 
                {
                    AlertAndExit($"Model={model.Name} had no Experiments.");
                    goto DoPlan;
                }
                Logit($"Info: Located model={modelName}. Using the first Experiment.");

                IExperiment experiment = model.Experiments[0]; // get the first experiment
                if (experiment != null)
                {
                    Logit($"Executing First Experiment. Name={experiment.Name}");
                    try
                    {

                        experiment.ScenarioStarted += (s, e) =>
                        {
                            Logit($"Info:   Started Scenarion={e.Scenario.Name}");
                        };
                        experiment.ScenarioEnded += (s, e) =>
                        {
                            Logit($"Info:   Ended Scenarion={e.Scenario.Name}");
                        };

                        experiment.ReplicationStarted += (s, e) =>
                        {
                            Logit($"Info:   Started Replication={e.ReplicationNumber} (Scenario={e.Scenario.Name})");

                        };
                        experiment.ReplicationEnded += (s, e) =>
                        {
                            Logit($"Info:   Ended Replication={e.ReplicationNumber} (Scenario={e.Scenario.Name})");
                            if (e.ReplicationEndedState != ExperimentationStatus.Completed )
                            {
                                Logit($"Warning: Replication ended. State={e.ReplicationEndedState} Message={e.ErrorMessage}");
                            }
                        };

                        Logit($"Info: Experiment={experiment.Name} resetting..");
                        experiment.Reset();
                        Logit($"Info: Experiment={experiment.Name} starting run...");
                        experiment.Run();
                        Logit($"Info: Experiment={experiment.Name} finished run.");
                    }
                    catch (Exception ex)
                    {
                        Logit($"Experiment={experiment.Name} failed. Error={ex.Message}");
                    }
                }
                else
                    Logit($"Could not locate an Experiment");
            }
            else
            {
                Logit($"Could not locate model={modelName}");
            }

            string savePath = Path.Combine(workPath, $"{projectName}-Saved.spfx");
            Logit($"Info: Saving project to={savePath}...");
            projectExperiment.SupportSourceControl = saveWithSourceControlSupport;

            SimioProjectFactory.SaveProject(projectExperiment, savePath, out string[] saveWarnings);
            if ( saveWarnings?.Length > 0)
            {
                int count = 0;
                foreach ( var warning in saveWarnings )
                {
                    Logit($"Save Warning #{++count}={warning}");
                }
            }
            Logit($"Info: Project (with Experiment) saved to={savePath}.");

            //========== Now do a Plan ===============
            DoPlan:

            projectPath = Path.Combine(workPath, Settings.Default.PlanProjectFilename);
            if (!File.Exists(projectPath))
            {
                AlertAndExit($"Cannot locate Plan Project Path={projectPath}. Model Name (from Settings)={Settings.Default.ExperimentProjectFilename}");
            }

            projectName = Path.GetFileNameWithoutExtension(projectPath);

            Logit($"Info: Project Path={projectPath} exists. ProjectName={projectName} Loading...");
            ISimioProject projectWithPlan = SimioProjectFactory.LoadProject(projectPath, out string[] loadWarnings2);
            if (loadWarnings2?.Length > 0)
            {
                Logit($"Found {loadWarnings?.Length} Warnings while loading={projectPath}");
                int count = 0;
                foreach (var warning in loadWarnings2)
                {
                    Logit($"Warning: Load Warning #{++count}={warning}");
                }
            }

            modelName = Settings.Default.PlanModelName;
            model = projectWithPlan.Models[modelName];

            if (model != null)
            {
                Logit($"Info: Project={projectName} Plan: Located model={modelName}. ");

                IPlan plan = model.Plan; // get the plan
                if (plan != null)
                {
                    try
                    {
                        Logit($"Info: Plan starting run...");
                        RunPlanOptions runOptions = new RunPlanOptions();
                        runOptions.AllowDesignErrors = false;
                        plan.RunPlan(runOptions);
                        Logit($"Info: Plan finished run. {plan.TargetResults}.");

                        ITargetResults results = plan.TargetResults;
                        if (results != null) 
                        {
                            Logit($" There are {results.Count} results:");
                            int resultCount = 0;
                            foreach (var result in results)
                            {
                                Logit($"  {resultCount++}. TargetName={result.Target.Name} PlanValue={result.PlanValue}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Logit($"Plan={plan} Error={ex.Message}");
                    }
                }
                else
                    Logit($"Could not locate a Plan");
            }
            else
            {
                Logit($"Could not locate model={modelName}");
            }

            savePath = Path.Combine(workPath, $"{projectName}-Saved.spfx");
            projectWithPlan.SupportSourceControl = saveWithSourceControlSupport;
            Logit($"Info: Saving project to={savePath}...");
            SimioProjectFactory.SaveProject(projectWithPlan, savePath, out saveWarnings);
            if (saveWarnings?.Length > 0)
            {
                int count = 0;
                foreach (var warning in saveWarnings)
                {
                    Logit($"Save Warning #{++count}={warning}");
                }
            }
            Logit($"Info: Project (with Plan) saved to={savePath}.");

            Logit($"Run completed. Press ENTER to exit.");
            Console.ReadLine();

        } // main

        /// <summary>
        /// Put message and exit app
        /// </summary>
        /// <param name="messsage"></param>
        /// <returns></returns>
        private static void AlertAndExit(string messsage)
        {
            Console.Write($"{DateTime.Now:HH:mm:ss.fff}: {messsage}");
            Environment.Exit(-1);
        }

        /// <summary>
        /// Put out to console with timestamp
        /// </summary>
        /// <param name="messsage"></param>
        /// <returns></returns>
        private static void Logit(string messsage)
        {
            Console.WriteLine($"{DateTime.Now:HH:mm:ss.fff}: {messsage}");
        }
        private static void Alert(string messsage)
        {
            Console.WriteLine($"{DateTime.Now:HH:mm:ss.fff}: {messsage}");
            string _ = Console.ReadLine();
        }
    } // class Program
}
