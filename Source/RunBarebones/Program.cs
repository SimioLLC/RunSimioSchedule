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

            string simioSubfolder = Properties.Settings.Default.SimioSubFolder; // If you keep multiple versions around, this could be something like 'Simio 248'

            // Where to look for example Simio projects
            string rootPath = Path.Combine(programFiles, "Simio LLC", simioSubfolder);  
            Logit($"Info: RootPath={rootPath}");

            Logit($"Info: Simio Subfolder={simioSubfolder}");

            //========== Common items (e.g. setting up ExtensionsFolderpath) ===============
            string extensionsFolderpath = Path.Combine( rootPath, "UserExtensions");
            Logit($"Info: Setting (User) Extensions path to={extensionsFolderpath}");
            SimioProjectFactory.SetExtensionsPath(extensionsFolderpath);

            //========== Do an Experiment ===============
            string projectName = "HospitalEmergencyDepartment";
            string loadPath = Path.Combine(rootPath, "Examples", $"{projectName}.spfx");
            if ( !File.Exists(loadPath) ) 
            {
                Alert($"Project Path={loadPath} does not exist. Exiting...");
                Environment.Exit(-1);
            }

            Logit($"Info: Project={loadPath} exists. Loading...");
            ISimioProject projectExperiment = SimioProjectFactory.LoadProject(loadPath, out string[] loadWarnings);
            if (loadWarnings?.Length > 0)
            {
                int count = 0;
                foreach (var warning in loadWarnings)
                {
                    Logit($"Warning: Load Warning #{++count}={warning}");
                }
            }

            string modelName = "Model";
            IModel model = projectExperiment.Models[modelName];

            if ( model != null )
            {
                Logit($"Info: Located model={modelName}. Using the first Experiment.");

                IExperiment experiment = model.Experiments[0]; // get the first experiment
                if (experiment != null)
                {
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
                Logit($"Could not located model={modelName}");
            }


            string savePath = Path.Combine(rootPath, $"{projectName}-Saved.spfx");
            Logit($"Info: Saving project to={savePath}...");
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

            projectName = "SchedulingDiscretePartProduction";
            loadPath = Path.Combine(rootPath, $"{projectName}.spfx");
            if (!File.Exists(loadPath))
            {
                Logit($"Project={loadPath} does not exist. Exiting...");
                Environment.Exit(-1);
            }

            Logit($"Info: Project={loadPath} exists. Loading...");
            ISimioProject projectWithPlan = SimioProjectFactory.LoadProject(loadPath, out string[] loadWarnings2);
            if (loadWarnings2?.Length > 0)
            {
                int count = 0;
                foreach (var warning in loadWarnings2)
                {
                    Logit($"Warning: Load Warning #{++count}={warning}");
                }
            }

            modelName = "Model";
            model = projectWithPlan.Models[modelName];

            if (model != null)
            {
                Logit($"Info: Located model={modelName}. Using the first Experiment.");

                IPlan plan = model.Plan; // get the plan
                if (plan != null)
                {
                    try
                    {
                        Logit($"Info: Plan starting run...");
                        RunPlanOptions runOptions = new RunPlanOptions();
                        runOptions.AllowDesignErrors = false;
                        plan.RunPlan(runOptions);
                        Logit($"Info: Plan finished run. {plan}.");
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

            savePath = Path.Combine(rootPath, $"{projectName}-Saved.spfx");
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
