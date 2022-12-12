using SimioAPI;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace RunSimioSchedule2
{


    /// <summary>
    /// Helper class.
    /// </summary>
    public static class HeadlessHelpers
    {

        /// <summary>
        /// Make sure there is a project file path and move it to Processed. 
        /// If all is well, then load the project and find our Model.
        /// Then use booleans to determine what to run.
        /// </summary>
        /// <param name="modelName"></param>
        public static bool RunScheduleAndSaveProject(RunContext runContext, string incomingFilepath)
        {

            string marker = "Begin";

            if (!File.Exists(incomingFilepath))
                throw new ApplicationException($"Simio Project={incomingFilepath} not found.");

            try
            {
                runContext.ProjectFilename = Path.GetFileName(incomingFilepath);

                HeadlessHelpers.MoveFileToNewFolder( incomingFilepath, runContext.ProcessingFolderpath);

                bool isFaulted = true; // default

                // Open project file and pay attention to any warnings.
                runContext.SimioProject = SimioProjectFactory.LoadProject(runContext.ProcessingFilepath, out string[] warnings);
                if (warnings.Length > 0)
                {
                    LogIt($"Loaded={runContext.ProcessingFilepath} but with {warnings.Length} warnings:");
                    int ii = 0;
                    foreach (string warning in warnings)
                    {
                        LogIt($"   Warning {++ii}: {warning}");
                    }
                    LogIt($" Warning were found. Model will not be run.");
                }
                else
                {
                    string modelName = runContext.ModelName; // default name

                    if (string.IsNullOrEmpty(modelName))
                        modelName = "Model"; // default

                    try
                    {
                        LogIt(marker = $"Looking for Model={modelName}");
                        var model = runContext.SimioProject.Models[modelName];
                        if (model == null)
                        {
                            LogIt($"Model={modelName} Not Found In Project");
                        }
                        else
                        {
                            // Delete Status
                            if (runContext.IsStatusToBeDeletedBeforeRun)
                            {
                                LogIt(marker = $"Deleting Status before Run.");
                                DeleteStatus();
                            }

                            // Start Plan
                            LogIt(marker = $"Run Plan? ({runContext.IsPlanToBeRun})");
                            if (runContext.IsPlanToBeRun)
                            {
                                RunPlanOptions options = new RunPlanOptions
                                {
                                    AllowDesignErrors = false
                                };

                                LogIt(marker = "Running Plan...");
                                model.Plan.RunPlan(options);
                            }

                            LogIt(marker = $"Run Completed. Optional Risk Analysis ({runContext.IsRiskAnalysisToBeRun})");
                            if (runContext.IsRiskAnalysisToBeRun)
                            {
                                LogIt(marker = "Running Risk Analysis...");
                                model.Plan.RunRiskAnalysis();
                            }

                            LogIt(marker = "Run Plan completed.");
                            isFaulted = false;

                            // Also look for experiments
                            LogIt(marker = $"Run Experiment? ({runContext.IsExperimentToBeRun})");
                            if (runContext.IsExperimentToBeRun && !isFaulted)
                            {
                                // Also check for any experiments, if we wish them to be run
                                string experimentName = runContext.ExperimentName; // default name

                                if (string.IsNullOrEmpty(experimentName))
                                    experimentName = "Experiment1"; // default

                                LogIt(marker = $"Looking within Model={model} for Experiment={experimentName}");
                                var experiment = model.Experiments[experimentName];
                                if (experiment == null)
                                {
                                    LogIt($"Experiment={experimentName} Not Found In Model={modelName}");
                                    isFaulted = true;
                                }
                                else // found Experiment
                                {
                                    // Run is synchronously
                                    LogIt(marker = $"Ready to Reset and Run Experiment={experimentName} of Model={modelName} ");
                                    experiment.Reset();
                                    experiment.Run();

                                    LogIt(marker = $"Experiment Completed. ");
                                    isFaulted = false;
                                }

                            } // experiment was run

                            LogIt(marker = $"Actions Complete. Optional Save. ({runContext.IsProjectToBeSaved})");
                            // Save Projects
                            if (runContext.IsProjectToBeSaved)
                            {
                                LogIt(marker = "Saving Project (with all Results)");
                                SaveProject(runContext);
                            }

                        } // model was found

                    }
                    catch (Exception ex)
                    {
                        LogIt($"Project={runContext.ProcessingFilepath}. Marker={marker}. Error={ex.Message}");
                    }

                } // load was ok

                // Oops. Something went wrong. Move the project file to the error folder.
                if (isFaulted)
                {
                    MoveFileToNewFolder(runContext.ProcessingFilepath, runContext.ErrorFolderpath);
                }
                else // Not faulted. Move to Saved
                {
                    MoveFileToNewFolder(runContext.ProcessingFilepath, runContext.SuccessFolderpath);
                }

                LogFlush();
                return isFaulted == false;
            }
            catch (Exception ex)
            {
                LogIt($"Marker={marker} Err={ex}");
                return false;
            }
        }

        /// <summary>
        /// Move a file in the sourceFilepath to a new folder using the same filename.
        /// If the target file already exists there, then delete it before moving the new one.
        /// Throw exception if there are issues.
        /// </summary>
        /// <returns></returns>
        public static void MoveFileToNewFolder(string sourceFilepath, string newFolderpath)
        {
            try
            {
                LogIt($"Info: Moving File={sourceFilepath} to folder={newFolderpath}");
                if (!File.Exists(sourceFilepath))
                    throw new ApplicationException($"No such file={sourceFilepath}");
                if (!Directory.Exists(newFolderpath))
                    throw new ApplicationException($"No such file={sourceFilepath}");

                string sourceFolderpath = Path.GetDirectoryName(sourceFilepath);
                string sourceFilename = Path.GetFileName(sourceFilepath);
                string targetFilePath = Path.Combine(newFolderpath, sourceFilename);

                if (File.Exists(targetFilePath))
                    File.Delete(targetFilePath);

                File.Move(sourceFilepath, targetFilePath);
                return;
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Source={sourceFilepath} TargetFolder={newFolderpath}, Err={ex.Message}");
            }
        }


        /// <summary>
        /// Use the SimEngine API to save the project to the same folder where it was processed.
        /// </summary>
        private static void SaveProject(RunContext runContext)
        {
            try
            {
                LogIt($"Info: Project saving to={runContext.ProcessingFilepath}...");
                SimioProjectFactory.SaveProject(runContext.SimioProject, runContext.ProcessingFilepath, out string[] warnings);
                if (warnings?.Count() > 0)
                {
                    LogIt($"Warning: Project saved to={runContext.SuccessFilepath} with {warnings?.Count()} warnings");

                    int ii = 0;
                    foreach (string warning in warnings)
                    {
                        LogIt($"   Warning {++ii}: {warning}");
                    }
                }
                else
                {
                    LogIt($"Info: Project Saved to={runContext.ProcessingFilepath}");
                }

            }
            catch (Exception ex)
            {
                throw new ApplicationException($"SavingProject={runContext.ProcessingFilepath}. Err={ex.Message}");
            }
        }

        /// <summary>
        /// Log message to a text file specified in the Properties
        /// The message is prepended with a timetamp
        /// and the lines are truncated to 32000 characters.
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="append message"></param>
        public static void LogIt( string msg, bool append)
        {
            try
            {
                using (StreamWriter file = new StreamWriter(Properties.Settings.Default.StatusFilepath, append))
                {
                    string timestamp = System.DateTime.Now.ToString("yyyy-MM-dd (MMM) HH:mm:ss");
                    msg = $"{timestamp}: {msg}";
                    if (msg.Length > 32000) msg = msg.Substring(0, 32000);
                    file.WriteLine(msg);

                }
            }
            catch { }
        }

        public static void LogIt( string msg)
        {
            LogIt(msg, true);
        }

        public static void LogFlush()
        {
            using (StreamWriter file = new StreamWriter(Properties.Settings.Default.StatusFilepath, true))
            {
                file.Flush();
            }
        }


        /// <summary>
        /// Delete the Status file
        /// </summary>
        private static void DeleteStatus()
        {
            try
            {
                if (File.Exists(Properties.Settings.Default.StatusFilepath))
                {
                    File.Delete(Properties.Settings.Default.StatusFilepath);
                }
            }
            catch { }
        }


        /// <summary>
        /// Convert the Resource Usage log 
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        private static DataTable ConvertResourceUsageLogAndTargetResultToDataTable(IModel model)
        {

            List<string[]> logList = new List<string[]>();

            // get all column names
            List<string> colNames = new List<string>();
            colNames.Add("OwnerName");
            colNames.Add("Probability");
            colNames.Add("ResourceName");
            colNames.Add("StartTime");
            colNames.Add("EndTime");
            logList.Add(colNames.ToArray());

            // each row in resource usage log
            foreach (var row in model.Plan.ResourceUsageLog)
            {
                List<string> thisRow = new List<string>();
                // get data
                thisRow.Add(row.OwnerName);
                Boolean foundFlag = false;
                if (row.OwnerId != null)
                {
                    foreach (var item in model.Plan.TargetResults)
                    {
                        if (item.Owner == row.OwnerId)
                        {
                            thisRow.Add(item.RiskWithinBoundsProbability.ToString());
                            foundFlag = true;
                            break;
                        }
                    }
                }
                if (foundFlag == false) thisRow.Add("");
                if (row.ResourceName != null) thisRow.Add(row.ResourceName);
                else thisRow.Add("");
                if (row.StartTime != null) thisRow.Add(row.StartTime.ToString("s"));
                else thisRow.Add("");
                if (row.EndTime != null) thisRow.Add(row.EndTime.ToString("s"));
                else thisRow.Add("");
                logList.Add(thisRow.ToArray());
            }

            // New table.
            var dataTable = new DataTable();

            // Get max columns.
            int columns = 0;
            foreach (var array in logList)
            {
                if (array.Length > columns)
                {
                    columns = array.Length;
                }
            }

            // Add columns.
            for (int ii = 0; ii < columns; ii++)
            {
                var array = logList[0];
                dataTable.Columns.Add(array[ii]);
            }

            // Remove Column Headings
            if (logList.Count > 0)
            {
                logList.RemoveAt(0);
            }

            // Add rows.
            foreach (var array in logList)
            {
                dataTable.Rows.Add(array);
            }

            return dataTable;
        }

        /// <summary>
        /// Convert a Simio Table to a Microsoft DataTable.
        /// Note: this is not used in this project but is left here as a reference.
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        private static DataTable ConvertTableToDataTable(ITable table)
        {
            List<string[]> tableList = new List<string[]>();
            int rowNumber = 0;

            // get all column names
            List<string> colNames = new List<string>();

            // get property column names
            List<string> propColNames = new List<string>();
            foreach (var col in table.Columns)
            {
                colNames.Add(col.Name);
                propColNames.Add(col.Name);
            }
            // get state column names
            List<string> stateColNames = new List<string>();
            foreach (var stateCol in table.StateColumns)
            {
                colNames.Add(stateCol.Name);
                stateColNames.Add(stateCol.Name);
            }
            tableList.Add(colNames.ToArray());

            // Get Row Data
            foreach (var row in table.Rows)
            {
                rowNumber++;
                List<string> thisRow = new List<string>();
                // get properties
                foreach (var array in propColNames)
                {
                    if (row.Properties[array.ToString()].Value != null) thisRow.Add(row.Properties[array.ToString()].Value);
                    else thisRow.Add("");
                }
                // get states
                foreach (var array in stateColNames)
                {
                    if (table.StateRows[rowNumber - 1].StateValues[array.ToString()].PlanValue != null) thisRow.Add(table.StateRows[rowNumber - 1].StateValues[array.ToString()].PlanValue.ToString());
                    else thisRow.Add("");
                }
                tableList.Add(thisRow.ToArray());
            }

            // New table.
            var dataTable = new DataTable();

            // Get max columns.
            int columns = 0;
            foreach (var array in tableList)
            {
                if (array.Length > columns)
                {
                    columns = array.Length;
                }
            }

            // Add columns.
            for (int ii = 0; ii < columns; ii++)
            {
                var array = tableList[0];
                dataTable.Columns.Add(array[ii]);
            }

            // Remove Column Headings
            if (tableList.Count > 0)
            {
                tableList.RemoveAt(0);
            }

            // sort rows
            //var sortedList = list.OrderBy(x => x[0]).ThenBy(x => x[3]).ToList();

            // Add rows.
            foreach (var array in tableList)
            {
                dataTable.Rows.Add(array);
            }

            return dataTable;
        }
    }

}

