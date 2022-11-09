using SimioAPI;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RunSimioSchedule
{

    /// <summary>
    /// Everything needed for Running a Schedule headless.
    /// </summary>
    public class RunContext
    {
        public ISimioProject SimioProject { get; set; }

        public string ModelName { get; set; }

        /// <summary>
        /// Full path to project (e.g. spfx) file
        /// </summary>
        public string ProjectFilepath { get; set; }

        /// <summary>
        /// Full path to the file that is looked for 
        /// </summary>
        public string EventFilepath { get; set; }

        /// <summary>
        /// Full path to the file containing the schedule export file.
        /// </summary>
        public string ExportScheduleFilepath { get; set; }

        /// <summary>
        /// Where the extensions path is set to during the contructor.
        /// </summary>
        public string ExtensionPath { get; private set; }

        public bool DeleteStatusBeforeRun { get; set; }

        public bool SaveProject { get; set; }

        public bool RunPlan { get; set; } = true;

        public bool RunRiskAnalysis { get; set; } = false;

        public bool ExportSchedule { get; set; } = false;


        /// <summary>
        /// Constructor. Set up the ISimioProject and checks directories.
        /// </summary>
        /// <param name="modelName"></param>
        /// <param name="projectPath"></param>
        /// <param name="eventPath"></param>
        /// <param name="exportSchedulePath"></param>
        public RunContext(string modelName, string projectPath, string eventPath, string exportSchedulePath)
        {
            string marker = "Begin.";
            try
            {
                ModelName = modelName;
                ProjectFilepath = projectPath;
                EventFilepath = eventPath;
                ExportScheduleFilepath = exportSchedulePath;

                marker = $"RunContext Project={projectPath}. Model={ModelName}";

                ExtensionPath = System.AppDomain.CurrentDomain.BaseDirectory;
                marker = $"Setting ExtensionPath to={ExtensionPath}";
                SimioProjectFactory.SetExtensionsPath(ExtensionPath);

                // Set event file.
                string eventFolder = Path.GetDirectoryName(EventFilepath);
                if (!Directory.Exists(eventFolder))
                {
                    throw new ApplicationException($"Event Folder ({eventFolder} not found.");
                }

                // If File Not Exist, Throw Exeption
                if (File.Exists(projectPath) == false)
                {
                    throw new ApplicationException($"Project Not Found={ProjectFilepath}");
                }

                // Open project file.
                SimioProject = SimioProjectFactory.LoadProject(ProjectFilepath, out string[] warnings);
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Project={projectPath} Marker={marker} Err={ex}");
            }
        } // constructor

    }  // RunContext class

    /// <summary>
    /// Helper class.
    /// </summary>
    public static class HeadlessHelpers
    {

        /// <summary>
        /// Look for the event file.
        /// When found, run the model.
        /// </summary>
        /// <param name="modelName"></param>
        public static void RunScheduleExportResultsAndSaveProject(RunContext runContext)
        {

            string marker = "Begin";

            if ( runContext.SimioProject == null )
            {
                throw new ApplicationException($"Simio Project not loaded.");
            }

            try
            {
                var exceptionsTable = new DataTable();
                LogIt(marker = $"Reading Event File={runContext.EventFilepath}");
                exceptionsTable = ReadCsvFileIntoDataTable(runContext.EventFilepath);

                LogIt(marker = $"Deleting Event File={runContext.EventFilepath}");
                File.Delete(runContext.EventFilepath);

                LogIt(marker = $"Getting Model={runContext.ModelName}");
                var model = runContext.SimioProject.Models[runContext.ModelName];
                if (model == null)
                {
                    LogIt($"Model={model} Not Found In Project");
                }
                else
                {
                    // Delete Status
                    if (runContext.DeleteStatusBeforeRun)
                    {
                        LogIt(marker = $"Deleting Status before Run.");
                        DeleteStatus();
                    }

                    if (exceptionsTable != null)
                    {
                        LogIt(marker = "Reading Resource Exceptions");
                        ImportDowntime(model, exceptionsTable);
                        // Save Projects
                        if (runContext.SaveProject)
                        {
                            LogIt(marker = "Saving Project Prior To Run Plan");
                            SaveProject(runContext);
                        }
                    }

                    // Start Plan
                    LogIt(marker = $"Run Plan?={runContext.RunPlan}");
                    if (runContext.RunPlan)
                    {
                        RunPlanOptions options = new RunPlanOptions 
                        {
                            AllowDesignErrors = false
                        };

                        model.Plan.RunPlan(options);
                    }

                    LogIt(marker = $"Run Risk Analysis?={runContext.RunRiskAnalysis}");
                    if (runContext.RunRiskAnalysis)
                    {
                        model.Plan.RunRiskAnalysis();
                    }

                    LogIt(marker = $"Export Schedule?={runContext.ExportSchedule}");
                    if ( runContext.ExportSchedule)
                    {
                        LogIt(marker = $"Analyze Risk Finished...Exporting Schedule={runContext.ExportScheduleFilepath}");
                        ExportSchedule(model, runContext.ExportScheduleFilepath);
                    }

                    LogIt(marker = "Actions Complete. Optional Save...");
                    // Save Projects
                    if (runContext.SaveProject)
                    {
                        LogIt(marker = "Saving Project After Schedule Run");
                        SaveProject(runContext);
                    }
                    LogIt(marker = "Run completed.");

                }

            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Marker={marker} Err={ex}");
            }
        }


        /// <summary>
        /// Import a downtime table
        /// </summary>
        /// <param name="model"></param>
        /// <param name="exceptionsTable"></param>
        /// <returns></returns>
        private static Int32 ImportDowntime(IModel model, DataTable exceptionsTable)
        {
            Int32 numberOfUpdates = 0;

            var resorucesTable = model.Tables["Resources"];
            if (resorucesTable == null)
            {
                LogIt("Info: Resources table does not exist in model. Not used.");
                return 0;
            }

            foreach (var resourcesRow in resorucesTable.Rows)
            {
                var intellObj = model.Facility.IntelligentObjects[resourcesRow.Properties["ResourceName"].Value];

                if (intellObj != null)
                {
                    foreach (IProperty prop in intellObj.Properties)
                    {
                        if (prop.Name == "WorkPeriodExceptions")
                        {
                            if (prop is IRepeatingProperty repeatProp)
                            {
                                repeatProp.Rows.Clear();
                                foreach (System.Data.DataRow dr in exceptionsTable.Rows)
                                {
                                    if (intellObj.ObjectName == dr[0].ToString())
                                    {
                                        var repeatRow = repeatProp.Rows.Create();
                                        repeatRow.Properties[0].Value = Convert.ToDateTime(dr[1]).ToString();
                                        repeatRow.Properties[1].Value = Convert.ToDateTime(dr[2]).ToString();
                                        repeatRow.Properties[3].Value = dr[3].ToString();
                                        numberOfUpdates++;
                                    }
                                } // for each data row
                            }
                        }
                    } // for each propert
                }
            } // for each row

            return numberOfUpdates;
        }

        /// <summary>
        /// Read a text CSV (Comma Separated Value) file where the first row is column names
        /// and return as a Microsoft DataTable
        /// </summary>
        /// <param name="fileNameAndPath"></param>
        /// <returns></returns>
        public static DataTable ReadCsvFileIntoDataTable(string fileNameAndPath)
        {
            var dtCsv = new DataTable();
            string Fulltext;

            try
            {
                using (StreamReader sr = new StreamReader(fileNameAndPath))
                {
                    while (!sr.EndOfStream)
                    {
                        Fulltext = sr.ReadToEnd().ToString(); //read full file text  
                        string[] rows = Fulltext.Split('\n'); //split full file text into rows  
                        for (int ii = 0; ii < rows.Count() - 1; ii++)
                        {
                            string[] rowValues = rows[ii].Split(','); //split each row with comma to get individual values  
                            {
                                if (ii == 0)
                                {
                                    for (int jj = 0; jj < rowValues.Count(); jj++)
                                    {
                                        dtCsv.Columns.Add(rowValues[jj]); // Assuming first row has column names 
                                    }
                                }
                                else
                                {
                                    DataRow dr = dtCsv.NewRow();
                                    for (int kk = 0; kk < rowValues.Count(); kk++)
                                    {
                                        dr[kk] = rowValues[kk].ToString();
                                    }
                                    dtCsv.Rows.Add(dr); //add other rows  
                                }
                            }
                        }
                    }
                } // using streamreader

                return dtCsv;
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"File={fileNameAndPath} Err={ex}");
            }

        }

        /// <summary>
        /// Export the schedule file.
        /// </summary>
        /// <param name="model"></param>
        /// <param name="exportScheduleFile"></param>
        private static void ExportSchedule(IModel model, string exportScheduleFile)
        {
            // Export Schedule
            string marker = $"Begin Exporting to={exportScheduleFile}";
            try
            {
                marker = $"Checking for ExportScheduleFile={exportScheduleFile}";
                if (System.IO.File.Exists(exportScheduleFile))
                {
                    marker = $"Checking for ExportScheduleFile={exportScheduleFile}";
                    File.Delete(exportScheduleFile);
                }

                marker = $"Converting Log and Results to DataTable";
                var dataSet = new DataSet();
                var dataTable = ConvertResourceUsageLogAndTargetResultToDataTable(model);
                dataTable.TableName = "ResourceUsageLog";
                dataSet.Tables.Add(dataTable);

                marker = $"Writing ExportScheduleFile={exportScheduleFile}";
                dataSet.WriteXml(exportScheduleFile);
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Marker={marker} Err={ex}");
            }
        }

        /// <summary>
        /// Use the API to save the project
        /// </summary>
        private static void SaveProject(RunContext runContext)
        {
            try
            {
                string[] warnings;
                SimioProjectFactory.SaveProject(runContext.SimioProject, runContext.ProjectFilepath, out warnings);
                if (warnings?.Count() > 0)
                    LogIt($"Warning: Project saved with {warnings?.Count()} warnings");

                LogIt($"Info: Project Saved.");
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"SavingProject={runContext.ProjectFilepath}. Err={ex}");
            }
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

        private static void DeleteStatus()
        {
            try
            {
                if (System.IO.File.Exists(Properties.Settings.Default.StatusFile))
                {
                    File.Delete(Properties.Settings.Default.StatusFile);
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
        /// Convert a Simio Table to a Microsoft DataTable
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

