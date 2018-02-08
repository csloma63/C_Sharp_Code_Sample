using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections;
using System.Configuration;

namespace ApaBuilders
{
    public class AMF : IApaBuilder
    {
        public readonly static string amFreightErrors = ConfigurationManager.AppSettings["AM_FREIGHT_APA_EMAIL"];

        CodeFinder codes;
        public string SpreadSheetFileName = string.Empty;// Holds the name of the Spreadsheet
        public string SpreadSheetPath = string.Empty;
        public int StartCount;
        public List<string> FinalList;
        DataTable dtPlates;
        DataTable finalSort;
        public string RunSheetPath { get; private set; }
        public string RunSheetFileName { get; private set; }
        public AMF(CodeFinder _codes){ codes = _codes; }
        public void DoWork()
        {
            GetInputSheet();
            GetDataFromRunSheet(out dtPlates);
            AlterDescriptionValue(dtPlates);
            AssignGroups(dtPlates);
            FilterAndConsolidateInput(dtPlates);
            CreateTextFiles();
            CreateApaFile();
        }
        private void GetInputSheet()
        {
            // FIND RUNSHEET TO PROCESS
            var ext = new List<string> { ".xls", ".xlsx" };
            List<string> runsheet = Directory.GetFiles(codes.Input, "*.*", SearchOption.AllDirectories).Where(s => ext.Contains(Path.GetExtension(s))).ToList<string>();
            if (runsheet.Count > 1 || runsheet.Count == 0)
            {
                File.AppendAllText(Utils.logFile, DateTime.Now.ToString("MM/dd/yyyy HH:mm ") + " :Check the American Freight InputDirectory, there is either no files in it or more than one.  Correct and rerun. \r\n");
                Utils.SendEmail(amFreightErrors, "AMF", "File path does not contain customer name. \r\n");
                Environment.Exit(0);
            }

            RunSheetPath = runsheet[0];
            //RunSheetFileName = Path.GetFileName(RunSheetPath);
        }
        private void GetDataFromRunSheet(out DataTable dtPlates)
        {
            string plates = string.Empty;
            //Gets a list of sheets names from inputsheet in case the sheet names get changed.  Returns an alphabetical list,
            //so as long a sheet starts with Plates... 
            List<String> getSheetNames = new List<String>();
            getSheetNames = Utils.GetExcelSheetNames(RunSheetPath);
            foreach (string sheetname in getSheetNames)
            {
                if (sheetname.ToLower().Contains("plates")) { plates = sheetname.Replace("'", "").ToString(); }
            }
            if (String.IsNullOrEmpty(plates))
            {
                File.AppendAllText(Utils.logFile, DateTime.Now.ToString("MM/dd/yyyy HH:mm ") + " :Check the American Freight InputDirectory, there is either no files in it or more than one.  Correct and rerun. \r\n");
                Utils.SendEmail(amFreightErrors, "AMF", "Can not find a input sheet tab with the word 'PLATES' on it.");//Message for failing to meet conditions
                Environment.Exit(0);
            }
            string query1 = "SELECT * FROM [" + plates + "A4:W65536] where 'Run Order' is not null";//Start at row 4 and read to the end.
            dtPlates = Utils.DtFromExcelSheet(RunSheetPath, query1);
            foreach (DataRow dr in dtPlates.Rows)
            {
                if (!String.IsNullOrEmpty(dr["Run Order"].ToString()))
                {
                    StartCount++;
                }
            }
            dtPlates.Rows.RemoveAt(StartCount);
        }
        private void AlterDescriptionValue(DataTable dtPlates)
        {
            foreach (DataRow dr in dtPlates.Rows)
            {
                string mystring = dr["Description"].ToString();
                dr["Description"] = mystring.Split(' ').Last();
                //dr["Description"] = mystring.Substring(Math.Max(0, mystring.Length - 4));
            }
        }
        private static void AssignGroups(DataTable dtPlates)
        {
            dtPlates.Columns.Add("Group", typeof(System.Int32));
            int groupNumber = 0;
            foreach (DataRow dr in dtPlates.Rows)
            {
                if (!dr.IsNull("Plates"))
                {
                    if (dr["Plates"].ToString().Equals("1")) { dr["Group"] = groupNumber; }
                    else { groupNumber++; dr["Group"] = groupNumber; }
                }
            }
        }
        private void FilterAndConsolidateInput(DataTable dtPlates)
        {
            try
            {
                int sets = Convert.ToInt32(dtPlates.Compute("max([Group])", string.Empty));//gets number of plate groups
                DataTable dtPlatesFirstSort = dtPlates.Select("", "Group, Run Quantity DESC").CopyToDataTable();//sorts by group then quantity desc
                finalSort = dtPlates.Clone();
                for (int i = 1; i <= sets; i++)
                {
                    DataTable resultSet = dtPlatesFirstSort.Select("Group = " + i).CopyToDataTable();//separates group to another datatable for futher sorting
                    if (resultSet.Rows.Count == 1) { finalSort.ImportRow(resultSet.Rows[0]); }//If there is only one entry in the plate group
                    else
                    {
                        for (int c = 0; c < resultSet.Rows.Count; c = c + 2)//iterate the spreadsheet by 2 to compare
                        {
                            if (resultSet.Rows.Count - c == 1) { finalSort.ImportRow(resultSet.Rows[c]); break; }//Takes care of the last line items if there is nothing for comparison
                            int value = Convert.ToInt32(resultSet.Rows[c]["Run Quantity"]);
                            int compareValue = Convert.ToInt32(resultSet.Rows[c + 1]["Run Quantity"]);
                            if (value - compareValue <= 15000)
                            {
                                resultSet.Rows[c]["Description"] = resultSet.Rows[c]["Description"] + "-" + resultSet.Rows[c + 1]["Description"];//concatinate description for items within 15000 qty
                                resultSet.Rows[c]["Run Quantity"] = value + compareValue;

                                DateTime date1 = DateTime.Parse(resultSet.Rows[c]["Ship Date"].ToString());
                                DateTime date2 = DateTime.Parse(resultSet.Rows[c + 1]["Ship Date"].ToString());
                                int result = DateTime.Compare(date1, date2);
                                if (result > 0) { resultSet.Rows[c]["Ship Date"] = resultSet.Rows[c + 1]["Ship Date"]; }// select closest ship date from comparison

                                int firstRunOrder = Convert.ToInt32(resultSet.Rows[c]["Run Order"]);
                                int secondRunOrder = Convert.ToInt32(resultSet.Rows[c + 1]["Run Order"]);
                                resultSet.Rows[c]["Run Order"] = Math.Min(firstRunOrder, secondRunOrder);

                                finalSort.ImportRow(resultSet.Rows[c]);
                            }
                            else
                            {
                                finalSort.ImportRow(resultSet.Rows[c]); c--;
                            }
                        }
                    }
                }
                //finalSort = finalSort.Select("", "Group, Ship Date").CopyToDataTable();//sorts by ship date
                finalSort = finalSort.Select("", "Group, Run Order").CopyToDataTable();//sorts by ship date
            }
            catch (Exception ex)
            {
                File.AppendAllText(Utils.logFile, DateTime.Now.ToString("MM/dd/yyyy HH:mm ") + " :Failure in the FilterAndConsolidateInput method: \r\n" + ex.Message);
                Utils.SendEmail(amFreightErrors, "AMF", "Failure in the FilterAndConsolidateInput method: \r\n" + ex.Message);
                Environment.Exit(0);
            }
        }
        private void CreateTextFiles()
        {
            FinalList = new List<string>();
            string output;
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < finalSort.Rows.Count; i++)
            {
                int number = i + 1;// need to pad the count numbers for things to work down the line
                if (finalSort.Rows.Count >= 100) { output = number.ToString("000") + "_" + finalSort.Rows[i]["Description"]; }
                else { output = number.ToString("00") + "_" + finalSort.Rows[i]["Description"]; }
                FinalList.Add(output);//create a List to use to build the APA file
                sb.Append(output + Environment.NewLine);
                Utils.MessageWriter(finalSort.Rows[i]["Ship Date"].ToString(), output + ".txt", codes.OutputFolder);
            }
            Utils.MessageWriter(sb.ToString(), "AmerFreight.txt", codes.OutputFolder);
            sb.Clear();
        }
        public void CreateApaFile()
        {
            StringBuilder sb = new StringBuilder();
            string APAfileName = "AMFAPA_OUT";
            sb.Append("!APA 1.0" + Environment.NewLine);//THIS IS A REQUIRED LINE. IT MUST BE THE FIRST LINE OF THE APA FILE!
            foreach (string set in FinalList)
            {
                if (set.Length <= 8)
                {
                    string first = set.Substring(Math.Max(0, set.Length - 4));
                    sb.Append(String.Format("ASSIGN= \"[$] {0}-[$].p[$page].pdf\" \"{1}\" [$page] 1", first, set) + Environment.NewLine);
                }
                else
                {
                    int index = set.IndexOf("_") + 1;
                    string doubled = set.Substring(index, 4);
                    string second = set.Substring(Math.Max(0, set.Length - 4));
                    sb.Append(String.Format("ASSIGN= \"[$] {0}-[$].p[$page].pdf\" \"{1}\" [$page] 1", doubled, set) + Environment.NewLine);
                    sb.Append(String.Format("ASSIGN= \"[$] {0}-[$].p[$page].pdf\" \"{1}\" [$page]+2 1", second, set) + Environment.NewLine);
                }
            }
            Utils.MessageWriter(sb.ToString(), APAfileName + ".apa", codes.OutputFolder); sb.Clear();
        }
    }
}
