using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ApaBuilders
{
    public class MND : IApaBuilder
    {
        CodeFinder codes;
        private DataTable dtExcel;
        private DataTable dtAllRunsheet;
        private DataTable dtExcelOptimized;
        List<string> gpc_versions;
        public readonly static string MNDErrors = ConfigurationManager.AppSettings["MND_APA_EMAIL"];

        public string RunSheetPath { get; private set; }
        public string SpreadSheetPath { get; private set; }
        public string RunSheetFileName { get; private set; }
        public string SpreadSheetFileName { get; private set; }

        public MND(CodeFinder _codes) { codes = _codes; }
        public void DoWork()
        {
            GetRunSheet();
            GetDataFromRunSheet(out gpc_versions);
            GetDataFromSpreadSheet(out dtExcel);
            CheckForMissingPages(dtExcel);
            ExtendDataTable(dtExcel);
        }
        private void ExtendDataTable(DataTable dtExcel)
        {
            // EXTEND THE DATATABLE TO ALLOW FOR PLATESET AND PRESS RUN OPTIMIZING
            dtExcel.Columns.Add("PlateSetID", typeof(System.Int32)); //Add the PlateSet Grouping column (this groups the data rows into PlateSets)
            dtExcel.Columns.Add("PlateSetOrder", typeof(System.Int32)); //Add the PlateSet order (within a group) column (added Just-In-Case, may not be needed)
            dtExcel.Columns.Add("PressRunOrder", typeof(System.Int32)); //Add the PressRun optimized order column (this orders the PlateSet Groups)
            foreach (DataRow row in dtExcel.Rows) //Initialize all cells to zero.
            {
                row["PlateSetID"] = 0;
                row["PlateSetOrder"] = 0;
                row["PressRunOrder"] = 0;
            }
        }

        public void GetRunSheet()
        {
            // FIND RUNSHEET TO PROCESS
            List<string> ext = new List<string> { ".xls", ".xlsx" };
            List<string> runsheets = Directory.GetFiles(codes.Input, "*.*", SearchOption.AllDirectories).Where(s => ext.Contains(Path.GetExtension(s))).ToList<string>();
            foreach (string file in runsheets)
            {
                if (file.Contains("Grandville")) { RunSheetPath = file; }
                if (file.Contains("Pagination")) { SpreadSheetPath = file; }
            }
            if (String.IsNullOrEmpty(RunSheetPath))
            {
                File.AppendAllText(Utils.logFile, DateTime.Now.ToString("MM/dd/yyyy HH:mm ") + " :MND Run sheet not found. \r\n");
                Utils.SendEmail(MNDErrors, "MND", "MND Run sheet not found. ");
            }
            else { RunSheetFileName = Path.GetFileName(RunSheetPath); }
            if (String.IsNullOrEmpty(SpreadSheetPath))
            {
                File.AppendAllText(Utils.logFile, DateTime.Now.ToString("MM/dd/yyyy HH:mm ") + " :MND Pagination sheet not found. \r\n");
                Utils.SendEmail(MNDErrors, "MND", "MND Pagination sheet not found. ");
            }
            else { SpreadSheetFileName = Path.GetFileName(SpreadSheetPath); }
        }
        private void GetDataFromRunSheet(out List<string> gpc_versions)
        {
            string location = string.Empty, preprint = string.Empty;
            //Gets a list of sheets names from runsheet in case the sheet names get changed.  Returns an alphabetical list,
            //so as long as sheet1 starts with Preprint... and sheet2 starts with Location... it will work OK.
            List<String> getSheetNames = new List<String>();
            getSheetNames = Utils.GetExcelSheetNames(RunSheetPath);
            foreach (string sheetname in getSheetNames)
            {
                if (sheetname.ToLower().Contains("preprint") || sheetname.ToLower().Contains("newspapers")) { preprint = sheetname.Replace("'", "").ToString(); }
                if (sheetname.ToLower().Contains("location")) { location = sheetname.Replace("'", "").ToString(); }
            }
            if (String.IsNullOrEmpty(location))
            {
                File.AppendAllText(Utils.logFile, DateTime.Now.ToString("MM/dd/yyyy HH:mm ") + " :Can not find a runsheet tab with the word 'Location' on it. \r\n");
                Utils.SendEmail(MNDErrors, "MND", "Can not find a runsheet tab with the word 'Location' on it. ");
                Environment.Exit(0);
            }
            if (String.IsNullOrEmpty(preprint))
            {
                File.AppendAllText(Utils.logFile, DateTime.Now.ToString("MM/dd/yyyy HH:mm ") + " :Can not find a runsheet tab with the word 'Preprint' or 'Newspapers' on it. \r\n");
                Utils.SendEmail(MNDErrors, "MND", "Can not find a runsheet tab with the word 'Preprint' or 'Newspapers' on it. ");
                Environment.Exit(0);
            }
            string query1 = "SELECT * FROM [" + preprint + "A8:W65536] WHERE Printer LIKE '%Grandville'";//Start at row 8 and read to the end.
            DataTable dtRunNewsPaper = Utils.DtFromExcelSheet(RunSheetPath, query1);
            string query2 = "SELECT * FROM [" + location + "A8:AA65536] WHERE Printer LIKE '%Grandville'";//Start at row 8 and read to the end.
            DataTable dtRunStores = Utils.DtFromExcelSheet(RunSheetPath, query2);

            // ***** Get all 'IMPT' *****
            gpc_versions = new List<string>();
            foreach (DataRow row in dtRunNewsPaper.Rows)
            {
                string value = (string)row["IMPT"];
                if (!gpc_versions.Contains(value)) { gpc_versions.Add((string)row["IMPT"]); }
            }
            foreach (DataRow row in dtRunStores.Rows)
            {
                string value = (string)row["IMPT"];
                if (!gpc_versions.Contains(value)) { gpc_versions.Add((string)row["IMPT"]); }
            }
        }
        public void GetDataFromSpreadSheet(out DataTable dtExcel)
        {
            string query1 = "SELECT * FROM [SHEET1$] WHERE Imprint LIKE '%V-%'";//<<<<<<If they want a specific Tab name on the spreadsheet, change it here.
            dtExcel = Utils.DtFromExcelSheet(SpreadSheetPath, query1);
            //remove rows that GPC does not print for.
            for (int i = dtExcel.Rows.Count - 1; i >= 0; i--)
            {
                DataRow recRow = dtExcel.Rows[i];
                bool remove = true;
                foreach (string impt in gpc_versions)
                {
                    string wtf = recRow[0].ToString();
                    if (impt.ToLower().Equals(wtf.ToLower()))
                    {
                        remove = false;
                    }
                }
                if (remove)
                {
                    dtExcel.Rows.Remove(recRow);
                    dtExcel.AcceptChanges();
                }
            }
            //remove unnecesaary columns.
            for (int i = dtExcel.Columns.Count - 1; i >= 0; i--)
            {
                DataColumn recCol = dtExcel.Columns[i];
                bool remove = true;
                if (recCol.ColumnName.Contains("Imprint") || recCol.ColumnName.Contains("PAGE") || recCol.ColumnName.Contains("WRAP")|| recCol.ColumnName.Contains("BPAG"))
                {
                    remove = false;
                }
                if (remove)
                {
                    dtExcel.Columns.Remove(recCol);
                    dtExcel.AcceptChanges();
                }
            }
        }
        private void CheckForMissingPages(DataTable dtExcel)
        {
            List<string> rtnErrors = Utils.FindMissingPages(dtExcel);
            if (rtnErrors.Count != 0)
            {
                StringBuilder output = new StringBuilder();
                output.AppendLine("There are missing pdf files in the MND dtExcel table, ApaBuilder process has been stopped. ");
                foreach (string row in rtnErrors)
                {
                    output.AppendLine(row);
                }
                File.AppendAllText(Utils.logFile, DateTime.Now.ToString("MM/dd/yyyy HH:mm ") + " :" + output.ToString() + "\r\n");
                Utils.SendEmail(MNDErrors, "MND", output.ToString());
                Environment.Exit(0);
            }
        }
    }
}
