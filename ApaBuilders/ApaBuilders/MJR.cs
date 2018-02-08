using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.IO;

namespace ApaBuilders
{
    public class MJR : IApaBuilder
    {
        CodeFinder codes;
        public readonly static string MJRErrors = ConfigurationManager.AppSettings["MJR_APA_EMAIL"];
        private List<string> allMarkets = new List<string>(ConfigurationManager.AppSettings["ALLMARKETS"].Split('-').ToList());
        private List<string> doNotRun = new List<string>(ConfigurationManager.AppSettings["DO_NOT_RUN"].Split('-').ToList());
        private List<string> paginationFiles;
        private List<string> textFiles;
        private List<string> TextFilesForSearch;

        private List<SortedDictionary<string, string>> pageChanges;
        private int pageCount = 0;
        private DataSet xlsFiles;
        private DataTable mainTable;

        public MJR() { }
        public MJR(CodeFinder _codes) { codes = _codes; }
        public void DoWork()
        {
            allMarkets = RemoveTownsFromAll(doNotRun, allMarkets);
            GetInputFilesFiles(out paginationFiles, out textFiles);
            AddTablesToDataset(out xlsFiles);
            UnwrapTextFiles(textFiles, out TextFilesForSearch);
            CleanAndProcessXlsInput();
            pageChanges = NewPageAssignments(xlsFiles);
            CreateMainTable(out mainTable);
            mainTable = AddPagesToMain(pageCount, mainTable);
            LetsFillMainTable(pageChanges);
            CreateApaFile();
        }
        private List<string> RemoveTownsFromAll(List<string> cutlist, List<string> otherMarkets)
        {
            foreach (string town in cutlist)
            {
                if (otherMarkets.Contains(town))
                {
                    int index = otherMarkets.FindIndex(x => x.StartsWith(town));
                    allMarkets.RemoveAt(index);
                }
            }
            return otherMarkets;
        }
        private void GetInputFilesFiles(out List<string> paginationFiles, out List<string> textFiles)
        {
            paginationFiles = null;
            textFiles = null;
            try
            {
                // FIND FILES TO PROCESS
                var ext = new List<string> { ".xls", ".xlsx" };
                paginationFiles = Directory.GetFiles(codes.Input, "*.*", SearchOption.AllDirectories).Where(s => ext.Contains(Path.GetExtension(s))).ToList<string>();
                textFiles = Directory.GetFiles(codes.Input, "*.txt", SearchOption.AllDirectories).ToList<string>();
                //textFiles = textFiles.OrderBy(s => new string(s.Substring(s.Length - 3).ToArray())).ToList();
            }
            catch (FileNotFoundException nf)
            {
                File.AppendAllText(Utils.logFile, DateTime.Now.ToString("MM/dd/yyyy HH:mm ") + " :Unable to find Excel files. " + nf.Message + " \r\n");
                Utils.SendEmail(MJRErrors, "MJR", "Unable to find Excel files. " + nf.Message);
            }
        }
        private void AddTablesToDataset(out DataSet xlsFiles)
        {
            xlsFiles = new DataSet();
            for (int i = 0; i < paginationFiles.Count; i++)
            {
                DataTable table = new DataTable("Table" + i);
                string query1 = "SELECT * FROM [SHEET1$]";//<<<<<<If they want a specific Tab name on the spreadsheet, change it here.
                try
                {
                    table = Utils.DtFromExcelSheet(paginationFiles[i], query1);
                }
                catch (Exception ex)
                {
                    File.AppendAllText(Utils.logFile, DateTime.Now.ToString("MM/dd/yyyy HH:mm ") + " :Failed to load data from spreadsheet in " + codes.Input + ". The import spreadsheet is not formatted properly. MJR ApaBuilder process stopped." + ex.Message + " \r\n");
                    Utils.SendEmail(MJRErrors, "MJR", "Failed to load data from spreadsheet in " + codes.Input + ". The import spreadsheet is not formatted properly. MJR ApaBuilder process stopped." + ex.Message);
                    Environment.Exit(0);
                }
                xlsFiles.Tables.Add(table);
            }
        }
        private void UnwrapTextFiles(List<string> textFiles, out List<string> textFilesForSearch)
        {
            textFilesForSearch = new List<string>();
            try
            {
                foreach (string TxtFile in textFiles)
                {
                    using (var streamReader = File.OpenText(TxtFile))
                    {
                        var lines = streamReader.ReadToEnd().Split("\r\n".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                        foreach (var line in lines)
                        {
                            textFilesForSearch.Add(line.Substring(0, line.IndexOf('<')));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText(Utils.logFile, DateTime.Now.ToString("MM/dd/yyyy HH:mm ") + " :Unable to retreive PDF names from text files. " + ex.Message + " \r\n");
                Utils.SendEmail(MJRErrors, "MJR", "Unable to retreive PDF names from text files. " + ex.Message);
            }
        }
        private void CleanAndProcessXlsInput()
        {
            foreach (DataTable table in xlsFiles.Tables)
            {
                table.Columns[0].ColumnName = "F1";
                List<int> IndicesToRemove = new List<int>();
                foreach (DataRow row in table.Rows)
                {
                    if (!Utils.IsNumeric(row[0].ToString()) && !row[0].ToString().Equals("File #"))
                    { IndicesToRemove.Add(table.Rows.IndexOf(row)); }
                }
                IndicesToRemove.Sort();
                for (int i = IndicesToRemove.Count - 1; i >= 0; i--) table.Rows.RemoveAt(IndicesToRemove[i]);
            }
        }
        private void CreateMainTable(out DataTable mainTable)
        {
            mainTable = new DataTable();
            mainTable.Columns.Add("towns", typeof(String));
            mainTable.PrimaryKey = new DataColumn[] { mainTable.Columns["towns"] };
            foreach (string town in allMarkets)
            {
                mainTable.Rows.Add(town);
            }
        }
        private DataTable AddPagesToMain(int pageCount, DataTable mainTable)
        {
            for (int i = 1; i <= pageCount; i++)
            {
                mainTable.Columns.Add("Page " + i, typeof(String));
            }
            return mainTable;
        }
        private void LetsFillMainTable(List<SortedDictionary<string, string>> listItem)
        {
            List<string> newAllMarkets = allMarkets.ToList();
            List<string> holdTowns = new List<string>();

            foreach (var item in listItem)
            {
                if (item.ContainsValue("endTable"))
                {
                    newAllMarkets = allMarkets.ToList();
                }
                holdTowns.Clear();

                if (!item.ContainsValue("endTable"))
                {
                    foreach (var dicItem in item)
                    {
                        if (dicItem.Value.Equals("endTable"))
                        {
                            newAllMarkets = allMarkets.ToList();
                            holdTowns = new List<string>();
                        }
                        else if (dicItem.Key.Equals("0") && !dicItem.Value.Equals("endTable"))
                        {
                            try
                            {
                                if (dicItem.Value.Contains("All Markets") || dicItem.Value.Contains("All Other Markets"))
                                {
                                    holdTowns = newAllMarkets.ToList();
                                }
                                else
                                {
                                    holdTowns = GetListOfTownsFromHeader(dicItem.Value);
                                    foreach (string town in holdTowns)
                                    {
                                        if (newAllMarkets.Contains(town))
                                        {
                                            newAllMarkets.RemoveAt(newAllMarkets.IndexOf(town));
                                        }
                                    }
                                }
                            }
                            catch
                            {
                                File.AppendAllText(Utils.logFile, DateTime.Now.ToString("MM/dd/yyyy HH:mm ") + " :Can not get list of MJR towns from excel spreadsheet headers. \r\n");
                                Utils.SendEmail(MJRErrors, "MJR", "Can not get list of towns from excel spreadsheet headers.");
                            }
                        }
                        else
                        {
                            foreach (string town in holdTowns)
                            {
                                if (allMarkets.Contains(town))
                                {
                                    AssignPDFs(town, dicItem);
                                }
                            }
                        }
                    }
                }
            }
        }
        private List<string> GetListOfTownsFromHeader(string towns)
        {
            return towns.Split(new Char[] { ',', '-', ' ', '_' }).ToList();
        }
        private List<SortedDictionary<string, string>> NewPageAssignments(DataSet xlsFiles)
        {
            int pageOffset = 0;
            pageChanges = new List<SortedDictionary<string, string>>();
            foreach (DataTable cTable in xlsFiles.Tables)
            {
                int count = 0;
                foreach (DataColumn col in cTable.Columns.Cast<DataColumn>().Skip(1))
                {
                    SortedDictionary<string, string> insert = new SortedDictionary<string, string>();
                    if (!cTable.Rows[0][col].ToString().Trim().Contains("Notes") && !String.IsNullOrEmpty(cTable.Rows[0][col].ToString()))
                    {
                        insert.Add("0", cTable.Rows[0][col].ToString());
                        foreach (DataRow row in cTable.Rows.Cast<DataRow>().Skip(1))
                        {
                            if (Utils.IsNumeric(row[col].ToString().TrimStart('0')) && !String.IsNullOrEmpty(row[col].ToString()))
                            {
                                //insert.Add((Convert.ToInt32(row[col]) + pageOffset).ToString().TrimStart('0'), row[0].ToString());
                                char pre = row[0].ToString()[0];
                                string key = (Convert.ToInt32(row[0].ToString().Remove(0, 1)) + pageOffset).ToString();
                                string value = pre.ToString() + row[col].ToString();
                                insert.Add(key, value);
                            }
                        }
                        if (count < Convert.ToInt32(insert.Keys.Count - 1))
                        {
                            count = Convert.ToInt32(insert.Keys.Count - 1);
                        }
                        pageChanges.Add(insert);
                    }
                    else
                    {
                        insert.Add("0", "endTable");
                        pageChanges.Add(insert);
                    }
                }
                pageOffset += count;
            }
            pageCount = pageOffset;
            return pageChanges;
        }
        private void AssignPDFs(string town, KeyValuePair<string, string> dicItem)
        {
            foreach (string pdf in TextFilesForSearch)
            {
                if (pdf.Contains(town) && pdf.Contains("_" + dicItem.Value + "_"))
                {
                    DataRow insertAt = mainTable.Rows.Find(town);
                    insertAt[Convert.ToInt32(dicItem.Key)] = pdf.Substring(0, pdf.IndexOf('.'));//***error with towns insert
                    break;
                }
            }
        }
        public void CreateApaFile()
        {
            List<string> rtnErrors = Utils.FindMissingPages(mainTable);
            if (rtnErrors.Count == 0)
            {
                //WRITE OUT THE APA DATA
                StringBuilder sb = new StringBuilder();
                string APAfileName = "MJRAPA_OUT";

                sb.Append("!APA 1.0" + Environment.NewLine);//THIS IS A REQUIRED LINE. IT MUST BE THE FIRST LINE OF THE APA FILE!
                foreach (DataRow row in mainTable.Rows)
                {
                    for (int i = 1; i < mainTable.Columns.Count; i++)
                    {
                        if (!String.IsNullOrEmpty(row[i].ToString()))
                        {
                            sb.Append("ASSIGN= \"" + row[i] + ".p1.pdf\" \"" + row[0] + "\" " + i + " 1" + Environment.NewLine);
                        }
                    }
                }
                try
                {
                    Utils.MessageWriter(sb.ToString(), APAfileName + ".apa", codes.OutputFolder); sb.Clear();
                }
                catch (Exception ex)
                {
                    File.AppendAllText(Utils.logFile, DateTime.Now.ToString("MM/dd/yyyy HH:mm ") + " :Failed to send MJR APA file to Output folder. " + ex.Message +"\r\n");
                    Utils.SendEmail(MJRErrors, "MJR", "Failed to send MJR APA file to Output folder. /r/n" + ex.Message);
                }
                Utils.MessageWriter(sb.ToString(), APAfileName + ".apa", codes.OutputFolder); sb.Clear();
            }
            else
            {
                StringBuilder output = new StringBuilder();
                output.AppendLine("There are missing text files in the MJR input folder, ApaBuilder process has been stopped. ");
                foreach (string row in rtnErrors)
                {
                    output.AppendLine(row);
                }
                File.AppendAllText(Utils.logFile, DateTime.Now.ToString("MM/dd/yyyy HH:mm ") + " :" + output.ToString() + "\r\n");
                Utils.SendEmail(MJRErrors, "MJR", output.ToString());
                Environment.Exit(0);
            }
        }
    }
}
