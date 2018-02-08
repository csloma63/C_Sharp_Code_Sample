using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Net.Mail;

namespace ApaBuilders
{
    public static class Utils
    {
        public readonly static string logFile = ConfigurationManager.AppSettings["LOG_FILE"];
        public static List<DataRow> FindNullValues(DataTable dt)
        {
            return FindNullValues(dt, new List<string>());
        }
        public static List<DataRow> FindNullValues(DataTable dt, List<string> fldExclusions)
        {
            List<DataRow> rtnErrors = new List<DataRow>();
            string filterExpresion = "";
            foreach (DataColumn dc in dt.Columns)
            {
                if (!fldExclusions.Contains(dc.ColumnName))
                {
                    string nextPart = (filterExpresion.Length == 0) ? "" : " OR ";
                    filterExpresion += nextPart + "( [" + dc.ColumnName + "] IS NULL  OR  trim([" + dc.ColumnName + "]) = '' )";
                }
            }
            rtnErrors = dt.Select(filterExpresion).ToList();
            return rtnErrors;
        }
        public static List<string> FindMissingPages(DataTable _mainTable)
        {
            List<string> rtnErrors = new List<string>();
            foreach (DataRow row in _mainTable.Rows)
            {
                foreach (DataColumn col in _mainTable.Columns)
                {
                    //test for null here
                    if (row[col] == DBNull.Value)
                    {
                        string error = String.Format("The item of {0} is missing {1}. ", row[0].ToString(), col.ColumnName.ToString());
                        rtnErrors.Add(error);
                    }
                }
            }
            return rtnErrors;
        }
        public static void MessageWriter(string message, string fileName, string outputFolder)
        {
            using (StreamWriter messagewriter = new StreamWriter(outputFolder + fileName, true))//use OutputFolder for metadata upgrade
            {
                messagewriter.Write(message);// can not be WriteLine.  Messes up the .apa builder
                messagewriter.Close();
            }
        } // send message to output file Imprint folder
        public static DataTable DtFromExcelSheet(string filePath, string query)
        {
            // **If the .exe will not start when run in Switch, verify that the compiling PC has the same ACE driver as the Switch server (currently 32-bit as of 3/30/2017).
            string excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=YES;ReadOnly=true;IMEX=1;ImportMixedTypes=Text;TypeGuessRows=0'";
            DataTable dtExcelData = new DataTable();
            excelConnect = string.Format(excelConnect, filePath);//final excel connection string
            using (OleDbConnection excelCon = new OleDbConnection(excelConnect))
            {
                excelCon.Open();
                OleDbDataAdapter oda = new OleDbDataAdapter(query, excelCon);
                oda.Fill(dtExcelData);
                foreach (DataColumn col in dtExcelData.Columns) col.ColumnName = col.ColumnName.Trim();//Makes sure customer did not add a space(s) to column name
                //clears empty rows from datatable
                dtExcelData = dtExcelData.AsEnumerable().Where(row => !row.ItemArray.All(f => f is DBNull || String.IsNullOrWhiteSpace(f.ToString()))).CopyToDataTable();
                excelCon.Close();
            }
            return dtExcelData;
        } // creates an in memory datatable from an excel spreadsheet
        public static List<String> GetExcelSheetNames(string filePath)//returns a alphabetical list of workbook spreadsheet names
        {
            // **If the .exe will not start when run in Switch, verify that the compiling PC has the same ACE driver as the Switch server (currently 32-bit as of 3/30/2017).
            string excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=YES;ReadOnly=true;IMEX=1;ImportMixedTypes=Text;TypeGuessRows=0'";
            DataTable dtExcelData = new DataTable();
            excelConnect = string.Format(excelConnect, filePath);//final excel connection string
            using (OleDbConnection excelCon = new OleDbConnection(excelConnect))
            {
                try
                {
                    excelCon.Open();
                    dtExcelData = excelCon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    if (dtExcelData == null) { return null; }
                    List<String> excelSheets = new List<String>();
                    // Add the sheet name to the string array.
                    foreach (DataRow row in dtExcelData.Rows)
                    {
                        if (!row["TABLE_NAME"].ToString().Contains("FilterDatabase"))
                        {
                            excelSheets.Add(row["TABLE_NAME".Trim()].ToString());
                        }
                    }
                    return excelSheets;
                }
                catch (Exception ex)
                {
                    //ErrorMessageWriter("Unable to obtain the sheet names from the runsheet.  Correct and rerun." + ex.Message);
                    return null;
                }
                finally
                {
                    // Clean up.
                    if (excelConnect != null)
                    {
                        excelCon.Close();
                        excelCon.Dispose();
                    }
                    if (dtExcelData != null)
                    {
                        dtExcelData.Dispose();
                    }
                }
            }
        }
        public static bool IsDirectoryEmpty(string path) => !Directory.EnumerateFileSystemEntries(path).Any();
        public static bool IsNumeric(string input)
        {
            if (string.IsNullOrEmpty(input)) { return false; }
            return int.TryParse(input, out int number);
        }
        /// <summary>
        /// Create and send emails to alert to program output or errors
        /// </summary>
        /// <param name="emails">List of receipients specified in the app.config</param>
        /// <param name="subject">What type of message or where from</param>
        /// <param name="body">information to send</param>
        /// <returns></returns>
        public static bool SendEmail(string emails, string subject, string body)
        {
            string[] recipients = emails.Split(',');
            try
            {
                //Create Smtp client to send email
                SmtpClient client = new SmtpClient("smtp.gpco.com");
                client.Port = 587;
                client.UseDefaultCredentials = false;
                client.Credentials = new System.Net.NetworkCredential("gpcadmin@gpco.com", "KM5dj1pc1-admin");
                //Create MailMessage(e-mail) to send
                MailMessage msg = new MailMessage();
                //Sender e-mail
                msg.From = new MailAddress("gpcadmin@gpco.com");

                //E-mail body
                switch (subject)
                {
                    case "args":
                        msg.Subject = "No file paths in args()";
                        msg.Body = "Unable to find file paths for the Apa Builder to run. " + body;
                        break;

                    case "MJR":
                        msg.Subject = "MJR ApaBuilder processing issue.";
                        msg.Body = body;
                        break;

                    case "AMF":
                        msg.Subject = "American Freight ApaBuilder processing issue.";
                        msg.Body = body;
                        break;

                    case "MND":
                        msg.Subject = "MND ApaBuilder processing issue.";
                        msg.Body = body;
                        break;

                    default: break;
                }
                //Recipient e-mail
                foreach (string recipient in recipients)
                {
                    msg.To.Add(recipient);
                }
                client.Send(msg);
                return true;
            }
            catch (Exception ex)
            {
                File.AppendAllText(logFile, "Error Sending Email. Exception: " + ex.ToString() + "\r\n");
                return false;
            }
        }
    }
}
