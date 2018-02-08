using System;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Runtime.Remoting;

namespace ApaBuilders
{
    public class CodeFinder : PathFinder
    {
        enum ApaCustomers { AMF, MJR, MND};

        protected readonly string fileName;
        protected readonly string customer;
        protected readonly string switchCode;
        protected readonly string metadataCode;
        protected readonly string outputFolder;

        public string FileName { get => fileName; set { FileName = value; } }
        public string Customer { get => customer; set { Customer = value; } }
        public string SwitchCode { get => switchCode; set { SwitchCode = value; } }
        public string MetadataCode { get => metadataCode; set { MetadataCode = value; } }
        public string OutputFolder { get => outputFolder; set { OutputFolder = value; } }

        public readonly static string undefinedArgs = ConfigurationManager.AppSettings["UNDEFINED_EMAIL"];

        public CodeFinder(string[] args):base(args)
        {
            GetCodes(out customer, out metadataCode, out switchCode, out fileName);
            ValidateCustomer();
            CreateOutputFolder(out outputFolder);
        }
        private void ValidateCustomer()
        {
            if (!Enum.IsDefined(typeof(ApaCustomers), Customer))
            {
                File.AppendAllText(Utils.logFile, DateTime.Now.ToString("MM/dd/yyyy HH:mm ") + " :File path " + FileName + " does not contain customer name. \r\n");
                Utils.SendEmail(undefinedArgs, "args", "File path does not contain customer name. \r\n");
                Environment.Exit(0);
            }
        }// checks that the customer is in the enum else sends an email / logs issue.
        private void GetCodes(out string customer, out string metadataCode, out string switchCode, out string fileName)
        {
            DirectoryInfo inputDirectory = new DirectoryInfo(Trigger);
            FileInfo[] files = inputDirectory.GetFiles("*.xml");
            string name = files[0].ToString();
            string file = Path.GetFileName(name);
            switchCode = file.Substring(0, 7);
            int i = file.IndexOf("_x");
            customer = file.Substring(7, i - 7);
            metadataCode = file.Substring(7, i);
            fileName = file.Substring(i + 8, file.Length - (i + 8));
        }
        private void CreateOutputFolder(out string outputFolder)
        {
            outputFolder = Output + metadataCode + "_ApaOut\\";//folder in output directory with process id
            try
            {
                if (!Output.Equals("C:\\"))//Need to make sure the output folder is empty before starting a new process
                {
                    DirectoryInfo di = new DirectoryInfo(Output);
                    foreach (FileInfo file in di.GetFiles()) { file.Delete(); }
                    foreach (DirectoryInfo info in di.GetDirectories()) { info.Delete(true); }
                    if (!Directory.Exists(outputFolder)) { Directory.CreateDirectory(outputFolder); }  // if it doesn't exist, create
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText(Utils.logFile, DateTime.Now.ToString("MM/dd/yyyy HH:mm ") + " :Failed to parse processId from input filename: " + Trigger + ": " + ex.Message + " \r\n");
                Utils.SendEmail(undefinedArgs, "args", "Failed to parse processId from input filename: " + Trigger + ": " + ex.Message + " \r\n");
                Environment.Exit(0);
            }
        }

    }
}
