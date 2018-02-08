using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ApaBuilders
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Count() == 3)
            {
                CodeFinder cf = new CodeFinder(args);
                IApaBuilder builder = ApaBuilderFactory.CreateBuilder(cf);
                builder.DoWork();
            }
            else
            {
                File.AppendAllText(Utils.logFile, DateTime.Now.ToString("MM/dd/yyyy HH:mm ") + " :Attempted to run an ApaBuilder, but missing directory references.\r\n");
            }
            Environment.Exit(0);
        }
    }
}
