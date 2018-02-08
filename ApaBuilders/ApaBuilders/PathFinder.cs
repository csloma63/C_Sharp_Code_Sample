using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ApaBuilders
{
    public abstract class PathFinder
    {
        private string input = string.Empty;
        private string output = string.Empty;
        private string trigger = string.Empty;
        public PathFinder(string[] args)
        {
            input = args[0].ToString();
            output = args[1].ToString();
            trigger = args[2].ToString();
        }
        public string Input { get { return input; } set { Input = value; } }
        public string Output { get { return output; } set { Output = value; } }
        public string Trigger { get { return trigger; } set { Trigger = value; } }
    }
}
