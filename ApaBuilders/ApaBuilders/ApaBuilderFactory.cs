using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ApaBuilders
{
    class ApaBuilderFactory
    {
        static public IApaBuilder CreateBuilder(CodeFinder cf)
        {
            IApaBuilder Customer = null;
            switch (cf.Customer)
            {
                case "MJR":
                    Customer = new MJR(cf);
                    break;

                case "MND":
                    Customer = new MND(cf);
                    break;

                case "AMF":
                    Customer = new AMF(cf);
                    break;
            }
            return Customer;
        }
}
}
