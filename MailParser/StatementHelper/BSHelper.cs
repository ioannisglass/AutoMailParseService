using Logger;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StatementHelper
{
    public class BSHelper
    {
        private List<BSParser> bs_parsers = new List<BSParser>();

        public BSHelper()
        {
            bs_parsers.Add(new BSParserWF());
            bs_parsers.Add(new BSParserChase());
            bs_parsers.Add(new BSParserCOB());
            bs_parsers.Add(new BSParserCustomers());
        }
        public bool parse_pdf(string pdf_file)
        {
            int i = 0;
            for (i = 0; i < bs_parsers.Count; i++)
            {
                if (bs_parsers[i].parse(pdf_file))
                    break;
            }
            if (i >= bs_parsers.Count)
            {
                MyLogger.Info($"Invalid Bank Statement File : {pdf_file}");
                return false;
            }

            return true;
        }
    }
}
