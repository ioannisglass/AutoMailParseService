using BaseModule;
using MailHelper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailParser
{
    public class WorkerParam
    {
        public string status = "";

        public KReportBase card;
        public List<string> link_urls;
    }
}
