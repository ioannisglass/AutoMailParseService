using MailParser;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailHelper
{
    public class ZScrapParam
    {
        public readonly KReportBase report;
        public readonly string link;
        public string mailto;
        public string retailer;
        public float value;

        public int status;

        public ZScrapParam(KReportBase _report, string _link, string _mailto = "", string _retailer = "", float _value = 0)
        {
            report = _report;
            link = _link;
            mailto = _mailto;
            retailer = _retailer;
            value = _value;
            status = ConstEnv.SCRAP_READY;
        }
    }
}
