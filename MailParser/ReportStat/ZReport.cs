using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportStat
{
    public class ZReport
    {
        public enum ReportVer
        {
            // 4_retailers.
            v1_for_4_retailers = 0
        }

        public readonly ReportVer ver;
        public string retailer;
        public string order;
        public string status;
        public DateTime time;
        public ZReport(ReportVer _ver)
        {
            ver = _ver;
            retailer = "";
            order = "";
            status = "";
            time = DateTime.MinValue;
        }
        public ZReport(ReportVer _ver, string _retailer, string _order, string _status, DateTime _time)
        {
            ver = _ver;
            retailer = _retailer;
            order = _order;
            status = _status;
            time = _time;
        }
    }
}
