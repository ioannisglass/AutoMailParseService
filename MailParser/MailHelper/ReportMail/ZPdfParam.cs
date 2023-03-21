using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailHelper
{
    public class ZPdfParam
    {
        public readonly int mail_id;
        public readonly string order_id;
        public readonly string retailer;

        public ZPdfParam(int _mail_id, string _order_id, string _retailer)
        {
            mail_id = _mail_id;
            order_id = _order_id;
            retailer = _retailer;
        }
    }
}
