using MailHelper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportStat
{
    public class ZSalesTaxPayData
    {
        public readonly DateTime purchase_time;
        public readonly string order_id;
        public readonly string retailer;
        public readonly float total;
        public readonly float tax;
        public readonly List<ZPaymentCard> payments;

        public ZSalesTaxPayData(DateTime _purchase_time, string _order_id, string _retailer, float _total, float _tax, List<ZPaymentCard> _payments)
        {
            purchase_time = _purchase_time;
            order_id = _order_id;
            retailer = _retailer;
            total = _total;
            tax = _tax;
            payments = _payments;
        }
    }
}
