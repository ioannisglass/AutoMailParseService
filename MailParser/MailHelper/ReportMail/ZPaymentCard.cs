using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailHelper
{
    public class ZPaymentCard
    {
        public string payment_type;
        public string last_4_digit;
        public float price;

        public ZPaymentCard()
        {
            payment_type = "";
            last_4_digit = "";
            price = 0;
        }
        public ZPaymentCard(string _payment_type, string _last_4_digit, float _price)
        {
            payment_type = _payment_type;
            last_4_digit = _last_4_digit;
            price = _price;
        }
    }
}
