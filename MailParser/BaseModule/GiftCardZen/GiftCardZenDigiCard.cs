using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BaseModule
{
    public class GiftCardZenDigiCard
    {
        public string retailer;
        public string value;
        public string cost;
        public string discount;
        public string card_number;
        public string pin;

        public string status;

        public GiftCardZenDigiCard()
        {
            retailer = string.Empty;
            value = string.Empty;
            cost = string.Empty;
            discount = string.Empty;
            card_number = string.Empty;
            pin = string.Empty;
            status = string.Empty;
        }
    }
}
