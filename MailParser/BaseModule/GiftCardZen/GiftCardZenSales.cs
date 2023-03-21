using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BaseModule
{
    public class GiftCardZenSales
    {
        public string order_ID;
        public string date_of_purchase;
        public string total;

        public List<GiftCardZenDigiCard> lst_digi_cards;

        public GiftCardZenSales()
        {
            order_ID = string.Empty;
            date_of_purchase = string.Empty;
            total = string.Empty;
            lst_digi_cards = new List<GiftCardZenDigiCard>();
        }
    }
}
