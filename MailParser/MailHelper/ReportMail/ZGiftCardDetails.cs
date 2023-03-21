using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailHelper
{
    public class ZGiftCardDetails
    {
        public string m_retailer;
        public float m_value;
        public float m_cost;
        public string m_gift_card;
        public string m_pin;

        public ZGiftCardDetails()
        {
            m_retailer = "";
            m_value = 0;
            m_cost = 0;
            m_gift_card = "";
            m_pin = "";
        }
        public ZGiftCardDetails(string retailer, float value, float cost, string gift_card, string pin)
        {
            m_retailer = retailer;
            m_value = value;
            m_cost = cost;
            m_gift_card = gift_card;
            m_pin = pin;
        }
    }
    public class ZGiftCardDetails_V1
    {
        public string m_retailer;
        public float m_value;
        public float m_cost;

        public ZGiftCardDetails_V1()
        {
            m_retailer = "";
            m_value = 0;
            m_cost = 0;
        }
        public ZGiftCardDetails_V1(string retailer, float value, float cost)
        {
            m_retailer = retailer;
            m_value = value;
            m_cost = cost;
        }
    }
    public class ZGiftCardDetails_V2
    {
        public string m_gift_card;
        public string m_pin;

        public ZGiftCardDetails_V2()
        {
            m_gift_card = "";
            m_pin = "";
        }
        public ZGiftCardDetails_V2(string gift_card, string pin)
        {
            m_gift_card = gift_card;
            m_pin = pin;
        }
    }
}
