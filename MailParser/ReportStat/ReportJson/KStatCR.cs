using MailParser;
using MailHelper;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportStat
{
    public class KStatCR
    {
        [JsonProperty] public int[] m_mail_ids;
        [JsonProperty] public KReportBase.MailType m_vendor_type;
        [JsonProperty] public string m_order_id;
        [JsonProperty] public float m_total;
        [JsonProperty] public DateTime m_purchase_date;
        [JsonProperty] public ZPaymentCard m_payment_card_info;

        [JsonProperty] public List<ZGiftCardDetails> m_card_details; // for vendor 1 and vendor 4
        [JsonProperty] public List<ZGiftCardDetails_V1> m_card_details_v1; // for vendor 2, vendor 3, vendor 7
        [JsonProperty] public List<ZGiftCardDetails_V2> m_card_details_v2; // for vendor 2, vendor 3, vendor 7

        [JsonProperty] public float m_discount;

        [JsonProperty] public List<float> m_instant_cashback;
        public KStatCR()
        {
            m_mail_ids = null;
            m_vendor_type = KReportBase.MailType.Mail_Unknown;
            m_order_id = "";
            m_total = 0;
            m_payment_card_info = new ZPaymentCard();
            m_card_details = new List<ZGiftCardDetails>();
            m_card_details_v1 = new List<ZGiftCardDetails_V1>();
            m_card_details_v2 = new List<ZGiftCardDetails_V2>();
            m_discount = 0;
            m_instant_cashback = new List<float>();
        }
        public static KStatCR LoadFromText(string josn_text)
        {
            KStatCR me = new KStatCR();
            me = JsonConvert.DeserializeObject<KStatCR>(josn_text);
            return me;
        }
    }
}
