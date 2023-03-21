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
    public class KStatCC
    {
        [JsonProperty] public int[] m_mail_ids;
        [JsonProperty] public KReportBase.MailType m_vendor_type;
        [JsonProperty] public string m_cc_order_id;
        [JsonProperty] public float m_cc_total;
        [JsonProperty] public string m_cc_receiver;
        [JsonProperty] public string m_cc_retailer;
        [JsonProperty] public DateTime m_cc_date;
        [JsonProperty] public List<ZProduct> m_cc_items;
        [JsonProperty] public float m_cc_tax;
        [JsonProperty] public List<ZPaymentCard> m_cc_purchased_card_list;
        [JsonProperty] public string m_cc_status;
        public KStatCC()
        {
            m_mail_ids = null;
            m_vendor_type = KReportBase.MailType.Mail_Unknown;
            m_cc_order_id = "";
            m_cc_total = 0;
            m_cc_receiver = "";
            m_cc_retailer = "";
            m_cc_date = DateTime.MinValue;
            m_cc_items = new List<ZProduct>();
            m_cc_tax = 0;
            m_cc_purchased_card_list = new List<ZPaymentCard>();
            m_cc_status = ConstEnv.REPORT_ORDER_STATUS_CANCELED;
        }
        public static KStatCC LoadFromText(string josn_text)
        {
            KStatCC me = new KStatCC();
            me = JsonConvert.DeserializeObject<KStatCC>(josn_text);
            return me;
        }
    }
}
