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
    public class KStatSC
    {
        [JsonProperty] public int[] m_mail_ids;
        [JsonProperty] public KReportBase.MailType m_vendor_type;
        [JsonProperty] public string m_sc_order_id;
        [JsonProperty] public float m_total;
        [JsonProperty] public string m_sc_retailer;
        [JsonProperty] public DateTime m_sc_ship_date;
        [JsonProperty] public List<ZProduct> m_sc_items;
        [JsonProperty] public string m_sc_post_type;
        [JsonProperty] public string m_sc_tracking;
        [JsonProperty] public DateTime m_sc_expected_deliver_date;
        public KStatSC()
        {
            m_mail_ids = null;
            m_vendor_type = KReportBase.MailType.Mail_Unknown;
            m_sc_order_id = "";
            m_total = 0;
            m_sc_retailer = "";
            m_sc_ship_date = DateTime.MinValue;
            m_sc_items = new List<ZProduct>();
            m_sc_post_type = "";
            m_sc_tracking = "";
            m_sc_expected_deliver_date = DateTime.MinValue;
        }
        public static KStatSC LoadFromText(string josn_text)
        {
            KStatSC me = new KStatSC();
            me = JsonConvert.DeserializeObject<KStatSC>(josn_text);
            return me;
        }
    }
}
