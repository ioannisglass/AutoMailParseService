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
    public class KStatOP
    {
        [JsonProperty] public int[] m_mail_ids;
        [JsonProperty] public KReportBase.MailType m_vendor_type;
        [JsonProperty] public string m_op_order_id;
        [JsonProperty] public float m_op_order_total;
        [JsonProperty] public string m_op_receiver;
        [JsonProperty] public string m_op_retailer;
        [JsonProperty] public DateTime m_op_purchase_date;
        [JsonProperty] public List<ZProduct> m_op_items;
        [JsonProperty] public float m_op_tax;
        [JsonProperty] public List<ZPaymentCard> m_op_purchased_card_list;
        public KStatOP()
        {
            m_mail_ids = null;
            m_vendor_type = KReportBase.MailType.Mail_Unknown;
            m_op_order_id = "";
            m_op_order_total = 0;
            m_op_receiver = "";
            m_op_retailer = "";
            m_op_purchase_date = DateTime.MinValue;
            m_op_items = new List<ZProduct>();
            m_op_tax = 0;
            m_op_purchased_card_list = new List<ZPaymentCard>();
        }
        public static KStatOP LoadFromText(string josn_text)
        {
            KStatOP me = new KStatOP();
            me = JsonConvert.DeserializeObject<KStatOP>(josn_text);
            return me;
        }
    }
}
