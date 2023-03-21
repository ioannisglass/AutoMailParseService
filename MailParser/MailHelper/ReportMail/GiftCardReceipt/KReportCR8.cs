using MailParser;
using Logger;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailHelper
{
    public class KReportCR8 : KReportCR
    {
        public KReportCR8(int mail_seq_num) : base()
        {
            m_mail_type = MailType.CR_8;
        }
        public KReportCR8(MailType mail_type) : base()
        {
            if (mail_type != MailType.CR_8)
                throw new Exception($"Invalid mail type for CR_8 : {mail_type.ToString()}");
            m_mail_type = MailType.CR_8;
        }
        public override string make_json_text()
        {
            string mail_address = Program.g_user.get_mailaddress_by_account_id(m_mail_account_id);
            JObject jsonObject =
                new JObject(
                        new JProperty("mail_address", mail_address),
                        new JProperty("mail_time", m_mail_sent_date.ToString("yyyy-MM-dd HH:mm:ss")),
                        new JProperty("m_mail_id", m_mail_id),
                        new JProperty("m_report_id", m_report_id),
                        new JProperty("m_mail_type", m_mail_type.ToString()),
                        new JProperty("m_order_id", m_order_id),
                        new JProperty("m_retailer", m_retailer),
                        new JProperty("m_total", m_total),
                        new JProperty("m_card_details_v2", JToken.FromObject(m_giftcard_details_v2))
            );

            return JsonConvert.SerializeObject(jsonObject, Newtonsoft.Json.Formatting.Indented);
        }
    }
}
