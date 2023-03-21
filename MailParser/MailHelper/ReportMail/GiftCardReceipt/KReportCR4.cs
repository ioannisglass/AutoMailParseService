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
    public class KReportCR4 : KReportCR
    {
        public KReportCR4(int mail_seq_num) : base()
        {
            if (mail_seq_num == 0)
                m_mail_type = MailType.CR_4;
            else if (mail_seq_num == 1)
                m_mail_type = MailType.CR_4_1;
            else if (mail_seq_num == 2)
                m_mail_type = MailType.CR_4_2;
            else if (mail_seq_num == 3)
                m_mail_type = MailType.CR_4_3;
            else if (mail_seq_num == 4)
                m_mail_type = MailType.CR_4_4;
            else
                throw new Exception($"Invalid CR4 mail seq num. {mail_seq_num}");
        }
        public KReportCR4(MailType mail_type) : base()
        {
            m_mail_type = mail_type;
        }
        public override string make_json_text()
        {
            JObject jsonObject = null;
            string mail_address = Program.g_user.get_mailaddress_by_account_id(m_mail_account_id);
            if (m_mail_type == MailType.CR_4)
            {
                int[] child_ids = Program.g_db.get_child_report_ids(m_report_id);
                jsonObject =
                    new JObject(
                            new JProperty("mail_address", mail_address),
                            new JProperty("mail_time", m_mail_sent_date.ToString("yyyy-MM-dd HH:mm:ss")),
                            new JProperty("m_mail_ids", child_ids),
                            new JProperty("m_report_id", m_report_id),
                            new JProperty("m_mail_type", m_mail_type.ToString()),
                            new JProperty("m_order_id", m_order_id),
                            new JProperty("m_purchase_date", m_purchase_date),
                            new JProperty("m_total", m_total),
                            new JProperty("m_cr_payment_type", JToken.FromObject(m_cr_payment_type)),
                            new JProperty("m_cr_payment_id", JToken.FromObject(m_cr_payment_id)),
                            new JProperty("m_giftcard_details", JToken.FromObject(m_giftcard_details)),
                            new JProperty("m_instant_cashback", m_instant_cashback),
                            new JProperty("status", ConstEnv.REPORT_STATUS_ALL_RECEIVED)
                );
            }
            else if (m_mail_type == MailType.CR_4_1)
            {
                jsonObject =
                    new JObject(
                            new JProperty("mail_address", mail_address),
                            new JProperty("mail_time", m_mail_sent_date.ToString("yyyy-MM-dd HH:mm:ss")),
                            new JProperty("m_mail_id", m_mail_id),
                            new JProperty("m_report_id", m_report_id),
                            new JProperty("m_mail_type", m_mail_type.ToString()),
                            new JProperty("m_order_id", m_order_id),
                            new JProperty("m_purchase_date", m_purchase_date),
                            new JProperty("m_total", m_total),
                            new JProperty("m_cr_payment_type", JToken.FromObject(m_cr_payment_type)),
                            new JProperty("m_cr_payment_id", JToken.FromObject(m_cr_payment_id))
                );
            }
            else if (m_mail_type == MailType.CR_4_2)
            {
                jsonObject =
                    new JObject(
                            new JProperty("mail_address", mail_address),
                            new JProperty("mail_time", m_mail_sent_date.ToString("yyyy-MM-dd HH:mm:ss")),
                            new JProperty("m_mail_id", m_mail_id),
                            new JProperty("m_report_id", m_report_id),
                            new JProperty("m_mail_type", m_mail_type.ToString()),
                            new JProperty("m_order_id", m_order_id),
                            new JProperty("m_total", m_total),
                            new JProperty("m_instant_cashback", m_instant_cashback),
                            new JProperty("m_cr_payment_type", JToken.FromObject(m_cr_payment_type)),
                            new JProperty("m_cr_payment_id", JToken.FromObject(m_cr_payment_id))
                );
            }
            else if (m_mail_type == MailType.CR_4_3)
            {
                jsonObject =
                    new JObject(
                            new JProperty("mail_address", mail_address),
                            new JProperty("mail_time", m_mail_sent_date.ToString("yyyy-MM-dd HH:mm:ss")),
                            new JProperty("m_mail_id", m_mail_id),
                            new JProperty("m_report_id", m_report_id),
                            new JProperty("m_mail_type", m_mail_type.ToString()),
                            new JProperty("m_order_id", m_order_id),
                            new JProperty("m_instant_cashback", m_instant_cashback),
                            new JProperty("m_giftcard_details", JToken.FromObject(m_giftcard_details))
                );
            }
            else if (m_mail_type == MailType.CR_4_4)
            {
                jsonObject =
                    new JObject(
                            new JProperty("mail_address", mail_address),
                            new JProperty("mail_time", m_mail_sent_date.ToString("yyyy-MM-dd HH:mm:ss")),
                            new JProperty("m_mail_id", m_mail_id),
                            new JProperty("m_report_id", m_report_id),
                            new JProperty("m_mail_type", m_mail_type.ToString()),
                            new JProperty("m_instant_cashback", m_instant_cashback),
                            new JProperty("status", ConstEnv.REPORT_STATUS_CR4_CASHBACK)
                );
            }

            if (jsonObject != null)
                return JsonConvert.SerializeObject(jsonObject, Newtonsoft.Json.Formatting.Indented);

            return "";
        }
        public override void load_from_json(string json_text)
        {
            IDictionary<string, JToken> dict = Newtonsoft.Json.Linq.JObject.Parse(json_text);

            string mail_address = (string)dict["mail_address"];
            m_mail_account_id = Program.g_user.get_account_id(mail_address);

            m_mail_sent_date = DateTime.Parse((string)dict["mail_time"]);
            m_report_id = int.Parse((string)dict["m_report_id"]);

            string mail_type = (string)dict["m_mail_type"];
            m_mail_type = (MailType)Enum.Parse(typeof(MailType), mail_type, true);

            m_order_id = (string)dict["m_order_id"];

            if (m_mail_type == MailType.CR_4)
            {
                m_total = float.Parse((string)dict["m_total"]);
                m_purchase_date = DateTime.Parse((string)dict["m_purchase_date"]);
                m_cr_payment_type = (string)dict["m_cr_payment_type"];
                m_cr_payment_id = (string)dict["m_cr_payment_id"];
                var instant_cashbacks = dict["m_instant_cashback"].ToList();
                m_instant_cashback = new List<float>();
                foreach (float f in instant_cashbacks)
                    m_instant_cashback.Add(f);

                m_order_status = (string)dict["status"];

                var giftcard_details = dict["m_giftcard_details"].ToList();
                foreach (IDictionary<string, JToken> giftcard_detail in giftcard_details)
                {
                    ZGiftCardDetails item = new ZGiftCardDetails()
                    {
                        m_retailer = (string)giftcard_detail["m_retailer"],
                        m_gift_card = (string)giftcard_detail["m_gift_card"],
                        m_pin = (string)giftcard_detail["m_pin"],
                        m_value = float.Parse((string)giftcard_detail["m_value"]),
                        m_cost = float.Parse((string)giftcard_detail["m_cost"])
                    };
                    m_giftcard_details.Add(item);
                }
            }
            if (m_mail_type == MailType.CR_4_1)
            {
                m_mail_id = int.Parse((string)dict["m_mail_id"]);
                m_total = float.Parse((string)dict["m_total"]);
                m_purchase_date = DateTime.Parse((string)dict["m_purchase_date"]);
                m_cr_payment_type = (string)dict["m_cr_payment_type"];
                m_cr_payment_id = (string)dict["m_cr_payment_id"];
            }
            if (m_mail_type == MailType.CR_4_2)
            {
                m_mail_id = int.Parse((string)dict["m_mail_id"]);
                m_total = float.Parse((string)dict["m_total"]);
                m_cr_payment_type = (string)dict["m_cr_payment_type"];
                m_cr_payment_id = (string)dict["m_cr_payment_id"];
                var instant_cashbacks = dict["m_instant_cashback"].ToList();
                m_instant_cashback = new List<float>();
                foreach (float f in instant_cashbacks)
                    m_instant_cashback.Add(f);
            }
            if (m_mail_type == MailType.CR_4_3)
            {
                m_mail_id = int.Parse((string)dict["m_mail_id"]);
                var instant_cashbacks = dict["m_instant_cashback"].ToList();
                m_instant_cashback = new List<float>();
                foreach (float f in instant_cashbacks)
                    m_instant_cashback.Add(f);

                var giftcard_details = dict["m_giftcard_details"].ToList();
                foreach (IDictionary<string, JToken> giftcard_detail in giftcard_details)
                {
                    ZGiftCardDetails item = new ZGiftCardDetails()
                    {
                        m_retailer = (string)giftcard_detail["m_retailer"],
                        m_gift_card = (string)giftcard_detail["m_gift_card"],
                        m_pin = (string)giftcard_detail["m_pin"],
                        m_value = float.Parse((string)giftcard_detail["m_value"]),
                        m_cost = float.Parse((string)giftcard_detail["m_cost"])
                    };
                    m_giftcard_details.Add(item);
                }
            }
            if (m_mail_type == MailType.CR_4_4)
            {
                m_mail_id = int.Parse((string)dict["m_mail_id"]);
                var instant_cashbacks = dict["m_instant_cashback"].ToList();
                m_instant_cashback = new List<float>();
                foreach (float f in instant_cashbacks)
                    m_instant_cashback.Add(f);
            }
        }
        public override KReportBase merge_multi_reports()
        {
            try
            {
                if (m_mail_type != MailType.CR_4_3)
                    return null;

                // Find child reports.

                int seq_1_id = Program.g_db.find_matched_cr4_1_report_by_order(m_order_id, m_mail_account_id);
                if (seq_1_id == -1)
                    return null;

                KReportCR4 seq_1_report = new KReportCR4(1);
                seq_1_report.make_report_from_db(seq_1_id);

                int seq_2_id = Program.g_db.find_matched_cr4_2_report(seq_1_report);
                if (seq_2_id == -1)
                    return null;

                KReportCR4 seq_2_report = new KReportCR4(1);
                seq_2_report.make_report_from_db(seq_2_id);

                // Create a parent card.

                KReportCR4 parent_report = new KReportCR4(0);

                parent_report.set_order_id(m_order_id);
                parent_report.m_mail_account_id = seq_1_report.m_mail_account_id;
                parent_report.m_purchase_date = seq_1_report.m_purchase_date;
                parent_report.set_total(seq_1_report.m_total);
                parent_report.m_cr_payment_type = seq_1_report.m_cr_payment_type;
                parent_report.m_cr_payment_id = seq_1_report.m_cr_payment_id;

                parent_report.m_instant_cashback = m_instant_cashback;
                parent_report.m_giftcard_details = seq_1_report.m_giftcard_details;
                parent_report.m_mail_sent_date = m_mail_sent_date;

                int parent_report_id = parent_report.insert_report_to_db(-1);
                if (parent_report_id == -1)
                    return null;

                // update parent id of all children.
                Program.g_db.update_report_parent_id(seq_1_id, parent_report_id);
                Program.g_db.update_report_parent_id(seq_2_id, parent_report_id);
                Program.g_db.update_report_parent_id(m_report_id, parent_report_id);

                return parent_report;

            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                return null;
            }
        }
    }
}
