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
    public class KReportCR : KReportBase
    {
        public KReportCR() : base()
        {
            m_purchase_date = DateTime.MinValue;
            m_cr_payment_type = "";
            m_cr_payment_id = "";
            m_giftcard_details = new List<ZGiftCardDetails>();
            m_giftcard_details_v1 = new List<ZGiftCardDetails_V1>();
            m_giftcard_details_v2 = new List<ZGiftCardDetails_V2>();
            m_discount = 0;
            m_instant_cashback = new List<float>();
        }

        public override void report()
        {
            try
            {
                base.report();

                KReportBase parent_report = merge_multi_reports();
                if (parent_report == null)
                    return;

                parent_report.send_to_crm();
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
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
                        new JProperty("m_purchase_date", m_purchase_date),
                        new JProperty("m_total", m_total),
                        new JProperty("m_cr_payment_type", JToken.FromObject(m_cr_payment_type)),
                        new JProperty("m_cr_payment_id", JToken.FromObject(m_cr_payment_id)),
                        new JProperty("m_card_details", JToken.FromObject(m_giftcard_details)),
                        new JProperty("m_card_details_v1", JToken.FromObject(m_giftcard_details_v1)),
                        new JProperty("m_card_details_v2", JToken.FromObject(m_giftcard_details_v2)),
                        new JProperty("m_discount", m_discount),
                        new JProperty("m_instant_cashback", m_instant_cashback)
            );

            return JsonConvert.SerializeObject(jsonObject, Newtonsoft.Json.Formatting.Indented);
        }
        public override void load_from_json(string json_text)
        {
            IDictionary<string, JToken> dict = Newtonsoft.Json.Linq.JObject.Parse(json_text);

            string mail_address = (string)dict["mail_address"];
            m_mail_account_id = Program.g_user.get_account_id(mail_address);

            m_mail_sent_date = DateTime.Parse((string)dict["mail_time"]);
            m_report_id = int.Parse((string)dict["m_report_id"]);

            m_mail_id = int.Parse((string)dict["m_mail_id"]);

            string mail_type = (string)dict["m_mail_type"];
            m_mail_type = (MailType)Enum.Parse(typeof(MailType), mail_type, true);

            m_order_id = (string)dict["m_order_id"];

            m_total = float.Parse((string)dict["m_total"]);
            m_discount = float.Parse((string)dict["m_discount"]);
            m_purchase_date = DateTime.Parse((string)dict["m_purchase_date"]);
            m_cr_payment_type = (string)dict["m_cr_payment_type"];
            m_cr_payment_id = (string)dict["m_cr_payment_id"];
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

            var giftcard_details_v1 = dict["m_giftcard_details_v1"].ToList();
            foreach (IDictionary<string, JToken> giftcard_detail in giftcard_details_v1)
            {
                ZGiftCardDetails_V1 item = new ZGiftCardDetails_V1()
                {
                    m_retailer = (string)giftcard_detail["m_retailer"],
                    m_value = float.Parse((string)giftcard_detail["m_value"]),
                    m_cost = float.Parse((string)giftcard_detail["m_cost"])
                };
                m_giftcard_details_v1.Add(item);
            }

            var giftcard_details_v2 = dict["m_giftcard_details_v2"].ToList();
            foreach (IDictionary<string, JToken> giftcard_detail in giftcard_details_v2)
            {
                ZGiftCardDetails_V2 item = new ZGiftCardDetails_V2()
                {
                    m_gift_card = (string)giftcard_detail["m_gift_card"],
                    m_pin = (string)giftcard_detail["m_pin"],
                };
                m_giftcard_details_v2.Add(item);
            }
        }
        public override void set_total(float total)
        {
            if (m_total == 0)
            {
                m_total = total;
            }
            else if (m_total != total)
            {
                MyLogger.Info($"*** TOTAL CHANGED *** {m_total} -> {total}");
                m_total = total;
            }
            if (m_total == 0)
                MyLogger.Info($"*** ZERO TOTAL ***");
        }
        public void add_giftcard_details(ZGiftCardDetails info)
        {
            m_giftcard_details.Add(info);
        }
        public void add_giftcard_details(string retailer, float value, float cost, string gift_card, string pin)
        {
            add_giftcard_details(new ZGiftCardDetails(retailer, value, cost, gift_card, pin));
        }
        public void add_giftcard_details_v1(ZGiftCardDetails_V1 info)
        {
            m_giftcard_details_v1.Add(info);
        }
        public void add_giftcard_details_v1(string retailer, float value, float cost)
        {
            add_giftcard_details_v1(new ZGiftCardDetails_V1(retailer, value, cost));
        }
        public void add_giftcard_details_v2(ZGiftCardDetails_V2 info)
        {
            m_giftcard_details_v2.Add(info);
        }
        public void add_giftcard_details_v2(string gift_card, string pin)
        {
            add_giftcard_details_v2(new ZGiftCardDetails_V2(gift_card, pin));
        }
        #region Process by DB Data
        public override int insert_report_to_db(int mail_id)
        {
            int report_id = -1;

            try
            {
                report_id = base.insert_report_to_db(mail_id);
                if (report_id == -1)
                {
                    MyLogger.Error($"Failed to insert card to main table.");
                    return report_id;
                }

                Program.g_db.insert_cr_report_to_db(report_id, this);
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                return -1;
            }
            return report_id;
        }
        public override void make_report_from_db(int report_id)
        {
            base.make_report_from_db(report_id);
            Program.g_db.get_cr_report_from_db(this, report_id);
        }
        #endregion Process by DB Data
    }
}
