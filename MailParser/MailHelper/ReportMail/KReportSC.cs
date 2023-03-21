using MailParser;
using Logger;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailHelper
{
    public class KReportSC : KReportBase
    {
        public KReportSC() : base()
        {
            m_order_status = ConstEnv.REPORT_ORDER_STATUS_SHIPPED;
            m_sc_ship_date = DateTime.MinValue;
            m_sc_post_type = "";
            m_sc_tracking = "";
            m_sc_expected_deliver_date = DateTime.MinValue;
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
                        new JProperty("m_retailer", m_retailer),
                        new JProperty("m_order_id", m_order_id),
                        new JProperty("m_product_items", JToken.FromObject(m_product_items)),
                        new JProperty("m_sc_ship_date", m_sc_ship_date),
                        new JProperty("m_sc_post_type", m_sc_post_type),
                        new JProperty("m_sc_tracking", m_sc_tracking),
                        new JProperty("m_sc_expected_deliver_date", m_sc_expected_deliver_date),
                        new JProperty("status", m_order_status)
            );

            return JsonConvert.SerializeObject(jsonObject, Newtonsoft.Json.Formatting.Indented);
        }
        public override string make_canceled_json_text()
        {
            string mail_address = Program.g_user.get_mailaddress_by_account_id(m_mail_account_id);


            List<ZProduct> canceld_items = new List<ZProduct>();
            foreach (ZProduct product in m_product_items)
            {
                if (m_order_status == ConstEnv.REPORT_ORDER_STATUS_CANCELED || product.status == ConstEnv.REPORT_ORDER_STATUS_CANCELED)
                    canceld_items.Add(product);
            }
            if (canceld_items.Count == 0)
                return "";

            JObject jsonObject =
                new JObject(
                        new JProperty("mail_address", mail_address),
                        new JProperty("mail_time", m_mail_sent_date.ToString("yyyy-MM-dd HH:mm:ss")),
                        new JProperty("m_mail_id", m_mail_id),
                        new JProperty("m_report_id", m_report_id),
                        new JProperty("m_mail_type", m_mail_type.ToString()),
                        new JProperty("m_retailer", m_retailer),
                        new JProperty("m_order_id", m_order_id),
                        new JProperty("canceld_items", JToken.FromObject(canceld_items))
            );

            return JsonConvert.SerializeObject(jsonObject, Newtonsoft.Json.Formatting.Indented);
        }
        public override void load_from_json(string json_text)
        {
            IDictionary<string, JToken> dict = Newtonsoft.Json.Linq.JObject.Parse(json_text);

            string mail_address = (string)dict["mail_address"];
            m_mail_account_id = Program.g_user.get_account_id(mail_address);

            m_mail_sent_date = DateTime.Parse((string)dict["mail_time"]);
            m_mail_id = int.Parse((string)dict["m_mail_id"]);
            m_report_id = int.Parse((string)dict["m_report_id"]);

            string mail_type = (string)dict["m_mail_type"];
            m_mail_type = (MailType)Enum.Parse(typeof(MailType), mail_type, true);

            m_retailer = (string)dict["m_retailer"];
            m_order_id = (string)dict["m_order_id"];
            m_sc_ship_date = DateTime.Parse((string)dict["m_sc_ship_date"]);
            m_sc_post_type = (string)dict["m_sc_post_type"];
            m_sc_tracking = (string)dict["m_sc_tracking"];
            m_sc_expected_deliver_date = DateTime.Parse((string)dict["m_sc_expected_deliver_date"]);
            m_order_status = (string)dict["status"];

            var product_items = dict["m_product_items"].ToList();
            foreach (IDictionary<string, JToken> product in product_items)
            {
                ZProduct item = new ZProduct()
                {
                    title = (string)product["title"],
                    sku = (string)product["sku"],
                    qty = int.Parse((string)product["qty"]),
                    price = float.Parse((string)product["price"])
                };
                if (product.Keys.Contains("status"))
                    item.status = (string)product["status"];
                m_product_items.Add(item);
            }
        }
        public void set_tracking(string tracking)
        {
            tracking = tracking.Trim();
            if (tracking == "")
                return;
            if (m_sc_tracking == "")
            {
                m_sc_tracking = tracking;
                return;
            }

            string[] trakings = tracking.Split(',');
            foreach (string s in trakings)
            {
                if (m_sc_tracking.IndexOf(s) == -1)
                    m_sc_tracking += "," + s;
            }
        }
        #region Process by DB Data
        public override int insert_report_to_db(int mail_id)
        {
            int card_id = -1;

            try
            {
                card_id = base.insert_report_to_db(mail_id);
                if (card_id == -1)
                {
                    MyLogger.Error($"Failed to insert card to main table.");
                    return card_id;
                }

                Program.g_db.insert_sc_report_to_db(card_id, this);
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                return -1;
            }
            return card_id;
        }
        public override void make_report_from_db(int card_id)
        {
            base.make_report_from_db(card_id);
            Program.g_db.get_sc_report_info_from_db(this, card_id);
        }
        #endregion Process by DB Data
    }
}
