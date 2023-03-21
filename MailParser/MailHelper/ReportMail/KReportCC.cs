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
    public class KReportCC : KReportBase
    {
        public KReportCC() : base()
        {
            m_order_status = ConstEnv.REPORT_ORDER_STATUS_CANCELED;
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
                        new JProperty("m_total", m_total),
                        new JProperty("m_tax", m_tax),
                        new JProperty("m_product_items", JToken.FromObject(m_product_items)),
                        new JProperty("m_payment_card_list", JToken.FromObject(m_payment_card_list))
            );

            return JsonConvert.SerializeObject(jsonObject, Newtonsoft.Json.Formatting.Indented);
        }
        public override string make_canceled_json_text()
        {
            return make_json_text();
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
            m_total = float.Parse((string)dict["m_total"]);
            m_tax = float.Parse((string)dict["m_tax"]);

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

            var payment_cards = dict["m_payment_card_list"].ToList();
            foreach (IDictionary<string, JToken> payment_card in payment_cards)
            {
                ZPaymentCard item = new ZPaymentCard()
                {
                    payment_type = (string)payment_card["payment_type"],
                    last_4_digit = (string)payment_card["last_4_digit"],
                    price = float.Parse((string)payment_card["price"])
                };
                m_payment_card_list.Add(item);
            }
        }

        public override bool add_payment_card_info(ZPaymentCard c)
        {
            if (c.price < 0)
                c.price = -1 * c.price;
            if (!m_payment_card_list.Contains(c))
            {
                m_payment_card_list.Add(c);
                return true;
            }
            return false;
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

                Program.g_db.insert_cc_report_to_db(card_id, this);
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
            Program.g_db.get_cc_report_info_from_db(this, card_id);
        }
        #endregion Process by DB Data
    }
}
