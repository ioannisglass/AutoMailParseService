using MailParser;
using Logger;
using MimeKit;
using Newtonsoft.Json;
using ReportStat;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WebAuto;

namespace MailHelper
{
    [JsonObject(MemberSerialization.OptIn)]
    public class KReportBase
    {
        public enum MailType
        {
            Mail_Unknown,
            CR_START,
            CR_1,
            CR_1_1,
            CR_1_2,
            CR_1_3,
            CR_2,
            CR_2_1,
            CR_2_2,
            CR_3,
            CR_3_1,
            CR_3_2,
            CR_4,
            CR_4_1,
            CR_4_2,
            CR_4_3,
            CR_4_4,
            CR_5,
            CR_6,
            CR_7,
            CR_7_1,
            CR_7_2,
            CR_7_3,
            CR_7_4,
            CR_7_5,
            CR_8,
            CR_END,

            OP_START,
            OP_1,
            OP_2,
            OP_3,
            OP_4,
            OP_5,
            OP_6,
            OP_7,
            OP_8,
            OP_9,
            OP_10,
            OP_11,
            OP_12,
            OP_13,
            OP_14,
            OP_15,
            OP_16,
            OP_17,
            OP_18,
            OP_19,
            OP_20,
            OP_21,
            OP_END,

            SC_START,
            SC_1,
            SC_2,
            SC_3,
            SC_4,
            SC_5,
            SC_6,
            SC_7,
            SC_8,
            SC_9,
            SC_10,
            SC_11,
            SC_12,
            SC_END,

            CC_START,
            CC_1,
            CC_2,
            CC_3,
            CC_4,
            CC_5,
            CC_6,
            CC_7,
            CC_8,
            CC_9,
            CC_10,
            CC_11,
            CC_12,
            CC_END
        }

        static public readonly string PAYMENT_TYPE_CC = "CC";

        static public readonly string POST_TYPE_FEDEX = "FEDEX";
        static public readonly string POST_TYPE_UPS = "UPS";
        static public readonly string POST_TYPE_USPS = "USPS";

        #region Report Common
        [JsonProperty] public MailType m_mail_type;
        [JsonProperty] public string m_order_id { get; protected set; }
        [JsonProperty] public float m_total { get; protected set; }
        [JsonProperty] public float m_tax;
        [JsonProperty] public string m_pdf_file;
        [JsonProperty] public List<ZProduct> m_product_items;
        [JsonProperty] public string m_receiver;
        [JsonProperty] public string m_retailer;
        [JsonProperty] public List<ZPaymentCard> m_payment_card_list;
        public DateTime m_mail_sent_date; // only valid for single mail report.
        [JsonProperty] public string m_order_status;
        #endregion Report Common


        #region Gift Card Recipt Information
        [JsonProperty] public DateTime m_purchase_date;
        [JsonProperty] public string m_cr_payment_type;
        [JsonProperty] public string m_cr_payment_id;


        [JsonProperty] public List<ZGiftCardDetails> m_giftcard_details; // for vendor 1 and vendor 4
        [JsonProperty] public List<ZGiftCardDetails_V1> m_giftcard_details_v1; // for vendor 2, vendor 3, vendor 7
        [JsonProperty] public List<ZGiftCardDetails_V2> m_giftcard_details_v2; // for vendor 2, vendor 3, vendor 7

        [JsonProperty] public float m_discount;

        [JsonProperty] public List<float> m_instant_cashback;
        #endregion Gift Card Recipt Information

        #region Order Confirmation / Purchased

        [JsonProperty] public DateTime m_op_purchase_date;
        [JsonProperty] public string m_op_ship_address;
        [JsonProperty] public string m_op_ship_address_state;

        #endregion Order Confirmation / Purchased

        #region Shipping Confirmation

        [JsonProperty] public DateTime m_sc_ship_date;
        [JsonProperty] public string m_sc_post_type;
        [JsonProperty] public string m_sc_tracking { get; protected set; }
        [JsonProperty] public DateTime m_sc_expected_deliver_date;

        #endregion Shipping Confirmation

        #region Cancellation Confirmation / Purchased


        #endregion Cancellation Confirmation / Purchased

        public KReportBase next_report;

        #region Parsing Information

        public int m_mail_account_id;
        [JsonProperty] public int m_mail_id;
        [JsonProperty] public int m_report_id;
        public List<ZScrapParam> m_scrap_params;

        #endregion Parsing Information


        public KReportBase()
        {
            m_mail_type = MailType.Mail_Unknown;

            m_order_id = "";
            m_total = 0;
            m_pdf_file = "";

            m_mail_account_id = -1;
            m_mail_id = -1;
            m_report_id = -1;
            m_scrap_params = null;
            next_report = null;
            m_mail_sent_date = DateTime.MinValue;
            m_receiver = "";
            m_retailer = "";
            m_product_items = new List<ZProduct>();
            m_tax = 0;
            m_payment_card_list = new List<ZPaymentCard>();
            m_order_status = "";
        }
        public bool is_4_retailers()
        {
            if (m_mail_type == MailType.OP_1 || m_mail_type == MailType.SC_1 || m_mail_type == MailType.CC_1)
                return true;
            if (m_mail_type == MailType.OP_2 || m_mail_type == MailType.SC_2 || m_mail_type == MailType.CC_2)
                return true;
            if (m_mail_type == MailType.OP_10 || m_mail_type == MailType.SC_10 || m_mail_type == MailType.CC_10)
                return true;
            if (m_mail_type == MailType.OP_12 || m_mail_type == MailType.SC_12 || m_mail_type == MailType.CC_12)
                return true;
            return false;
        }
        public bool is_order_report()
        {
            if (m_mail_type > MailType.OP_START && m_mail_type < MailType.OP_END)
                return true;
            if (m_mail_type > MailType.SC_START && m_mail_type < MailType.SC_END)
                return true;
            if (m_mail_type > MailType.CC_START && m_mail_type < MailType.CC_END)
                return true;
            return false;
        }
        public virtual void set_order_id(string order_id)
        {
            order_id = order_id.Trim();
            order_id = XMailHelper.html2text(order_id);
            if (order_id == "")
                return;
            if (m_order_id == "")
            {
                m_order_id = order_id;
            }
            else
            {
                string[] orders = m_order_id.Split(',');
                if (!orders.Contains(order_id))
                    m_order_id += "," + order_id;
            }
        }
        public virtual void set_total(float total)
        {
            if (total < 0)
                total *= -1;

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
        public virtual bool add_payment_card_info(ZPaymentCard c)
        {
            throw new Exception("[add_card_info] NOT IMPLEMENTED");
        }
        public bool add_payment_card_info(string payment_type, string last_4_digit, float price)
        {
            ZPaymentCard c = new ZPaymentCard(payment_type, last_4_digit, price);
            return add_payment_card_info(c);
        }
        protected bool is_order_confirmation_mail()
        {
            if (this.GetType() == typeof(KReportOP))
                return true;
            return false;
        }
        public void add_web_link(List<string> url_list)
        {
            if (m_scrap_params == null)
                m_scrap_params = new List<ZScrapParam>();

            foreach (string url in url_list)
            {
                m_scrap_params.Add(new ZScrapParam(this, url));
            }
        }
        public void add_web_link(string url)
        {
            if (m_scrap_params == null)
                m_scrap_params = new List<ZScrapParam>();
            m_scrap_params.Add(new ZScrapParam(this, url));
        }
        public void add_web_link(string url, string mailto)
        {
            if (m_scrap_params == null)
                m_scrap_params = new List<ZScrapParam>();
            m_scrap_params.Add(new ZScrapParam(this, url, mailto));
        }
        public void add_web_link(string url, string retailer, float value)
        {
            if (m_scrap_params == null)
                m_scrap_params = new List<ZScrapParam>();
            m_scrap_params.Add(new ZScrapParam(this, url, "", retailer, value));
        }
        public void set_address(string full_address)
        {
            string state_address = XMailHelper.get_address_state_name(full_address);
            if (state_address != "")
                set_address(full_address, state_address);
        }
        public void set_address(string full_address, string state_address)
        {
            if (m_op_ship_address == "")
            {
                m_op_ship_address = full_address;
            }
            else if (m_op_ship_address.ToUpper() != full_address.ToUpper())
            {
                MyLogger.Info($"*** ADDRESS is changed *** : {m_op_ship_address} -> {full_address}");
                m_op_ship_address = full_address;
            }

            if (m_op_ship_address_state == "")
            {
                m_op_ship_address_state = state_address;
            }
            else if (m_op_ship_address_state.ToUpper() != state_address.ToUpper())
            {
                MyLogger.Info($"*** STATE NAME is changed *** : {m_op_ship_address_state} -> {state_address}");
                m_op_ship_address_state = state_address;
            }
        }
        public virtual string make_json_text()
        {
            return "";
        }
        public virtual string make_canceled_json_text()
        {
            return "";
        }
        static public KReportBase generate_report_from_json(string json_text)
        {
            try
            {
                int pos = json_text.IndexOf("\"m_mail_type\":");
                if (pos == -1)
                    return null;
                string temp = json_text.Substring(pos + "\"m_mail_type\":".Length);
                pos = temp.IndexOf(",");
                if (pos == -1)
                    return null;
                temp = temp.Substring(0, pos).Trim();
                temp = temp.Replace("\"", "");
                MailType mail_type = (MailType)Enum.Parse(typeof(MailType), temp, true);

                KReportBase report = null;
                if (mail_type >= MailType.CR_1 && mail_type <= MailType.CR_1_2)
                    report = new KReportCR1(mail_type);
                else if (mail_type >= MailType.CR_2 && mail_type <= MailType.CR_2_2)
                    report = new KReportCR2(mail_type);
                else if (mail_type >= MailType.CR_3 && mail_type <= MailType.CR_3_2)
                    report = new KReportCR3(mail_type);
                else if (mail_type >= MailType.CR_4 && mail_type <= MailType.CR_4_4)
                    report = new KReportCR3(mail_type);
                else if (mail_type > MailType.OP_START && mail_type < MailType.OP_END)
                    report = new KReportOP();
                else if (mail_type > MailType.SC_START && mail_type < MailType.SC_END)
                    report = new KReportSC();
                else if (mail_type > MailType.CC_START && mail_type < MailType.CC_END)
                    report = new KReportCC();

                if (report == null)
                    return null;

                report.load_from_json(json_text);
                return report;
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
            return null;
        }
        public virtual void load_from_json(string json_text)
        {
        }
        static private object lock_out_file = new object();
        public virtual void send_to_crm()
        {
            lock (lock_out_file)
            {
                // Normal report

                string json_text = make_json_text() + ",";

                string path = ConstEnv.get_output_file_path();

                File.AppendAllText(path, json_text + Environment.NewLine);

                MyLogger.Info("************ >>> WEBPOST REPORT");
                MyLogger.Info($"{json_text}");
                MyLogger.Info("************ <<< WEBPOST REPORT");

                // Canceled report

                json_text = make_canceled_json_text();
                if (json_text != "")
                {
                    path = ConstEnv.get_canceled_output_file_path();

                    File.AppendAllText(path, json_text + Environment.NewLine);

                    MyLogger.Info("************ >>> WEBPOST CANCELD REPORT");
                    MyLogger.Info($"{json_text}");
                    MyLogger.Info("************ <<< WEBPOST CANCELD REPORT");
                }
            }
        }
        #region Process by DB Data
        public virtual int insert_report_to_db(int mail_id)
        {
            int report_id = -1;

            try
            {
                report_id = Program.g_db.insert_report_to_main_table_db(mail_id, this);
                if (report_id != -1)
                    m_report_id = report_id;
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                return -1;
            }
            return report_id;
        }
        public virtual void make_report_from_db(int report_id)
        {
            Program.g_db.get_report_main_info_from_db(this, report_id);
        }
        static protected KReportBase create_report_by_report_id(int report_id)
        {
            MailType mail_type = Program.g_db.get_mail_type_by_report_Id(report_id);
            if (mail_type == MailType.Mail_Unknown)
            {
                MyLogger.Error($"Unknown card type : id = {report_id}");
                return null;
            }
            KReportBase report = null;

            if (mail_type > MailType.CR_START && mail_type < MailType.CR_END)
            {
                if (mail_type == MailType.CR_1 || mail_type == MailType.CR_1_1 || mail_type == MailType.CR_1_2 || mail_type == MailType.CR_1_3)
                {
                    report = new KReportCR1(mail_type);
                }
                else if (mail_type == MailType.CR_2 || mail_type == MailType.CR_2_1 || mail_type == MailType.CR_2_2)
                {
                    report = new KReportCR2(mail_type);
                }
                else if (mail_type == MailType.CR_3 || mail_type == MailType.CR_3_1 || mail_type == MailType.CR_3_2)
                {
                    report = new KReportCR3(mail_type);
                }
                else if (mail_type == MailType.CR_4 || mail_type == MailType.CR_4_1 || mail_type == MailType.CR_4_2 || mail_type == MailType.CR_4_3 || mail_type == MailType.CR_4_4)
                {
                    report = new KReportCR4(mail_type);
                }
                else if (mail_type == MailType.CR_7 || mail_type == MailType.CR_7_1 || mail_type == MailType.CR_7_2 || mail_type == MailType.CR_7_3 || mail_type == MailType.CR_7_4 || mail_type == MailType.CR_7_5)
                {
                    report = new KReportCR7(mail_type);
                }
                else if (mail_type == MailType.CR_8)
                {
                    report = new KReportCR8(mail_type);
                }
            }
            else if (mail_type > MailType.OP_START && mail_type < MailType.OP_END)
            {
                report = new KReportOP();
            }
            else if (mail_type > MailType.SC_START && mail_type < MailType.SC_END)
            {
                report = new KReportSC();
            }
            else if (mail_type > MailType.CC_START && mail_type < MailType.CC_END)
            {
                report = new KReportCC();
            }

            return report;
        }
        public virtual void report()
        {
            try
            {
                send_to_crm();
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
        }
        public virtual KReportBase merge_multi_reports()
        {
            return null;
        }
        #endregion Process by DB Data
    }
}
