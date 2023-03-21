using Logger;
using Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using MimeKit;
using MailParser;

namespace MailHelper
{
    public class KMailCR2 : KMailBaseCR
    {
        public KMailCR2() : base()
        {
        }
        #region override functions
        public override bool check_valid_mail(int work_mode, string subject, string sender, out int mail_seq_num)
        {
            mail_seq_num = 0;

            if (Program.g_user.is_report_mode_gs(work_mode))
                return false;

            if (sender != "orders@cardpool.com")
                return false;

            if (subject.StartsWith("Cardpool Order Confirmation (Order #") && subject.EndsWith(")"))
            {
                mail_seq_num = 1;
                return true;
            }
            else if (subject.StartsWith("Cardpool Electronic Gift Card Delivery (Order #") && subject.EndsWith(")"))
            {
                mail_seq_num = 2;
                return true;
            }

            return false;
        }
        protected override bool parse_mail(MimeMessage mail, int mail_order, KReportBase base_report)
        {
            try
            {
                KReportCR2 report = base_report as KReportCR2;

                if (mail_order == 1)
                    parse_mail_cr_2_1(mail, report);
                else if (mail_order == 2)
                    parse_mail_cr_2_2(mail, report);
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                return false;
            }
            return true;
        }
        #endregion override functions

        #region class specific functions
        private void parse_mail_cr_2_1(MimeMessage mail, KReportCR2 report)
        {
            string str_chk_order = "Order #";
            string subject = XMailHelper.get_subject(mail);

            report.m_purchase_date = XMailHelper.get_sentdate(mail);

            int idx_order_first = subject.IndexOf(str_chk_order, StringComparison.InvariantCultureIgnoreCase) + str_chk_order.Length;
            string temp = subject.Substring(idx_order_first);
            temp = temp.Substring(0, temp.IndexOf(")", StringComparison.InvariantCultureIgnoreCase));
            report.set_order_id(temp.Trim());
            MyLogger.Info($"... 1st mail order = {report.m_order_id}");

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.IndexOf("You saved", StringComparison.InvariantCultureIgnoreCase) > -1 &&
                        line.IndexOf("off your order by buying", StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    temp = line.Substring(line.IndexOf("You saved", StringComparison.InvariantCultureIgnoreCase) + 9);
                    temp = temp.Substring(0, temp.IndexOf("off your order by buying"));
                    temp = temp.Trim();
                    if (temp.EndsWith("%"))
                        temp = temp.Substring(0, temp.Length - 1);
                    report.m_discount = Str_Utils.string_to_float(temp.Trim());
                    MyLogger.Info($"... 1st mail discount = {report.m_discount}%");

                    temp = line.Substring(line.IndexOf("items for only", StringComparison.InvariantCultureIgnoreCase) + 14);
                    temp = temp.Trim();
                    if (temp.StartsWith("$"))
                        temp = temp.Substring(1);
                    if (temp.EndsWith("."))
                        temp = temp.Substring(0, temp.Length - 1);
                    temp = temp.Trim();
                    report.set_total(Str_Utils.string_to_float(temp));
                    MyLogger.Info($"... 1st mail total = {report.m_total}");
                }

                if (line.IndexOf("Electronic Gift Card, sold at", StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    temp = line.Substring(line.IndexOf("$"));
                    temp = temp.Substring(0, temp.IndexOf(" "));
                    if (temp.StartsWith("$"))
                        temp = temp.Substring(1);
                    if (temp.EndsWith("."))
                        temp = temp.Substring(0, temp.Length - 1);
                    temp = temp.Trim();
                    string value = temp;

                    temp = line.Substring(line.IndexOf("$") + 1 + value.Length);
                    temp = temp.Substring(0, temp.IndexOf("Electronic Gift Card, sold at", StringComparison.InvariantCultureIgnoreCase));
                    if (temp[0] == ' ')
                        temp = temp.Substring(1);
                    if (temp[temp.Length - 1] == ' ')
                        temp = temp.Substring(0, temp.Length - 1);
                    string retailer = temp;

                    temp = line.Substring(line.IndexOf("discount for", StringComparison.InvariantCultureIgnoreCase) + 12);
                    temp = temp.Trim();
                    if (temp.StartsWith("$"))
                        temp = temp.Substring(1);
                    if (temp.EndsWith("."))
                        temp = temp.Substring(0, temp.Length - 1);
                    temp = temp.Trim();
                    string cost = temp;

                    float c = Str_Utils.string_to_float(cost);
                    float v = Str_Utils.string_to_float(value);

                    MyLogger.Info($"... 1st mail add cost     = {c}");
                    MyLogger.Info($"... 1st mail add value    = {v}");
                    MyLogger.Info($"... 1st mail add retailer = {retailer}");

                    KReportCR v_card = report as KReportCR;
                    v_card.add_giftcard_details(retailer, v, c, "", "");

                }
            }
        }
        private void parse_mail_cr_2_2(MimeMessage mail, KReportCR2 report)
        {
            parse_mail_cr_2_2_for_htmltext(mail, report);
        }
        private void parse_mail_cr_2_2_for_htmltext(MimeMessage mail, KReportCR2 report)
        {
            string subject = XMailHelper.get_subject(mail);

            if (!subject.StartsWith("Cardpool Electronic Gift Card Delivery (Order #") || !subject.EndsWith(")"))
                throw new Exception($"Invalid Vendor2 2nd mail subject. {subject}");

            string order = subject.Substring("Cardpool Electronic Gift Card Delivery (Order #".Length);
            order = order.Substring(0, order.Length - 1);

            report.set_order_id(order);
            MyLogger.Info($"... 2nd mail order = {report.m_order_id}");

            List<string> web_links = new List<string>();

            string html_text = mail.HtmlBody;

            int next_pos;
            string temp;
            next_pos = html_text.IndexOf("You may redeem these codes online", StringComparison.CurrentCultureIgnoreCase);
            if (next_pos != -1)
            {
                temp = html_text.Substring(next_pos);
                next_pos = temp.IndexOf("</p>");
                if (next_pos != -1)
                    temp = temp.Substring(next_pos + "</p>".Length);

                string card_part = XMailHelper.find_html_part(temp, "p", out next_pos);
                card_part = XMailHelper.html2text(card_part);
                card_part.Replace("\r\n", "\n");
                while (card_part.IndexOf("Number:", StringComparison.CurrentCultureIgnoreCase) != -1 && card_part.IndexOf("Pin:", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    if (card_part.IndexOf(" ") == -1)
                        break;
                    string temp1 = card_part.Substring(0, card_part.IndexOf(" "));
                    float value = Str_Utils.string_to_currency(temp1);

                    string retailer = "";
                    temp1 = card_part.Substring(card_part.IndexOf(" ") + 1).Trim();
                    if (temp1.IndexOf("Electronic Gift Card", StringComparison.CurrentCultureIgnoreCase) != -1)
                        retailer = temp1.Substring(0, temp1.IndexOf("Electronic Gift Card", StringComparison.CurrentCultureIgnoreCase)).Trim();

                    temp1 = card_part.Substring(card_part.IndexOf("Number:", StringComparison.CurrentCultureIgnoreCase) + "Number:".Length).Trim();
                    string card_number = temp1.Substring(0, temp1.IndexOf("Pin:", StringComparison.CurrentCultureIgnoreCase)).Trim();
                    if (card_number.EndsWith(","))
                        card_number = card_number.Substring(0, card_number.Length - 1).Trim();
                    string pin = temp1.Substring(temp1.IndexOf("Pin:", StringComparison.CurrentCultureIgnoreCase) + "Pin:".Length).Trim();

                    temp = temp.Substring(next_pos);
                    card_part = XMailHelper.find_html_part(temp, "p", out next_pos);
                    card_part = XMailHelper.html2text(card_part);
                    card_part.Replace("\r\n", "\n");

                    report.add_giftcard_details(retailer, value, 0, card_number, pin);

                    MyLogger.Info($"... 2nd mail add retailer       = {retailer}");
                    MyLogger.Info($"... 2nd mail Add value          = {value}");
                    MyLogger.Info($"... 2nd mail Add card number    = {card_number}");
                    MyLogger.Info($"... 2nd mail Add pin            = {pin}");
                }
            }

            next_pos = html_text.IndexOf("You may print these out and redeem them in store:", StringComparison.CurrentCultureIgnoreCase);
            if (next_pos != -1)
            {
                temp = html_text.Substring(next_pos);
                next_pos = temp.IndexOf("</p>");
                if (next_pos != -1)
                    temp = temp.Substring(next_pos + "</p>".Length);

                string weblink_part = XMailHelper.find_html_part(temp, "p", out next_pos);
                while (weblink_part.StartsWith("$") && weblink_part.IndexOf("Electronic Gift Card", StringComparison.CurrentCultureIgnoreCase) != -1 && weblink_part.IndexOf("<a href=") != -1)
                {
                    if (weblink_part.IndexOf(" ") == -1)
                        break;
                    string temp1 = weblink_part.Substring(0, weblink_part.IndexOf(" "));
                    float value = Str_Utils.string_to_currency(temp1);

                    string retailer = "";
                    temp1 = weblink_part.Substring(weblink_part.IndexOf(" ") + 1).Trim();
                    if (temp1.IndexOf("Electronic Gift Card", StringComparison.CurrentCultureIgnoreCase) != -1)
                        retailer = temp1.Substring(0, temp1.IndexOf("Electronic Gift Card", StringComparison.CurrentCultureIgnoreCase)).Trim();

                    temp1 = weblink_part.Substring(weblink_part.IndexOf("<a href=\"") + "<a href=\"".Length).Trim();
                    string link = temp1.Substring(0, temp1.IndexOf("\">")).Trim();

                    temp = temp.Substring(next_pos);
                    weblink_part = XMailHelper.find_html_part(temp, "p", out next_pos);

                    report.add_web_link(link, retailer, value);

                    MyLogger.Info($"... 2nd mail Add scrap link     = {link}");
                    MyLogger.Info($"... 2nd mail Add retailer       = {retailer}");
                    MyLogger.Info($"... 2nd mail Add value          = {value}");
                }
            }
        }
        #endregion class specific functions
    }
}
