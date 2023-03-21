using Logger;
using Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MimeKit;
using MailParser;

namespace MailHelper
{
    public class KMailCR4 : KMailBaseCR
    {
        public KMailCR4() : base()
        {
        }
        #region override functions
        public override bool check_valid_mail(int work_mode, string subject, string sender, out int mail_seq_num)
        {
            mail_seq_num = 0;

            if (Program.g_user.is_report_mode_gs(work_mode))
                return false;

            if (sender.StartsWith("receipts") && sender.EndsWith("@stripe.com") && subject.StartsWith("Your ‪RetailMeNot, Inc receipt [#") && subject.EndsWith("]"))
            {
                mail_seq_num = 1;
                return true;
            }
            else if (/*sender == "mail@t.retailmenot.com" && */subject.StartsWith("Transaction receipt for order "))
            {
                mail_seq_num = 2;
                return true;
            }
            else if (sender == "mail@t.retailmenot.com" && subject.StartsWith("Your ") && subject.EndsWith(" Gift Cards are here!"))
            {
                mail_seq_num = 3;
                return true;
            }
            else if (sender == "service@paypal.com" && subject.IndexOf("RetailMeNot paid you ") != -1)
            {
                mail_seq_num = 4;
                return true;
            }

            return false;
        }
        protected override bool parse_mail(MimeMessage mail, int mail_order, KReportBase base_report)
        {
            try
            {
                KReportCR4 report = base_report as KReportCR4;

                if (mail_order == 1)
                    parse_mail_4_1(mail, report);
                else if (mail_order == 2)
                    parse_mail_4_2(mail, report);
                else if (mail_order == 3)
                    parse_mail_4_3(mail, report);
                else if (mail_order == 4)
                    parse_mail_4_4(mail, report);
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
        private void parse_mail_4_1(MimeMessage mail, KReportCR4 report)
        {
            if (!XMailHelper.is_bodytext_existed(mail))
            {
                parse_mail_4_1_for_htmltext(mail, report);
                return;
            }

            report.m_purchase_date = XMailHelper.get_sentdate(mail);
            MyLogger.Info($"... 1st mail date = {report.m_purchase_date.ToString()}");

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if (line.IndexOf("Payment method", StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    string temp = lines[i + 1];
                    temp = temp.Replace("\r", "");
                    if (temp[0] != ' ')
                    {
                        string payment_type = "";
                        string last_4_digit = "";

                        string temp1 = temp;
                        temp = temp.Substring(0, temp.IndexOf("–"));
                        temp = temp.Trim();
                        if (temp[0] == '[')
                            temp = temp.Substring(1);
                        if (temp.EndsWith("]"))
                            temp = temp.Substring(0, temp.Length - 1);
                        payment_type = temp;

                        temp = temp1;
                        temp = temp.Substring(temp.IndexOf("–"));
                        temp = temp.Trim();
                        last_4_digit = temp.Substring(1);

                        report.m_cr_payment_type = payment_type;
                        report.m_cr_payment_id = last_4_digit;
                        MyLogger.Info($"... 1st mail payment type = {payment_type} payment id = {last_4_digit}");
                    }
                }

                if (line.IndexOf("Order #:", StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    string temp = line.Substring(line.IndexOf("RMN-", StringComparison.InvariantCultureIgnoreCase) + 4);
                    temp = temp.Substring(0, temp.IndexOf(" "));
                    report.set_order_id(temp);
                    MyLogger.Info($"... 1st mail order = {report.m_order_id}");
                }

                if (line.IndexOf("Amount paid", StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    string temp = line.Substring(line.IndexOf("Amount paid", StringComparison.InvariantCultureIgnoreCase) + 11);
                    if (temp == "")
                        continue;
                    temp = temp.Replace("\r", "");
                    temp = temp.Trim();
                    if (temp[0] == '$')
                        temp = temp.Substring(1);
                    report.set_total(Str_Utils.string_to_float(temp));
                    MyLogger.Info($"... 1st mail total = {report.m_total}");
                }
            }
        }
        private void parse_mail_4_1_for_htmltext(MimeMessage mail, KReportCR4 report)
        {
            report.m_purchase_date = XMailHelper.get_sentdate(mail);
            MyLogger.Info($"... 1st mail date = {report.m_purchase_date.ToString()}");

            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.IndexOf("Order #:", StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    string temp = line.Substring(line.IndexOf("RMN-", StringComparison.InvariantCultureIgnoreCase) + 4);
                    temp = temp.Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... 1st mail order = {report.m_order_id}");
                }

                if (line.IndexOf("Amount paid") != -1 && lines[++i].Trim()[0] == '$')
                {
                    string temp = lines[i].Trim();
                    report.set_total(Str_Utils.string_to_currency(temp));
                    MyLogger.Info($"... 1st mail total = {report.m_total}");
                }
            }

            string html_text = mail.HtmlBody;
            if (html_text != null && html_text.IndexOf("Payment method") != -1)
            {
                string payment_type = "";
                string last_4_digit = "";

                string temp = html_text.Substring(html_text.IndexOf("Payment method"));
                temp = temp.Substring(temp.IndexOf("</tr>") + "</tr>".Length);
                temp = temp.Substring(0, temp.IndexOf("</tbody>"));

                int next_pos;
                string span_part = XMailHelper.find_html_part(temp, "span", out next_pos);
                if (span_part == "")
                    return;
                span_part = span_part.Substring(span_part.IndexOf("alt=\"") + "alt=\"".Length);
                span_part = span_part.Substring(0, span_part.IndexOf("\""));
                payment_type = span_part;
                temp = temp.Substring(next_pos);

                span_part = XMailHelper.find_html_part(temp, "span", out next_pos);
                if (span_part == "")
                    return;
                span_part = span_part.Substring(span_part.IndexOf("–") + 1).Trim();
                last_4_digit = span_part;

                report.m_cr_payment_type = payment_type;
                report.m_cr_payment_id = last_4_digit;
                MyLogger.Info($"... 1st mail payment type = {payment_type} last_4_digit = {last_4_digit}");
            }
        }
        private void parse_mail_4_2(MimeMessage mail, KReportCR4 report)
        {
            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');

            string payment_type = "";
            string last_4_digit = "";

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if (line.IndexOf("Amount:", StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    string temp = line.Substring(line.IndexOf("Amount:", StringComparison.InvariantCultureIgnoreCase) + "Amount:".Length);
                    temp = temp.Replace("\r", "");
                    temp = temp.Replace(" ", "");
                    if (temp[0] == '$')
                        temp = temp.Substring(1);
                    if (temp.EndsWith("USD"))
                        temp = temp.Substring(0, temp.Length - 3);
                    temp = temp.Trim();
                    report.set_total(Str_Utils.string_to_float(temp));
                    MyLogger.Info($"... 2nd mail total = {report.m_total}");
                    continue;
                }

                if (line.IndexOf("Card Type:", StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    string temp = line.Substring(line.IndexOf("Card Type:", StringComparison.InvariantCultureIgnoreCase) + "Card Type:".Length);
                    temp = temp.Trim();
                    payment_type = temp;
                    MyLogger.Info($"... 2nd mail payment type = {temp}");
                    continue;
                }

                if (line.IndexOf("Credit Card Ends With:", StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    string temp = line.Substring(line.IndexOf("Credit Card Ends With:", StringComparison.InvariantCultureIgnoreCase) + "Credit Card Ends With:".Length);
                    temp = temp.Replace("\r", "");
                    temp = temp.Trim();
                    last_4_digit = temp;
                    MyLogger.Info($"... 2nd mail payment id = {temp}");
                    continue;
                }
            }
            report.m_cr_payment_type = payment_type;
            report.m_cr_payment_id = last_4_digit;
        }
        private void parse_mail_4_3(MimeMessage mail, KReportCR4 report)
        {
            if (!XMailHelper.is_bodytext_existed(mail))
            {
                parse_mail_4_3_for_htmltext(mail, report);
                return;
            }
            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if (line.IndexOf("Order Number:", StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    string temp = line.Substring(line.IndexOf("Order Number:", StringComparison.InvariantCultureIgnoreCase) + "Order Number:".Length);
                    temp = temp.Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... 3rd mail order = {report.m_order_id}");
                    continue;
                }
                if (line.StartsWith("Instant Cash Back") && line != "Instant Cash Back")
                {
                    string temp = line.Substring("Instant Cash Back".Length);
                    temp = temp.Trim();
                    if (temp[0] == '$')
                        temp = temp.Substring(1);
                    temp = temp.Trim();
                    float f = Str_Utils.string_to_float(temp);
                    report.m_instant_cashback.Add(f);
                    MyLogger.Info($"... 3rd mail instant cashback = {f}");
                    continue;
                }
                if (line.IndexOf("View Gift Card", StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    string temp = line.Substring(line.IndexOf("View Gift Card", StringComparison.InvariantCultureIgnoreCase) + "View Gift Card".Length);
                    if (temp.IndexOf("<") != -1)
                    {
                        temp = XMailHelper.trim_link(temp);
                        report.add_web_link(temp);
                        MyLogger.Info($"... 3rd mail add web link = {temp}");
                    }
                    continue;
                }
                if (line.IndexOf("View Gift Cards", StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    string temp = line.Substring(line.IndexOf("View Gift Cards", StringComparison.InvariantCultureIgnoreCase) + "View Gift Cards".Length);
                    if (temp.IndexOf("<") != -1)
                    {
                        temp = XMailHelper.trim_link(temp);
                        report.add_web_link(temp);
                        MyLogger.Info($"... 3rd mail add web link = {temp}");
                    }
                    continue;
                }
            }
        }
        private void parse_mail_4_3_for_htmltext(MimeMessage mail, KReportCR4 report)
        {
            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if (line.IndexOf("Order Number:", StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    string temp = line.Substring(line.IndexOf("Order Number:", StringComparison.InvariantCultureIgnoreCase) + "Order Number:".Length);
                    temp = temp.Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... 3rd mail order = {report.m_order_id}");
                    continue;
                }
                if (line.StartsWith("Instant Cash Back") && line != "Instant Cash Back")
                {
                    string temp = line.Substring("Instant Cash Back".Length);
                    temp = temp.Trim();
                    if (temp[0] == '$')
                        temp = temp.Substring(1);
                    temp = temp.Trim();
                    float f = Str_Utils.string_to_float(temp);
                    report.m_instant_cashback.Add(f);
                    MyLogger.Info($"... 3rd mail instant cashback = {f}");
                    continue;
                }
            }
            string html_text = mail.HtmlBody;
            if (html_text != null && (html_text.IndexOf("View Gift Cards</a>") != -1 || html_text.IndexOf("View Gift Card</a>") != -1))
            {
                string temp = "";
                if (html_text.IndexOf("View Gift Cards</a>") != -1)
                    temp = html_text.Substring(0, html_text.IndexOf("View Gift Cards</a>"));
                else if (html_text.IndexOf("View Gift Card</a>") != -1)
                    temp = html_text.Substring(0, html_text.IndexOf("View Gift Card</a>"));
                if (temp == "")
                    return;
                temp = temp.Substring(temp.LastIndexOf("<a") + 2);
                if (temp.IndexOf("href=\"") == -1)
                    return;
                temp = temp.Substring(temp.IndexOf("href=\"") + "href=\"".Length);
                temp = temp.Substring(0, temp.IndexOf("\""));
                report.add_web_link(temp.Trim());
                MyLogger.Info($"... 3rd mail add web link = {temp}");
            }
        }
        private void parse_mail_4_4(MimeMessage mail, KReportCR4 report)
        {
            string temp = XMailHelper.get_subject(mail);
            temp = temp.Replace("\r", "");
            temp = temp.Trim();
            if (temp.IndexOf("RetailMeNot paid you ", StringComparison.InvariantCultureIgnoreCase) > -1)
            {
                temp = temp.Substring(temp.IndexOf("RetailMeNot paid you ", StringComparison.InvariantCultureIgnoreCase) + "RetailMeNot paid you ".Length);
                temp = temp.Trim();
                if (temp[0] == '$')
                    temp = temp.Substring(1);
                if (temp.IndexOf("(") != -1)
                    temp = temp.Substring(0, temp.IndexOf("("));
                temp = temp.Trim();
                float f = Str_Utils.string_to_float(temp); // instant cash back : float
                report.m_instant_cashback.Add(f);
                MyLogger.Info($"... 4th mail instant cashback = {f}");
            }

        }

        #endregion class specific functions
    }
}
