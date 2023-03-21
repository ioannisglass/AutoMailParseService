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
    public class KMailCR7 : KMailBaseCR
    {
        public KMailCR7() : base()
        {
        }
        #region override functions
        public override bool check_valid_mail(int work_mode, string subject, string sender, out int mail_order)
        {
            mail_order = 0;

            if (Program.g_user.is_report_mode_gs(work_mode))
                return false;

            if (sender == "orders@oe.target.com" && subject.StartsWith("Please accept a ") && subject.EndsWith("eGiftCard with our apologies."))
            {
                mail_order = 1;
                return true;
            }
            else if (sender == "sears@account.sears.com" && subject == "Here is your refund on a Sears eGift Card")
            {
                mail_order = 2;
                return true;
            }
            else if (sender == "homedepotgiftcards@cashstar.com" && subject.StartsWith("You've received ") && subject.IndexOf("Home Depot") != -1)
            {
                mail_order = 3;
                return true;
            }
            else if (sender == "help@walmartcorp.com" && subject.StartsWith("CS Team has sent you ") && subject.IndexOf("Walmart") != -1)
            {
                mail_order = 4;
                return true;
            }
            else if (sender == "dellgiftcards@cashstar.com" && subject == "Dell Promo eGift Cards sent you a Promotional eGift Card for Dell")
            {
                mail_order = 5;
                return true;
            }

            return false;
        }
        protected override bool parse_mail(MimeMessage mail, int mail_order, KReportBase base_report)
        {
            try
            {
                KReportCR7 report = base_report as KReportCR7;

                if (mail_order == 1)
                    parse_mail_cr_7_1(mail, report);
                else if (mail_order == 2)
                    parse_mail_cr_7_2(mail, report);
                else if (mail_order == 3)
                    parse_mail_cr_7_3(mail, report);
                else if (mail_order == 4)
                    parse_mail_cr_7_4(mail, report);
                else if (mail_order == 5)
                    parse_mail_cr_7_5(mail, report);
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
        private void parse_mail_cr_7_1(MimeMessage mail, KReportCR7 report)
        {
            report.m_retailer = ConstEnv.RETAILER_TARGET;

            if (!XMailHelper.is_bodytext_existed(mail))
            {
                parse_mail_cr_7_1_for_htmltext(mail, report);
                return;
            }

            string subject = XMailHelper.get_subject(mail);
            string amount = "";

            report.m_purchase_date = XMailHelper.get_sentdate(mail);
            MyLogger.Info($"... 1st mail date = {report.m_purchase_date.ToString()}");

            if (subject.StartsWith("Please accept a ") && subject.IndexOf("eGiftCard with our apologies.") != -1)
            {
                string temp = subject.Substring("Please accept a ".Length);
                temp = temp.Substring(0, temp.IndexOf("eGiftCard with our apologies."));
                temp = temp.Trim();

                if (temp.IndexOf(" ") != -1)
                {
                    amount = temp.Substring(0, temp.IndexOf(" "));
                    if (amount[0] == '$')
                        amount = amount.Substring(1);
                    string retailer = temp.Substring(temp.IndexOf(" ") + 1);

                    report.add_giftcard_details_v1(retailer, Str_Utils.string_to_float(amount), Str_Utils.string_to_float(amount));

                    MyLogger.Info($"... 1st mail cost     = {amount}");
                    MyLogger.Info($"... 1st mail value    = {amount}");
                    MyLogger.Info($"... 1st mail retailer = {retailer}");
                }
            }

            bool has_amount = false;
            List<string> cn_list = new List<string>();
            List<string> pin_list = new List<string>();
            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Card number:"))
                {
                    string temp = line.Substring("Card number:".Length);
                    temp = temp.Trim(); // Card Number - Gift Card #
                    cn_list.Add(temp);
                    MyLogger.Info($"... 1st mail add gift card = {temp}");
                }
                if (line.StartsWith("Access number:"))
                {
                    string temp = line.Substring("Access number:".Length);
                    temp = temp.Trim(); // Access Number - Pin
                    pin_list.Add(temp);
                    MyLogger.Info($"... 1st mail add pin = {temp}");
                }

                if (line.StartsWith("Order #"))
                {
                    string temp = line.Substring("Order #".Length);
                    temp = temp.Substring(0, temp.IndexOf("<"));
                    temp = temp.Trim(); // order
                    report.set_order_id(temp);
                    MyLogger.Info($"... 1st mail order = {temp}");
                }

                if (amount != "" && line == "$" + amount)
                    has_amount = true;
            }
            int n = Math.Min(cn_list.Count, pin_list.Count);
            for (int i = 0; i < n; i++)
            {
                report.add_giftcard_details_v2(cn_list[i], pin_list[i]);
            }
            if (!has_amount)
            {
                // To Do.
                MyLogger.Error($"... NO amount in mail body of 1st mail");
            }
        }
        private void parse_mail_cr_7_1_for_htmltext(MimeMessage mail, KReportCR7 report)
        {
            string subject = XMailHelper.get_subject(mail);
            string amount = "";

            report.m_purchase_date = XMailHelper.get_sentdate(mail);
            MyLogger.Info($"... 1st mail date = {report.m_purchase_date.ToString()}");

            if (subject.StartsWith("Please accept a ") && subject.IndexOf("eGiftCard with our apologies.") != -1)
            {
                string temp = subject.Substring("Please accept a ".Length);
                temp = temp.Substring(0, temp.IndexOf("eGiftCard with our apologies."));
                temp = temp.Trim();

                if (temp.IndexOf(" ") != -1)
                {
                    amount = temp.Substring(0, temp.IndexOf(" "));
                    if (amount[0] == '$')
                        amount = amount.Substring(1);
                    string retailer = temp.Substring(temp.IndexOf(" ") + 1);

                    report.add_giftcard_details_v1(retailer, Str_Utils.string_to_float(amount), Str_Utils.string_to_float(amount));

                    MyLogger.Info($"... 1st mail cost     = {amount}");
                    MyLogger.Info($"... 1st mail value    = {amount}");
                    MyLogger.Info($"... 1st mail retailer = {retailer}");
                }
            }

            bool has_amount = false;
            List<string> cn_list = new List<string>();
            List<string> pin_list = new List<string>();
            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Card number:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp = line.Substring("Card number:".Length);
                    temp = temp.Trim(); // Card Number - Gift Card #
                    cn_list.Add(temp);
                    MyLogger.Info($"... 1st mail add gift card = {temp}");
                }
                if (line.StartsWith("Access number:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp = line.Substring("Access number:".Length);
                    temp = temp.Trim(); // Access Number - Pin
                    pin_list.Add(temp);
                    MyLogger.Info($"... 1st mail add pin = {temp}");
                }

                if (line.StartsWith("Order #", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp = line.Substring("Order #".Length);
                    report.set_order_id(temp);
                    MyLogger.Info($"... 1st mail order = {temp}");
                }

                if (amount != "" && line == "$" + amount)
                    has_amount = true;
            }
            int n = Math.Min(cn_list.Count, pin_list.Count);
            for (int i = 0; i < n; i++)
            {
                report.add_giftcard_details_v2(cn_list[i], pin_list[i]);
            }
            if (!has_amount)
            {
                // To Do.
                MyLogger.Error($"... NO amount in mail body of 1st mail");
            }
        }
        private void parse_mail_cr_7_2(MimeMessage mail, KReportCR7 report)
        {
            string subject = XMailHelper.get_subject(mail);

            if (subject.IndexOf("Sears") == -1)
                throw new Exception($"Invalid mail 2 of vendor 7. subject = {subject}");

            report.m_purchase_date = XMailHelper.get_sentdate(mail);
            report.m_retailer = ConstEnv.RETAILER_SEARS;

            MyLogger.Info($"... 2nd mail date = {report.m_purchase_date.ToString()}");

            bool is_link_got = false;
            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Your Sears Refund:"))
                {
                    string temp = line.Substring("Your Sears Refund:".Length);
                    temp = temp.Trim();
                    if (temp[0] == '$')
                        temp = temp.Substring(1);
                    temp = temp.Trim();
                    float f = Str_Utils.string_to_float(temp); // value and cost

                    report.add_giftcard_details_v1(ConstEnv.RETAILER_SEARS, f, f);

                    MyLogger.Info($"... 2nd mail cost     = {f}");
                    MyLogger.Info($"... 2nd mail value    = {f}");
                }

                if (line.StartsWith("REFUND FOR RETURN|"))
                {
                    string temp = line.Substring("REFUND FOR RETURN|".Length);
                    temp = temp.Trim();
                    report.set_order_id(temp); // order
                    MyLogger.Info($"... 2nd mail order = {report.m_order_id}");
                }

                if (line.StartsWith("[REDEEM/PRINT]"))
                {
                    string temp = line.Substring("[REDEEM/PRINT]".Length);
                    temp = temp.Replace("\r", "");
                    temp = XMailHelper.trim_link(temp);

                    report.add_web_link(temp);
                    MyLogger.Info($"... 2nd mail add web link = {temp}");
                    is_link_got = true;
                }
            }

            if (!is_link_got)
            {
                string[] html_lines = XMailHelper.get_htmltext(mail).Replace("\r", "").Split('\n');
                int blue_btn_count = 0;
                foreach(string html_line in html_lines)
                {
                    if (!html_line.Contains("alt=\"REDEEM/PRINT\""))
                        continue;
                    
                    string temp = html_line.Trim();
                    temp = temp.Substring(temp.IndexOf("href=\"") + "href=\"".Length).Trim();
                    temp = temp.Substring(0, temp.IndexOf("\"")).Trim();
                    report.add_web_link(temp);
                    MyLogger.Info($"... 2nd mail add web link = {temp}");
                    blue_btn_count++;
                }
                if (blue_btn_count > 1)
                    MyLogger.Info($"... 2nd mail blue button Count - {blue_btn_count}");
            }
        }
        private void parse_mail_cr_7_3(MimeMessage mail, KReportCR7 report)
        {
            string subject = XMailHelper.get_subject(mail);

            if (subject.IndexOf("Home Depot") == -1)
                throw new Exception($"Invalid mail 3 of vendor 7. subject = {subject}");

            report.m_retailer = ConstEnv.RETAILER_HOMEDEPOT;
            report.m_purchase_date = XMailHelper.get_sentdate(mail);

            string mailto = XMailHelper.get_mailto(mail);

            MyLogger.Info($"... 3rd mail date = {report.m_purchase_date.ToString()}");

            bool is_link_got = false;

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("You've received a "))
                {
                    string temp = line.Substring("You've received a ".Length);
                    temp = temp.Trim();
                    if (temp.IndexOf(" USD") != -1)
                    {
                        temp = temp.Substring(0, temp.IndexOf(" USD"));
                    }
                    else if (temp.IndexOf(" ") != -1)
                    {
                        temp = temp.Substring(0, temp.IndexOf(" "));
                    }
                    else
                    {
                        continue;
                    }
                    temp = temp.Trim(); //  “value” and “cost”
                    if (temp[0] == '$')
                        temp = temp.Substring(1);
                    float f = Str_Utils.string_to_float(temp); // value and cost

                    report.add_giftcard_details_v1(ConstEnv.RETAILER_HOMEDEPOT, f, f);

                    MyLogger.Info($"... 3rd mail cost     = {f}");
                    MyLogger.Info($"... 3rd mail value    = {f}");
                }
                if (line.StartsWith("VIEW MY ESTORE CREDIT"))
                {
                    string temp = line.Substring("VIEW MY ESTORE CREDIT".Length);
                    temp = XMailHelper.trim_link(temp);
                    report.add_web_link(temp, mailto);
                    MyLogger.Info($"... 3rd mail add web link = {temp}");
                    is_link_got = true;
                }
            }

            if (!is_link_got)
            {
                if (!is_link_got)
                {
                    string link = ExtractLinkFromHTML(mail);
                    report.add_web_link(link, mailto);
                    MyLogger.Info($"... 3th mail add web link = {link}");
                }
            }
        }
        private void parse_mail_cr_7_4(MimeMessage mail, KReportCR7 report)
        {
            string subject = XMailHelper.get_subject(mail);

            if (subject.IndexOf("Walmart") == -1)
                throw new Exception($"Invalid mail 4 of vendor 7. subject = {subject}");

            report.m_retailer = ConstEnv.RETAILER_WALMART;
            report.m_purchase_date = XMailHelper.get_sentdate(mail);

            string mailto = XMailHelper.get_mailto(mail);

            MyLogger.Info($"... 4th mail date = {report.m_purchase_date.ToString()}");

            bool is_link_got = false;
            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("CS Team sent you a "))
                {
                    string temp = line.Substring("CS Team sent you a ".Length);
                    temp = temp.Trim();
                    if (temp.IndexOf(" USD") != -1)
                    {
                        temp = temp.Substring(0, temp.IndexOf(" USD"));
                    }
                    else if (temp.IndexOf(" ") != -1)
                    {
                        temp = temp.Substring(0, temp.IndexOf(" "));
                    }
                    else
                    {
                        continue;
                    }
                    temp = temp.Trim(); //  “value” and “cost”
                    if (temp[0] == '$')
                        temp = temp.Substring(1);
                    float f = Str_Utils.string_to_float(temp); // value and cost

                    report.add_giftcard_details_v1(ConstEnv.RETAILER_WALMART, f, f);

                    MyLogger.Info($"... 4th mail cost     = {f}");
                    MyLogger.Info($"... 4th mail value    = {f}");
                }
                if (line.IndexOf(" View My eGift Card ") != -1)
                {
                    string temp = line.Substring(line.IndexOf(" View My eGift Card ") + " View My eGift Card ".Length);
                    temp = XMailHelper.trim_link(temp);
                    report.add_web_link(temp, mailto);
                    is_link_got = true;
                    MyLogger.Info($"... 4th mail add web link = {temp}");
                }
            }

            if (!is_link_got)
            {
                string link = ExtractLinkFromHTML(mail);
                report.add_web_link(link, mailto);
                MyLogger.Info($"... 4th mail add web link = {link}");
            }
        }
        private void parse_mail_cr_7_5(MimeMessage mail, KReportCR7 report)
        {
            string subject = XMailHelper.get_subject(mail);

            if (subject != "Dell Promo eGift Cards sent you a Promotional eGift Card for Dell")
                throw new Exception($"Invalid mail 5 of vendor 7. subject = {subject}");

            report.m_retailer = ConstEnv.RETAILER_DELL;
            report.m_purchase_date = XMailHelper.get_sentdate(mail);

            MyLogger.Info($"... 5th mail date = {report.m_purchase_date.ToString()}");

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.IndexOf("Code value:") != -1)
                {
                    string temp = line.Substring("Code value:".Length);
                    temp = temp.Trim();
                    if (temp.IndexOf(" USD") != -1)
                    {
                        temp = temp.Substring(0, temp.IndexOf(" USD"));
                    }
                    temp = temp.Trim(); //  “value” and “cost”
                    if (temp[0] == '$')
                        temp = temp.Substring(1);
                    float f = Str_Utils.string_to_float(temp); // value and cost

                    report.add_giftcard_details_v1(ConstEnv.RETAILER_DELL, f, f);

                    MyLogger.Info($"... 5th mail cost     = {f}");
                    MyLogger.Info($"... 5th mail value    = {f}");
                }
                if (line.IndexOf("View My Promotional eGift Card") != -1)
                {
                    string temp = line.Substring(line.IndexOf("View My Promotional eGift Card") + "View My Promotional eGift Card".Length);
                    temp = XMailHelper.trim_link(temp);
                    report.add_web_link(temp);
                    MyLogger.Info($"... 5th mail add web link = {temp}");
                }
            }
        }

        #endregion class specific functions

        private string ExtractLinkFromHTML(MimeMessage mail)
        {
            string html_txt = XMailHelper.get_htmltext(mail);
            string[] html_lines = html_txt.Replace("\r", "").Split('\n');
            string link = string.Empty;

            foreach (string html_line in html_lines)
            {
                if (!html_line.Contains("href=\""))
                    continue;
                string temp = html_line.Trim();
                temp = temp.Substring(temp.IndexOf("href=\"") + "href=\"".Length);
                link = temp.Substring(0, temp.IndexOf("\"")).Trim();
                temp = temp.Substring(temp.IndexOf("\"")).Trim();
                if (!temp.Contains(link))
                    continue;

                break;
            }
            return link;
        }
    }
}
