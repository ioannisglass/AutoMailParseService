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
    public class KMailCR3 : KMailBaseCR
    {
        public KMailCR3() : base()
        {
        }
        #region override functions
        public override bool check_valid_mail(int work_mode, string subject, string sender, out int mail_seq_num)
        {
            mail_seq_num = 0;

            if (sender != "noreply@giftcardspread.com")
                return false;

            if (Program.g_user.is_report_mode_gs(work_mode))
                return false;

            if (subject.EndsWith(": Thank you for placing your order with us"))
            {
                mail_seq_num = 1;
                return true;
            }
            else if (subject.StartsWith("Your order # ") && subject.EndsWith(" has been shipped"))
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
                KReportCR3 report = base_report as KReportCR3;

                if (mail_order == 1)
                    parse_mail_cr_3_1(mail, report);
                else if (mail_order == 2)
                    parse_mail_cr_3_2(mail, report);
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
        private void parse_mail_cr_3_1(MimeMessage mail, KReportCR3 report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.set_order_id(subject.Substring(0, subject.IndexOf(":")));
            report.m_purchase_date = XMailHelper.get_sentdate(mail);
            MyLogger.Info($"... 1st mail order = {report.m_order_id}");
            MyLogger.Info($"... 1st mail date  = {report.m_purchase_date.ToString()}");

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length;)
            {
                string line = lines[i].Trim();

                int i1 = line.IndexOf("Merchant", StringComparison.InvariantCultureIgnoreCase);
                int i2 = line.IndexOf("Delivery Type", StringComparison.InvariantCultureIgnoreCase);
                int i3 = line.IndexOf("Current Value", StringComparison.InvariantCultureIgnoreCase);
                int i4 = line.IndexOf("Paid Amount", StringComparison.InvariantCultureIgnoreCase);

                if (line.IndexOf("Merchant", StringComparison.InvariantCultureIgnoreCase) > -1 &&
                    line.IndexOf("Delivery Type", StringComparison.InvariantCultureIgnoreCase) > -1 &&
                    line.IndexOf("Current Value", StringComparison.InvariantCultureIgnoreCase) > -1 &&
                    line.IndexOf("Paid Amount", StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    while (true)
                    {
                        i++;
                        line = lines[i];

                        if (line.IndexOf("$") == -1)
                            continue;
                        if (line.IndexOf("SUB TOTAL :", StringComparison.InvariantCultureIgnoreCase) > -1)
                            break;

                        string retailer = line.Substring(0, line.IndexOf("  "));
                        string temp = line.Substring(line.IndexOf(" $"));
                        temp = temp.Trim();
                        string value = temp.Substring(0, temp.IndexOf(" "));
                        value = value.Trim();

                        temp = temp.Substring(2);
                        temp = temp.Substring(temp.IndexOf(" $"));
                        temp = temp.Trim();
                        string cost = temp;

                        if (cost[0] == '$')
                            cost = cost.Substring(1);
                        if (value[0] == '$')
                            value = value.Substring(1);

                        report.add_giftcard_details(retailer, Str_Utils.string_to_float(value), Str_Utils.string_to_float(cost), "", "");

                        MyLogger.Info($"... 1st mail cost     = {cost}");
                        MyLogger.Info($"... 1st mail value    = {value}");
                        MyLogger.Info($"... 1st mail retailer = {retailer}");
                    }
                }

                if (line.IndexOf("Order Number:", StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    string temp = line.Substring(line.IndexOf("Order Number:", StringComparison.InvariantCultureIgnoreCase) + 13);
                    temp = temp.Trim();
                    if (report.m_order_id != temp)
                        throw new Exception($"Card order between subject and mail contents is mismatched. {report.m_order_id} != {temp}");
                }

                if (line.IndexOf("Order Amount:", StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    string temp = line.Substring(line.IndexOf("Order Amount:", StringComparison.InvariantCultureIgnoreCase) + 13);
                    temp = temp.Trim();
                    if (temp[0] == '$')
                        temp = temp.Substring(1);
                    temp = temp.Trim();
                    report.set_total(Str_Utils.string_to_float(temp));
                    MyLogger.Info($"... 1st mail total = {report.m_total}");
                }

                i++;
            }
        }
        private void parse_mail_cr_3_2(MimeMessage mail, KReportCR3 report)
        {
            string subject = XMailHelper.get_subject(mail);

            if (!subject.StartsWith("Your order # ") || !subject.EndsWith(" has been shipped"))
                throw new Exception($"Invalid Vendor3 2nd mail subject. {subject}");

            string order = subject.Substring("Your order # ".Length);
            order = order.Substring(0, order.Length - " has been shipped".Length);
            report.set_order_id(order);
            MyLogger.Info($"... 2nd mail order = {report.m_order_id}");

            report.add_web_link("https://www.giftcardspread.com/login");
            MyLogger.Info($"... add 2nd mail web link = {report.m_scrap_params[0].link}");
        }

        #endregion class specific functions
    }
}
