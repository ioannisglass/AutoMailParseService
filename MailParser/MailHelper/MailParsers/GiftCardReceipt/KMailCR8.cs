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
    public class KMailCR8 : KMailBaseCR
    {
        public KMailCR8() : base()
        {
        }
        #region override functions
        public override bool check_valid_mail(int work_mode, string subject, string sender, out int mail_order)
        {
            mail_order = 0;

            if (Program.g_user.is_report_mode_gs(work_mode))
                return false;

            if (sender == "gc-orders@gc.email.amazon.com" && subject.IndexOf("sent you a Gift Card for", StringComparison.CurrentCultureIgnoreCase) != -1)
            {
                mail_order = 0;
                return true;
            }

            return false;
        }
        protected override bool parse_mail(MimeMessage mail, int mail_order, KReportBase base_report)
        {
            try
            {
                KReportCR8 report = base_report as KReportCR8;

                string subject = XMailHelper.get_subject(mail);
                if (subject.IndexOf("sent you a Gift Card for", StringComparison.CurrentCultureIgnoreCase) == -1)
                    return false;

                report.m_retailer = subject.Substring(subject.IndexOf("sent you a Gift Card for", StringComparison.CurrentCultureIgnoreCase) + "sent you a Gift Card for".Length).Trim();
                if (report.m_retailer == "")
                {
                    MyLogger.Error($"Invalid CR_8 subject. no retailer : {subject}");
                    return false;
                }
                MyLogger.Info($"CR_8 retailer = {report.m_retailer}");
                report.m_mail_type = KReportBase.MailType.CR_8;

                string html_text = XMailHelper.get_htmltext(mail);
                string temp = html_text;

                if (temp.IndexOf("You've received a ", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    temp = temp.Substring(temp.IndexOf("You've received a ", StringComparison.CurrentCultureIgnoreCase) + "You've received a ".Length).Trim();
                    if (temp.IndexOf(" ") != -1)
                    {
                        temp = temp.Substring(0, temp.IndexOf(" ")).Trim();
                        float total = Str_Utils.string_to_currency(temp);
                        report.set_total(total);

                        MyLogger.Info($"CR_8 value = {total}");
                    }
                }

                temp = html_text;
                if (temp.IndexOf("Order Number:", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    temp = temp.Substring(temp.IndexOf("Order Number:", StringComparison.CurrentCultureIgnoreCase) + "Order Number:".Length).Trim();
                    if (temp.IndexOf("<u></u>") != -1)
                    {
                        temp = temp.Substring(0, temp.IndexOf("<u></u>"));

                        if (temp.IndexOf(">") != -1)
                        {
                            temp = temp.Substring(temp.LastIndexOf(">") + 1).Trim();

                            report.set_order_id(temp);
                            MyLogger.Info($"CR_8 order number = {temp}");
                        }
                    }
                }

                temp = html_text;
                if (temp.IndexOf("Selecting this button will take you to", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    temp = temp.Substring(0, temp.IndexOf("Selecting this button will take you to", StringComparison.CurrentCultureIgnoreCase)).Trim();
                    if (temp.IndexOf("<a href=\"") != -1)
                    {
                        temp = temp.Substring(temp.LastIndexOf("<a href=\"") + "<a href=\"".Length).Trim();
                        if (temp.IndexOf("\"") != -1)
                        {
                            temp = temp.Substring(0, temp.IndexOf("\"")).Trim();
                            temp = XMailHelper.html2text(temp);
                            report.add_web_link(temp);
                            MyLogger.Info($"CR_8 add web link = {temp}");
                        }
                    }
                }
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

        #endregion class specific functions

    }
}
