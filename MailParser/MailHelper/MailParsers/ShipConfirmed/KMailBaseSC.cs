using MailParser;
using Logger;
using MimeKit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailHelper
{
    partial class KMailBaseSC : KMailBaseParser
    {
        public KMailBaseSC() : base()
        {
        }
        private string get_post_type(string src)
        {
            string post_type = src.ToLower();

            if (post_type == "fedex express" || post_type == "fedex" || post_type == "fed" || post_type.IndexOf("fedex") != -1)
                return KReportBase.POST_TYPE_FEDEX;
            if (post_type == "united parcel service" || post_type == "ups" || post_type.IndexOf("ups") != -1)
                return KReportBase.POST_TYPE_UPS;
            if (post_type == "united states postal service" || post_type == "usps" || post_type.IndexOf("usps") != -1)
                return KReportBase.POST_TYPE_USPS;
            return src;
        }
        #region override functions
        public override bool check_valid_mail(int work_mode, string subject, string sender, out int mail_order)
        {
            mail_order = 0;

            if (sender == "orders@basspronews.com" && subject == "Bass Pro Shops Order Has Shipped")
            {
                mail_order = 1;
                return true;
            }
            else if (sender == "orders@oe.target.com" && (subject.StartsWith("Good news! An item from order") && subject.EndsWith("has shipped.")) || (subject.StartsWith("Good news! Items from order") && subject.EndsWith("have shipped.")))
            {
                mail_order = 2;
                return true;
            }
            else if (sender == "BestBuyInfo@emailinfo.bestbuy.com" && subject.StartsWith("Your order") && subject.EndsWith(" has shipped"))
            {
                if (Program.g_user.is_report_mode_gs(work_mode))
                    return false;
                mail_order = 3;
                return true;
            }
            else if ((sender == "shipment-tracking@amazon.com" && subject.StartsWith("Your AmazonSmile order") && subject.EndsWith("has shipped"))
                || (sender == "fba-noreply@amazon.com" && subject.StartsWith("Your order has shipped")))
            {
                if (Program.g_user.is_report_mode_gs(work_mode))
                    return false;
                mail_order = 4;
                return true;
            }
            else if (sender == "HomeDepot@order.homedepot.com" && subject.StartsWith("Your order") && subject.IndexOf("just shipped") != -1)
            {
                if (Program.g_user.is_report_mode_gs(work_mode))
                    return false;
                mail_order = 5;
                return true;
            }
            else if (sender == "from@notifications.dcsg.com" && subject == "Your shipment is on its way!")
            {
                if (Program.g_user.is_report_mode_gs(work_mode))
                    return false;
                mail_order = 6;
                return true;
            }
            else if (sender == "transaction@samsclub.com" && subject == "Your SamsClub.com order was delivered")
            {
                if (Program.g_user.is_report_mode_gs(work_mode))
                    return false;
                mail_order = 7;
                return true;
            }
            else if (sender == "transaction@samsclub.com" && subject == "Your SamsClub.com order has shipped")
            {
                if (Program.g_user.is_report_mode_gs(work_mode))
                    return false;
                mail_order = 8;
                return true;
            }
            else if (sender == "help@walmart.com" && subject.StartsWith("Shipped and arriving"))
            {
                if (Program.g_user.is_report_mode_gs(work_mode))
                    return false;
                mail_order = 9;
                return true;
            }
            else if (sender == "orders@order.cabelas.com" && subject == "Update on your order")
            {
                mail_order = 10;
                return true;
            }
            else if (sender == "sears@account.sears.com" && subject.StartsWith("\"Where's my stuff?\"") && subject.EndsWith("Open and find out!"))
            {
                if (Program.g_user.is_report_mode_gs(work_mode))
                    return false;
                mail_order = 11;
                return true;
            }
            else if (sender == "CustomerCare@lowes.com" && subject.StartsWith("Your Item(s) Are Waiting! Don't Forget to Pick Them Up #"))
            {
                mail_order = 12;
                return true;
            }
            return false;
        }
        protected override bool parse_mail(MimeMessage mail, int mail_order, KReportBase base_card)
        {
            try
            {
                KReportSC card = base_card as KReportSC;

                if (mail_order == 1)
                    parse_mail_sc_1(mail, card);
                else if (mail_order == 2)
                    parse_mail_sc_2(mail, card);
                else if (mail_order == 3)
                    parse_mail_sc_3(mail, card);
                else if (mail_order == 4)
                    parse_mail_sc_4(mail, card);
                else if (mail_order == 5)
                    parse_mail_sc_5(mail, card);
                else if (mail_order == 6)
                    parse_mail_sc_6(mail, card);
                else if (mail_order == 7)
                    parse_mail_sc_7(mail, card);
                else if (mail_order == 8)
                    parse_mail_sc_8(mail, card);
                else if (mail_order == 9)
                    parse_mail_sc_9(mail, card);
                else if (mail_order == 10)
                    parse_mail_sc_10(mail, card);
                else if (mail_order == 11)
                    parse_mail_sc_11(mail, card);
                else if (mail_order == 12)
                    parse_mail_sc_12(mail, card);
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
