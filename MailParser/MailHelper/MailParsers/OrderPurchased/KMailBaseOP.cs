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
    partial class KMailBaseOP : KMailBaseParser
    {
        public KMailBaseOP() : base()
        {
        }
        #region override functions
        public override bool check_valid_mail(int work_mode, string subject, string sender, out int mail_order)
        {
            mail_order = 0;

            if ((sender == "orders@basspronews.com" && subject == "Bass Pro Shops Order Confirmation")
                || (sender == "bassproshops@t.basspronews.com" && subject.StartsWith("Great News", StringComparison.CurrentCultureIgnoreCase) && subject.EndsWith("Your Order is Being Processed!", StringComparison.CurrentCultureIgnoreCase))
                )
            {
                mail_order = 1;
                return true;
            }
            else if (sender == "orders@oe.target.com" && subject.StartsWith("Thanks for shopping with us! Here's your order #:"))
            {
                mail_order = 2;
                return true;
            }
            else if (sender == "BestBuyInfo@emailinfo.bestbuy.com" && subject.StartsWith("We've received your order #"))
            {
                if (Program.g_user.is_report_mode_gs(work_mode))
                    return false;
                mail_order = 3;
                return true;
            }
            else if (sender == "auto-confirm@amazon.com" && subject.StartsWith("Your Amazon.com order"))
            {
                if (Program.g_user.is_report_mode_gs(work_mode))
                    return false;
                mail_order = 4;
                return true;
            }
            else if (sender == "HomeDepot@order.homedepot.com" && subject == "We received your order!")
            {
                if (Program.g_user.is_report_mode_gs(work_mode))
                    return false;
                mail_order = 5;
                return true;
            }
            else if (sender == "from@notifications.dcsg.com" && subject == "Thank you for your order!")
            {
                if (Program.g_user.is_report_mode_gs(work_mode))
                    return false;
                mail_order = 6;
                return true;
            }
            else if (sender == "transaction@samsclub.com" && subject == "SamsClub.com order confirmation")
            {
                if (Program.g_user.is_report_mode_gs(work_mode))
                    return false;
                mail_order = 7;
                return true;
            }
            else if (sender == "help@walmart.com" && subject.StartsWith("Order received."))
            {
                if (Program.g_user.is_report_mode_gs(work_mode))
                    return false;
                mail_order = 9;
                return true;
            }
            else if (sender == "orders@order.cabelas.com" && subject == "Thank you for your Cabela's order!")
            {
                mail_order = 10;
                return true;
            }
            else if (sender == "sears@account.sears.com" && subject.StartsWith("Hey,") && subject.EndsWith("we just received your order! Details inside"))
            {
                if (Program.g_user.is_report_mode_gs(work_mode))
                    return false;
                mail_order = 11;
                return true;
            }
            else if (sender == "CustomerCare@lowes.com" && subject == "Your Order is in Process")
            {
                mail_order = 12;
                return true;
            }
            else if (sender == "ebay@ebay.com" && subject.IndexOf("ORDER CONFIRMED:") != -1)
            {
                if (Program.g_user.is_report_mode_gs(work_mode))
                    return false;
                mail_order = 14;
                return true;
            }
            else if (sender == "support@orders.staples.com" && subject.StartsWith("Confirmation of Staples Order: #"))
            {
                if (Program.g_user.is_report_mode_gs(work_mode))
                    return false;
                mail_order = 15;
                return true;
            }
            else if (sender == "myhpsales@hp.com" && subject.StartsWith("Your HP order") && subject.EndsWith(" has been received"))
            {
                if (Program.g_user.is_report_mode_gs(work_mode))
                    return false;
                mail_order = 16;
                return true;
            }
            else if (sender == "OfficeDepotOrders@officedepot.com" && subject.StartsWith("Order Confirmation #"))
            {
                if (Program.g_user.is_report_mode_gs(work_mode))
                    return false;
                mail_order = 17;
                return true;
            }
            else if (sender == "Kohls@t.kohls.com" && subject.StartsWith("Thanks for your order,"))
            {
                if (Program.g_user.is_report_mode_gs(work_mode))
                    return false;
                mail_order = 18;
                return true;
            }
            else if (sender == "dell_automated_email@dell.com" && subject.StartsWith("Dell Order Has Been Received for Dell Purchase ID:"))
            {
                if (Program.g_user.is_report_mode_gs(work_mode))
                    return false;
                mail_order = 19;
                return true;
            }
            else if (sender == "confirm@order.dell.com" && subject == "Your Dell Order has been Confirmed!")
            {
                if (Program.g_user.is_report_mode_gs(work_mode))
                    return false;
                mail_order = 20;
                return true;
            }
            else if (sender == "express-orders@google.com" && subject.StartsWith("Thanks for your order"))
            {
                if (Program.g_user.is_report_mode_gs(work_mode))
                    return false;
                mail_order = 21;
                return true;
            }
            return false;
        }
        protected override bool parse_mail(MimeMessage mail, int mail_order, KReportBase base_card)
        {
            try
            {

                KReportOP card = base_card as KReportOP;

                if (mail_order == 1)
                    parse_mail_op_1(mail, card);
                else if (mail_order == 2)
                    parse_mail_op_2(mail, card);
                else if (mail_order == 3)
                    parse_mail_op_3(mail, card);
                else if (mail_order == 4)
                    parse_mail_op_4(mail, card);
                else if (mail_order == 5)
                    parse_mail_op_5(mail, card);
                else if (mail_order == 6)
                    parse_mail_op_6(mail, card);
                else if (mail_order == 7)
                    parse_mail_op_7(mail, card);

                else if (mail_order == 9)
                    parse_mail_op_9(mail, card);
                else if (mail_order == 10)
                    parse_mail_op_10(mail, card);
                else if (mail_order == 11)
                    parse_mail_op_11(mail, card);
                else if (mail_order == 12)
                    parse_mail_op_12(mail, card);

                else if (mail_order == 14)
                    parse_mail_op_14(mail, card);
                else if (mail_order == 15)
                    parse_mail_op_15(mail, card);
                else if (mail_order == 16)
                    parse_mail_op_16(mail, card);
                else if (mail_order == 17)
                    parse_mail_op_17(mail, card);
                else if (mail_order == 18)
                    parse_mail_op_18(mail, card);
                else if (mail_order == 19)
                    parse_mail_op_19(mail, card);
                else if (mail_order == 20)
                    parse_mail_op_20(mail, card);
                else if (mail_order == 21)
                    parse_mail_op_21(mail, card);
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
