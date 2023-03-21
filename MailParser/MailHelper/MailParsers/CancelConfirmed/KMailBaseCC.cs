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
    partial class KMailBaseCC : KMailBaseParser
    {
        public KMailBaseCC() : base()
        {
        }
        #region override functions
        public override bool check_valid_mail(int work_mode, string subject, string sender, out int mail_seq_num)
        {
            mail_seq_num = 0;

            if (sender == "bassproshops@t.basspronews.com" && subject == "Order Update - Your Order Has Been Canceled")
            {
                mail_seq_num = 1;
                return true;
            }
            else if ((sender == "orders@oe.target.com" && (subject == "Sorry, we had to cancel." || subject == "Your cancellation is complete." || subject.StartsWith("You've successfully canceled an item from your order ending in") || subject.StartsWith("Sorry, we had to cancel an item from your order ending in")))
                    || (sender == "orders@services.target.com" && subject == "Your order was canceled.")
                    || (sender == "orders@target.com" && subject == "Something from your order has been canceled.")
                    )
            {
                mail_seq_num = 2;
                return true;
            }
            else if (sender == "BestBuyInfo@emailinfo.bestbuy.com")
            {
                if (subject.IndexOf("Your Item Has Been Canceled", StringComparison.CurrentCultureIgnoreCase) != -1
                    || subject.IndexOf("Your order has been canceled", StringComparison.CurrentCultureIgnoreCase) != -1
                    || subject.IndexOf(" Delayed item(s) canceled", StringComparison.CurrentCultureIgnoreCase) != -1
                    || subject.IndexOf("Your Items Have Been Canceled", StringComparison.CurrentCultureIgnoreCase) != -1
                    || (subject.StartsWith("Order ", StringComparison.CurrentCultureIgnoreCase) && subject.EndsWith("Item(s) canceled", StringComparison.CurrentCultureIgnoreCase))
                    )
                {
                    if (Program.g_user.is_report_mode_gs(work_mode))
                        return false;
                    mail_seq_num = 3;
                    return true;
                }
            }
            else if (sender == "fba-noreply@amazon.com" || sender == "order-update@amazon.com" || sender == "payments-messages@amazon.com" || sender == "payments-update@amazon.com" || sender == "qla@amazon.com" || sender == "seller-notification@amazon.com")
            {
                if (subject.StartsWith("Your Basic Fulfillment order was cancelled", StringComparison.CurrentCultureIgnoreCase)
                    || subject.StartsWith("Item canceled for your Amazon.com order", StringComparison.CurrentCultureIgnoreCase)
                    || subject.StartsWith("Item canceled for your AmazonSmile order", StringComparison.CurrentCultureIgnoreCase)
                    || subject.StartsWith("Partial item(s) cancellation from your Amazon.com order", StringComparison.CurrentCultureIgnoreCase)
                    || (subject.StartsWith("Successful cancellation", StringComparison.CurrentCultureIgnoreCase) && subject.IndexOf("from your Amazon.com order", StringComparison.CurrentCultureIgnoreCase) != -1)
                    || (subject.StartsWith("Your Amazon.com order", StringComparison.CurrentCultureIgnoreCase) && subject.IndexOf("has been canceled", StringComparison.CurrentCultureIgnoreCase) != -1)
                    || (subject.StartsWith("Your AmazonSmile order", StringComparison.CurrentCultureIgnoreCase) && subject.IndexOf("has been canceled", StringComparison.CurrentCultureIgnoreCase) != -1)
                    || subject.IndexOf("item has been canceled from your AmazonSmile order", StringComparison.CurrentCultureIgnoreCase) != -1
                    || subject.IndexOf("Your AmazonSmile order has been canceled", StringComparison.CurrentCultureIgnoreCase) != -1
                    || subject.IndexOf("Your Amazon.com Order Has Been Canceled", StringComparison.CurrentCultureIgnoreCase) != -1
                    || subject.IndexOf("Amazon Listing Canceled", StringComparison.CurrentCultureIgnoreCase) != -1
                    || subject.StartsWith("Notice of Cancellation. ID", StringComparison.CurrentCultureIgnoreCase)
                    )
                {
                    if (Program.g_user.is_report_mode_gs(work_mode))
                        return false;
                    mail_seq_num = 4;
                    return true;
                }
            }
            else if (sender.ToUpper() == "HOMEDEPOT@HOMEDEPOT.COM" || sender.ToUpper() == "HOMEDEPOT@ORDER.HOMEDEPOT.COM" || sender.ToUpper() == "HOMEDEPOT@ORDERS.HOMEDEPOT.COM")
            {
                if (subject.StartsWith("Cancellation Confirmation", StringComparison.CurrentCultureIgnoreCase)
                    || subject.StartsWith("Order Cancellation ", StringComparison.CurrentCultureIgnoreCase)
                    || subject.StartsWith("Order Cancelation", StringComparison.CurrentCultureIgnoreCase)
                    || subject.StartsWith("Cancelation Confirmation", StringComparison.CurrentCultureIgnoreCase)
                    || subject.StartsWith("The Home Depot Order Cancellation for ", StringComparison.CurrentCultureIgnoreCase)
                    )
                {
                    if (Program.g_user.is_report_mode_gs(work_mode))
                        return false;
                    mail_seq_num = 5;
                    return true;
                }
            }
            else if (sender == "transaction@samsclub.com" && subject.StartsWith("Your recent order # ", StringComparison.CurrentCultureIgnoreCase) && subject.EndsWith(" has been cancelled", StringComparison.CurrentCultureIgnoreCase))
            {
                if (Program.g_user.is_report_mode_gs(work_mode))
                    return false;
                mail_seq_num = 7;
                return true;
            }
            else if (sender == "help@walmart.com" || sender.EndsWith("@relay.walmart.com", StringComparison.CurrentCultureIgnoreCase))
            {
                if (subject.StartsWith("Message from Walmart.com Customer:  Cancel Order", StringComparison.CurrentCultureIgnoreCase)
                    || subject.StartsWith("Item(s) canceled from your Walmart.com Order", StringComparison.CurrentCultureIgnoreCase)
                    || subject.StartsWith("Item(s) from Your Walmart.com Order Has Been Cancelled", StringComparison.CurrentCultureIgnoreCase)
                    || subject.StartsWith("Items have been canceled from your order", StringComparison.CurrentCultureIgnoreCase)
                    || subject.StartsWith("Your Walmart.com order has been canceled", StringComparison.CurrentCultureIgnoreCase)
                    )
                {
                    if (Program.g_user.is_report_mode_gs(work_mode))
                        return false;
                    mail_seq_num = 9;
                    return true;
                }
            }
            else if (sender == "sears@account.sears.com" || sender == "sears2@value.sears.com")
            {
                if (subject == "We've canceled your order" || subject == "We've canceled your order" || subject == "We've canceled items"
                    || subject.EndsWith("has been canceled", StringComparison.CurrentCultureIgnoreCase) || subject.EndsWith("have been canceled", StringComparison.CurrentCultureIgnoreCase))
                {
                    if (Program.g_user.is_report_mode_gs(work_mode))
                        return false;
                    mail_seq_num = 11;
                    return true;
                }
            }
            return false;
        }
        protected override bool parse_mail(MimeMessage mail, int mail_seq_num, KReportBase base_card)
        {
            try
            {
                KReportCC card = base_card as KReportCC;

                if      (mail_seq_num == 1)  parse_mail_cc_1 (mail, card);
                else if (mail_seq_num == 2)  parse_mail_cc_2 (mail, card);
                else if (mail_seq_num == 3)  parse_mail_cc_3 (mail, card);
                else if (mail_seq_num == 4)  parse_mail_cc_4 (mail, card);
                else if (mail_seq_num == 5)  parse_mail_cc_5 (mail, card);
                else if (mail_seq_num == 7)  parse_mail_cc_7 (mail, card);
                else if (mail_seq_num == 9)  parse_mail_cc_9 (mail, card);
                else if (mail_seq_num == 11) parse_mail_cc_11(mail, card);
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
