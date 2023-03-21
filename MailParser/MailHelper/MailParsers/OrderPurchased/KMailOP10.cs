using MailParser;
using Logger;
using MimeKit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Utils;

namespace MailHelper
{
    partial class KMailBaseOP : KMailBaseParser
    {
        private void parse_mail_op_10(MimeMessage mail, KReportOP report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_10;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = ConstEnv.RETAILER_CABELAS;

            MyLogger.Info($"... OP-10 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-10 m_op_retailer = {report.m_retailer}");

            string[] lines;

            lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.IndexOf("Order Number:", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    string temp = line.Substring(line.IndexOf("Order Number:", StringComparison.CurrentCultureIgnoreCase) + "Order Number:".Length).Trim();
                    if (temp == "")
                        continue;
                    string order_id = "";
                    if (temp.IndexOf("Order Date:", StringComparison.CurrentCultureIgnoreCase) != -1)
                    {
                        order_id = temp.Substring(0, temp.IndexOf("Order Date:", StringComparison.CurrentCultureIgnoreCase));
                        temp = temp.Substring(temp.IndexOf("Order Date:", StringComparison.CurrentCultureIgnoreCase) + "Order Date:".Length).Trim();
                        var match = Regex.Match(temp, @"([0-9]|1[012])[- /.]([0-9]|[12][0-9]|3[01])[- /.](19|20)\d\d");
                        if (match.Success)
                        {
                            DateTime time = DateTime.Parse(match.Value);

                            report.m_op_purchase_date = time;
                            MyLogger.Info($"... OP-10 order date = {time}");
                        }
                    }
                    else
                    {
                        order_id = temp;
                    }
                    report.set_order_id(order_id);
                    MyLogger.Info($"... OP-10 order id = {order_id}");
                    continue;
                }
                if (line.StartsWith("Order Date:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp = line.Substring("Order Date:".Length).Trim();
                    DateTime time = DateTime.Parse(temp);

                    report.m_op_purchase_date = time;
                    MyLogger.Info($"... OP-10 order date = {time}");
                }

                if (line.ToUpper() == "SHIPPING & HANDLING")
                {
                    float sub_total = 0;
                    float shipping = 0;
                    float tax = 0;
                    float gc_applied = 0;
                    float club_points = 0;
                    float total = 0;

                    string temp = lines[i + 1].Trim();
                    shipping = Str_Utils.string_to_currency(temp);

                    if (lines[i - 2].Trim().ToUpper() == "SUBTOTAL")
                    {
                        temp = lines[i - 1].Trim();
                        sub_total = Str_Utils.string_to_currency(temp);
                    }
                    int k = i + 2;
                    if (lines[k].Trim().StartsWith("Tax"))
                    {
                        temp = lines[k + 1].Trim();
                        tax = Str_Utils.string_to_currency(temp);
                        k += 2;
                    }
                    if (lines[k].Trim().ToUpper() == "GIFT CARDS APPLIED")
                    {
                        temp = lines[k + 1].Trim();
                        gc_applied = Str_Utils.string_to_currency(temp);
                        k += 2;
                    }
                    if (lines[k].Trim().ToUpper() == "CLUB POINTS APPLIED")
                    {
                        temp = lines[k + 1].Trim();
                        club_points = Str_Utils.string_to_currency(temp);
                        k += 2;
                    }
                    if (lines[k].Trim().ToUpper() == "TOTAL")
                    {
                        temp = lines[k + 1].Trim();
                        total = Str_Utils.string_to_currency(temp);
                    }

                    MyLogger.Info($"... OP-10 Subtotal            : {sub_total}");
                    MyLogger.Info($"... OP-10 Shipping & Handling : {shipping}");
                    MyLogger.Info($"... OP-10 Gift Cards Applied  : {gc_applied}");
                    MyLogger.Info($"... OP-10 CLUB Points Applied : {club_points}");
                    MyLogger.Info($"... OP-10 Tax                 : {tax}");
                    MyLogger.Info($"... OP-10 Total               : {total}");

                    if (gc_applied < 0)
                        gc_applied *= -1;

                    total = sub_total + tax + shipping;

                    if (gc_applied != 0)
                    {
                        report.add_payment_card_info(new ZPaymentCard(ConstEnv.GIFT_CARD, "", gc_applied));
                        MyLogger.Info($"... OP-10 Gift card = {gc_applied}");
                    }

                    report.m_tax = tax;
                    report.set_total(total);
                    MyLogger.Info($"... OP-10 tax = {tax}, total = {total}");

                    i = k + 1;
                    continue;
                }

                if (line.ToUpper() == "SHIP TO:")
                {
                    string temp = lines[++i].Trim();
                    string full_address = "";
                    string state_address = "";
                    int k = 0;
                    while (k++ < 5)
                    {
                        full_address += " " + temp;
                        temp = lines[++i].Trim();

                        state_address = XMailHelper.get_address_state_name(full_address);
                        if (state_address != "")
                            break;
                    }
                    full_address = full_address.Trim();
                    state_address = state_address.Trim();
                    if (state_address != "")
                    {
                        report.set_address(full_address, state_address);
                        MyLogger.Info($"... OP-10 full_address = {full_address}, state_address = {state_address}");
                    }
                    continue;
                }
            }

            get_op10_products(mail, report);
            get_op10_payments(mail, report);
        }
        private void get_op10_payments(MimeMessage mail, KReportOP report)
        {
            string htmlbody = XMailHelper.get_htmltext(mail);
            int start_pos = htmlbody.IndexOf("<!-- START PAYMENT DETAILS [order_payment] -->");
            int end_pos = htmlbody.IndexOf("<!--END PAYMENT DETAILS [order_payment]-->");
            if (start_pos == -1 || end_pos == -1 || start_pos > end_pos)
                return;

            start_pos += "<!-- START PAYMENT DETAILS [order_payment] -->".Length;
            string temp = htmlbody.Substring(start_pos, end_pos - start_pos).Trim();
            temp = XMailHelper.html2text(temp);
            temp = temp.Replace("\r\n", " ");
            temp = temp.Replace("\n", " ");
            if (temp.StartsWith("Payment", StringComparison.CurrentCultureIgnoreCase))
                temp = temp.Substring("Payment".Length).Trim();
            if (temp.IndexOf(" ") == -1)
                return;

            string payment_type = temp.Substring(0, temp.IndexOf(" ")).Trim();
            string last_4_digits = temp.Substring(temp.IndexOf(" ") + 1).Trim();

            report.add_payment_card_info(new ZPaymentCard(payment_type, last_4_digits, 0));
            MyLogger.Info($"... OP-10 payment_type = {payment_type}, last_4_digits = {last_4_digits}");
        }
        private void get_op10_products(MimeMessage mail, KReportOP report)
        {
            string htmlbody = XMailHelper.get_htmltext(mail);
            int cancelled_pos = htmlbody.IndexOf("These Item(s) Have Been Canceled");
            int item_pos = htmlbody.IndexOf("Item:");
            string temp = htmlbody;
            int next_pos;
            while (item_pos != -1)
            {
                string title = "";
                string sku = "";
                int qty = 0;
                float price = 0;
                string status = "";

                if (cancelled_pos != -1 && item_pos > cancelled_pos) // it's a canceled item.
                {
                    status = ConstEnv.REPORT_ORDER_STATUS_PARTIAL_CANCELED;
                }

                temp = htmlbody.Substring(0, item_pos);
                int find1 = temp.LastIndexOf("<td");
                if (find1 == -1)
                    break;
                string temp1 = temp.Substring(0, find1);
                if (temp1.LastIndexOf("<td") == -1)
                    break;
                temp1 = temp.Substring(temp1.LastIndexOf("<td"));
                string td_part = XMailHelper.find_html_part(temp1, "td", out next_pos);
                if (td_part == "")
                    break;
                temp1 = XMailHelper.html2text(td_part);
                temp1 = temp1.Replace("\n", " ").Trim();
                title = temp1;

                temp = htmlbody.Substring(find1);
                td_part = XMailHelper.find_html_part(temp, "td", out next_pos);
                if (td_part == "")
                    break;
                temp1 = XMailHelper.html2text(td_part);
                if (temp1.IndexOf("Item:") == -1)
                    break;
                temp1 = temp1.Substring(temp1.IndexOf("Item:") + "Item:".Length).Trim();
                if (temp1.IndexOf("|") != -1)
                    temp1 = temp1.Substring(0, temp1.IndexOf("|")).Trim();
                sku = temp1;

                temp = temp.Substring(next_pos).Trim();
                td_part = XMailHelper.find_html_part(temp, "td", out next_pos);
                if (td_part == "")
                    break;
                temp1 = XMailHelper.html2text(td_part);
                temp1 = temp1.Replace("\n", " ").Trim();
                if (temp1[0] == 'x')
                {
                    temp1 = temp1.Substring(1).Trim();
                    if (temp1.IndexOf("$") != -1)
                    {
                        string qty_part = temp1.Substring(0, temp1.IndexOf("$")).Trim();
                        qty = Str_Utils.string_to_int(qty_part);
                        temp1 = temp1.Substring(temp1.IndexOf("$")).Trim();
                    }
                }
                if (temp1[0] == '$')
                {
                    if (temp1.IndexOf("(") != -1 && temp1.IndexOf(")") != -1)
                    {
                        temp1 = temp1.Substring(temp1.IndexOf("(") + 1).Trim();
                        temp1 = temp1.Substring(0, temp1.IndexOf(")")).Trim();
                        if (temp1.IndexOf("each") != -1)
                            temp1 = temp1.Substring(0, temp1.IndexOf("each")).Trim();
                        price = Str_Utils.string_to_currency(temp1);
                    }
                    else
                    {
                        price = Str_Utils.string_to_currency(temp1);
                    }
                }

                ZProduct product = new ZProduct();
                product.price = price;
                product.sku = sku;
                product.title = title;
                product.qty = qty;
                product.status = status;
                report.m_product_items.Add(product);

                MyLogger.Info($"... OP-10 qty = {qty}, price = {price}, sku = {sku}, item title = {title}, status = {status}");

                item_pos = htmlbody.IndexOf("Item:", item_pos + "Item:".Length);
            }
        }
    }
}
