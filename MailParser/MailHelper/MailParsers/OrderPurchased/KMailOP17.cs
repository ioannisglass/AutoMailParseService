using MailParser;
using Logger;
using MimeKit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Utils;

namespace MailHelper
{
    partial class KMailBaseOP : KMailBaseParser
    {
        private void parse_mail_op_17(MimeMessage mail, KReportOP report)
        {
            parse_mail_op_17_for_htmltext(mail, report);
        }
        private void parse_mail_op_17_for_htmltext(MimeMessage mail, KReportOP report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_17;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = ConstEnv.RETAILER_OFFICEDEPOT;

            MyLogger.Info($"... OP-17 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-17 m_op_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line == "Order Number:")
                {
                    string temp = lines[++i].Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... OP-17 order id = {temp}");
                    continue;
                }
                if (line == "Order Date:")
                {
                    string temp = lines[++i].Trim();
                    DateTime date = DateTime.Parse(temp);
                    report.m_op_purchase_date = date;
                    MyLogger.Info($"... OP-17 order date = {date}");
                    continue;
                }
                if (line.StartsWith("Payment info:"))
                {
                    string temp = "";
                    if (line == "Payment info:")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Payment info:".Length).Trim();
                    if (temp.IndexOf("last 4 digits:") != -1)
                    {
                        string temp1 = temp.Substring(0, temp.IndexOf("last 4 digits:")).Trim();
                        if (temp1.EndsWith(","))
                            temp1 = temp1.Substring(0, temp1.Length - 1);
                        string payment_type = temp1;
                        temp = temp.Substring(temp.IndexOf("last 4 digits:") + "last 4 digits:".Length).Trim();
                        string last_digit = temp;
                        float price = 0;
                        ZPaymentCard c = new ZPaymentCard(payment_type, last_digit, price);
                        report.add_payment_card_info(c);
                        MyLogger.Info($"... OP-17 card type = {payment_type}, last_digit = {last_digit}, price = {price}");
                    }
                    else if (temp.StartsWith("Amt. applied to"))
                    {
                        string payment_type = "";
                        float price = 0;
                        string last_digit = "";

                        temp = temp.Substring("Amt. applied to".Length).Trim();
                        if (temp.IndexOf(":") != -1)
                        {
                            payment_type = temp.Substring(0, temp.IndexOf(":")).Trim();
                            temp = temp.Substring(temp.IndexOf(":") + 1).Trim();
                            price = Str_Utils.string_to_currency(temp);
                        }
                        else
                        {
                            payment_type = temp;
                        }
                        ZPaymentCard c = new ZPaymentCard(payment_type, last_digit, price);
                        report.add_payment_card_info(c);
                        MyLogger.Info($"... OP-17 card type = {payment_type}, last_digit = {last_digit}, price = {price}");
                    }
                    else
                    {
                        string payment_type = "";
                        float price = 0;
                        string last_digit = "";

                        if (temp.EndsWith("."))
                            temp = temp.Substring(0, temp.Length - 1);
                        payment_type = temp;
                        ZPaymentCard c = new ZPaymentCard(payment_type, last_digit, price);
                        report.add_payment_card_info(c);
                        MyLogger.Info($"... OP-17 card type = {payment_type}, last_digit = {last_digit}, price = {price}");
                    }
                    continue;
                }
                if (line.ToUpper() == "SUBTOTAL:")
                {
                    float sub_total = 0;
                    float shipping = 0;
                    float tax = 0;
                    float gc_applied = 0;
                    float misc = 0;
                    float total = 0;

                    string temp = lines[i + 1].Trim();
                    sub_total = Str_Utils.string_to_currency(temp);

                    int k = i + 2;
                    if (lines[k].Trim().ToUpper() == "TAX:")
                    {
                        temp = lines[k + 1].Trim();
                        tax = Str_Utils.string_to_currency(temp);
                        k += 2;
                    }
                    if (lines[k].Trim().ToUpper() == "DELIVERY FEE:" || lines[k].Trim().ToUpper() == "DELIVERY CHARGE:")
                    {
                        temp = lines[k + 1].Trim();
                        shipping = Str_Utils.string_to_currency(temp);
                        k += 2;
                    }
                    if (lines[k].Trim().ToUpper() == "MISC.:")
                    {
                        temp = lines[k + 1].Trim();
                        misc = Str_Utils.string_to_currency(temp);
                        k += 2;
                    }
                    if (lines[k].Trim().ToUpper() == "GIFT/REWARD CARD")
                    {
                        temp = lines[k + 1].Trim();
                        if (temp.StartsWith("(") && temp.EndsWith(")"))
                            temp = temp.Substring(1, temp.Length - 2).Trim();
                        gc_applied = Str_Utils.string_to_currency(temp);
                        k += 2;
                    }
                    if (lines[k].Trim().ToUpper() == "TOTAL:")
                    {
                        temp = lines[k + 1].Trim();
                        total = Str_Utils.string_to_currency(temp);
                    }

                    MyLogger.Info($"... OP-10 Subtotal            : {sub_total}");
                    MyLogger.Info($"... OP-10 Delivery Fee        : {shipping}");
                    MyLogger.Info($"... OP-10 Gift/Reward Card    : {gc_applied}");
                    MyLogger.Info($"... OP-10 Misc                : {misc}");
                    MyLogger.Info($"... OP-10 Tax                 : {tax}");
                    MyLogger.Info($"... OP-10 Total               : {total}");

                    if (gc_applied < 0)
                        gc_applied *= -1;

                    total = sub_total + tax + shipping + misc;

                    report.m_tax = tax;
                    report.set_total(total);
                    MyLogger.Info($"... OP-17 tax = {tax}, total = {total}");

                    i = k + 1;
                    continue;
                }
                if (line.ToUpper() == "STATUS:")
                {
                    string temp = lines[++i].Trim();
                    MyLogger.Info($"... OP-17 status = {temp}");

                    if (temp.ToUpper() == "CANCELLED" || temp.ToUpper() == "DELETED BY CUSTOMER" || temp.ToUpper() == "DELETED-INVENTORY REASONS")
                    {
                        report.m_order_status = ConstEnv.REPORT_ORDER_STATUS_CANCELED;
                    }
                    if (temp.ToUpper() == "SHIPPED")
                    {
                        //report.m_order_status = ConstEnv.REPORT_ORDER_STATUS_SHIPPED;
                    }
                }
                if (line.ToUpper() == "SHIPPING TO:")
                {
                    string full_address = "";
                    string state_address = "";
                    int k = i + 1;
                    while (k < i + 6)
                    {
                        full_address += " " + lines[k].Trim();
                        state_address = XMailHelper.get_address_state_name(full_address);
                        if (state_address != "")
                            break;
                        k++;
                    }
                    full_address = full_address.Trim();
                    state_address = state_address.Trim();
                    if (state_address != "")
                    {
                        report.set_address(full_address, state_address);
                        MyLogger.Info($"... OP-17 full_address = {full_address}, state_address = {state_address}");
                    }
                    i = k;
                    continue;
                }
            }

            string html_text = mail.HtmlBody;
            if (html_text != null && (html_text.IndexOf("EXTENDED PRICE</font>") != -1 || html_text.IndexOf("Qty</font>") != -1))
            {

                int next_pos;
                string temp;

                // parse header

                List<string> headers = new List<string>();
                int find;
                if (html_text.IndexOf("EXTENDED PRICE</font>") != -1)
                    find = html_text.IndexOf("EXTENDED PRICE</font>");
                else
                    find = html_text.IndexOf("Qty</font>");
                temp = html_text.Substring(0, find);
                find = temp.LastIndexOf("<tr");
                if (find == -1)
                    return;
                temp = html_text.Substring(find);
                string hdr_tr = XMailHelper.find_html_part(temp, "tr", out next_pos);
                if (hdr_tr == "")
                    return;
                temp = temp.Substring(next_pos).Trim();

                while (hdr_tr != "")
                {
                    string td_part = XMailHelper.find_html_part(hdr_tr, "td", out next_pos);
                    if (td_part == "")
                        break;
                    hdr_tr = hdr_tr.Substring(next_pos);

                    string temp1 = XMailHelper.html2text(td_part);
                    temp1 = temp1.Replace("\n", " ");
                    headers.Add(temp1);
                }
                if (headers.Count == 0)
                    return;

                // get product part.

                if (temp.StartsWith("\r\n"))
                    temp = temp.Substring(2);
                if (temp.StartsWith("<tbody"))
                {
                    string tbody_part = XMailHelper.find_html_part(temp, "tbody", out next_pos);
                    temp = tbody_part;
                }
                else
                {
                    temp = XMailHelper.find_html_part(temp, "tr", out next_pos);
                    if (temp == "")
                        return;
                    temp = "<tr>\r\n" + temp + "\r\n</tr>";
                }

                // get item information

                while (temp != "")
                {
                    string tr_part = XMailHelper.find_html_part(temp, "tr", out next_pos);
                    if (tr_part == "")
                        break;
                    temp = temp.Substring(next_pos);

                    string title = "";
                    string sku = "";
                    float price = 0;
                    int qty = 0;

                    for (int i = 0; i < headers.Count; i++)
                    {
                        string hdr = headers[i];

                        string td_part = XMailHelper.find_html_part(tr_part, "td", out next_pos);
                        if (td_part == "")
                            break;
                        tr_part = tr_part.Substring(next_pos);

                        if (hdr.ToUpper() == "DESCRIPTION" || hdr.ToUpper() == "ITEM DESCRIPTION")
                        {
                            string temp1 = XMailHelper.html2text(td_part);
                            if (temp1.IndexOf("(") != -1)
                            {
                                title = temp1.Substring(0, temp1.LastIndexOf("(")).Trim();
                                temp1 = temp1.Substring(temp1.LastIndexOf("(") + 1).Trim();
                                if (temp1.IndexOf(")") != -1)
                                    sku = temp1.Substring(0, temp1.IndexOf(")")).Trim();
                            }
                            else
                            {
                                title = temp1;
                            }
                            continue;
                        }
                        if (hdr.ToUpper() == "QTY")
                        {
                            string temp1 = XMailHelper.html2text(td_part);
                            qty = Str_Utils.string_to_int(temp1);
                            continue;
                        }
                        if (hdr.ToUpper() == "UNIT PRICE")
                        {
                            string temp1 = XMailHelper.html2text(td_part);
                            price = Str_Utils.string_to_currency(temp1);
                            continue;
                        }
                        if (hdr.ToUpper() == "ITEM NUMBER")
                        {
                            string temp1 = XMailHelper.html2text(td_part);
                            sku = temp1;
                            continue;
                        }
                    }

                    if (qty > 0)
                    {
                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        report.m_product_items.Add(product);

                        MyLogger.Info($"... OP-17 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                    }
                }
            }
        }
    }
}
