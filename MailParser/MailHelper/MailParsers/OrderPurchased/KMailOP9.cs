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
        private void parse_mail_op_9(MimeMessage mail, KReportOP report)
        {
            if (!XMailHelper.is_bodytext_existed(mail))
            {
                parse_mail_op_9_for_htmltext(mail, report);
                return;
            }
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_9;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = ConstEnv.RETAILER_WALMART;

            MyLogger.Info($"... OP-9 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-9 m_op_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Order #:"))
                {
                    string temp = line.Substring("Order #:".Length).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... OP-9 order id = {temp}");
                }
                if (line.Replace(" ", "") == "ItemQtyTotal")
                {
                    string next_line = lines[++i].Trim();
                    while (i < lines.Length && (next_line != "" && !next_line.StartsWith("__")))
                    {
                        string title = "";
                        string sku = "";
                        int qty = 0;
                        float price = 0;

                        string temp = next_line;
                        if (temp.IndexOf("<") != -1)
                            temp = temp.Substring(0, temp.IndexOf("<"));
                        title = temp;

                        temp = lines[++i].Trim();

                        string price_part = temp.Substring(0, temp.IndexOf(" "));
                        price = Str_Utils.string_to_currency(price_part);
                        temp = temp.Substring(temp.IndexOf(" ") + 1).Trim();

                        string qty_part = temp.Substring(0, temp.IndexOf(" "));
                        qty = Str_Utils.string_to_int(qty_part);

                        if (qty > 0)
                        {
                            ZProduct product = new ZProduct();
                            product.price = price;
                            product.sku = sku;
                            product.title = title;
                            product.qty = qty;
                            report.m_product_items.Add(product);

                            MyLogger.Info($"... OP-9 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                        }
                        next_line = lines[++i].Trim();
                    }
                    continue;
                }
                if (line.StartsWith("Order total:"))
                {
                    string temp = line.Substring("Order total:".Length);
                    float total = Str_Utils.string_to_currency(temp);
                    report.set_total(total);
                    MyLogger.Info($"... OP-9 total = {total}");
                    continue;
                }
                if (line == "Payment method(s)")
                {
                    string temp = lines[++i].Trim(); // "________________________________"
                    temp = lines[++i].Trim();
                    if (temp.IndexOf("ending in") != -1)
                    {
                        string last_digit = "";
                        float card_price = 0;
                        string payment_type = temp.Substring(0, temp.IndexOf("ending in")).Trim();

                        last_digit = temp.Substring(temp.IndexOf("ending in") + "ending in".Length).Trim();

                        ZPaymentCard c = new ZPaymentCard(payment_type, last_digit, card_price);
                        report.add_payment_card_info(c);

                        MyLogger.Info($"... OP-9 payment_type = {payment_type}, last_digit = {last_digit}, price = {card_price}");
                    }
                    continue;
                }
            }
        }
        private void parse_mail_op_9_for_htmltext(MimeMessage mail, KReportOP report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_9;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = ConstEnv.RETAILER_WALMART;

            MyLogger.Info($"... OP-9 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-9 m_op_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Order #:"))
                {
                    string temp;
                    if (line == "Order #:")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Order #:".Length).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... OP-9 order id = {temp}");
                    continue;
                }
                if (line.StartsWith("Order number:"))
                {
                    string temp;
                    if (line == "Order number:")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Order number:".Length).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... OP-9 order id = {temp}");
                    continue;
                }
                if (line.Replace(" ", "") == "ItemQtyTotal")
                {
                    string next_line = lines[++i].Trim();
                    while (i < lines.Length && (next_line != "" && !next_line.StartsWith("__")))
                    {
                        string title = "";
                        string sku = "";
                        int qty = 0;
                        float price = 0;

                        string temp = next_line;
                        if (temp.IndexOf("<") != -1)
                            temp = temp.Substring(0, temp.IndexOf("<"));
                        title = temp;

                        temp = lines[++i].Trim();

                        string price_part = temp.Substring(0, temp.IndexOf(" "));
                        price = Str_Utils.string_to_currency(price_part);
                        temp = temp.Substring(temp.IndexOf(" ") + 1).Trim();

                        string qty_part = temp.Substring(0, temp.IndexOf(" "));
                        qty = Str_Utils.string_to_int(qty_part);

                        if (qty > 0)
                        {
                            ZProduct product = new ZProduct();
                            product.price = price;
                            product.sku = sku;
                            product.title = title;
                            product.qty = qty;
                            report.m_product_items.Add(product);

                            MyLogger.Info($"... OP-9 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                        }
                        next_line = lines[++i].Trim();
                    }
                    continue;
                }
                if (line == "Item" && lines[i + 1].Trim() == "Qty" && lines[i + 2].Trim() == "Price" && lines[i + 3].Trim() == "Total")
                {
                    i += 4;
                    string next_line = lines[i].Trim();
                    while (i < lines.Length && (next_line != "" && next_line != "Items may arrive in multiple boxes on different days." && next_line != "New! Store pickup made easy"))
                    {
                        string title = "";
                        string sku = "";
                        int qty = 0;
                        float price = 0;
                        int k = i + 2;

                        string temp = "";

                        while (lines[k].Trim()[0] != '$')
                            k++;

                        for (int m = i; m < k - 1; m++)
                            temp += lines[m].Trim() + " ";
                        temp = temp.Trim();
                        title = temp;

                        temp = lines[k - 1].Trim();
                        qty = Str_Utils.string_to_int(temp);

                        temp = lines[k].Trim();
                        price = Str_Utils.string_to_currency(temp);

                        while (lines[k].Trim()[0] == '$')
                            k++;
                        i = k;

                        if (qty > 0)
                        {
                            ZProduct product = new ZProduct();
                            product.price = price;
                            product.sku = sku;
                            product.title = title;
                            product.qty = qty;
                            report.m_product_items.Add(product);

                            MyLogger.Info($"... OP-9 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                        }
                        next_line = lines[i].Trim();
                    }
                    continue;
                }
                if (line == "Item" && lines[i + 1].Trim() == "Qty" && lines[i + 2].Trim() == "Total")
                {
                    i += 3;
                    string next_line = lines[i].Trim();
                    while (i < lines.Length && next_line != "")
                    {
                        if (next_line == "Items may arrive in multiple boxes on different days.")
                            break;
                        if (next_line == "New! Store pickup made easy")
                            break;
                        if (next_line == "Order summary")
                            break;
                        if (next_line == "Order subtotal")
                            break;
                        if (next_line.StartsWith("*This item is covered by the Walmart Technical Support Program."))
                            break;
                        string title = "";
                        string sku = "";
                        int qty = 0;
                        float price = 0;
                        int k = i + 1;

                        string temp = "";

                        while (lines[k].Trim()[0] != '$')
                            k++;

                        for (int m = i; m < k; m++)
                            temp += lines[m].Trim() + " ";
                        temp = temp.Trim();
                        title = temp;

                        temp = lines[k].Trim();
                        if (lines[k + 1].Trim()[0] == '$')
                            temp = lines[++k].Trim();
                        price = Str_Utils.string_to_currency(temp);

                        temp = lines[k + 1].Trim();
                        qty = Str_Utils.string_to_int(temp);

                        k += 2;
                        while (lines[k].Trim()[0] == '$')
                            k++;
                        i = k;

                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        report.m_product_items.Add(product);

                        MyLogger.Info($"... OP-9 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");

                        next_line = lines[i].Trim();
                    }
                    continue;
                }
                if (line.StartsWith("Order total:"))
                {
                    string temp;
                    if (line == "Order total:")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Order total:".Length).Trim();
                    float total = Str_Utils.string_to_currency(temp);
                    report.set_total(total);
                    MyLogger.Info($"... OP-9 total = {total}");
                    continue;
                }
                if (line == "Payment method(s)" && lines[i + 1].Trim().StartsWith("_________"))
                {
                    string temp = lines[++i].Trim(); // "________________________________"
                    temp = lines[++i].Trim();
                    if (temp.IndexOf("ending in") != -1)
                    {
                        string last_digit = "";
                        float card_price = 0;
                        string payment_type = temp.Substring(0, temp.IndexOf("ending in")).Trim();

                        last_digit = temp.Substring(temp.IndexOf("ending in") + "ending in".Length).Trim();

                        ZPaymentCard c = new ZPaymentCard(payment_type, last_digit, card_price);
                        report.add_payment_card_info(c);

                        MyLogger.Info($"... OP-9 payment_type = {payment_type}, last_digit = {last_digit}, price = {card_price}");
                    }
                    continue;
                }
                if (line.EndsWith("card information:"))
                {
                    string temp = lines[++i].Trim();
                    if (temp.IndexOf("#:") != -1)
                    {
                        string last_digit = "";
                        float card_price = 0;
                        string payment_type = temp.Substring(0, temp.IndexOf("#:")).Trim();

                        last_digit = temp.Substring(temp.IndexOf("#:") + "#:".Length).Trim();

                        ZPaymentCard c = new ZPaymentCard(payment_type, last_digit, card_price);
                        report.add_payment_card_info(c);

                        MyLogger.Info($"... OP-9 payment_type = {payment_type}, last_digit = {last_digit}, price = {card_price}");
                    }
                    continue;
                }
                if (line.ToUpper() == "WE'LL SEND AN EMAIL WITH TRACKING INFO WHEN YOUR ORDER SHIPS." && lines[i + 1].Trim().ToUpper() != "SHIPPING TO:")
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
                        MyLogger.Info($"... OP-9 full_address = {full_address}, state_address = {state_address}");
                    }
                    continue;
                }
                if (line.ToUpper() == "SHIPPING TO:")
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
                        MyLogger.Info($"... OP-9 full_address = {full_address}, state_address = {state_address}");
                    }
                    continue;
                }
            }
        }
    }
}
