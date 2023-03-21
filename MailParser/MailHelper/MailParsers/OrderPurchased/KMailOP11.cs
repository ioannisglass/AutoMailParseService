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
        private void parse_mail_op_11(MimeMessage mail, KReportOP report)
        {
            if (!XMailHelper.is_bodytext_existed(mail))
            {
                parse_mail_op_11_for_htmltext(mail, report);
                return;
            }
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_11;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = ConstEnv.RETAILER_SEARS;

            MyLogger.Info($"... OP-11 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-11 m_op_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("We're processing your order #"))
                {
                    string temp = line.Substring("We're processing your order #".Length).Trim();
                    if (temp.IndexOf("<") != -1)
                        temp = temp.Substring(0, temp.IndexOf("<")).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... OP-11 order id = {temp}");
                    continue;
                }
                if (line.EndsWith("Order Date"))
                {
                    string temp = lines[++i].Trim();
                    DateTime date = DateTime.Parse(temp);
                    report.m_op_purchase_date = date;
                    MyLogger.Info($"... OP-11 Date Of Purchase = {date}");
                    continue;
                }
                if (line == "ITEM DETAILS")
                {
                    string temp = lines[++i].Trim();
                    if (temp.IndexOf("]") != -1)
                        temp = temp.Substring(temp.IndexOf("]") + 1).Trim();
                    string title = temp;

                    string next_line = lines[++i].Trim();
                    while (i < lines.Length && next_line.IndexOf("UPC:") == -1)
                    {
                        next_line = lines[++i].Trim();
                    }
                    string sku = "";
                    if (next_line.IndexOf("UPC:") != -1)
                    {
                        temp = next_line.Substring("UPC:".Length).Trim();
                        sku = temp;
                    }

                    float price = 0;
                    int qty = 0;
                    temp = "";
                    next_line = lines[++i].Trim();
                    while (i < lines.Length)
                    {
                        temp = next_line.Replace(" ", "");
                        if (temp == "PRICEQTYITEMSUBTOTAL")
                            break;
                        next_line = lines[++i].Trim();
                    }
                    if (temp == "PRICEQTYITEMSUBTOTAL")
                    {
                        // first line or item price.

                        temp = lines[++i].Trim();
                        if (temp.StartsWith("Sale"))
                        {
                            temp = temp.Substring("Sale".Length).Trim();
                            price = Str_Utils.string_to_currency(temp);
                            temp = lines[++i].Trim();
                            if (temp.StartsWith("Reg."))
                            {
                                temp = temp.Substring("Reg.".Length).Trim();
                                if (temp.IndexOf(" ") != -1)
                                {
                                    temp = temp.Substring(temp.IndexOf(" ") + 1).Trim();
                                }
                            }
                        }
                        else
                        {
                            if (temp.IndexOf(" ") != -1)
                            {
                                string price_part = temp.Substring(0, temp.IndexOf(" ")).Trim();
                                price = Str_Utils.string_to_currency(price_part);
                                temp = temp.Substring(temp.IndexOf(" ") + 1).Trim();
                            }
                            else
                            {
                                price = Str_Utils.string_to_currency(temp);
                                temp = lines[++i].Trim();
                            }
                        }
                        if (temp.IndexOf(" ") != -1)
                        {
                            string qty_part = temp.Substring(0, temp.IndexOf(" ")).Trim();
                            qty = Str_Utils.string_to_int(qty_part);
                            temp = temp.Substring(temp.IndexOf(" ") + 1).Trim();
                        }
                        else
                        {
                            qty = Str_Utils.string_to_int(temp);
                            temp = lines[++i].Trim();
                        }

                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        report.m_product_items.Add(product);

                        MyLogger.Info($"... OP-11 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                    }
                    continue;
                }
                if (line.StartsWith("Total Due"))
                {
                    string temp = line.Substring("Total Due".Length).Trim();
                    float total = Str_Utils.string_to_currency(temp);
                    report.set_total(total);
                    MyLogger.Info($"... OP-11 order total = {total}");
                    continue;
                }
                if (line.StartsWith("Sales Tax"))
                {
                    string temp = line.Substring("Sales Tax".Length).Trim();
                    float tax = Str_Utils.string_to_currency(temp);
                    report.m_tax = tax;
                    MyLogger.Info($"... OP-11 tax = {tax}");
                    continue;
                }
                if (line == "How You Paid")
                {
                    string next_line = lines[++i].Trim();
                    while (next_line == "")
                    {
                        next_line = lines[++i].Trim();
                    }

                    while (i < lines.Length && next_line.StartsWith("Total Deducted from") && next_line.IndexOf("ending in") != -1 && next_line.IndexOf("**") != -1)
                    {
                        string temp = next_line.Substring("Total Deducted from".Length).Trim();
                        string payment_type = temp.Substring(0, temp.IndexOf("ending in")).Trim();
                        temp = temp.Substring(temp.LastIndexOf("*") + 1);
                        string last_digit = "";
                        float price = 0;
                        if (temp.IndexOf(" ") != -1)
                        {
                            last_digit = temp.Substring(0, temp.IndexOf(" ")).Trim();
                            temp = temp.Substring(temp.IndexOf(" ") + 1).Trim();
                            if (temp[0] == '-')
                                temp = temp.Substring(1);
                            price = Str_Utils.string_to_currency(temp);
                        }
                        ZPaymentCard c = new ZPaymentCard(payment_type, last_digit, price);
                        report.add_payment_card_info(c);
                        MyLogger.Info($"... OP-11 card type = {payment_type}, last_digit = {last_digit}, price = {price}");

                        next_line = lines[++i].Trim();
                    }

                    continue;
                }
            }
        }
        private void parse_mail_op_11_for_htmltext(MimeMessage mail, KReportOP report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_11;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = ConstEnv.RETAILER_SEARS;

            MyLogger.Info($"... OP-11 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-11 m_op_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("We're processing your order #"))
                {
                    string temp = line.Substring("We're processing your order #".Length).Trim();
                    if (temp.IndexOf("<") != -1)
                        temp = temp.Substring(0, temp.IndexOf("<")).Trim();
                    if (temp.IndexOf(",") != -1)
                        temp = temp.Substring(0, temp.IndexOf(",")).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... OP-11 order id = {temp}");
                    continue;
                }
                if (line.EndsWith("Order Date"))
                {
                    string temp = lines[++i].Trim();
                    DateTime date = DateTime.Parse(temp);
                    report.m_op_purchase_date = date;
                    MyLogger.Info($"... OP-11 Date Of Purchase = {date}");
                    continue;
                }
                if (line == "ITEM DETAILS")
                {
                    string title = "";
                    string sku = "";
                    int qty = 0;
                    float price = 0;

                    title = lines[++i].Trim();
                    while (i < lines.Length && !lines[i].Trim().StartsWith("UPC:"))
                        i++;
                    string temp = lines[i].Trim().Substring("UPC:".Length).Trim();
                    sku = temp;

                    while (i < lines.Length && lines[i].Trim() != "ITEM SUBTOTAL")
                        i++;

                    temp = lines[++i].Trim();
                    if (temp.StartsWith("Member Price", StringComparison.CurrentCultureIgnoreCase))
                    {
                        temp = temp.Substring("Member Price".Length).Trim();
                        price = Str_Utils.string_to_currency(temp);

                        temp = lines[++i].Trim();
                        if (temp.StartsWith("Sale"))
                            ++i;
                        temp = lines[++i].Trim();
                        if (temp.StartsWith("Reg."))
                            ++i;
                    }
                    else if (temp.StartsWith("Sale"))
                    {
                        temp = temp.Substring("Sale".Length).Trim();
                        price = Str_Utils.string_to_currency(temp);

                        temp = lines[++i].Trim();
                        if (temp.StartsWith("Reg."))
                            ++i;
                    }
                    else
                    {
                        price = Str_Utils.string_to_currency(temp);
                        temp = lines[++i].Trim();
                    }

                    temp = lines[i].Trim();
                    qty = Str_Utils.string_to_int(temp);

                    ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        report.m_product_items.Add(product);

                        MyLogger.Info($"... OP-11 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                    continue;
                }
                if (line.StartsWith("Total Due"))
                {
                    string temp = line.Substring("Total Due".Length).Trim();
                    if (temp == "")
                        temp = lines[++i].Trim();
                    float total = Str_Utils.string_to_currency(temp);
                    report.set_total(total);
                    MyLogger.Info($"... OP-11 order total = {total}");
                    continue;
                }
                if (line.StartsWith("Sales Tax"))
                {
                    string temp = line.Substring("Sales Tax".Length).Trim();
                    if (temp == "")
                        temp = lines[++i].Trim();
                    float tax = Str_Utils.string_to_currency(temp);
                    report.m_tax = tax;
                    MyLogger.Info($"... OP-11 tax = {tax}");
                    continue;
                }
                if (line == "How You Paid")
                {
                    string next_line = lines[++i].Trim();
                    while (next_line == "")
                    {
                        next_line = lines[++i].Trim();
                    }

                    while (i < lines.Length && next_line.StartsWith("Total Deducted from") && next_line.IndexOf("ending in") != -1 && next_line.IndexOf("**") != -1)
                    {
                        string temp = next_line.Substring("Total Deducted from".Length).Trim();
                        string payment_type = temp.Substring(0, temp.IndexOf("ending in")).Trim();
                        temp = temp.Substring(temp.LastIndexOf("*") + 1).Trim();
                        string last_digit = "";
                        float price = 0;
                        if (temp.IndexOf(" ") != -1)
                        {
                            last_digit = temp.Substring(0, temp.IndexOf(" ")).Trim();
                            temp = temp.Substring(temp.IndexOf(" ") + 1).Trim();
                            if (temp[0] == '-')
                                temp = temp.Substring(1);
                            price = Str_Utils.string_to_currency(temp);
                        }
                        else
                        {
                            last_digit = temp;
                            temp = lines[++i].Trim();
                            if (temp[0] == '-')
                                temp = temp.Substring(1);
                            price = Str_Utils.string_to_currency(temp);
                        }
                        ZPaymentCard c = new ZPaymentCard(payment_type, last_digit, price);
                        report.add_payment_card_info(c);
                        MyLogger.Info($"... OP-11 card type = {payment_type}, last_digit = {last_digit}, price = {price}");

                        next_line = lines[++i].Trim();
                    }

                    continue;
                }
                if (line.ToUpper() == "SHIPPING ADDRESS")
                {
                    string temp = lines[++i].Trim();
                    string full_address = "";
                    string state_address = "";
                    int k = 0;
                    while (k++ < 10)
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
                        MyLogger.Info($"... OP-11 full_address = {full_address}, state_address = {state_address}");
                    }
                    continue;
                }
            }
            if (report.m_order_id == "")
            {
                string htmltext = XMailHelper.get_htmltext(mail);
                if (htmltext.IndexOf("We're processing your order") != -1)
                {
                    string temp = htmltext.Substring(0, htmltext.IndexOf("We're processing your order"));
                    int find = temp.LastIndexOf("<td");
                    temp = htmltext.Substring(find);
                    if (find != -1)
                    {
                        int next_pos;
                        temp = XMailHelper.find_html_part(temp, "td", out next_pos);
                        temp = XMailHelper.html2text(temp);
                        temp = temp.Replace("\n", " ");
                        if (temp.IndexOf("We're processing your order #") != -1)
                        {
                            temp = temp.Substring(temp.IndexOf("We're processing your order #") + "We're processing your order #".Length).Trim();
                            if (temp.IndexOf("placed on") != -1)
                                temp = temp.Substring(0, temp.IndexOf("placed on")).Trim();
                            if (temp.IndexOf(",") != -1)
                                temp = temp.Substring(0, temp.IndexOf(",")).Trim();
                            report.set_order_id(temp);
                            MyLogger.Info($"... OP-11 order id = {temp}");
                        }
                    }
                }
            }
        }
    }
}
