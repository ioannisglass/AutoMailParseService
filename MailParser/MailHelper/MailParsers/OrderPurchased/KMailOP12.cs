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
        private void parse_mail_op_12(MimeMessage mail, KReportOP report)
        {
            if (!XMailHelper.is_bodytext_existed(mail))
            {
                parse_mail_op_12_for_htmltext(mail, report);
                return;
            }
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_12;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = ConstEnv.RETAILER_LOWES;

            MyLogger.Info($"... OP-12 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-12 m_op_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Order #"))
                {
                    string temp = line.Substring("Order #".Length).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... OP-12 order id = {temp}");
                }
                if (line == "Pickup Item(s)")
                {
                    string next_line = lines[++i].Trim();
                    while (i < lines.Length && next_line.IndexOf("Original Price") == -1)
                    {
                        string title = "";
                        string sku = "";
                        int qty = 0;
                        float price = 0;

                        string temp = next_line; // empty line

                        temp = lines[++i].Trim();
                        title = temp;

                        temp = lines[++i].Trim(); // empty line
                        temp = lines[++i].Trim();
                        if (temp.StartsWith("Item #:"))
                        {
                            temp = temp.Substring("Item #:".Length).Trim();
                            temp = temp.Substring(0, temp.IndexOf("|")).Trim();
                            sku = temp;
                        }

                        temp = lines[++i].Trim(); // empty line
                        temp = lines[++i].Trim();
                        if (temp.StartsWith("Unit Price") && temp.IndexOf("|") != -1)
                        {
                            temp = temp.Substring(0, temp.IndexOf("|")).Trim();
                            temp = temp.Substring("Unit Price".Length);

                            price = Str_Utils.string_to_currency(temp);
                        }

                        temp = lines[++i].Trim(); // empty line
                        temp = lines[++i].Trim();
                        if (temp == "QTY")
                        {
                            temp = lines[++i].Trim(); // empty line
                            temp = lines[++i].Trim();
                            qty = Str_Utils.string_to_int(temp);
                        }

                        if (qty > 0)
                        {
                            ZProduct product = new ZProduct();
                            product.price = price;
                            product.sku = sku;
                            product.title = title;
                            product.qty = qty;
                            report.m_product_items.Add(product);

                            MyLogger.Info($"... OP-12 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                        }
                        next_line = lines[++i].Trim(); // empty line
                        next_line = lines[++i].Trim();
                    }

                    continue;
                }
                if (line.StartsWith("Total Tax"))
                {
                    string temp = line.Substring("Total Tax".Length);
                    float tax = Str_Utils.string_to_currency(temp);
                    report.m_tax = tax;
                    MyLogger.Info($"... OP-12 tax = {tax}");
                    continue;
                }
                if (line.StartsWith("Order Total"))
                {
                    string temp = line.Substring("Order Total".Length);
                    float total = Str_Utils.string_to_currency(temp);
                    report.set_total(total);
                    MyLogger.Info($"... OP-12 total = {total}");

                    temp = lines[++i].Trim();
                    if (temp.StartsWith("Payment"))
                    {
                        string next_line = temp.Substring("Payment".Length).Trim();
                        while (i < lines.Length && next_line != "" && next_line != "Billing Information")
                        {
                            temp = next_line;
                            if (temp.IndexOf("ending in") != -1)
                            {
                                string last_digit = "";
                                float card_price = 0;
                                string payment_type = temp.Substring(0, temp.IndexOf("ending in")).Trim();

                                temp = temp.Substring(temp.IndexOf("ending in") + "ending in".Length).Trim();
                                if (temp.IndexOf(" ") != -1)
                                {
                                    last_digit = temp.Substring(0, temp.IndexOf(" ")).Trim();
                                    temp = temp.Substring(temp.IndexOf(" ")).Trim();
                                    card_price = Str_Utils.string_to_currency(temp);
                                }

                                ZPaymentCard c = new ZPaymentCard(payment_type, last_digit, card_price);
                                report.add_payment_card_info(c);

                                MyLogger.Info($"... OP-12 payment_type = {payment_type}, last_digit = {last_digit}, price = {card_price}");
                            }
                            next_line = lines[++i].Trim();
                        }
                        continue;
                    }
                    continue;
                }
            }
        }
        private void parse_mail_op_12_for_htmltext(MimeMessage mail, KReportOP report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_12;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = ConstEnv.RETAILER_LOWES;

            MyLogger.Info($"... OP-12 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-12 m_op_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Order #:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "ORDER #:")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Order #".Length).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... OP-12 order id = {temp}");
                    continue;
                }
                if (line.StartsWith("Order #"))
                {
                    string temp;
                    if (line.ToUpper() == "ORDER #")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Order #".Length).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... OP-12 order id = {temp}");
                    continue;
                }
                if (line == "Pickup Item(s)" || line == "Shipping Item(s)")
                {
                    string next_line = lines[++i].Trim();
                    while (i < lines.Length)
                    {
                        string title = "";
                        string sku = "";
                        int qty = 0;
                        float price = 0;

                        string temp;

                        title = next_line;

                        temp = lines[++i].Trim();
                        if (temp.StartsWith("Item #:"))
                        {
                            temp = temp.Substring("Item #:".Length).Trim();
                            temp = temp.Substring(0, temp.IndexOf("|")).Trim();
                            sku = temp;
                        }

                        temp = lines[++i].Trim();
                        if (temp.StartsWith("Unit Price"))
                        {
                            if (temp.IndexOf("|") != -1)
                            {
                                temp = temp.Substring(0, temp.IndexOf("|")).Trim();
                                temp = temp.Substring("Unit Price".Length);
                                price = Str_Utils.string_to_currency(temp);
                            }
                            else
                            {
                                temp = temp.Substring("Unit Price".Length);
                                price = Str_Utils.string_to_currency(temp);

                                if (lines[i + 1].Trim().IndexOf("Subtotal") != -1)
                                    i++;
                            }
                        }

                        temp = lines[++i].Trim();
                        if (temp == "QTY")
                        {
                            temp = lines[++i].Trim();
                            qty = Str_Utils.string_to_int(temp);
                        }

                        if (qty > 0)
                        {
                            ZProduct product = new ZProduct();
                            product.price = price;
                            product.sku = sku;
                            product.title = title;
                            product.qty = qty;
                            report.m_product_items.Add(product);

                            MyLogger.Info($"... OP-12 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                        }

                        if (lines[i + 1].IndexOf("Original Price") != -1)
                            break;
                        if (lines[i + 1].StartsWith("$") && lines[i + 2].IndexOf("Original Price") != -1)
                            break;
                        if (lines[i + 1].StartsWith("If you have a question", StringComparison.CurrentCultureIgnoreCase))
                            break;
                        next_line = lines[++i].Trim();
                    }
                    continue;
                }
                if (line.ToUpper() == "PRODUCTS ORDERED" && lines[i + 1].ToUpper().Trim() == "UNIT PRICE" && lines[i + 2].ToUpper().Trim() == "QTY" && lines[i + 3].ToUpper().Trim() == "TOTAL")
                {
                    i += 4;

                    string next_line = lines[i].Trim();
                    while (i < lines.Length && next_line.ToUpper() != "SHIP TO" && next_line.ToUpper() != "DELIVER TO" && next_line.ToUpper() != "PICKUP LOCATION")
                    {
                        string title = "";
                        string sku = "";
                        int qty = 0;
                        float price = 0;

                        string temp = next_line;

                        while (!lines[i + 1].Trim().StartsWith("Item #:", StringComparison.CurrentCultureIgnoreCase))
                        {
                            temp += " " + lines[i + 1].Trim();
                            i++;
                        }
                        title = temp.Trim();

                        temp = lines[++i].Trim();
                        temp = temp.Substring("Item #:".Length).Trim();
                        if (temp == "")
                            temp = lines[++i].Trim();
                        if (temp.IndexOf("|") != -1)
                            temp = temp.Substring(0, temp.IndexOf("|")).Trim();
                        if (temp.IndexOf(" ") != -1)
                            temp = temp.Substring(0, temp.IndexOf(" ")).Trim();
                        if (temp.IndexOf("Model #:") != -1)
                            temp = temp.Substring(0, temp.IndexOf("Model #:")).Trim();
                        sku = temp;

                        temp = lines[++i].Trim();
                        if (temp.IndexOf("Model #:", StringComparison.CurrentCultureIgnoreCase) != -1)
                        {
                            if (temp.ToUpper() == "MODEL #:")
                            {
                                i += 2;
                                temp = lines[i].Trim();
                            }
                            else
                            {
                                temp = lines[++i].Trim();
                            }
                        }
                        price = Str_Utils.string_to_currency(temp);

                        temp = lines[++i].Trim();
                        if (temp == "")
                            temp = lines[++i].Trim();
                        qty = Str_Utils.string_to_int(temp);

                        ++i; // total

                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        report.m_product_items.Add(product);

                        MyLogger.Info($"... OP-12 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");

                        temp = lines[i + 1].Trim();
                        if (temp.ToUpper() == "DELIVER TO" || temp.ToUpper() == "PICKUP LOCATION")
                            break;
                        while (lines[i + 1].Trim()[0] == '$')
                            i += 2;

                        next_line = lines[++i].Trim();
                    }
                    if (next_line.ToUpper() == "DELIVER TO" || next_line.ToUpper() == "PICKUP LOCATION" || next_line.ToUpper() == "SHIP TO")
                    {
                        i--;
                    }
                    continue;
                }
                if (line.StartsWith("Total Tax:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "TOTAL TAX:")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Total Tax".Length).Trim();
                    float tax = Str_Utils.string_to_currency(temp);
                    report.m_tax = tax;
                    MyLogger.Info($"... OP-12 tax = {tax}");
                    continue;
                }
                if (line.StartsWith("Total Tax", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "TOTAL TAX")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Total Tax".Length).Trim();
                    float tax = Str_Utils.string_to_currency(temp);
                    report.m_tax = tax;
                    MyLogger.Info($"... OP-12 tax = {tax}");
                    continue;
                }
                if (line.StartsWith("Order Total:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "ORDER TOTAL:")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Order Total".Length).Trim();
                    float total = Str_Utils.string_to_currency(temp);
                    report.set_total(total);
                    MyLogger.Info($"... OP-12 total = {total}");
                    continue;
                }
                if (line.StartsWith("Order Total", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "ORDER TOTAL")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Order Total".Length).Trim();
                    float total = Str_Utils.string_to_currency(temp);
                    report.set_total(total);
                    MyLogger.Info($"... OP-12 total = {total}");

                    temp = lines[++i].Trim();
                    if (temp.StartsWith("Payment", StringComparison.CurrentCultureIgnoreCase))
                    {
                        string next_line;
                        if (temp.ToUpper() == "PAYMENT")
                            next_line = lines[++i].Trim();
                        else
                            next_line = temp.Substring("Payment".Length).Trim();
                        while (i < lines.Length && next_line != "" && next_line != "Billing Information")
                        {
                            temp = next_line;
                            while (temp.IndexOf("ending in", StringComparison.CurrentCultureIgnoreCase) != -1)
                            {
                                string last_digit = "";
                                float card_price = 0;
                                string payment_type = temp.Substring(0, temp.IndexOf("ending in", StringComparison.CurrentCultureIgnoreCase)).Trim();

                                temp = temp.Substring(temp.IndexOf("ending in", StringComparison.CurrentCultureIgnoreCase) + "ending in".Length).Trim();
                                if (temp.IndexOf(" ") != -1)
                                {
                                    last_digit = temp.Substring(0, temp.IndexOf(" ")).Trim();
                                    temp = temp.Substring(temp.IndexOf(" ")).Trim();
                                    card_price = Str_Utils.string_to_currency(temp);
                                }

                                ZPaymentCard c = new ZPaymentCard(payment_type, last_digit, card_price);
                                report.add_payment_card_info(c);

                                MyLogger.Info($"... OP-12 payment_type = {payment_type}, last_digit = {last_digit}, price = {card_price}");

                                int n = 0;
                                while (n < temp.Length && (temp[n] == '-' || temp[n] == '$' || temp[n] == '.' || char.IsDigit(temp[n])))
                                    n++;
                                if (n == temp.Length)
                                    break;
                                temp = temp.Substring(n);
                            }
                            next_line = lines[++i].Trim();
                        }
                        continue;
                    }
                    continue;
                }
                if (line.ToUpper() == "BILLING SUMMARY")
                {
                    i++;
                    string temp;
                    while (lines[i].Trim().ToUpper() != "SUBTOTAL:")
                    {
                        string last_digit = "";
                        float card_price = 0;
                        string payment_type = "";

                        if (lines[i].Trim().EndsWith(":"))
                        {
                            temp = lines[i].Trim();
                            temp = temp.Substring(0, temp.Length - 1).Trim();
                            if (temp.LastIndexOf(" ") == -1)
                                break;
                            last_digit = temp.Substring(temp.LastIndexOf(" ") + 1).Trim();
                            if (last_digit.Length != 4)
                                break;
                            payment_type = temp.Substring(0, temp.LastIndexOf(" ")).Trim();

                            i++;
                        }
                        else if (lines[i + 2].Trim() == ":")
                        {
                            payment_type = lines[i].Trim();
                            last_digit = lines[i + 1].Trim();
                            i += 3;
                        }
                        else
                        {
                            break;
                        }

                        temp = lines[i].Trim();
                        if (temp[0] != '$')
                            break;
                        card_price = Str_Utils.string_to_currency(temp);

                        ZPaymentCard c = new ZPaymentCard(payment_type, last_digit, card_price);
                        report.add_payment_card_info(c);

                        MyLogger.Info($"... OP-12 payment_type = {payment_type}, last_digit = {last_digit}, price = {card_price}");

                        i++;
                    }
                }
                if (line.ToUpper() == "SHIPPING INFORMATION")
                {
                    string temp = lines[++i].Trim();
                    if (temp.ToUpper() == "ESTIMATED")
                        temp = lines[++i].Trim();
                    if (temp.ToUpper() == "ARRIVAL DATE")
                        temp = lines[++i].Trim();
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
                        MyLogger.Info($"... OP-12 full_address = {full_address}, state_address = {state_address}");
                    }
                    continue;
                }
                if (line.ToUpper() == "ADDRESS" || line.ToUpper() == "PICKUP LOCATION" || line.ToUpper() == "DELIVER TO")
                {
                    string temp = lines[++i].Trim();
                    if (temp.ToUpper() == "DELIVERY FROM")
                        temp = lines[++i].Trim();
                    if (temp.ToUpper() == "ESTIMATED PICKUP DATE")
                        temp = lines[++i].Trim();
                    if (temp.ToUpper() == "DELIVERY DATE")
                        temp = lines[++i].Trim();
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
                        MyLogger.Info($"... OP-12 full_address = {full_address}, state_address = {state_address}");
                    }
                    continue;
                }
            }
        }
    }
}
