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
        private void extract_op1_item_info_from_html(MimeMessage mail, string sku, out string title, out int qty, out float price)
        {
            title = "";
            qty = 0;
            price = 0;

            string htmlbody = XMailHelper.get_htmltext(mail);
            if (htmlbody.IndexOf($">{sku}</font></td>") == -1)
                return;
            string temp = htmlbody.Substring(htmlbody.IndexOf($">{sku}</font></td>") + $">{sku}</font></td>".Length).Trim();

            if (temp.IndexOf("</font></td>") == -1)
                return;

            string temp1 = temp.Substring(0, temp.IndexOf("</font></td>"));
            if (temp1.LastIndexOf(">") != -1)
            {
                temp1 = temp1.Substring(temp1.LastIndexOf(">") + 1);
                qty = Str_Utils.string_to_int(temp1);
            }
            temp = temp.Substring(temp.IndexOf("</font></td>") + "</font></td>".Length).Trim();

            temp1 = temp.Substring(0, temp.IndexOf("</font></td>"));
            if (temp1.LastIndexOf(">") != -1)
            {
                temp1 = temp1.Substring(temp1.LastIndexOf(">") + 1);
                title = temp1;
            }
            temp = temp.Substring(temp.IndexOf("</font></td>") + "</font></td>".Length).Trim();

            temp1 = temp.Substring(0, temp.IndexOf("</font></td>"));
            if (temp1.LastIndexOf(">") != -1)
            {
                temp1 = temp1.Substring(temp1.LastIndexOf(">") + 1);
                price = Str_Utils.string_to_currency(temp1);
            }
            temp = temp.Substring(temp.IndexOf("</font></td>") + "</font></td>".Length).Trim();
        }
        private void parse_mail_op_1(MimeMessage mail, KReportOP report)
        {
            bool get_items = false;
            bool get_tax = false;
            bool get_total = false;
            bool get_purchase_date = false;
            bool get_address = false;
            bool get_payment = false;

            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_1;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = ConstEnv.RETAILER_BASS;

            MyLogger.Info($"... OP-1 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-1 m_op_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line == "Qty" && i < lines.Length - 4 && lines[i + 1] == "Description" && lines[i + 2] == "Unit Cost" && lines[i + 3] == "Status")
                {
                    i += 4;
                    line = lines[i].Trim();

                    if (line.IndexOf("[") != -1 && line.IndexOf("]") != -1)
                    {
                        string temp = line;

                        temp = temp.Substring(temp.IndexOf("]") + 1);
                        temp = temp.Trim();

                        string sku = temp.Substring(0, temp.IndexOf(" "));

                        string title;
                        int qty;
                        float price;
                        extract_op1_item_info_from_html(mail, sku, out title, out qty, out price);

                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        report.m_product_items.Add(product);
                        MyLogger.Info($"... OP-1 sku = {sku}, title = {title}, qty = {qty}, price = {price}");
                        get_items = true;
                    }
                    continue;
                }
                if (line.StartsWith("Item #:"))
                {
                    string temp = "";
                    if (line == "Item #:")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring(line.IndexOf("Item #:") + "Item #:".Length).Trim();

                    string sku = temp;
                    string title;
                    int qty;
                    float price;
                    extract_op1_item_info_from_html(mail, sku, out title, out qty, out price);

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    report.m_product_items.Add(product);

                    MyLogger.Info($"... OP-1 sku = {sku}, title = {title}, qty = {qty}, price = {price}");
                    get_items = true;
                    continue;
                }

                if (line.IndexOf("Order #:", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    string temp;
                    if (line == "Order #:")
                    {
                        if (lines[i + 1].IndexOf("Descriptive", StringComparison.CurrentCultureIgnoreCase) != -1)
                            continue;
                        temp = lines[++i].Trim();
                    }
                    else
                    {
                        temp = line.Substring(line.IndexOf("Order #:") + "Order #:".Length).Trim();
                    }
                    if (temp.IndexOf("]") != -1)
                    {
                        temp = temp.Substring(temp.IndexOf("]") + 1).Trim();
                        if (temp.IndexOf("[") != -1)
                        {
                            temp = temp.Substring(0, temp.IndexOf("]")).Trim();
                            report.set_order_id(temp);
                        }
                    }
                    else
                    {
                        report.set_order_id(temp);
                    }
                    MyLogger.Info($"... OP-1 order id = {temp}");
                    continue;
                }
                if (line.IndexOf("Order Number:") != -1 || line.IndexOf("Order number:") != -1)
                {
                    string temp;
                    string order_id = "";

                    if (line == "Order Number:" || line == "Order number:")
                        temp = lines[++i].Trim();
                    else if (line.IndexOf("Order Number:") != -1)
                        temp = line.Substring(line.IndexOf("Order Number:") + "Order Number:".Length).Trim();
                    else
                        temp = line.Substring(line.IndexOf("Order number:") + "Order number:".Length).Trim();
                    if (temp.IndexOf("]") != -1)
                    {
                        temp = temp.Substring(temp.IndexOf("]") + 1).Trim();
                        if (temp.IndexOf("[") != -1)
                        {
                            temp = temp.Substring(0, temp.IndexOf("]")).Trim();
                        }
                    }
                    if (temp.IndexOf("Order Date:", StringComparison.CurrentCultureIgnoreCase) != -1)
                    {
                        order_id = temp.Substring(0, temp.IndexOf("Order Date:", StringComparison.CurrentCultureIgnoreCase)).Trim();
                        temp = temp.Substring(temp.IndexOf("Order Date:", StringComparison.CurrentCultureIgnoreCase) + "Order Date:".Length).Trim();
                        DateTime date = DateTime.Parse(temp);
                        report.m_op_purchase_date = date;
                        MyLogger.Info($"... OP-1 Date Of Purchase = {date}");
                    }
                    else
                    {
                        order_id = temp;
                    }
                    report.set_order_id(order_id);
                    MyLogger.Info($"... OP-1 order id = {order_id}");
                    continue;
                }
                if (line.IndexOf("Order Date:") != -1)
                {
                    string temp;
                    if (line == "Order Date:")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring(line.IndexOf("Order Date:") + "Order Date:".Length).Trim();
                    if (temp.IndexOf("]") != -1)
                        temp = temp.Substring(temp.IndexOf("]") + 1).Trim();
                    DateTime date = DateTime.Parse(temp);
                    report.m_op_purchase_date = date;
                    MyLogger.Info($"... OP-1 Date Of Purchase = {date}");
                    get_purchase_date = true;

                    continue;
                }

                if (line.ToUpper() == "MERCHANDISE TOTAL:" &&
                    lines[i + 1].ToUpper() == "GIFT CARD PURCHASED:" &&
                    lines[i + 2].ToUpper() == "SHIPPING:" &&
                    lines[i + 3].ToUpper() == "TAX:" &&
                    lines[i + 4].ToUpper() == "CREDITS/DISCOUNTS:" &&
                    (lines[i + 5].ToUpper() == "TOTAL:" || lines[i + 5].ToUpper() == "TOTAL TO BE CHARGED:")
                    )
                {
                    i += 6;

                    string temp = lines[i].Trim();
                    float merchandise_total = Str_Utils.string_to_currency(temp);

                    temp = lines[i + 1].Trim();
                    float gc_purchased = Str_Utils.string_to_currency(temp);

                    temp = lines[i + 2].Trim();
                    float shipping_price = Str_Utils.string_to_currency(temp);

                    temp = lines[i + 3].Trim();
                    float tax = Str_Utils.string_to_currency(temp);

                    temp = lines[i + 4].Trim();
                    float credit = Str_Utils.string_to_currency(temp);

                    temp = lines[i + 5].Trim();
                    float total = Str_Utils.string_to_currency(temp);

                    MyLogger.Info($"... OP-1 Merchandise Total   : {merchandise_total}");
                    MyLogger.Info($"... OP-1 Gift Card Purchased : {gc_purchased}");
                    MyLogger.Info($"... OP-1 Shipping            : {shipping_price}");
                    MyLogger.Info($"... OP-1 Tax                 : {tax}");
                    MyLogger.Info($"... OP-1 Credits/Discounts   : {credit}");
                    MyLogger.Info($"... OP-1 Total               : {total}");

                    if (gc_purchased < 0)
                        gc_purchased *= -1;
                    if (credit < 0)
                        credit *= -1;

                    if (credit != 0)
                    {
                        report.add_payment_card_info(new ZPaymentCard(ConstEnv.CREDIT_CARD, "", credit));
                        MyLogger.Info($"... OP-1 Credit card = {credit}");
                        get_payment = true;
                    }

                    total = merchandise_total + tax + shipping_price;
                    if (total == 0)
                        total = credit;

                    report.m_tax = tax;
                    report.set_total(total);
                    MyLogger.Info($"... OP-1 tax = {tax}, total = {total}");
                    get_tax = true;
                    get_total = true;
                    i += 5;
                    continue;
                }
                if (line.StartsWith("Merchandise Total:", StringComparison.CurrentCultureIgnoreCase) &&
                    lines[i + 1].Trim().StartsWith("Gift Card Purchased:", StringComparison.CurrentCultureIgnoreCase) &&
                    lines[i + 2].Trim().StartsWith("Shipping:", StringComparison.CurrentCultureIgnoreCase) &&
                    lines[i + 3].Trim().StartsWith("Tax:", StringComparison.CurrentCultureIgnoreCase) &&
                    (lines[i + 4].Trim().StartsWith("Credits/Discounts:", StringComparison.CurrentCultureIgnoreCase) || lines[i + 4].Trim().StartsWith("Adjustments**:", StringComparison.CurrentCultureIgnoreCase)) &&
                    (lines[i + 5].Trim().StartsWith("Total:", StringComparison.CurrentCultureIgnoreCase) || lines[i + 5].Trim().StartsWith("Total to be charged:", StringComparison.CurrentCultureIgnoreCase))
                    )
                {
                    string temp = line.Substring("Merchandise Total:".Length).Trim();
                    float merchandise_total = Str_Utils.string_to_currency(temp);

                    temp = lines[i + 1].Trim().Substring("Gift Card Purchased:".Length).Trim();
                    float gc_purchased = Str_Utils.string_to_currency(temp);

                    temp = lines[i + 2].Trim().Substring("Shipping:".Length).Trim();
                    float shipping_price = Str_Utils.string_to_currency(temp);

                    temp = lines[i + 3].Trim().Substring("Tax:".Length).Trim();
                    float tax = Str_Utils.string_to_currency(temp);

                    if (lines[i + 4].Trim().StartsWith("Credits/Discounts:", StringComparison.CurrentCultureIgnoreCase))
                        temp = lines[i + 4].Trim().Substring("Credits/Discounts:".Length).Trim();
                    else
                        temp = lines[i + 4].Trim().Substring("Adjustments**:".Length).Trim();
                    float credit = Str_Utils.string_to_currency(temp);

                    if (lines[i + 5].Trim().StartsWith("Total:", StringComparison.CurrentCultureIgnoreCase))
                        temp = lines[i + 5].Trim().Substring("Total:".Length).Trim();
                    else
                        temp = lines[i + 5].Trim().Substring("Total to be charged:".Length).Trim();
                    float total = Str_Utils.string_to_currency(temp);

                    MyLogger.Info($"... OP-1 Merchandise Total   : {merchandise_total}");
                    MyLogger.Info($"... OP-1 Gift Card Purchased : {gc_purchased}");
                    MyLogger.Info($"... OP-1 Shipping            : {shipping_price}");
                    MyLogger.Info($"... OP-1 Tax                 : {tax}");
                    MyLogger.Info($"... OP-1 Credits/Discounts   : {credit}");
                    MyLogger.Info($"... OP-1 Total               : {total}");

                    if (gc_purchased < 0)
                        gc_purchased *= -1;
                    if (credit < 0)
                        credit *= -1;

                    if (credit != 0)
                    {
                        report.add_payment_card_info(new ZPaymentCard(ConstEnv.CREDIT_CARD, "", credit));
                        MyLogger.Info($"... OP-1 Credit card = {credit}");
                        get_payment = true;
                    }

                    total = merchandise_total + tax + shipping_price;
                    if (total == 0)
                        total = credit;

                    report.m_tax = tax;
                    report.set_total(total);
                    MyLogger.Info($"... OP-1 tax = {tax}, total = {total}");
                    get_tax = true;
                    get_total = true;
                    i += 5;
                    continue;
                }
                if (line.ToUpper() == "MERCHANDISE TOTAL:" &&
                    lines[i + 2].ToUpper() == "GIFT CARD PURCHASED:" &&
                    lines[i + 4].ToUpper() == "SHIPPING:" &&
                    lines[i + 6].ToUpper() == "TAX:" &&
                    (lines[i + 8].ToUpper() == "CREDITS/DISCOUNTS:" || lines[i + 8].ToUpper() == "CREDITS:") &&
                    (lines[i + 10].ToUpper() == "TOTAL:" || lines[i + 10].ToUpper() == "TOTAL TO BE CHARGED:")
                    )
                {
                    string temp = lines[i + 1].Trim();
                    float merchandise_total = Str_Utils.string_to_currency(temp);

                    temp = lines[i + 3].Trim();
                    float gc_purchased = Str_Utils.string_to_currency(temp);

                    temp = lines[i + 5].Trim();
                    float shipping_price = Str_Utils.string_to_currency(temp);

                    temp = lines[i + 7].Trim();
                    float tax = Str_Utils.string_to_currency(temp);

                    temp = lines[i + 9].Trim();
                    float credit = Str_Utils.string_to_currency(temp);

                    temp = lines[i + 11].Trim();
                    float total = Str_Utils.string_to_currency(temp);

                    MyLogger.Info($"... OP-1 Merchandise Total   : {merchandise_total}");
                    MyLogger.Info($"... OP-1 Gift Card Purchased : {gc_purchased}");
                    MyLogger.Info($"... OP-1 Shipping            : {shipping_price}");
                    MyLogger.Info($"... OP-1 Tax                 : {tax}");
                    MyLogger.Info($"... OP-1 Credits/Discounts   : {credit}");
                    MyLogger.Info($"... OP-1 Total               : {total}");

                    if (gc_purchased < 0)
                        gc_purchased *= -1;
                    if (credit < 0)
                        credit *= -1;

                    if (credit != 0)
                    {
                        report.add_payment_card_info(new ZPaymentCard(ConstEnv.CREDIT_CARD, "", credit));
                        MyLogger.Info($"... OP-1 Credit card = {credit}");
                        get_payment = true;
                    }

                    total = merchandise_total + tax + shipping_price;
                    if (total == 0)
                        total = credit;

                    report.m_tax = tax;
                    report.set_total(total);
                    MyLogger.Info($"... OP-1 tax = {tax}, total = {total}");
                    get_tax = true;
                    get_total = true;
                    i += 11;
                    continue;
                }
            }

            lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n'); // Add Nov 18, 2019 for new mails
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (!get_tax && line.ToUpper() == "TAX")
                {
                    string temp = lines[++i].Trim();
                    float tax = Str_Utils.string_to_currency(temp);
                    report.m_tax = tax;
                    get_tax = true;
                    MyLogger.Info($"... OP-1 tax = {tax}");
                    continue;
                }

                if (!get_total && line.ToUpper() == "TOTAL")
                {
                    string temp = lines[++i].Trim();
                    float total = Str_Utils.string_to_currency(temp);
                    report.set_total(total);
                    MyLogger.Info($"... OP-1 total = {total}");
                    get_total = true;
                    continue;
                }

                if (line.IndexOf("Ending in", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    string temp = line.Substring(0, line.IndexOf("Ending in", StringComparison.CurrentCultureIgnoreCase)).Trim();
                    string payment_type = temp;

                    temp = line.Substring(line.IndexOf("Ending in", StringComparison.CurrentCultureIgnoreCase) + "Ending in".Length).Trim();
                    string last_4_digit = temp;

                    temp = lines[++i].Trim();
                    float price = Str_Utils.string_to_currency(temp);

                    report.add_payment_card_info(new ZPaymentCard(payment_type, last_4_digit, price));

                    MyLogger.Info($"... OP-1 payment_type = {payment_type}, last_4_digit = {last_4_digit}, price = {price}");
                    continue;
                }

                if (!get_items && line.IndexOf("SKU:", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    string title = "";
                    string sku = "";
                    int qty = 0;
                    float price = 0;

                    int k = i - 1;
                    string temp = "";

                    while (lines[k].Trim().ToUpper() != "IN PROCESS" && !lines[k].Trim().StartsWith("$"))
                    {
                        temp = lines[k].Trim() + " " + temp;
                        k--;
                    }
                    temp = temp.Trim();
                    if (!line.StartsWith("SKU:", StringComparison.CurrentCultureIgnoreCase))
                        temp = temp + " " + line.Substring(0, line.IndexOf("SKU:", StringComparison.CurrentCultureIgnoreCase));
                    title = temp;

                    if (line.StartsWith("SKU:", StringComparison.CurrentCultureIgnoreCase) && line.ToUpper() != "SKU:")
                    {
                        sku = line.Substring("SKU:".Length).Trim();
                    }
                    else if (line.EndsWith("SKU:", StringComparison.CurrentCultureIgnoreCase))
                    {
                        sku = lines[++i].Trim();
                    }

                    i++;
                    if (lines[i].Trim().ToUpper() == "QUANTITY" && lines[i + 1].Trim().ToUpper() == "PRICE")
                    {
                        i += 2;
                        temp = lines[i].Trim();
                        if (temp.IndexOf("$") == -1)
                            temp += lines[++i].Trim();

                        string qty_part = temp.Substring(0, temp.IndexOf("$")).Trim();
                        qty = Str_Utils.string_to_int(qty_part);

                        temp = temp.Substring(temp.IndexOf("$") + 1).Trim();
                        price = Str_Utils.string_to_currency(qty_part);

                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        report.m_product_items.Add(product);
                        MyLogger.Info($"... OP-1 sku = {sku}, title = {title}, qty = {qty}, price = {price}");
                        get_items = true;
                    }
                    continue;
                }

                if (line.StartsWith("Shipping Address", StringComparison.CurrentCultureIgnoreCase))
                {
                    string full_address = "";

                    i++;
                    while (!lines[i].Trim().StartsWith("____") && !lines[i].Trim().StartsWith("SKU") && !lines[i].Trim().StartsWith("Billing Address", StringComparison.CurrentCultureIgnoreCase))
                    {
                        if (lines[i].Trim() != "")
                        {
                            full_address += " " + lines[i].Trim();
                        }
                        i++;
                    }
                    full_address = full_address.Trim();
                    string state_address = XMailHelper.get_address_state_name(full_address);
                    if (state_address != "")
                    {
                        report.set_address(full_address, state_address);
                        MyLogger.Info($"... OP-1 full_address = {full_address}, state_address = {state_address}");
                        get_address = true;
                    }
                }
            }

            string htmlbody = XMailHelper.get_htmltext(mail);
            if (!get_purchase_date)
            {
                if (htmlbody.IndexOf("Order Date:</b></font></td>") != -1)
                {
                    string temp = htmlbody.Substring(htmlbody.IndexOf("Order Date:</b></font></td>") + "Order Date:</b></font></td>".Length).Trim();
                    int next_pos;
                    temp = XMailHelper.find_html_part(temp, "td", out next_pos);
                    temp = XMailHelper.html2text(temp);
                    DateTime date = DateTime.Parse(temp);
                    report.m_op_purchase_date = date;
                    MyLogger.Info($"... OP-1 Date Of Purchase = {date}");
                }
            }
            if (!get_payment)
            {
                if (htmlbody.IndexOf("Credits:") != -1)
                {
                    string temp = htmlbody.Substring(htmlbody.IndexOf("Credits:"));
                    if (temp.IndexOf("</td>") != -1)
                    {
                        temp = temp.Substring(temp.IndexOf("</td>") + "</td>".Length);
                        int next_pos;
                        temp = XMailHelper.find_html_part(temp, "td", out next_pos);
                        temp = XMailHelper.html2text(temp);
                        float credit = Str_Utils.string_to_currency(temp);
                        if (credit != 0)
                        {
                            report.add_payment_card_info(new ZPaymentCard(ConstEnv.CREDIT_CARD, "", credit));
                            MyLogger.Info($"... OP-1 Credit card price = {credit}");
                            get_payment = true;
                        }
                        MyLogger.Info($"... OP-1 credit = {credit}");
                    }
                }

            }
            if (!get_items)
            {
                if (htmlbody.IndexOf("Qty") != -1)
                {
                    string temp = htmlbody.Substring(htmlbody.IndexOf("Qty"));
                    if (temp.IndexOf("</table>") != -1)
                    {
                        temp = temp.Substring(temp.IndexOf("</table>") + "</table>".Length);

                        int next_pos;
                        temp = XMailHelper.find_html_part(temp, "tbody", out next_pos);

                        while (temp.IndexOf("</font></td>") != -1)
                        {
                            string sku;
                            string title;
                            int qty;
                            float price;

                            string temp1 = temp.Substring(0, temp.IndexOf("</font></td>")).Trim();
                            if (temp1.LastIndexOf(">") == -1)
                                break;
                            temp1 = temp1.Substring(temp1.LastIndexOf(">") + 1).Trim();
                            sku = temp1;

                            temp = temp.Substring(temp.IndexOf("</font></td>") + "</font></td>".Length);
                            if (temp.IndexOf("</font></td>") == -1)
                                break;
                            temp1 = temp.Substring(0, temp.IndexOf("</font></td>")).Trim();
                            if (temp1.LastIndexOf(">") == -1)
                                break;
                            temp1 = temp1.Substring(temp1.LastIndexOf(">") + 1).Trim();
                            qty = Str_Utils.string_to_int(temp1);

                            temp = temp.Substring(temp.IndexOf("</font></td>") + "</font></td>".Length);
                            if (temp.IndexOf("</font></td>") == -1)
                                break;
                            temp1 = temp.Substring(0, temp.IndexOf("</font></td>")).Trim();
                            if (temp1.LastIndexOf(">") == -1)
                                break;
                            temp1 = temp1.Substring(temp1.LastIndexOf(">") + 1).Trim();
                            title = temp1;

                            temp = temp.Substring(temp.IndexOf("</font></td>") + "</font></td>".Length);
                            if (temp.IndexOf("</font></td>") == -1)
                                break;
                            temp1 = temp.Substring(0, temp.IndexOf("</font></td>")).Trim();
                            if (temp1.LastIndexOf(">") == -1)
                                break;
                            temp1 = temp1.Substring(temp1.LastIndexOf(">") + 1).Trim();
                            price = Str_Utils.string_to_currency(temp1);

                            temp = temp.Substring(temp.IndexOf("</font></td>") + "</font></td>".Length);
                            if (temp.IndexOf("</font></td>") == -1)
                                break;
                            temp = temp.Substring(temp.IndexOf("</font></td>") + "</font></td>".Length);

                            ZProduct product = new ZProduct();
                            product.price = price;
                            product.sku = sku;
                            product.title = title;
                            product.qty = qty;
                            report.m_product_items.Add(product);

                            MyLogger.Info($"... OP-1 sku = {sku}, title = {title}, qty = {qty}, price = {price}");
                        }
                    }
                }
            }
            if (!get_address)
            {
                int next_pos = htmlbody.IndexOf("Shipping Address", StringComparison.CurrentCultureIgnoreCase);
                do
                {
                    if (next_pos == -1)
                        break;
                    string temp = htmlbody.Substring(next_pos);
                    next_pos = temp.IndexOf("</table>", StringComparison.CurrentCultureIgnoreCase);
                    if (next_pos == -1)
                        break;
                    temp = temp.Substring(next_pos + "</table>".Length);

                    XMailHelper.find_html_part(temp, "table", out next_pos);
                    if (next_pos == -1)
                        break;
                    temp = temp.Substring(next_pos);

                    temp = XMailHelper.find_html_part(temp, "table", out next_pos);
                    if (temp == "")
                        break;
                    temp = XMailHelper.html2text(temp);
                    temp = temp.Replace("\r\n", " ");
                    temp = temp.Replace("\n", " ");
                    string full_address = temp.Trim();
                    string state_address = XMailHelper.get_address_state_name(full_address);
                    if (state_address != "")
                    {
                        report.set_address(full_address, state_address);
                        MyLogger.Info($"... OP-1 full_address = {full_address}, state_address = {state_address}");
                        get_address = true;
                    }

                } while (false);
            }
        }
    }
}
