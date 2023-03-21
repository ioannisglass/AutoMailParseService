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
        private void parse_mail_op_16(MimeMessage mail, KReportOP report)
        {
            if (!XMailHelper.is_bodytext_existed(mail))
            {
                parse_mail_op_16_for_htmltext(mail, report);
                return;
            }
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_16;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = ConstEnv.RETAILER_HP;

            MyLogger.Info($"... OP-16 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-16 m_op_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.IndexOf("Qty") != -1 && line.IndexOf("Price") != -1)
                {
                    string temp = lines[++i].Trim();
                    string title = temp.Trim();

                    temp = lines[++i].Trim();
                    string sku = temp.Trim();

                    i += 3;
                    temp = lines[i++].Trim();

                    int qty = 0;
                    float price = 0;

                    if (temp.IndexOf(" ") != -1)
                    {
                        string qty_part = temp.Substring(0, temp.IndexOf(" ")).Trim();
                        qty = Str_Utils.string_to_int(qty_part);
                        string price_part = temp.Substring(temp.IndexOf(" ") + 1).Trim();
                        price = Str_Utils.string_to_currency(price_part);
                    }
                    else if (temp.IndexOf("$") != -1)
                    {
                        string qty_part = temp.Substring(0, temp.IndexOf("$")).Trim();
                        qty = Str_Utils.string_to_int(qty_part);
                        string price_part = temp.Substring(temp.IndexOf("$")).Trim();
                        price = Str_Utils.string_to_currency(price_part);
                    }

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    report.m_product_items.Add(product);

                    MyLogger.Info($"... OP-16 item sku = {sku}, title = {title}, qty = {qty}, price = {price}");

                    continue;
                }
                if (line.IndexOf("*") != -1 && line.IndexOf("Your order number is") != -1)
                {
                    string temp = line.Substring(line.IndexOf("Your order number is") + "Your order number is".Length).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... OP-16 order id = {temp}");
                    continue;
                }
                if (line.IndexOf("*") != -1 && line.IndexOf("Order date:") != -1)
                {
                    string temp = line.Substring(line.IndexOf("Order date:") + "Order date:".Length).Trim();
                    DateTime date = DateTime.Parse(temp);
                    report.m_op_purchase_date = date;
                    MyLogger.Info($"... OP-16 Date Of Purchase = {date}");
                    continue;
                }
                if (line == "Total:")
                {
                    i += 3;
                    string temp = lines[i];
                    float tax = Str_Utils.string_to_currency(temp);
                    report.m_tax = tax;
                    MyLogger.Info($"... OP-16 tax = {tax}");
                    continue;
                }
                if (line.StartsWith("Ship to address:"))
                {
                    string temp = lines[i - 1];
                    float total = Str_Utils.string_to_currency(temp);
                    report.set_total(total);
                    MyLogger.Info($"... OP-16 order total = {total}");
                    continue;
                }
                if (line == "Payment Information:")
                {
                    List<string> payment_type_list = new List<string>();
                    List<float> price_list = new List<float>();

                    string next_line = lines[++i].Trim();
                    while (next_line == "")
                    {
                        next_line = lines[++i].Trim();
                    }

                    while (i < lines.Length)
                    {
                        if (next_line.EndsWith(":"))
                        {
                            string payment_type = next_line.Substring(0, next_line.Length - 1).Trim();
                            payment_type_list.Add(payment_type);
                        }
                        next_line = lines[++i].Trim();
                        if (next_line == "")
                            break;
                    }

                    while (next_line == "")
                    {
                        next_line = lines[++i].Trim();
                    }
                    while (next_line != "")
                    {
                        next_line = lines[++i].Trim();
                    }
                    while (next_line == "")
                    {
                        next_line = lines[++i].Trim();
                    }

                    while (i < lines.Length)
                    {
                        float price = Str_Utils.string_to_currency(next_line);
                        price_list.Add(price);
                        next_line = lines[++i].Trim();
                        if (next_line == "")
                            break;
                    }

                    for (int k = 0; k < payment_type_list.Count; k++)
                    {
                        ZPaymentCard c = new ZPaymentCard(payment_type_list[k], "", price_list[k]);
                        report.add_payment_card_info(c);
                        MyLogger.Info($"... OP-16 card type = {payment_type_list[k]}, price = {price_list[k]}");
                    }

                    continue;
                }
            }
        }
        private void parse_mail_op_16_for_htmltext(MimeMessage mail, KReportOP report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_16;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = ConstEnv.RETAILER_HP;

            MyLogger.Info($"... OP-16 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-16 m_op_retailer = {report.m_retailer}");

            int op16_type = 1;

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if (line == "QtyPrice")
                {
                    op16_type = 2;
                    break;
                }
            }

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (op16_type == 1)
                {
                    if (line.ToUpper() == "PRODUCT DETAILS" && lines[i + 1].Trim().ToUpper() == "STATUS" && lines[i + 2].Trim().ToUpper() == "QTY" && lines[i + 3].Trim().ToUpper() == "EXTENDED PRICE")
                    {
                        i += 4;
                        string temp = "";

                        while (i < lines.Length && !lines[i].Trim().StartsWith("Summary of savings", StringComparison.CurrentCultureIgnoreCase))
                        {
                            string title = "";
                            string sku = "";
                            int qty = 0;
                            float price = 0;

                            while (lines[i].Trim().IndexOf("Product number:", StringComparison.CurrentCultureIgnoreCase) == -1)
                            {
                                temp += " " + lines[i].Trim();
                                i++;
                            }
                            if (temp.IndexOf("-->") != -1)
                                temp = temp.Substring(temp.IndexOf("-->") + "-->".Length).Trim();
                            title = temp;

                            temp = lines[i].Trim();
                            temp = temp.Substring(temp.IndexOf("Product number:", StringComparison.CurrentCultureIgnoreCase) + "Product number:".Length).Trim();
                            sku = temp;

                            i++;
                            while (lines[i].Trim().StartsWith("•"))
                            {
                                title += " " + lines[i].Trim();
                                i++;
                            }
                            title = title.Trim();

                            temp = lines[++i].Trim();
                            if (temp.StartsWith("Estimated shipment:", StringComparison.CurrentCultureIgnoreCase))
                            {
                                if (temp.ToUpper() == "ESTIMATED SHIPMENT:")
                                    temp = lines[++i].Trim();
                                else
                                    temp = temp.Substring("Estimated shipment:".Length).Trim();
                                DateTime date = DateTime.Parse(temp);
                                MyLogger.Info($"... OP-16 Estimated shipment = {date.ToString()}");

                                temp = lines[++i].Trim();
                            }
                            if (temp.StartsWith("Estimated ship date:", StringComparison.CurrentCultureIgnoreCase))
                            {
                                if (temp.ToUpper() == "ESTIMATED SHIP DATE:")
                                    temp = lines[++i].Trim();
                                else
                                    temp = temp.Substring("Estimated ship date:".Length).Trim();
                                DateTime date = DateTime.Parse(temp);
                                MyLogger.Info($"... OP-16 Estimated shipment = {date.ToString()}");

                                temp = lines[++i].Trim();
                            }

                            qty = Str_Utils.string_to_int(temp);

                            temp = lines[++i].Trim();
                            if (temp.IndexOf("$") == -1)
                                break;
                            price = Str_Utils.string_to_currency(temp);

                            ZProduct product = new ZProduct();
                            product.price = price;
                            product.sku = sku;
                            product.title = title;
                            product.qty = qty;
                            report.m_product_items.Add(product);

                            MyLogger.Info($"... OP-16 item sku = {sku}, title = {title}, qty = {qty}, price = {price}");

                            i++;
                        }
                    }
                }
                else if (op16_type == 2)
                {
                    if (line.ToUpper() == "PRODUCT DETAILS" && lines[i + 1].Trim().ToUpper() == "STATUS" && lines[i + 2].Trim().ToUpper() == "QTYPRICE")
                    {
                        i += 3;
                        string temp = "";

                        while (i < lines.Length && !lines[i].Trim().StartsWith("Summary of savings", StringComparison.CurrentCultureIgnoreCase))
                        {
                            string title = "";
                            string sku = "";
                            int qty = 0;
                            float price = 0;

                            while (lines[i].Trim().IndexOf("#") == -1)
                            {
                                temp += " " + lines[i].Trim();
                                i++;
                            }
                            if (temp.IndexOf("-->") != -1)
                                temp = temp.Substring(temp.IndexOf("-->") + "-->".Length).Trim();
                            title = temp.Trim();

                            temp = lines[i].Trim();
                            sku = temp;

                            i++; // status

                            temp = lines[++i].Trim();
                            if (temp.StartsWith("Estimated shipment:", StringComparison.CurrentCultureIgnoreCase))
                            {
                                temp = temp.Substring("Estimated shipment:".Length).Trim();
                                DateTime date = DateTime.Parse(temp);
                                MyLogger.Info($"... OP-16 Estimated shipment = {date.ToString()}");

                                temp = lines[++i].Trim();
                            }

                            if (temp.IndexOf("$") == -1)
                                break;
                            string qty_part = temp.Substring(0, temp.IndexOf("$")).Trim();
                            qty = Str_Utils.string_to_int(qty_part);

                            temp = temp.Substring(temp.IndexOf("$")).Trim();
                            price = Str_Utils.string_to_currency(temp);

                            ZProduct product = new ZProduct();
                            product.price = price;
                            product.sku = sku;
                            product.title = title;
                            product.qty = qty;
                            report.m_product_items.Add(product);

                            MyLogger.Info($"... OP-16 item sku = {sku}, title = {title}, qty = {qty}, price = {price}");

                            i++;
                        }
                    }
                }
                if (line.IndexOf("Qty") != -1 && line.IndexOf("Price") != -1)
                {
                    string temp = lines[++i].Trim();
                    string title = temp.Trim();

                    temp = lines[++i].Trim();
                    string sku = temp.Trim();

                    i += 3;
                    temp = lines[i++].Trim();

                    int qty = 0;
                    float price = 0;

                    if (temp.IndexOf(" ") != -1)
                    {
                        string qty_part = temp.Substring(0, temp.IndexOf(" ")).Trim();
                        qty = Str_Utils.string_to_int(qty_part);
                        string price_part = temp.Substring(temp.IndexOf(" ") + 1).Trim();
                        price = Str_Utils.string_to_currency(price_part);
                    }
                    else if (temp.IndexOf("$") != -1)
                    {
                        string qty_part = temp.Substring(0, temp.IndexOf("$")).Trim();
                        qty = Str_Utils.string_to_int(qty_part);
                        string price_part = temp.Substring(temp.IndexOf("$")).Trim();
                        price = Str_Utils.string_to_currency(price_part);
                    }

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    report.m_product_items.Add(product);

                    MyLogger.Info($"... OP-16 item sku = {sku}, title = {title}, qty = {qty}, price = {price}");

                    continue;
                }
                if (line.StartsWith("Your order number is", StringComparison.CurrentCultureIgnoreCase) && !line.StartsWith("Your order number is your invoice number", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "YOUR ORDER NUMBER IS")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Your order number is".Length).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... OP-16 order id = {temp}");
                    continue;
                }
                if (line.StartsWith("Estimated shipment:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "ESTIMATED SHIPMENT:")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Estimated shipment:".Length).Trim();
                    DateTime date = DateTime.Parse(temp);
                    MyLogger.Info($"... OP-16 Estimated shipment = {date.ToString()}");
                    continue;
                }
                if (line.StartsWith("Order date:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "ORDER DATE:")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Order date:".Length).Trim();
                    DateTime date = DateTime.Parse(temp);
                    report.m_op_purchase_date = date;
                    MyLogger.Info($"... OP-16 Date Of Purchase = {date}");
                    continue;
                }
                if (op16_type == 1)
                {
                    if (line.IndexOf("Tax:", StringComparison.CurrentCultureIgnoreCase) != -1)
                    {
                        string temp;
                        if (line.EndsWith("Tax:", StringComparison.CurrentCultureIgnoreCase))
                            temp = lines[++i].Trim();
                        else
                            temp = line.Substring(line.IndexOf("Tax:", StringComparison.CurrentCultureIgnoreCase) + "Tax:".Length).Trim();
                        float tax = Str_Utils.string_to_currency(temp);   
                        report.m_tax = tax;
                        MyLogger.Info($"... OP-16 tax = {tax}");
                        continue;
                    }
                    if (line.StartsWith("Total:", StringComparison.CurrentCultureIgnoreCase))
                    {
                        string temp;
                        if (line.ToUpper() == "TOTAL:")
                            temp = lines[++i].Trim();
                        else
                            temp = line.Substring("Total:".Length).Trim();
                        float total = Str_Utils.string_to_currency(temp);
                        report.set_total(total);
                        MyLogger.Info($"... OP-16 order total = {total}");
                        continue;
                    }
                }
                else if (op16_type == 2)
                {
                    if (line.ToUpper() == "NJ TAX:" && lines[i + 1].Trim().ToUpper() == "TOTAL:")
                    {
                        /**
                         *      NJ Tax:
                         *      Total:
                         *      $199.99Free$0.00
                         *      $199.99
                         **/

                        i += 2;
                        string temp = lines[i].Trim();
                        if (temp.LastIndexOf("$") != -1)
                        {
                            temp = temp.Substring(temp.LastIndexOf("$") + 1).Trim();
                            float tax = Str_Utils.string_to_currency(temp);
                            report.m_tax = tax;
                            MyLogger.Info($"... OP-16 tax = {tax}");
                        }
                        temp = lines[++i].Trim();
                        if (temp.StartsWith("$"))
                        {
                            float total = Str_Utils.string_to_currency(temp);
                            report.set_total(total);
                            MyLogger.Info($"... OP-16 order total = {total}");
                        }
                    }
                }
                if (line.ToUpper() == "PAYMENT INFORMATION:" || line.ToUpper() == "PAYMENT INFORMATION")
                {
                    /**
                     * case op5_type == 1 :
                     * 
                     *      Payment information
                     *      Gift Card:
                     *      $91.51
                     *      Gift Card:
                     *      $168.47
                     *      pace cohen
                     *      PACE COHEN PRIME ELECTRONICS
                     *      49 CHELSEA RD
                     *      JACKSON, NJ 08527
                     * 
                     * case op5_type == 2 :
                     * 
                     *      Payment Information:
                     *      Discover Network:
                     *      Gift Card:
                     *      Gift Card:
                     *      pace cohenPACE COHEN PRIME ELECTRONICS
                     *      49 Chelsea RdJackson, NJ 08527
                     *      $49.96$250.00$100.02
                     * 
                     **/

                    List<string> payment_type_list = new List<string>();
                    List<float> price_list = new List<float>();

                    if (op16_type == 1)
                    {
                        i++;

                        while (i < lines.Length && lines[i].Trim().EndsWith(":"))
                        {
                            string temp = lines[i].Trim();
                            string payment_type = temp.Substring(0, temp.Length - 1);

                            temp = lines[++i].Trim();
                            if (!temp.StartsWith("$"))
                                break;
                            float price = Str_Utils.string_to_currency(temp);

                            payment_type_list.Add(payment_type);
                            price_list.Add(price);

                            i++;
                        }
                    }
                    else if (op16_type == 2)
                    {
                        i++;

                        string temp;
                        while (i < lines.Length && lines[i].Trim().EndsWith(":"))
                        {
                            temp = lines[i].Trim();
                            string payment_type = temp.Substring(0, temp.Length - 1);

                            payment_type_list.Add(payment_type);

                            i++;
                        }

                        while (!lines[i].Trim().StartsWith("$"))
                            i++;

                        temp = lines[i].Trim();
                        string[] prices = temp.Split('$');
                        foreach (string price_part in prices)
                        {
                            if (price_part == "")
                                continue;
                            float price = Str_Utils.string_to_currency(price_part);
                            price_list.Add(price);
                        }
                        if (price_list.Count != payment_type_list.Count)
                            continue;
                    }

                    if (payment_type_list.Count == price_list.Count)
                    {
                        for (int k = 0; k < payment_type_list.Count; k++)
                        {
                            ZPaymentCard c = new ZPaymentCard(payment_type_list[k], "", price_list[k]);
                            report.add_payment_card_info(c);
                            MyLogger.Info($"... OP-16 card type = {payment_type_list[k]}, price = {price_list[k]}");
                        }
                    }

                    continue;
                }
                if (line.ToUpper() == "SHIP TO ADDRESS" || line.ToUpper() == "SHIP TO ADDRESS:")
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
                        MyLogger.Info($"... OP-16 full_address = {full_address}, state_address = {state_address}");
                    }
                    i--;
                    continue;
                }
            }
        }
    }
}
