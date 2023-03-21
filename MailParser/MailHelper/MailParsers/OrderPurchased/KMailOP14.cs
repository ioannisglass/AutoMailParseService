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
        private void parse_mail_op_14(MimeMessage mail, KReportOP report)
        {
            parse_mail_op_14_for_htmltext(mail, report);
        }
        private void parse_mail_op_14_for_htmltext(MimeMessage mail, KReportOP report)
        {
            string subject = XMailHelper.get_subject(mail);
            string item_title = "";
            string item_sku = "";

            report.m_mail_type = KReportBase.MailType.OP_14;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = ConstEnv.RETAILER_EBAY;

            MyLogger.Info($"... OP-14 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-14 m_op_retailer = {report.m_retailer}");

            List<string> title_list = new List<string>();
            List<string> sku_list = new List<string>();
            List<int> qty_list = new List<int>();
            List<float> price_list = new List<float>();

            int mail_type = 0;

            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Order number:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp = line.Substring("Order number:".Length);
                    temp = temp.Trim();
                    report.set_order_id(temp);

                    MyLogger.Info($"... OP-14 order id = {temp}");
                    continue;
                }

                if (line.StartsWith("Item ID:", StringComparison.CurrentCultureIgnoreCase) && lines[i + 1].Trim().StartsWith("Quantity:", StringComparison.CurrentCultureIgnoreCase))
                {
                    /**
                     *          -->
                     *          HP N270h 27" Edge to Edge Full HD Gaming Monito...
                     *          Item ID: 122820047918
                     *          Quantity: 1
                     *          Estimated delivery: Fri. Jun. 15
                     *          Paid: $132.99 with Gift cards
                     **/

                    float price = 0;

                    string temp = line.Substring("Item ID:".Length).Trim();
                    item_sku = temp.Trim();

                    int k = i - 1;
                    temp = "";
                    while (lines[k].Trim() != "-->")
                    {
                        temp = lines[k] + " " + temp;
                        k--;
                    }
                    item_title = temp;

                    temp = lines[i + 1].Trim();
                    temp = temp.Substring("Quantity:".Length).Trim();
                    int qty = Str_Utils.string_to_int(temp);

                    i += 2;

                    temp = lines[i].Trim();
                    if (temp.StartsWith("Estimated delivery:", StringComparison.CurrentCultureIgnoreCase))
                    {
                        i++;
                        temp = lines[i].Trim();
                    }
                    if (temp.StartsWith("Paid:", StringComparison.CurrentCultureIgnoreCase))
                    {
                        temp = temp.Substring("Paid:".Length).Trim();
                        if (temp.IndexOf("with", StringComparison.CurrentCultureIgnoreCase) != -1)
                        {
                            string payment = temp.Substring(temp.IndexOf("with", StringComparison.CurrentCultureIgnoreCase) + "with".Length).Trim();
                            temp = temp.Substring(0, temp.IndexOf("with", StringComparison.CurrentCultureIgnoreCase));
                            price = Str_Utils.string_to_currency(temp);
                            if (price > 0)
                            {
                                report.add_payment_card_info(new ZPaymentCard(payment, "", price));
                                MyLogger.Info($"... OP-14 Total charged to {payment}, price = {price}");
                            }
                        }
                        else
                        {
                            price = Str_Utils.string_to_currency(temp);
                        }
                    }

                    title_list.Add(item_title);
                    sku_list.Add(item_sku);
                    qty_list.Add(qty);
                    price_list.Add(price);

                    mail_type = 1;

                    continue;
                }

                if (line.StartsWith("Item ID:", StringComparison.CurrentCultureIgnoreCase) && line != "Item ID:") // It follows by the Order summary
                {
                    /**
                     *          -->
                     *          HP 23ER Frameless Silver/White 23" IPS Widescreen LCD/LED Monitors, HDMI...
                     *          Total: $0.00
                     *       >  Item ID: 292322334712
                     *          To complement your purchase
                     *          
                     *          
                     *          View order details
                     *          Browse deals
                     *          Ninja Intelli-Sense Kitchen System with Smart Vessel Reco...
                     *          Ninja Intelli-Sense Kitchen System with Smart Vessel Recognition (CT680SS)
                     *          Total: $210.47
                     *       >  Item ID: 132908675413
                     *          To complement your purchase
                     *          
                     *          
                     *          Order summary
                     *          Belkin | Boost Up Fast Wireless Charging Pad for Qi Devic...
                     *          Belkin | Boost Up Fast Wireless Charging Pad for Qi Devices | Brand New
                     *          Total: $0.00
                     *          Order number: 23-03675-43577
                     *       >  Item ID: 254243043277
                     *          To complement your purchase
                     *          
                     *          
                     *          View order details
                     *          EVGA SuperNOVA 1000 G2, 80&#43; GOLD 1000W, Fully Mod...
                     *          EVGA SuperNOVA 1000 G2, 80&#43; GOLD 1000W, Fully Modular, 120-G2-1000-XR
                     *          -->
                     *          Item price: $99.99
                     *          Est. delivery to your address: Fri, Sep 14
                     *       >  Item ID: 263657161160
                     *          Seller: evga_official (1339)
                     *          View
                     *          order
                     *          details
                     *          View order details
                     *          Bushnell 119875C 24MP Trophy Cam HD Aggressor Trail Scout...
                     *          Bushnell 119875C 24MP Trophy Cam HD Aggressor Trail Scouting Game Camera Camo
                     *          -->
                     *          Item price: $103.99
                     *          Est. delivery to your address: Wed, Sep 12
                     *       >  Item ID: 142899526190
                     *          Seller: bigfishbuddy (132511)
                     *
                     **/

                    string temp = line.Substring("Item ID:".Length);
                    item_sku = temp.Trim();

                    int k = i - 1;
                    if (lines[k].Trim().StartsWith("Order number:", StringComparison.CurrentCultureIgnoreCase) || lines[k].Trim().IndexOf("delivery to", StringComparison.CurrentCultureIgnoreCase) != -1)
                        k--;
                    temp = lines[k].Trim();
                    if (temp.StartsWith("Total:", StringComparison.CurrentCultureIgnoreCase))
                    {
                        temp = lines[k - 1].Trim();
                        sku_list.Add(item_sku);
                        title_list.Add(temp);
                    }
                    else if (temp.StartsWith("Item price:", StringComparison.CurrentCultureIgnoreCase) && lines[k - 1].Trim() == "-->")
                    {
                        temp = temp.Substring("Item price:".Length).Trim();
                        price_list.Add(Str_Utils.string_to_currency(temp));

                        temp = lines[k - 2].Trim();
                        sku_list.Add(item_sku);
                        title_list.Add(temp);
                        qty_list.Add(1);

                        mail_type = 2;
                    }
                    else
                    {
                        sku_list.Add(item_sku);
                    }
                    continue;
                }

                if (mail_type == 0 && line.ToUpper() == "ORDER TOTAL:" && lines[i + 1].Trim().StartsWith("Price", StringComparison.CurrentCultureIgnoreCase))
                {
                    /**
                     *          Order total:
                     *          Price
                     *          $359.99
                     *          Shipping
                     *          Free
                     *          -->
                     *          Total:
                     *          Total:
                     *          $4.97
                     *          
                     *          
                     *          Order total:
                     *          Price (3 x $174.99)
                     *          $524.97
                     *          Shipping
                     *          Free
                     *          Sales tax
                     *          $34.78
                     *          Gift cards, coupons
                     *          -$559.75
                     *          -->
                     *          Total charged to
                     *          Total charged to
                     *          $0.00
                     **/

                    float total_price = 0;
                    float shipping = 0;
                    float tax = 0;
                    float gc = 0;
                    float total = 0;

                    string temp = lines[++i].Trim();
                    temp = temp.Substring("Price".Length).Trim();
                    if (temp.IndexOf("(") != -1 && temp.IndexOf(")") != -1)
                    {
                        temp = temp.Substring(temp.IndexOf("(") + 1);
                        temp = temp.Trim();
                        temp = temp.Substring(0, temp.IndexOf(")"));
                        temp = temp.Trim();

                        if (temp.IndexOf(" x ") != -1)
                        {
                            string num_part = temp.Substring(0, temp.IndexOf(" x "));
                            int qty = Str_Utils.string_to_int(num_part);

                            string price_part = temp.Substring(temp.IndexOf(" x ") + " x ".Length);
                            float price = Str_Utils.string_to_currency(price_part);

                            qty_list.Add(qty);
                            price_list.Add(price);
                        }
                        temp = lines[++i].Trim();
                        total_price = Str_Utils.string_to_currency(temp);
                    }
                    else
                    {
                        temp = lines[++i].Trim();
                        float price = Str_Utils.string_to_currency(temp);
                        total_price = price;

                        qty_list.Add(1);
                        price_list.Add(price);
                    }

                    int k = i;
                    if (lines[k].Trim().ToUpper() == "SHIPPING")
                    {
                        temp = lines[++k].Trim();
                        if (temp.ToUpper() == "FREE")
                            shipping = 0;
                        else
                            shipping = Str_Utils.string_to_currency(temp);
                        k++;
                    }
                    if (lines[k].Trim().ToUpper() == "SALES TAX")
                    {
                        temp = lines[++k].Trim();
                        tax = Str_Utils.string_to_currency(temp);
                        k++;
                    }
                    if (lines[k].Trim().ToUpper() == "GIFT CARDS, COUPONS")
                    {
                        temp = lines[++k].Trim();
                        gc = Str_Utils.string_to_currency(temp);
                        if (gc < 0)
                            gc *= -1;
                        k++;
                    }
                    if ((lines[k].Trim().ToUpper() == "TOTAL CHARGED TO" && lines[k + 1].Trim().ToUpper() == "TOTAL CHARGED TO") ||
                        (lines[k].Trim().ToUpper() == "TOTAL:" && lines[k + 1].Trim().ToUpper() == "TOTAL:")
                        )
                    {
                        k += 2;
                        temp = lines[k].Trim();
                        total = Str_Utils.string_to_currency(temp);
                    }

                    MyLogger.Info($"... OP-14 Price               : {total_price}");
                    MyLogger.Info($"... OP-14 Shipping            : {shipping}");
                    MyLogger.Info($"... OP-14 GIFT CARDS, COUPONS : {gc}");
                    MyLogger.Info($"... OP-14 Tax                 : {tax}");
                    MyLogger.Info($"... OP-14 Total               : {total}");

                    total = total_price + tax + shipping;

                    if (gc != 0)
                    {
                        report.add_payment_card_info(new ZPaymentCard(ConstEnv.GIFT_CARD, "", gc));
                        MyLogger.Info($"... OP-10 Gift card = {gc}");
                    }

                    report.m_tax = tax;
                    report.set_total(total);
                    MyLogger.Info($"... OP-14 tax = {tax}, total = {total}");

                    i = k;
                    continue;
                }

                if (line.StartsWith("Total:", StringComparison.CurrentCultureIgnoreCase) && line.IndexOf("$") != -1)
                {
                    string temp = line.Substring("Total:".Length);
                    report.set_total(Str_Utils.string_to_currency(temp));

                    MyLogger.Info($"... OP-14 order total = {report.m_total}");
                    continue;
                }
                if (line.ToUpper() == "TOTAL CHARGED TO" &&
                    ((lines[i + 1].Trim().IndexOf("x -") != -1 && (lines[i + 2].Trim().StartsWith("$") || lines[i + 2].Trim().StartsWith("-$"))) ||
                    (lines[i + 1].Trim().StartsWith("$") || lines[i + 1].Trim().StartsWith("-$"))))
                {
                    string last_4_digits = "";
                    string temp = lines[i + 1].Trim();
                    if (temp.IndexOf("x -") != -1)
                    {
                        last_4_digits = temp.Substring(temp.IndexOf("x -") + "x -".Length).Trim();
                        temp = lines[i + 2].Trim();
                    }
                    float price = Str_Utils.string_to_currency(temp);
                    if (price == 0)
                        continue;

                    string payment = "";
                    string htmltext = XMailHelper.get_htmltext(mail);
                    int pos = htmltext.IndexOf("Total charged to <img src=");
                    if (pos != -1)
                    {
                        temp = htmltext.Substring(pos);
                        temp = temp.Substring(0, temp.IndexOf(">"));
                        if (temp.IndexOf("alt=\"") != -1)
                        {
                            temp = temp.Substring(temp.IndexOf("alt=\"") + "alt=\"".Length);
                            temp = temp.Substring(0, temp.IndexOf("\"")).Trim();
                            payment = temp;
                        }
                    }

                    report.add_payment_card_info(new ZPaymentCard(payment, last_4_digits, price));

                    MyLogger.Info($"... OP-14 Total charged to {payment}, last_4_digits = {last_4_digits}, price = {price}");

                    if (report.m_total == 0)
                    {
                        report.set_total(price);
                        MyLogger.Info($"... OP-14 total (from Total charged to) = {report.m_total}");
                    }
                    continue;
                }

                if (line.StartsWith("Sales tax", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp = line.Substring("Sales tax".Length);
                    if (temp.IndexOf("$") != -1)
                        report.m_tax = Str_Utils.string_to_currency(temp);
                    else
                        report.m_tax = Str_Utils.string_to_currency(lines[++i].Trim());

                    MyLogger.Info($"... OP-14 tax = {report.m_tax}");
                    continue;
                }
                if (line.ToUpper() == "YOUR ORDER WILL SHIP TO:")
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
                        MyLogger.Info($"... OP-14 full_address = {full_address}, state_address = {state_address}");
                    }
                    continue;
                }
                if (line.IndexOf("It will ship to", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    string full_address = "";
                    string state_address = "";
                    string temp = line.Substring(line.IndexOf("It will ship to", StringComparison.CurrentCultureIgnoreCase) + "It will ship to".Length).Trim();
                    int k = i + 1;
                    while (k - i < 5)
                    {
                        full_address += " " + temp;
                        temp = lines[k].Trim();

                        state_address = XMailHelper.get_address_state_name(full_address);
                        if (state_address != "")
                            break;
                        if (temp.IndexOf("View order details", StringComparison.CurrentCultureIgnoreCase) != -1)
                            break;
                        k++;
                    }
                    full_address = full_address.Trim();
                    state_address = state_address.Trim();
                    if (state_address != "")
                    {
                        report.set_address(full_address, state_address);
                        MyLogger.Info($"... OP-14 full_address = {full_address}, state_address = {state_address}");
                    }
                    i = k - 1;
                    continue;
                }
            }

            int n = 0;
            if (mail_type == 0)
            {
                n = title_list.Count;

                if (n != price_list.Count)
                {
                    if (price_list.Count == 1)
                    {
                        for (int i = 1; i < n; i++)
                        {
                            price_list.Add(price_list[0]);
                            qty_list.Add(qty_list[0]);
                        }
                    }
                    else if (price_list.Count == 0)
                    {
                        for (int i = 0; i < n; i++)
                        {
                            price_list.Add(0);
                            qty_list.Add(1);
                        }
                    }
                    else
                    {
                        n = Math.Min(n, price_list.Count);
                        n = Math.Min(n, qty_list.Count);
                    }
                }
            }
            else
            {
                n = title_list.Count;
            }

            for (int i = 0; i < n; i++)
            {
                ZProduct product = new ZProduct();
                product.price = price_list[i];
                product.sku = sku_list[i];
                product.title = title_list[i];
                product.qty = qty_list[i];
                report.m_product_items.Add(product);

                MyLogger.Info($"... OP-14 qty = {qty_list[i]}, price = {price_list[i]}, sku = {sku_list[i]}, title = {title_list[i]}");
            }
        }
    }
}
