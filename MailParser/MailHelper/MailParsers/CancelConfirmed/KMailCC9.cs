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
    partial class KMailBaseCC : KMailBaseParser
    {
        private void parse_mail_cc_9(MimeMessage mail, KReportCC card)
        {
            if (!XMailHelper.is_bodytext_existed(mail))
            {
                analyze_mail_order_9_for_htmltext(mail, card);
                return;
            }

            string subject = XMailHelper.get_subject(mail);

            card.m_mail_type = KReportBase.MailType.CC_9;

            card.m_retailer = ConstEnv.RETAILER_WALMART;

            MyLogger.Info($"... CC-9 m_cc_retailer = {card.m_retailer}");

            float total = 0;
            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Order number:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;

                    if (line.ToUpper() == "ORDER NUMBER:")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Order number:".Length).Trim();
                    card.set_order_id(temp);
                    MyLogger.Info($"... CC-9 order id = {temp}");
                }

                if ((line.Replace(" ", "").ToUpper() == "ITEMDESCRIPTIONQTYPRICETOTAL" || line.Replace(" ", "").ToUpper() == "ITEMQTYPRICETOTAL") && lines[i + 1].Trim().StartsWith("---------------") && lines[i + 1].Trim().EndsWith("-"))
                {
                    /**
                     *      ITEM DESCRIPTION                                   QTY     PRICE     TOTAL
                     *      --------------------------------------------------------------------------
                     *      --
                     *      Garmin nuvi 40 4.3" Portable GPS                  2       $68.00
                     *      $136.00
                     *      ==========================================================================
                     *      ==
                     *      
                     *      
                     *      ITEM                                             QTY   PRICE   TOTAL
                     *      -----------------------------------------------------------------------
                     *      
                     *      Mainstays 6' Centerfold Table, Multiple Colors;   1     $38.88  $38.88
                     *      Color: White
                     *      
                     *      Mainstays 6' Centerfold Table, Multiple Colors;   1     $38.88  $38.88
                     *      Color: White
                     *      
                     *      =======================================================================
                     * 
                     *
                     **/

                    i += 2;
                    if (lines[i].Trim().StartsWith("-"))
                        i++;
                    while (lines[i].Trim() == "")
                        i++;

                    while (!lines[i].Trim().StartsWith("========"))
                    {
                        if (lines[i].Trim() == "")
                        {
                            i++;
                            continue;
                        }

                        string title = "";
                        string sku = "";
                        int qty = 1;
                        float price = 0;

                        string temp = lines[i].Trim();
                        int count = temp.Count(c => c == '$');
                        if (count == 1)
                        {
                            string price_part = temp.Substring(temp.IndexOf("$") + 1).Trim();
                            price = Str_Utils.string_to_currency(price_part);

                            temp = temp.Substring(0, temp.IndexOf("$")).Trim();
                            if (temp.LastIndexOf(" ") != -1)
                            {
                                string qty_part = temp.Substring(temp.LastIndexOf(" ") + 1).Trim();
                                qty = Str_Utils.string_to_int(qty_part);

                                title = temp.Substring(0, temp.LastIndexOf(" ")).Trim();
                            }
                            else
                            {
                                title = temp;
                            }

                            temp = lines[++i].Trim(); // subtotal
                            float t = Str_Utils.string_to_currency(temp);
                            total += t;
                            i++;
                        }
                        else if (count == 2)
                        {
                            string subtotal_part = temp.Substring(temp.LastIndexOf("$") + 1).Trim();
                            float t = Str_Utils.string_to_currency(subtotal_part);
                            total += t;
                            temp = temp.Substring(0, temp.LastIndexOf("$")).Trim();

                            string price_part = temp.Substring(temp.IndexOf("$") + 1).Trim();
                            price = Str_Utils.string_to_currency(price_part);
                            temp = temp.Substring(0, temp.IndexOf("$")).Trim();
                            if (temp.LastIndexOf(" ") != -1)
                            {
                                string qty_part = temp.Substring(temp.LastIndexOf(" ") + 1).Trim();
                                qty = Str_Utils.string_to_int(qty_part);

                                title = temp.Substring(0, temp.LastIndexOf(" ")).Trim();
                            }
                            else
                            {
                                title = temp;
                            }

                            i++;
                            while (lines[i].Trim() != "" && lines[i].Trim()[0] != '=')
                            {
                                title += " " + lines[i].Trim();
                                i++;
                            }
                        }
                        else
                        {
                            // Unknown
                            break;
                        }

                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        card.m_product_items.Add(product);

                        MyLogger.Info($"... CC-9 qty = {qty}, price = {price}, sku = {sku}, title = {title}");
                    }
                }
            }
            if (card.m_order_id == "")
            {
                analyze_mail_order_9_for_htmltext(mail, card);
                return;
            }
            card.set_total(total);
            MyLogger.Info($"... CC-9 total = {total}");

        }
        private void analyze_mail_order_9_for_htmltext(MimeMessage mail, KReportCC card)
        {
            string subject = XMailHelper.get_subject(mail);

            card.m_mail_type = KReportBase.MailType.CC_9;

            card.m_retailer = ConstEnv.RETAILER_WALMART;

            MyLogger.Info($"... CC-9 m_cc_retailer = {card.m_retailer}");

            float total = 0;
            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Sales Order ID:", StringComparison.CurrentCultureIgnoreCase)
                    && lines[i + 1].Trim().StartsWith("Item ID:", StringComparison.CurrentCultureIgnoreCase)
                    && lines[i + 2].Trim().StartsWith("Item:", StringComparison.CurrentCultureIgnoreCase)
                    )
                {
                    /**
                     * These type mails have the following features. 
                     * 
                     *     1) sender   : c-1746E2BE928948F1A5290FDFFE93DEE6@relay.walmart.com
                     *     2) Subject  : Message from Walmart.com Customer:  Cancel Order, 6171867186778
                     *     3) In the body text
                     * 
                     *          Sales Order ID: 6171867186778
                     *          Item ID: 4PYIMO59G4SH
                     *          Item: Kobalt 227-Piece Standard/Metric Mechanics Tool Set with Case 86756
                     *          
                     **/

                    string temp = line.Substring("Sales Order ID:".Length).Trim();
                    card.set_order_id(temp);
                    MyLogger.Info($"... CC-9 order id = {temp}");

                    string title = "";
                    string sku = "";
                    int qty = 1;
                    float price = 0;

                    temp = lines[++i].Trim();
                    temp = temp.Substring("Item ID:".Length).Trim();
                    sku = temp;

                    temp = lines[++i].Trim();
                    temp = temp.Substring("Item:".Length).Trim();
                    title = temp;

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    card.m_product_items.Add(product);

                    MyLogger.Info($"... CC-9 qty = {qty}, price = {price}, sku = {sku}, title = {title}");

                    continue;
                }

                if (line.StartsWith("Order number:", StringComparison.CurrentCultureIgnoreCase)
                    && lines[i + 1].Trim().StartsWith("Item name/number:", StringComparison.CurrentCultureIgnoreCase)
                    )
                {
                    /**
                     *          Order number: 2721984280658
                     *          Item name/number: Garmin 010-01540-01 Drivesmart 60lmt 6" Gps Navigator With Bluetooth & Free Lifetime Maps & Traffic Updates - 49369937 / Garmin 010-01540-01 Drivesmart 60lmt 6" Gps Navigator With Bluetooth & Free Lifetime Maps & Traffic Updates - 49369937
                     *          
                     *          Order number: 4961881796150
                     *          Item name/number: 159176736 / RYOBI 18-VOLT ONE&#43; AIRSTRIKE CORDLESS BRAD NAILER, 18-GAUGE, TOOL ONLY /
                     **/

                    string temp = line.Substring("Order number:".Length).Trim();
                    card.set_order_id(temp);
                    MyLogger.Info($"... CC-9 order id = {temp}");

                    temp = lines[++i].Trim();
                    temp = temp.Substring("Item name/number:".Length).Trim();

                    if (temp.EndsWith("/"))
                    {
                        if (temp.IndexOf(" / ") != -1)
                        {
                            string title = "";
                            string sku = "";
                            int qty = 1;
                            float price = 0;

                            temp = temp.Substring(0, temp.Length - 1).Trim();
                            sku = temp.Substring(0, temp.IndexOf(" / ")).Trim();
                            title = temp.Substring(temp.IndexOf(" / ") + temp.IndexOf(" / ") + 1).Trim();

                            ZProduct product = new ZProduct();
                            product.price = price;
                            product.sku = sku;
                            product.title = title;
                            product.qty = qty;
                            card.m_product_items.Add(product);

                            MyLogger.Info($"... CC-9 qty = {qty}, price = {price}, sku = {sku}, title = {title}");
                        }
                    }
                    else
                    {
                        string[] items = temp.Split(new string[] { " / " }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (string item in items)
                        {
                            string title = "";
                            string sku = "";
                            int qty = 1;
                            float price = 0;

                            temp = item.Trim();
                            if (temp.LastIndexOf(" - ") != -1)
                            {
                                title = temp.Substring(0, temp.LastIndexOf(" - ")).Trim();
                                sku = temp.Substring(temp.LastIndexOf(" - ") + " - ".Length).Trim();

                                if (title == "" && sku == "")
                                    continue;
                            }
                            else if (temp.EndsWith(" -"))
                            {
                                title = temp.Substring(0, temp.Length - 2);
                                if (title == "")
                                    continue;
                            }
                            else
                            {
                                continue;
                            }

                            ZProduct product = new ZProduct();
                            product.price = price;
                            product.sku = sku;
                            product.title = title;
                            product.qty = qty;
                            card.m_product_items.Add(product);

                            MyLogger.Info($"... CC-9 qty = {qty}, price = {price}, sku = {sku}, title = {title}");
                        }
                    }

                    continue;
                }

                if (line.StartsWith("Order number:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;

                    if (line.ToUpper() == "ORDER NUMBER:")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Order number:".Length).Trim();
                    card.set_order_id(temp);
                    MyLogger.Info($"... CC-9 order id = {temp}");
                }

                if (line.ToUpper() == "ITEM" && lines[i + 1].Trim().ToUpper() == "QTY" && lines[i + 2].Trim().ToUpper() == "PRICE" && lines[i + 3].Trim().ToUpper() == "TOTAL")
                {
                    /**
                     *      > + Item
                     *          Qty
                     *          Price
                     *        - Total
                     *  Items + Dell Inspiron i55655850GRY 15.6" Laptop, Touchscreen, Windows 10 Home, AMD FX-9800P Processor, 16GB RAM, 1TB Hard Drive
                     *          1
                     *          $499.00
                     *          $499.00
                     *        + AT&T Alcatel Ideal GoPhone Prepaid Smartphone
                     *          2
                     *          $29.88
                     *        - $59.76
                     *  total + $558.76
                     *   end  + Cancellation and refund information
                     **/

                    i += 4;

                    string temp;

                    while (!lines[i].Trim().StartsWith("Cancellation", StringComparison.CurrentCultureIgnoreCase) && !lines[i].Trim().StartsWith("Canceled", StringComparison.CurrentCultureIgnoreCase))
                    {
                        string title = "";
                        string sku = "";
                        int qty = 1;
                        float price = 0;

                        int qty_tmp;
                        temp = "";
                        while (lines[i].Trim()[0] != '$' && !int.TryParse(lines[i].Trim(), out qty_tmp))
                        {
                            temp += " " + lines[i].Trim();
                            i++;
                        }
                        title = temp.Trim();

                        temp = lines[i].Trim();
                        if (temp[0] == '$')
                        {
                            qty = 0;
                        }
                        else
                        {
                            qty = int.Parse(temp);
                            i++;
                        }

                        temp = lines[i].Trim();
                        if (temp[0] != '$')
                            break;
                        price = Str_Utils.string_to_currency(temp);
                        i++;

                        temp = lines[i].Trim();
                        if (temp[0] != '$') // subtotal
                            break;
                        i++;

                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        card.m_product_items.Add(product);

                        MyLogger.Info($"... CC-9 qty = {qty}, price = {price}, sku = {sku}, title = {title}");

                        temp = lines[i].Trim();
                        if (temp[0] == '$') // total. It means it is a last line.
                        {
                            total = Str_Utils.string_to_currency(temp);
                            card.set_total(total);
                            MyLogger.Info($"... CC-9 total = {total}");
                            break;
                        }
                    }
                }
                if (line.ToUpper() == "ITEM" && lines[i + 1].Trim().ToUpper() == "QTY" && lines[i + 2].Trim().ToUpper() == "TOTAL")
                {
                    /**
                     *      > + Item
                     *          Qty
                     *        - Total
                     *        + -->
                     *          Magnavox HD DVR/HDD 500GB
                     *          ATSC Tuner
                     *          1
                     *          $289.99
                     *        + -->
                     *          RCA Galileo Pro 11.5"
                     *          32GB 2-in-1 Tablet with Keyboard Case Android 6.0 (Marshmallow)
                     *          3
                     *          $239.94
                     *        + -->
                     *          RCA Voyager 7" 16GB Tablet
                     *          with Keyboard Case Android 6.0 (Marshmallow)
                     *          3
                     *          $134.94
                     *        + -->
                     *          RCA Voyager 7" 16GB Tablet
                     *          Android 6.0 (Marshmallow)
                     *          3
                     *          $113.94
                     *        - -->
                     *   end  + Cancellation and refund information
                     *   
                     *        + Item
                     *          Qty
                     *        - Total
                     *        + Braun Series 9 9090cc Premium Shaver &#43; Advanced Clean & Charge Station 6 pc Box
                     *          $249.00
                     *          1
                     *        - $249.00
                     *          Cancellation and refund information
                     **/

                    i += 3;

                    string temp = lines[i].Trim();
                    if (temp == "-->")
                        i++;

                    while (!lines[i].Trim().StartsWith("Cancellation", StringComparison.CurrentCultureIgnoreCase) && !lines[i].Trim().StartsWith("Canceled", StringComparison.CurrentCultureIgnoreCase))
                    {
                        string title = "";
                        string sku = "";
                        int qty = 1;
                        float price = 0;

                        int qty_tmp;
                        temp = "";
                        while (lines[i].Trim()[0] != '$' && !int.TryParse(lines[i].Trim(), out qty_tmp))
                        {
                            temp += " " + lines[i].Trim();
                            i++;
                        }
                        title = temp.Trim();

                        temp = lines[i].Trim();
                        if (temp[0] == '$')
                        {
                            if (!int.TryParse(lines[i + 1].Trim(), out qty_tmp))
                                break;
                            temp = lines[++i].Trim();
                            qty = int.Parse(temp);
                            i++;
                        }
                        else
                        {
                            qty = int.Parse(temp);
                            i++;
                        }

                        temp = lines[i].Trim();
                        if (temp[0] != '$')
                            break;
                        price = Str_Utils.string_to_currency(temp);
                        price /= qty;
                        i++;

                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        product.status = ConstEnv.REPORT_ORDER_STATUS_CANCELED;
                        card.m_product_items.Add(product);

                        MyLogger.Info($"... CC-9 qty = {qty}, price = {price}, sku = {sku}, title = {title}");

                        temp = lines[i].Trim();
                        if (temp != "-->")
                            break;
                        i++;
                    }
                }
                if (line.ToUpper() == "ITEM" && lines[i + 1].Trim().ToUpper() == "QTY")
                {
                    /**
                     *            + Item
                     *            - Qty
                     *            + Linksys E2500 N600 Dual-Band Wi-Fi Router
                     *            - 5
                     *              -->
                     *              Canceled Items - Walmart
                     *            + Item
                     *            - Qty
                     *            + Acer Swift 1 13.3" Full HD Ultra-Thin Notebook , Intel Pentium N4200, Intel UHD Graphics, 4GB, 64GB HDD, SF113-31-P5CK
                     *            - 1
                     *              Cancellation and refund information
                     **/

                    i += 2;

                    string temp;

                    string title = "";
                    string sku = "";
                    int qty = 1;
                    float price = 0;

                    int qty_tmp;
                    temp = "";
                    while (!int.TryParse(lines[i].Trim(), out qty_tmp))
                    {
                        temp += " " + lines[i].Trim();
                        i++;
                    }
                    title = temp.Trim();

                    temp = lines[i].Trim();
                    qty = int.Parse(temp);

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    card.m_product_items.Add(product);

                    MyLogger.Info($"... CC-9 qty = {qty}, price = {price}, sku = {sku}, title = {title}");
                }
            }
            card.set_total(total);
            MyLogger.Info($"... CC-9 total = {total}");
        }
    }
}
