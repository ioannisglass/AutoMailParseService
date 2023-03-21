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
        private void parse_mail_cc_5(MimeMessage mail, KReportCC card)
        {
            string subject = XMailHelper.get_subject(mail);

            card.m_mail_type = KReportBase.MailType.CC_5;

            card.m_retailer = ConstEnv.RETAILER_HOMEDEPOT;

            MyLogger.Info($"... CC-5 m_cc_retailer = {card.m_retailer}");

            int items_type = 0;

            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.ToUpper() == "ITEM" && lines[i + 1].Trim().ToUpper() == "UNIT PRICE" && lines[i + 2].Trim().ToUpper() == "QTY" && lines[i + 3].Trim().ToUpper() == "ITEM TOTAL")
                {
                    items_type = 0;
                    break;
                }
                if (line.ToUpper() == "PRODUCT" && lines[i + 1].Trim().ToUpper() == "CANCELLED" && lines[i + 2].Trim().ToUpper() == "QUANTITY" && lines[i + 3].Trim().ToUpper() == "UNIT"
                     && lines[i + 4].Trim().ToUpper() == "PRICE" && lines[i + 5].Trim().ToUpper() == "CANCELLED" && lines[i + 6].Trim().ToUpper() == "AMOUNT")
                {
                    items_type = 1;
                    break;
                }
                if (line.ToUpper() == "PRODUCT" && lines[i + 1].Trim().ToUpper() == "CANCELLED QUANTITY" && lines[i + 2].Trim().ToUpper() == "UNIT PRICE" && lines[i + 3].Trim().ToUpper() == "CANCELLED AMOUNT")
                {
                    items_type = 1;
                    break;
                }
                if (line.ToUpper() == "QTY" && lines[i + 1].Trim().ToUpper() == "ORDERED" && lines[i + 2].Trim().ToUpper() == "INTERNET #" && lines[i + 3].Trim().ToUpper() == "PRODUCT DESCRIPTION"
                     && lines[i + 4].Trim().ToUpper() == "UNIT PRICE" && lines[i + 5].Trim().ToUpper() == "AMOUNT")
                {
                    items_type = 2;
                    break;
                }
                if (line.ToUpper() == "QTY ORDERED" && lines[i + 1].Trim().ToUpper() == "INTERNET #" && lines[i + 2].Trim().ToUpper() == "PRODUCT DESCRIPTION"
                     && lines[i + 3].Trim().ToUpper() == "UNIT PRICE" && lines[i + 4].Trim().ToUpper() == "AMOUNT")
                {
                    items_type = 2;
                    break;
                }
                if (line.ToUpper() == "PRODUCT DESCRIPTION" && lines[i + 1].Trim().ToUpper() == "UNIT PRICE" && lines[i + 2].Trim().ToUpper() == "QTY" && lines[i + 3].Trim().ToUpper() == "ITEM TOTAL")
                {
                    items_type = 3;
                    break;
                }
            }

            float total = 0;
            lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.ToUpper() == "ORDER NUMBER:" && lines[i + 1].Trim().ToUpper() == "ORDER DATE:")
                {
                    string temp;

                    i += 2;
                    temp = lines[i].Trim();
                    card.set_order_id(temp);
                    MyLogger.Info($"... CC-5 order id = {temp}");
                    continue;
                }
                if (line.StartsWith("Order Number:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "ORDER NUMBER:")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Order Number:".Length).Trim();
                    card.set_order_id(temp);
                    MyLogger.Info($"... CC-5 order id = {temp}");
                    continue;
                }
                if (line.StartsWith("Order Number", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "ORDER NUMBER")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Order Number".Length).Trim();
                    card.set_order_id(temp);
                    MyLogger.Info($"... CC-5 order id = {temp}");
                    continue;
                }
                if (line.StartsWith("Order ID:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "Order ID:")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Order ID:".Length).Trim();
                    card.set_order_id(temp);
                    MyLogger.Info($"... CC-5 order id = {temp}");
                    continue;
                }
                if (items_type == 0)
                {
                    /**
                     *                  + Item
                     *                    Unit Price
                     *                    Qty
                     *                  - Item Total
                     *                  + M12 12-Volt Lithium-Ion Cordless 1000 Lumens ROVER LED Compact Flood Light with M12 Jobsite Speaker and 3.0 Ah Battery
                     *                    Internet # 309029372
                     *                    $227.00
                     *                    2
                     *                  - $454.00
                     *  repeated        + Item
                     *  (no visible)      M12 12-Volt Lithium-Ion Cordless 1000 Lumens ROVER LED Compact Flood Light with M12 Jobsite Speaker and 3.0 Ah Battery
                     *                    Internet # 309029372
                     *                    Unit Price   $227.00
                     *                  > Qty   2
                     *                  - Item Total   $454.00
                     * 
                     * 
                     *                  + Item
                     *                    Unit Price
                     *                    Qty
                     *                  - Item Total
                     *                  + M12 FUEL 12-Volt Li-Ion Brushless Cordless Hammer Drill and Impact Driver Combo Kit (2-Tool)w/ Free M12 3/8 in. Ratchet
                     *                    Internet #
                     *                    305883849
                     *                    Unit Price
                     *                    $348.00
                     *                    Qty
                     *                    5.0
                     *                    Item Total
                     *                  - $1,740.00
                     *                  
                     *                  
                     *                  + Item
                     *                    UnitPrice
                     *                    Qty
                     *                  - Item Total
                     *                  + M18 18-Volt Lithium-Ion Cordless HACKZALL Reciprocating Saw Kit with Free M18 4.0Ah Extended Capacity Battery
                     *                    Internet #: 202795928
                     *                    Qty: 5.0Unit Price: $318.00Item Total: $1,590.00
                     *                    $318.00
                     *                    5.0
                     *                  - $1,590.00
                     *
                     **/

                    if (line.StartsWith("Qty", StringComparison.CurrentCultureIgnoreCase) && line.IndexOf("Unit Price:", StringComparison.CurrentCultureIgnoreCase) != -1 && line.IndexOf("Item Total:", StringComparison.CurrentCultureIgnoreCase) != -1)
                    {
                        string title = "";
                        string sku = "";
                        int qty = 1;
                        float price = 0;
                        int k;

                        string temp;

                        temp = line.Substring("Qty:".Length).Trim();
                        qty = Str_Utils.string_to_int(temp.Substring(0, temp.IndexOf("Unit Price:", StringComparison.CurrentCultureIgnoreCase)).Trim());

                        temp = temp.Substring(temp.IndexOf("Unit Price:", StringComparison.CurrentCultureIgnoreCase) + "Unit Price:".Length).Trim();
                        price = Str_Utils.string_to_currency(temp.Substring(0, temp.IndexOf("Item Total:", StringComparison.CurrentCultureIgnoreCase)).Trim());

                        temp = temp.Substring(temp.IndexOf("Item Total:", StringComparison.CurrentCultureIgnoreCase) + "Item Total:".Length).Trim();
                        total += Str_Utils.string_to_currency(temp.Trim());

                        k = i - 1;
                        temp = lines[k].Trim();
                        if (temp.StartsWith("Internet #:", StringComparison.CurrentCultureIgnoreCase) && temp.ToUpper() != "INTERNET #:")
                        {
                            temp = temp.Substring("Internet #:".Length).Trim();
                            sku = temp;
                            k--;
                        }

                        temp = lines[k].Trim();
                        title = temp;

                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        card.m_product_items.Add(product);

                        MyLogger.Info($"... CC-5 qty = {qty}, price = {price}, sku = {sku}, title = {title}");
                        continue;
                    }
                    else if (line.StartsWith("Qty", StringComparison.CurrentCultureIgnoreCase) && line.ToUpper() != "QTY")
                    {
                        string title = "";
                        string sku = "";
                        int qty = 1;
                        float price = 0;
                        int k;

                        string temp;

                        temp = line.Substring("Qty".Length).Trim();
                        qty = Str_Utils.string_to_int(temp);

                        k = i - 1;
                        temp = lines[k].Trim();
                        if (temp.StartsWith("Unit Price", StringComparison.CurrentCultureIgnoreCase) && temp.ToUpper() != "UNIT PRICE")
                        {
                            temp = temp.Substring("UNIT PRICE".Length).Trim();
                            price = Str_Utils.string_to_currency(temp);
                            k--;
                        }
                        temp = lines[k].Trim();
                        if (temp.StartsWith("Internet #", StringComparison.CurrentCultureIgnoreCase) && temp.ToUpper() != "INTERNET #")
                        {
                            temp = temp.Substring("Internet #".Length).Trim();
                            sku = temp;
                            k--;
                        }
                        temp = lines[k].Trim();
                        if (temp.StartsWith("Store SKU #", StringComparison.CurrentCultureIgnoreCase) && temp.ToUpper() != "STORE SKU #")
                        {
                            temp = temp.Substring("Store SKU #".Length).Trim();
                            if (sku == "")
                                sku = temp;
                            k--;
                        }

                        while (k >= 0)
                        {
                            if (lines[k].Trim().ToUpper() == "ITEM")
                                break;
                            if (lines[k].Trim().IndexOf("$") != -1)
                                break;

                            title = lines[k].Trim() + " " + title;
                            k--;
                        }
                        title = title.Trim();

                        if (lines[i + 1].Trim().StartsWith("Item Total", StringComparison.CurrentCultureIgnoreCase))
                        {
                            temp = lines[++i].Trim();
                            temp = temp.Substring("Item Total".Length).Trim();
                            float f = Str_Utils.string_to_currency(temp);
                            total += f;
                        }

                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        card.m_product_items.Add(product);

                        MyLogger.Info($"... CC-5 qty = {qty}, price = {price}, sku = {sku}, title = {title}");
                        continue;
                    }
                    else if (line.ToUpper() == "UNIT PRICE" && lines[i + 2].Trim().ToUpper() == "QTY" && lines[i + 4].Trim().ToUpper() == "ITEM TOTAL")
                    {
                        string title = "";
                        string sku = "";
                        int qty = 1;
                        float price = 0;

                        string temp = lines[i + 1].Trim();
                        price = Str_Utils.string_to_currency(temp);

                        temp = lines[i + 3].Trim();
                        qty = Str_Utils.string_to_int(temp);

                        temp = lines[i + 5].Trim();
                        float f = Str_Utils.string_to_currency(temp);
                        total += f;

                        int k = i - 2;
                        temp = lines[k].Trim();
                        if (temp.StartsWith("INTERNET #", StringComparison.CurrentCultureIgnoreCase))
                        {
                            sku = lines[k + 1].Trim();
                            k -= 2;
                        }

                        temp = lines[k].Trim();
                        if (temp.StartsWith("Store SKU #", StringComparison.CurrentCultureIgnoreCase))
                        {
                            if (sku == "")
                                sku = lines[k + 1].Trim();
                            k -= 2;
                        }

                        temp = lines[k + 1].Trim();
                        title = temp;

                        i += 5;

                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        card.m_product_items.Add(product);

                        MyLogger.Info($"... CC-5 qty = {qty}, price = {price}, sku = {sku}, title = {title}");
                        continue;
                    }
                }
                else if (items_type == 1)
                {
                    if (line.ToUpper() == "INTERNET #:" && lines[i + 2].Trim().StartsWith("Cancelled Quantity:", StringComparison.CurrentCultureIgnoreCase)
                         && lines[i + 3].Trim().StartsWith("Unit Price:", StringComparison.CurrentCultureIgnoreCase)
                         && lines[i + 4].Trim().StartsWith("Cancelled Amount:", StringComparison.CurrentCultureIgnoreCase))
                    {
                        /**
                         *    + 1200 Series Single-Handle Pull-Down Sprayer Kitchen Faucet in Stainless Steel
                         *      Internet #:
                         *      204683817
                         *      Cancelled Quantity: 2
                         *      Unit Price: $179.00
                         *      Cancelled Amount: $358.00
                         *      2
                         *      $179.00
                         *      $358.00
                         *    * Promotion Discount:
                         *    - -$89.50
                         *    + 1-Spray 8 x 6 in. Rectangular Showerhead in Chrome
                         *      Internet #:
                         *      203893592
                         *      Cancelled Quantity: 1
                         *      Unit Price: $24.98
                         *      Cancelled Amount: $24.98
                         *      1
                         *      $24.98
                         *    - $24.98
                         *      You can
                         *      check your order status online at any time.
                         **/

                        string title = "";
                        string sku = "";
                        int qty = 1;
                        float price = 0;
                        int k = i - 1;

                        string temp;

                        sku = lines[i + 1].Trim();

                        temp = "";
                        while (k >= 0 && lines[k].Trim().IndexOf("$") == -1 && lines[k].Trim().ToUpper() != "AMOUNT")
                        {
                            temp = lines[k].Trim() + " " + temp;
                            k--;
                        }
                        if (k < 0)
                            continue;
                        title = temp.Trim();

                        temp = lines[i + 2].Trim();
                        temp = temp.Substring("Cancelled Quantity:".Length).Trim();
                        qty = Str_Utils.string_to_int(temp);

                        temp = lines[i + 3].Trim();
                        temp = temp.Substring("Unit Price:".Length).Trim();
                        price = Str_Utils.string_to_currency(temp);

                        temp = lines[i + 4].Trim();
                        temp = temp.Substring("Cancelled Amount:".Length).Trim();
                        float t = Str_Utils.string_to_currency(temp);
                        total += t;

                        i += 4;

                        if (lines[i + 4].Trim().EndsWith("Discount:", StringComparison.CurrentCultureIgnoreCase) && lines[i + 5].Trim().IndexOf("$") != -1)
                        {
                            temp = lines[i + 5].Trim();
                            t = Str_Utils.string_to_currency(temp);
                            total += t;
                            i += 5;
                        }

                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        card.m_product_items.Add(product);

                        MyLogger.Info($"... CC-5 qty = {qty}, price = {price}, sku = {sku}, title = {title}");
                        continue;
                    }
                }
                else if (items_type == 3)
                {
                    /**
                     *           Title    + Product Description
                     *                      Unit Price
                     *                      Qty
                     *                    - Item Total
                     *                    + 5-2
                     *                      Day Programmable Thermostat with Backlight
                     *                      Store SKU # 514261
                     *                      Internet # 203539496
                     *                    > Unit Price
                     *                      $32.97
                     *                      Qty
                     *                      5.00
                     *                      Item Total
                     *                      $164.85
                     *                      $24.98
                     *                      1.00
                     *                    - $24.98
                     *                    + 65W
                     *                      Equivalent Daylight BR30 Dimmable LED Light Bulb (6-Pack)
                     *                      Store SKU # 1001654247
                     *                      Internet # 206702062
                     *                      Unit Price
                     *                      $32.97
                     *                      Qty
                     *                      5.00
                     *                      Item Total
                     *                      $164.85
                     *                      $32.97
                     *                      5.00
                     *                      $164.85
                     * 
                     *
                     **/

                    if (line.ToUpper() == "UNIT PRICE" && lines[i + 2].Trim().ToUpper() == "QTY" && lines[i + 4].Trim().ToUpper() == "ITEM TOTAL")
                    {
                        string title = "";
                        string sku = "";
                        int qty = 1;
                        float price = 0;

                        int k = i - 1;
                        string temp = lines[k].Trim();
                        if (temp.StartsWith("Internet #", StringComparison.CurrentCultureIgnoreCase))
                        {
                            sku = temp.Substring("Internet #".Length).Trim();
                            k--;
                        }
                        temp = lines[k].Trim();
                        if (temp.StartsWith("Store SKU #", StringComparison.CurrentCultureIgnoreCase))
                        {
                            if (sku == "")
                                sku = temp.Substring("Store SKU #".Length).Trim();
                            k--;
                        }

                        temp = "";
                        while (k >= 0 && lines[k].Trim().IndexOf("$") == -1 && lines[k].Trim().ToUpper() != "ITEM TOTAL")
                        {
                            temp = lines[k].Trim() + " " + temp;
                            k--;
                        }
                        if (k < 0)
                            continue;
                        title = temp.Trim();

                        i += 5;

                        // We must ignore the first line.
                        i++;

                        temp = lines[i++].Trim();
                        price = Str_Utils.string_to_currency(temp);

                        temp = lines[i++].Trim();
                        qty = Str_Utils.string_to_int(temp);

                        temp = lines[i].Trim();
                        float f = Str_Utils.string_to_currency(temp);
                        total += f;

                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        card.m_product_items.Add(product);

                        MyLogger.Info($"... CC-5 qty = {qty}, price = {price}, sku = {sku}, title = {title}");
                        continue;
                    }
                }
            }

            if (items_type == 2)
            {
                get_cc5_items_for_type2(mail, card);
            }

            if (card.m_order_id == "")
            {
                /**
                 * Some mails does not have CrLf, namely is few lines body text, few lines body html. (I found 2 mails Oct 31, 2019 now)
                 * We will call that mail as type_5
                 **/

                get_cc5_items_for_type5(mail, card);
            }

            card.set_total(total);
            MyLogger.Info($"... CC-5 total = {total}");
        }
        private void get_cc5_items_for_type2(MimeMessage mail, KReportCC card)
        {
            string html_text = XMailHelper.get_htmltext(mail);
            int next_pos;
            string temp;
            int pos;
            float total = 0;

            pos = html_text.IndexOf("Qty", StringComparison.CurrentCultureIgnoreCase);
            if (pos == -1)
                return;

            pos = html_text.IndexOf("</tr>", pos);
            if (pos == -1)
                return;

            temp = html_text.Substring(pos + "</tr>".Length);
            pos = temp.IndexOf("</tbody>");
            if (pos == -1)
                return;
            temp = temp.Substring(0, pos);

            string tr_part = XMailHelper.find_html_part(temp, "tr", out next_pos);
            while (tr_part != "")
            {
                string title = "";
                string sku = "";
                int qty = 1;
                float price = 0;

                int next_pos1;
                string temp1 = tr_part;
                string td_part = XMailHelper.find_html_part(temp1, "td", out next_pos1);
                temp1 = temp1.Substring(next_pos1);
                if (td_part == "")
                    break;

                string td_text = XMailHelper.html2text(td_part);
                qty = Str_Utils.string_to_int(td_text);

                td_part = XMailHelper.find_html_part(temp1, "td", out next_pos1);
                temp1 = temp1.Substring(next_pos1);
                if (td_part == "")
                    break;
                td_text = XMailHelper.html2text(td_part);
                td_text = td_text.Replace("\r\n", " ");
                td_text = td_text.Replace("\n", " ");
                td_text = td_text.Trim();
                sku = td_text;

                td_part = XMailHelper.find_html_part(temp1, "td", out next_pos1);
                temp1 = temp1.Substring(next_pos1);
                if (td_part == "")
                    break;
                td_text = XMailHelper.html2text(td_part);
                td_text = td_text.Replace("\r\n", " ");
                td_text = td_text.Replace("\n", " ");
                td_text = td_text.Trim();
                title = td_text;

                td_part = XMailHelper.find_html_part(temp1, "td", out next_pos1);
                temp1 = temp1.Substring(next_pos1);
                if (td_part == "")
                    break;
                td_text = XMailHelper.html2text(td_part);
                price = Str_Utils.string_to_currency(td_text);

                td_part = XMailHelper.find_html_part(temp1, "td", out next_pos1);
                temp1 = temp1.Substring(next_pos1);
                if (td_part == "")
                    break;
                td_text = XMailHelper.html2text(td_part);
                float f = Str_Utils.string_to_currency(td_text);
                total += f;

                ZProduct product = new ZProduct();
                product.price = price;
                product.sku = sku;
                product.title = title;
                product.qty = qty;
                card.m_product_items.Add(product);

                MyLogger.Info($"... CC-5 qty = {qty}, price = {price}, sku = {sku}, title = {title}");

                temp = temp.Substring(next_pos);
                tr_part = XMailHelper.find_html_part(temp, "tr", out next_pos);
            }

            card.set_total(total);
            MyLogger.Info($"... CC-5 total = {total}");
        }
        private void get_cc5_items_for_type5(MimeMessage mail, KReportCC card)
        {
            string html_text = XMailHelper.get_htmltext(mail);
            int next_pos;
            string temp;
            int pos;
            float total = 0;

            // Find order id.

            pos = html_text.IndexOf("Order Number:", StringComparison.CurrentCultureIgnoreCase);
            if (pos == -1)
                return;
            temp = html_text.Substring(pos);
            temp = XMailHelper.find_html_part(temp, "span", out next_pos);
            if (temp == "")
                return;
            card.set_order_id(temp);
            MyLogger.Info($"... CC-5 Order Id = {card.m_order_id}");

            // Find items.
            // Html has the items twice. We must ignore second list.

            pos = html_text.IndexOf("<!-- Start Instore Order -->", StringComparison.CurrentCultureIgnoreCase);
            if (pos == -1)
                return;
            temp = html_text.Substring(pos + "<!-- Start Instore Order -->".Length);

            pos = temp.IndexOf("<div class=\"visible_mobile_view mobile_item_list\" style=\"display:none;width:100%;\">");
            if (pos == -1)
                return;
            html_text = temp.Substring(0, pos);

            int sku_pos = html_text.IndexOf("Internet #", StringComparison.CurrentCultureIgnoreCase);
            while (sku_pos != -1)
            {
                string title = "";
                string sku = "";
                int qty = 1;
                float price = 0;

                pos = html_text.LastIndexOf("<td", sku_pos);
                if (pos == -1)
                    break;
                html_text = html_text.Substring(pos);

                string td_part = XMailHelper.find_html_part(html_text, "td", out next_pos);
                string td_text = XMailHelper.html2text(td_part);
                td_text = td_text.Replace("\r\n", " ");
                td_text = td_text.Replace("\n", " ");
                td_text = td_text.Trim();
                if (td_text == "" || td_text.IndexOf("Internet #", StringComparison.CurrentCultureIgnoreCase) == -1)
                    break;

                title = td_text.Substring(0, td_text.IndexOf("Internet #", StringComparison.CurrentCultureIgnoreCase)).Trim();
                sku = td_text.Substring(td_text.IndexOf("Internet #", StringComparison.CurrentCultureIgnoreCase) + "Internet #".Length).Trim();

                html_text = html_text.Substring(next_pos);
                td_part = XMailHelper.find_html_part(html_text, "td", out next_pos);
                td_text = XMailHelper.html2text(td_part);
                if (td_text == "")
                    break;
                price = Str_Utils.string_to_currency(td_text);

                html_text = html_text.Substring(next_pos);
                td_part = XMailHelper.find_html_part(html_text, "td", out next_pos);
                td_text = XMailHelper.html2text(td_part);
                if (td_text == "")
                    break;
                qty = Str_Utils.string_to_int(td_text);

                html_text = html_text.Substring(next_pos);
                td_part = XMailHelper.find_html_part(html_text, "td", out next_pos);
                td_text = XMailHelper.html2text(td_part);
                if (td_text == "")
                    break;
                float f = Str_Utils.string_to_currency(td_text);
                total += f;

                ZProduct product = new ZProduct();
                product.price = price;
                product.sku = sku;
                product.title = title;
                product.qty = qty;
                card.m_product_items.Add(product);

                MyLogger.Info($"... CC-5 qty = {qty}, price = {price}, sku = {sku}, title = {title}");

                html_text = html_text.Substring(next_pos);
                sku_pos = html_text.IndexOf("Internet #", StringComparison.CurrentCultureIgnoreCase);
            }


            card.set_total(total);
            MyLogger.Info($"... CC-5 total = {total}");
        }
    }
}
