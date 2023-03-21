using MailParser;
using Logger;
using MimeKit;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Utils;

namespace MailHelper
{
    partial class KMailBaseOP : KMailBaseParser
    {
        private List<string> extract_op6_product_title(string html_text)
        {
            List<string> titles = new List<string>();

            int start_pos = 0;
            int find_pos = html_text.IndexOf("<span id=\"productTitle\">", start_pos);
            while (find_pos != -1)
            {
                start_pos = html_text.IndexOf("<span id=\"productTitle\">", start_pos) + "<span id=\"productTitle\">".Length;
                string temp = html_text.Substring(start_pos, html_text.IndexOf("</span>", start_pos) - start_pos).Trim();
                temp = temp.Replace("\r\n", " ");
                temp = temp.Replace("\r", "");
                temp = temp.Replace("\n", "");
                titles.Add(temp);

                start_pos = html_text.IndexOf("</span>", start_pos) + "</span>".Length;
                find_pos = html_text.IndexOf("<span id=\"productTitle\">", start_pos);
            }
            return titles;
        }
        private List<float> extract_op6_product_price(string html_text)
        {
            List<float> prices = new List<float>();

            int start_pos = 0;
            int find_pos = html_text.IndexOf("<span id=\"productDiscountedPrice\">", start_pos);
            while (find_pos != -1)
            {
                start_pos = html_text.IndexOf("<span id=\"productDiscountedPrice\">", start_pos) + "<span id=\"productDiscountedPrice\">".Length;
                string temp = html_text.Substring(start_pos, html_text.IndexOf("</span>", start_pos) - start_pos).Trim();
                temp = temp.Replace("\r\n", " ");
                temp = temp.Replace("\r", "");
                temp = temp.Replace("\n", "");
                float f = Str_Utils.string_to_currency(temp);
                prices.Add(f);

                start_pos = html_text.IndexOf("</span>", start_pos) + "</span>".Length;
                find_pos = html_text.IndexOf("<span id=\"productDiscountedPrice\">", start_pos);
            }
            return prices;
        }
        private List<int> extract_op6_product_qty(string html_text)
        {
            List<int> qtys = new List<int>();

            int start_pos = 0;
            int find_pos = html_text.IndexOf("<span id=\"productQuantity\">", start_pos);
            while (find_pos != -1)
            {
                start_pos = html_text.IndexOf("<span id=\"productQuantity\">", start_pos) + "<span id=\"productQuantity\">".Length;
                string temp = html_text.Substring(start_pos, html_text.IndexOf("</span>", start_pos) - start_pos).Trim();
                temp = temp.Replace("\r\n", " ");
                temp = temp.Replace("\r", "");
                temp = temp.Replace("\n", "");
                int qty = Str_Utils.string_to_int(temp);
                qtys.Add(qty);

                start_pos = html_text.IndexOf("</span>", start_pos) + "</span>".Length;
                find_pos = html_text.IndexOf("<span id=\"productQuantity\">", start_pos);
            }
            return qtys;
        }
        private List<string> extract_op6_product_sku(string html_text)
        {
            List<string> skus = new List<string>();

            int start_pos = 0;
            int find_pos = html_text.IndexOf("<span id=\"productSKU\">", start_pos);
            while (find_pos != -1)
            {
                start_pos = html_text.IndexOf("<span id=\"productSKU\">", start_pos) + "<span id=\"productSKU\">".Length;
                string temp = html_text.Substring(start_pos, html_text.IndexOf("</span>", start_pos) - start_pos).Trim();
                temp = temp.Replace("\r\n", " ");
                temp = temp.Replace("\r", "");
                temp = temp.Replace("\n", "");
                skus.Add(temp);

                start_pos = html_text.IndexOf("</span>", start_pos) + "</span>".Length;
                find_pos = html_text.IndexOf("<span id=\"productSKU\">", start_pos);
            }
            return skus;
        }
        private void parse_mail_op_6(MimeMessage mail, KReportOP report)
        {
            if (!XMailHelper.is_bodytext_existed(mail))
            {
                parse_mail_op_6_for_htmltext(mail, report);
                return;
            }
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_6;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = "DICK'S Sporting Goods";

            MyLogger.Info($"... OP-6 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-6 m_op_retailer = {report.m_retailer}");

            string html_text = XMailHelper.get_htmltext(mail);

            List<string> title_list = extract_op6_product_title(html_text);
            List<int> qty_list = extract_op6_product_qty(html_text);
            List<float> price_list = extract_op6_product_price(html_text);
            List<string> sku_list = extract_op6_product_sku(html_text);

            int n = Math.Min(title_list.Count, qty_list.Count);
            n = Math.Min(n, price_list.Count);
            n = Math.Min(n, sku_list.Count);
            for (int i = 0; i < n; i++)
            {

                ZProduct product = new ZProduct();
                product.price = price_list[i];
                product.sku = sku_list[i];
                product.title = title_list[i];
                product.qty = qty_list[i];
                report.m_product_items.Add(product);

                MyLogger.Info($"... OP-6 qty = {qty_list[i]}, price = {price_list[i]}, sku = {sku_list[i]}, item title = {title_list[i]}");
            }

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Order #"))
                {
                    string temp = line.Substring("Order #".Length).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... OP-6 order id = {temp}");
                    continue;
                }
                if (line.StartsWith("Order Placed"))
                {
                    string temp = line.Substring("Order Placed".Length).Trim();
                    DateTime date = DateTime.Parse(temp);
                    report.m_op_purchase_date = date;
                    MyLogger.Info($"... OP-6 order date = {date}");
                    continue;
                }
                if (line == "Estimated Tax")
                {
                    string temp = lines[++i].ToString();
                    float tax = Str_Utils.string_to_currency(temp);
                    report.m_tax = tax;
                    MyLogger.Info($"... OP-6 tax = {tax}");
                    continue;
                }
                if (line == "Estimated Order Total")
                {
                    string temp = lines[++i].ToString();
                    float total = Str_Utils.string_to_currency(temp);
                    report.set_total(total);
                    MyLogger.Info($"... OP-6 total = {total}");
                    continue;
                }
                if (line == "Payment Information")
                {
                    string next_line = lines[++i].Trim();
                    while (i < lines.Length && next_line != "")
                    {
                        string temp = next_line; // payment card image url

                        temp = lines[++i].Trim();
                        if (temp.IndexOf("x") != -1)
                        {
                            string last_digit = temp.Substring(temp.LastIndexOf("x") + 1).Trim();

                            ZPaymentCard c = new ZPaymentCard("", last_digit, 0);
                            report.add_payment_card_info(c);

                            MyLogger.Info($"... OP-6 payment_type = \"\", last_digit = {last_digit}, price = 0");
                        }
                        next_line = lines[++i].Trim();
                    }
                    continue;
                }
            }
        }
        private void parse_mail_op_6_for_htmltext(MimeMessage mail, KReportOP report)
        {
            bool get_order_id = false;
            bool get_date = false;
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_6;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = "DICK'S Sporting Goods";

            MyLogger.Info($"... OP-6 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-6 m_op_retailer = {report.m_retailer}");

            string html_text = XMailHelper.get_htmltext(mail);

            List<string> title_list = extract_op6_product_title(html_text);
            List<int> qty_list = extract_op6_product_qty(html_text);
            List<float> price_list = extract_op6_product_price(html_text);
            List<string> sku_list = extract_op6_product_sku(html_text);

            int n = Math.Min(title_list.Count, qty_list.Count);
            n = Math.Min(n, price_list.Count);
            n = Math.Min(n, sku_list.Count);
            for (int i = 0; i < n; i++)
            {

                ZProduct product = new ZProduct();
                product.price = price_list[i];
                product.sku = sku_list[i];
                product.title = title_list[i];
                product.qty = qty_list[i];
                report.m_product_items.Add(product);

                MyLogger.Info($"... OP-6 qty = {qty_list[i]}, price = {price_list[i]}, sku = {sku_list[i]}, item title = {title_list[i]}");
            }

            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Order #"))
                {
                    string temp = line.Substring("Order #".Length).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... OP-6 order id = {temp}");
                    get_order_id = true;
                    continue;
                }
                if (!get_order_id && line.IndexOf("Order #", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    string temp = line.Substring(line.IndexOf("Order #", StringComparison.CurrentCultureIgnoreCase) + "Order #".Length).TrimStart();
                    if (temp.IndexOf(" ") != -1)
                        temp = temp.Substring(0, temp.IndexOf(" "));
                    report.set_order_id(temp);
                    MyLogger.Info($"... OP-6 order id = {temp}");
                    get_order_id = true;
                    continue;
                }
                if (line.StartsWith("Order Placed"))
                {
                    string temp = line.Substring("Order Placed".Length).Trim();
                    DateTime date = DateTime.Parse(temp);
                    report.m_op_purchase_date = date;
                    MyLogger.Info($"... OP-6 order date = {date}");
                    get_date = true;
                    continue;
                }
                if (!get_date && line.IndexOf("Order Placed", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    string temp = line.Substring(line.IndexOf("Order Placed", StringComparison.CurrentCultureIgnoreCase) + "Order Placed".Length).Trim();
                    if (temp.IndexOf(" ") != -1)
                        temp = temp.Substring(0, temp.IndexOf(" "));
                    temp = temp.Substring(0, 8); // MM/dd/yy
                    DateTime date;
                    if (!DateTime.TryParse(temp, out date))
                    {
                        DateTimeFormatInfo formatProvider = new DateTimeFormatInfo();
                        formatProvider.Calendar.TwoDigitYearMax = DateTime.Now.Year;
                        if (!DateTime.TryParseExact(temp, "MM/dd/yy", formatProvider, DateTimeStyles.AllowLeadingWhite, out date))
                            continue;
                    }
                    report.m_op_purchase_date = date;
                    MyLogger.Info($"... OP-6 order date = {date}");
                    get_date = true;
                    continue;
                }
                if (line == "Estimated Tax")
                {
                    string temp = lines[++i].ToString();
                    float tax = Str_Utils.string_to_currency(temp);
                    report.m_tax = tax;
                    MyLogger.Info($"... OP-6 tax = {tax}");
                    continue;
                }
                if (line == "Estimated Order Total")
                {
                    string temp = lines[++i].ToString();
                    float total = Str_Utils.string_to_currency(temp);
                    report.set_total(total);
                    MyLogger.Info($"... OP-6 total = {total}");
                    continue;
                }
                if (line == "Payment Information")
                {
                    string next_line = lines[++i].Trim();
                    while (i < lines.Length)
                    {
                        string temp = next_line; // payment card image url

                        if (temp.IndexOf("x") == -1)
                            break;
                        string last_digit = temp.Substring(temp.LastIndexOf("x") + 1).Trim();

                        ZPaymentCard c = new ZPaymentCard("", last_digit, 0);
                        report.add_payment_card_info(c);

                        MyLogger.Info($"... OP-6 payment_type = \"\", last_digit = {last_digit}, price = 0");

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
                        MyLogger.Info($"... OP-6 full_address = {full_address}, state_address = {state_address}");
                    }
                    continue;
                }
                if (report.m_product_items.Count == 0)
                {
                    if (line == "Quantity" && lines[i - 1].Trim() == "Each" && lines[i - 2].Trim() == "Product" && lines[i + 1].Trim() == "Total")
                    {
                        string title = "";
                        string sku = "";
                        int qty = 0;
                        float price = 0;

                        // To Do : for multiple items.

                        i += 2;

                        string temp = lines[i].Trim();
                        while (temp[0] != '$')
                        {
                            title += temp + " ";
                            i++;
                            temp = lines[i].Trim();
                        }
                        title = title.Trim();
                        if (temp.IndexOf("$", 1) != -1)
                            temp = temp.Substring(1, temp.IndexOf("$", 1)).Trim();
                        price = Str_Utils.string_to_currency(temp);

                        i++;
                        if (lines[i].Trim()[0] == '$')
                            i++;
                        temp = lines[i].Trim();
                        qty = Str_Utils.string_to_int(temp);

                        i++;
                        temp = lines[i].Trim();
                        while (!temp.StartsWith("SKU:") && temp.IndexOf("Delivery") == -1)
                        {
                            i++;
                            temp = lines[i].Trim();
                        }
                        if (temp.StartsWith("SKU:"))
                            sku = temp.Substring("SKU:".Length).Trim();

                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        report.m_product_items.Add(product);

                        MyLogger.Info($"... OP-9 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");

                        continue;
                    }
                }
            }
        }
    }
}
