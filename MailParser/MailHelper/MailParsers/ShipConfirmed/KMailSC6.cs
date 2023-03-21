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
    partial class KMailBaseSC : KMailBaseParser
    {
        private void parse_mail_sc_6(MimeMessage mail, KReportSC report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.SC_6;

            report.m_retailer = "DICK's Sporting Goods";

            MyLogger.Info($"... SC-6 m_sc_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.IndexOf("Order #", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    string temp = line.Substring(line.IndexOf("Order #", StringComparison.CurrentCultureIgnoreCase) + "Order #".Length).Trim();
                    if (temp.IndexOf(" ") != -1)
                        temp = temp.Substring(0, temp.IndexOf(" ")).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... SC-6 order id = {temp}");
                    continue;
                }
            }

            string html_text = XMailHelper.get_htmltext(mail);

            DateTime date = extract_sc6_expected_delivery_date(html_text);
            if (date != DateTime.MinValue)
            {
                report.m_sc_expected_deliver_date = date;
                MyLogger.Info($"... SC-6 expected delivery date = {date}");
            }
            string tracking = extract_sc6_tracking(html_text);
            if (tracking != "")
            {
                report.set_tracking(tracking);
                MyLogger.Info($"... SC-6 tracking = {tracking}");
            }

            List<string> title_list = extract_sc6_product_title(html_text);
            List<int> qty_list = extract_sc6_product_qty(html_text);
            List<float> price_list = extract_sc6_product_price(html_text);
            List<string> sku_list = extract_sc6_product_sku(html_text);

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

                MyLogger.Info($"... SC-6 qty = {qty_list[i]}, price = {price_list[i]}, sku = {sku_list[i]}, item title = {title_list[i]}");
            }

            if (report.m_sc_tracking == "" || report.m_product_items.Count == 0)
                parse_sc6_from_bodytext(mail, report);
        }
        private void parse_sc6_from_bodytext(MimeMessage mail, KReportSC card)
        {
            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line == "Tracking Number:")
                {
                    string temp = lines[++i].Trim();
                    card.set_tracking(temp);
                    MyLogger.Info($"... SC-6 tracking = {temp}");
                    continue;
                }
                if (line.StartsWith("Threshold Delivery:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp = line.Substring("Threshold Delivery:".Length).Trim();
                    card.m_sc_expected_deliver_date = DateTime.Parse(temp);
                    MyLogger.Info($"... SC-6 (from bodytext) expected delivery date = {card.m_sc_expected_deliver_date}");
                    continue;
                }
                if (line == "Product" && lines[i + 1].Trim() == "Each" && lines[i + 2].Trim() == "Quantity" && lines[i + 3].Trim() == "Total")
                {
                    i += 4;
                    string nextline = lines[i].Trim();
                    while (i < lines.Length)
                    {
                        string title = "";
                        string sku = "";
                        int qty = 0;
                        float price = 0;

                        string temp = nextline;

                        if (temp.StartsWith("Earn $", StringComparison.CurrentCultureIgnoreCase))
                            break;
                        if (temp.StartsWith("Buy Online Pick Up ", StringComparison.CurrentCultureIgnoreCase))
                            break;
                        if (temp.StartsWith("Receive a", StringComparison.CurrentCultureIgnoreCase))
                            break;
                        if (temp.StartsWith("Flash Sale", StringComparison.CurrentCultureIgnoreCase))
                            break;
                        if (temp.StartsWith("Oversized shipping ", StringComparison.CurrentCultureIgnoreCase))
                            break;
                        if (temp.StartsWith("Standard Delivery", StringComparison.CurrentCultureIgnoreCase))
                            break;
                        if (temp.StartsWith("Tracking Number:", StringComparison.CurrentCultureIgnoreCase))
                            break;

                        title = temp;

                        if (lines[i + 1].Trim().IndexOf("$") != -1) // maybe it has no unit price, so we must check it.
                        {
                            temp = lines[++i].Trim();
                            if (temp[0] == '$')
                                temp = temp.Substring(1);
                            int pos = temp.IndexOf("$");
                            if (pos != -1)
                                temp = temp.Substring(0, pos);
                            price = Str_Utils.string_to_currency(temp);
                        }

                        if (lines[i + 1].Trim().IndexOf("$") == -1 && !lines[i + 1].Trim().StartsWith("SKU:", StringComparison.CurrentCultureIgnoreCase)) // maybe it has no qty, so we must check it.
                        {
                            temp = lines[++i].Trim();
                            qty = Str_Utils.string_to_int(temp);
                        }

                        if (lines[i + 1].Trim().IndexOf("$") != -1) // maybe it has no total, so we must check it.
                            temp = lines[++i].Trim(); // total

                        temp = lines[++i].Trim();

                        while (!temp.StartsWith("SKU:", StringComparison.CurrentCultureIgnoreCase)) // skip product attributes
                        {
                            if (temp.StartsWith("Earn $", StringComparison.CurrentCultureIgnoreCase))
                                break;
                            if (temp.StartsWith("Buy Online Pick Up ", StringComparison.CurrentCultureIgnoreCase))
                                break;
                            if (temp.StartsWith("Receive a", StringComparison.CurrentCultureIgnoreCase))
                                break;
                            if (temp.StartsWith("Flash Sale", StringComparison.CurrentCultureIgnoreCase))
                                break;
                            if (temp.StartsWith("Oversized shipping ", StringComparison.CurrentCultureIgnoreCase))
                                break;
                            if (temp.StartsWith("Standard Delivery:", StringComparison.CurrentCultureIgnoreCase))
                                break;
                            if (temp.StartsWith("Tracking Number:", StringComparison.CurrentCultureIgnoreCase))
                                break;
                            temp = lines[++i].Trim();
                        }

                        sku = temp.Substring("SKU:".Length).Trim();

                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        card.m_product_items.Add(product);

                        MyLogger.Info($"... SC-6 (from bodytext) qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                        nextline = lines[++i].Trim();
                    }
                }
            }
        }
        private List<string> extract_sc6_product_title(string html_text)
        {
            List<string> titles = new List<string>();

            int start_pos = 0;
            int find_pos = html_text.IndexOf("<span id=\"productTitle\">", start_pos);
            while (find_pos != -1)
            {
                start_pos = find_pos + "<span id=\"productTitle\">".Length;
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
        private List<float> extract_sc6_product_price(string html_text)
        {
            List<float> prices = new List<float>();

            int start_pos = 0;
            int find_pos = html_text.IndexOf("<span id=\"productDiscountedPrice\">", start_pos);
            while (find_pos != -1)
            {
                start_pos = find_pos + "<span id=\"productDiscountedPrice\">".Length;
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
        private List<int> extract_sc6_product_qty(string html_text)
        {
            List<int> qtys = new List<int>();

            int start_pos = 0;
            int find_pos = html_text.IndexOf("<span id=\"productQuantity\">", start_pos);
            while (find_pos != -1)
            {
                start_pos = find_pos + "<span id=\"productQuantity\">".Length;
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
        private List<string> extract_sc6_product_sku(string html_text)
        {
            List<string> skus = new List<string>();

            int start_pos = 0;
            int find_pos = html_text.IndexOf("<span id=\"productSKU\">", start_pos);
            while (find_pos != -1)
            {
                start_pos = find_pos + "<span id=\"productSKU\">".Length;
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
        private DateTime extract_sc6_expected_delivery_date(string html_text)
        {
            DateTime date = DateTime.MinValue;

            int start_pos = 0;
            int find_pos = html_text.IndexOf("<span id=\"estimatedDeliverydate\">", start_pos);
            if (find_pos != -1)
            {
                start_pos = find_pos + "<span id=\"estimatedDeliverydate\">".Length;
                string temp = html_text.Substring(start_pos, html_text.IndexOf("</span>", start_pos) - start_pos).Trim();
                temp = temp.Replace("\r\n", " ");
                temp = temp.Replace("\r", "");
                temp = temp.Replace("\n", "");

                date = DateTime.Parse(temp);
            }
            return date;
        }
        private string extract_sc6_tracking(string html_text)
        {
            string tracking = "";

            int start_pos = 0;
            int find_pos = html_text.IndexOf("id=\"trackingNum\"");
            if (find_pos != -1)
            {
                start_pos = find_pos + "id=\"trackingNum\"".Length;
                string temp = html_text.Substring(start_pos, html_text.IndexOf("</td>", start_pos) - start_pos).Trim();
                temp = temp.Substring(temp.IndexOf("<span") + "<span".Length).Trim();
                temp = temp.Substring(0, temp.IndexOf("</span>")).Trim();
                temp = temp.Substring(temp.IndexOf(">") + 1).Trim();
                temp = temp.Replace("\r\n", " ");
                temp = temp.Replace("\r", "");
                temp = temp.Replace("\n", "");
                tracking = temp;
            }
            return tracking;
        }
    }
}
