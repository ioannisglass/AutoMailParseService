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
        private string get_sc10_shipped_info(string html_order_part, string key_text, string[] exc_keys)
        {
            int next_pos;
            string temp = html_order_part;
            string ret = "";
            bool found = false;
            bool found_exc_key;

            // get shipped part.

            string table_part = XMailHelper.find_html_part(temp, "table", out next_pos);
            while (table_part != "")
            {
                string table_text_part = XMailHelper.html2text(table_part);

                if (!found && table_text_part.IndexOf(key_text, StringComparison.CurrentCultureIgnoreCase) != -1)
                    found = true;

                if (found)
                {
                    found_exc_key = false;
                    foreach (string exc_key in exc_keys)
                    {
                        if (table_text_part.IndexOf(exc_key, StringComparison.CurrentCultureIgnoreCase) != -1)
                        {
                            found_exc_key = true;
                            break;
                        }
                    }
                    if (found_exc_key)
                        break;
                }

                if (found)
                {
                    ret += "<table>\n" + table_part + "</table>\n";
                }

                temp = temp.Substring(next_pos);
                table_part = XMailHelper.find_html_part(temp, "table", out next_pos);
            }
            return ret;
        }
        private void get_sc10_product_info_from_html_part(string td_product_html, KReportSC card)
        {
            string temp = td_product_html;

            int next_pos;
            string tr_part = XMailHelper.find_html_part(temp, "tr", out next_pos);
            while (tr_part != "")
            {
                string title = "";
                string sku = "";
                int qty = 0;
                float price = 0;

                // details td.

                string temp1 = tr_part;
                int next_pos1, next_pos2;
                string td_part = XMailHelper.find_html_part(temp1, "td", out next_pos1);

                string title_part = XMailHelper.find_html_part(td_part, "tr", out next_pos2);
                td_part = td_part.Substring(next_pos2);
                string item_part = XMailHelper.find_html_part(td_part, "tr", out next_pos2);

                title = XMailHelper.html2text(title_part);
                title = title.Replace("\n", " ").Trim();

                sku = XMailHelper.html2text(item_part);
                sku = sku.Replace("\n", " ").Trim();
                if (sku.IndexOf("|") != -1)
                    sku = sku.Substring(0, sku.IndexOf("|")).Trim();
                if (sku.IndexOf("Item:", StringComparison.CurrentCultureIgnoreCase) != -1)
                    sku = sku.Substring(sku.IndexOf("Item:", StringComparison.CurrentCultureIgnoreCase) + "Item:".Length).Trim();

                // price

                temp1 = temp1.Substring(next_pos1);
                td_part = XMailHelper.find_html_part(temp1, "td", out next_pos1);
                temp1 = XMailHelper.html2text(temp1);
                temp1 = temp1.Replace("\n", " ").Trim();
                string qty_part = temp1.Substring(0, temp1.IndexOf("$")).Trim();
                if (qty_part.StartsWith("x ", StringComparison.CurrentCultureIgnoreCase))
                    qty_part = qty_part.Substring(2).Trim();
                qty = Str_Utils.string_to_int(qty_part);

                temp1 = temp1.Substring(temp1.IndexOf("$") + 1).Trim();
                if (temp1.IndexOf("(") != -1 && temp1.EndsWith(")"))
                {
                    temp1 = temp1.Substring(0, temp1.Length - 1).Trim();
                    temp1 = temp1.Substring(temp1.LastIndexOf("(") + 1).Trim();
                    if (temp1.EndsWith(" each", StringComparison.CurrentCultureIgnoreCase))
                        temp1 = temp1.Substring(0, temp1.Length - " each".Length).Trim();
                    price = Str_Utils.string_to_currency(temp1);
                }
                else
                {
                    price = Str_Utils.string_to_currency(temp1);
                    if (qty != 1)
                        price /= qty;
                }

                ZProduct product = new ZProduct();
                product.price = price;
                product.sku = sku;
                product.title = title;
                product.qty = qty;
                card.m_product_items.Add(product);

                temp = temp.Substring(next_pos);
                tr_part = XMailHelper.find_html_part(temp, "tr", out next_pos);
            }
        }
        private void get_sc10_shipped_date_from_html_part(string table_html, KReportSC card)
        {
            string text = XMailHelper.html2text(table_html);
            text = text.Replace("\n", " ").Trim();
            int pos = text.IndexOf("Shipped:", StringComparison.CurrentCultureIgnoreCase);
            if (pos != -1)
            {
                text = text.Substring(pos + "Shipped:".Length).Trim();
                card.m_sc_ship_date = DateTime.Parse(text);
            }
        }
        private void get_sc10_tracking_from_html_part(string table_html, KReportSC card)
        {
            do
            {
                int pos = table_html.IndexOf("Carrier:", StringComparison.CurrentCultureIgnoreCase);
                if (pos == -1)
                    break;
                pos = table_html.IndexOf("</tr>", pos, StringComparison.CurrentCultureIgnoreCase);
                string temp = table_html.Substring(pos + "</tr>".Length);
                int next_pos1;
                string td_part = XMailHelper.find_html_part(temp, "td", out next_pos1);
                string text = XMailHelper.html2text(td_part);
                text = text.Replace("\n", " ").Trim();
                card.m_sc_post_type = get_post_type(text);

                temp = temp.Substring(next_pos1);
                td_part = XMailHelper.find_html_part(temp, "td", out next_pos1);
                text = XMailHelper.html2text(td_part);
                text = text.Replace("\n", " ").Trim();
                string tracking = get_post_type(text);
                card.set_tracking(tracking);

            } while (false);
        }
        private int get_sc10_table_type(string table_html)
        {
            int table_type = 0;

            string text = XMailHelper.html2text(table_html);

            if (text.IndexOf("These Item(s) Have Shipped", StringComparison.CurrentCultureIgnoreCase) != -1)
                return 1;
            if (text.IndexOf("These Item(s) Have Not Yet Shipped", StringComparison.CurrentCultureIgnoreCase) != -1)
                return 1;
            if (text.IndexOf("These Item(s) Have Been Canceled", StringComparison.CurrentCultureIgnoreCase) != -1)
                return 1;

            if (text.IndexOf("Shipped:", StringComparison.CurrentCultureIgnoreCase) != -1)
                return 2;
            if (text.IndexOf("Guaranteed Express", StringComparison.CurrentCultureIgnoreCase) != -1)
                return 2;

            if (text.IndexOf("Carrier:", StringComparison.CurrentCultureIgnoreCase) != -1)
                return 3;
            if (text.IndexOf("Tracking Number:", StringComparison.CurrentCultureIgnoreCase) != -1)
                return 3;

            if (table_html.IndexOf("<td class=\"product-details\"", StringComparison.CurrentCultureIgnoreCase) != -1)
                return 4;

            return table_type;
        }
        private void parse_sc10_shipped_part(string shipped_html_part, KReportSC card)
        {
            /**
             * The shipped html part has 4 table.
             * 
             * 1) It contains "These Item(s) Have Shipped" clause.
             * 2) shipped date
             * 3) carrier, tracking number
             * 4) product information
             * repeat 3) and 4) by the number of the product items. (Single shipment)
             * Or repeat 2), 3) and 4) by the number of the product items. (Multi-shipment)
             * 
             **/

            int next_pos;
            string temp = shipped_html_part;

            string table_part = XMailHelper.find_html_part(temp, "table", out next_pos);

            // ignore 1st table.

            // 2nd table : get shipped date.

            temp = temp.Substring(next_pos);
            table_part = XMailHelper.find_html_part(temp, "table", out next_pos);
            while (table_part != "")
            {
                int table_type = get_sc10_table_type(table_part);
                if (table_type == 1)
                    break;

                if (table_type == 0)
                {
                    // ignore.
                }
                if (table_type == 2)
                {
                    get_sc10_shipped_date_from_html_part(table_part, card);
                }
                else if (table_type == 3)
                {
                    get_sc10_tracking_from_html_part(table_part, card);
                }
                else if (table_type == 4)
                {
                    do
                    {
                        int pos = table_part.IndexOf("<td class=\"product-details\"", StringComparison.CurrentCultureIgnoreCase);
                        if (pos == -1)
                            break;
                        string temp1 = table_part.Substring(pos);
                        int next_pos1;
                        temp1 = XMailHelper.find_html_part(temp1, "td", out next_pos1);
                        get_sc10_product_info_from_html_part(temp1, card);

                    } while (false);
                }

                temp = temp.Substring(next_pos);
                table_part = XMailHelper.find_html_part(temp, "table", out next_pos);
            }
        }
        private void parse_sc10_notshipped_part(string shipped_html_part, KReportSC card)
        {
            /**
             * The not-shipped html part has 3 table.
             * 
             * 1) It contains "These Item(s) Have Not Yet Shipped" clause.
             * 2) "Will Ship Guaranteed Express" clause.
             * 3) product information
             * repeat 2) and 3) by the number of the product items. (Single shipment)
             * 
             **/

            int next_pos;
            string temp = shipped_html_part;

            string table_part = XMailHelper.find_html_part(temp, "table", out next_pos);

            // ignore 1st table.

            temp = temp.Substring(next_pos);
            table_part = XMailHelper.find_html_part(temp, "table", out next_pos);
            while (table_part != "")
            {
                int table_type = get_sc10_table_type(table_part);
                if (table_type == 1)
                    break;

                if (table_type == 0)
                {
                    // ignore.
                }
                if (table_type == 2)
                {
                    // ignore.
                }
                else if (table_type == 3)
                {
                    // ignore.
                }
                else if (table_type == 4)
                {
                    do
                    {
                        int pos = table_part.IndexOf("<td class=\"product-details\"", StringComparison.CurrentCultureIgnoreCase);
                        if (pos == -1)
                            break;
                        string temp1 = table_part.Substring(pos);
                        int next_pos1;
                        temp1 = XMailHelper.find_html_part(temp1, "td", out next_pos1);
                        get_sc10_product_info_from_html_part(temp1, card);

                    } while (false);
                }

                temp = temp.Substring(next_pos);
                table_part = XMailHelper.find_html_part(temp, "table", out next_pos);
            }
        }
        private void parse_sc10_cancelled_part(string shipped_html_part, KReportSC card)
        {
            /**
             * The canceled html part has 2 table.
             * 
             * 1) It contains "These Item(s) Have Been Canceled" clause.
             * 2) product information
             * 
             **/

            int next_pos;
            string temp = shipped_html_part;

            string table_part = XMailHelper.find_html_part(temp, "table", out next_pos);

            // ignore 1st table.

            temp = temp.Substring(next_pos);
            table_part = XMailHelper.find_html_part(temp, "table", out next_pos);
            while (table_part != "")
            {
                int table_type = get_sc10_table_type(table_part);
                if (table_type == 1)
                    break;

                if (table_type == 0)
                {
                    // ignore.
                }
                if (table_type == 2)
                {
                    // ignore.
                }
                else if (table_type == 3)
                {
                    // ignore.
                }
                else if (table_type == 4)
                {
                    do
                    {
                        int pos = table_part.IndexOf("<td class=\"product-details\"", StringComparison.CurrentCultureIgnoreCase);
                        if (pos == -1)
                            break;
                        string temp1 = table_part.Substring(pos);
                        int next_pos1;
                        temp1 = XMailHelper.find_html_part(temp1, "td", out next_pos1);
                        get_sc10_product_info_from_html_part(temp1, card);

                    } while (false);
                }

                temp = temp.Substring(next_pos);
                table_part = XMailHelper.find_html_part(temp, "table", out next_pos);
            }
        }
        private void parse_mail_sc_10(MimeMessage mail, KReportSC report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.SC_10;

            report.m_retailer = ConstEnv.RETAILER_CABELAS;

            MyLogger.Info($"... SC-10 m_sc_retailer = {report.m_retailer}");

            string htmlbody = XMailHelper.get_htmltext(mail);
            int start_pos = htmlbody.IndexOf("<!-- START ORDER PRODUCT -->", StringComparison.CurrentCultureIgnoreCase);
            int end_pos = htmlbody.IndexOf("<!-- END ORDER PRODUCT -->", StringComparison.CurrentCultureIgnoreCase);
            if (start_pos != -1 && start_pos < end_pos)
            {
                string temp = htmlbody.Substring(start_pos + "<!-- START ORDER PRODUCT -->".Length, end_pos - start_pos - "<!-- START ORDER PRODUCT -->".Length);

                string shipped_part = get_sc10_shipped_info(temp, "These Item(s) Have Shipped", new string[] { "These Item(s) Have Not Yet Shipped", "These Item(s) Have Been Canceled" });
                string not_shipped_part = get_sc10_shipped_info(temp, "These Item(s) Have Not Yet Shipped", new string[] { "These Item(s) Have Shipped", "These Item(s) Have Been Canceled" });
                string cancelled_part = get_sc10_shipped_info(temp, "These Item(s) Have Been Canceled", new string[] { "These Item(s) Have Not Yet Shipped", "These Item(s) Have Shipped" });

                KReportSC shipped_report = new KReportSC();
                if (shipped_part != "")
                {
                    parse_sc10_shipped_part(shipped_part, shipped_report);

                    report.m_sc_post_type = shipped_report.m_sc_post_type;
                    report.set_tracking(shipped_report.m_sc_tracking);
                    report.m_sc_ship_date = shipped_report.m_sc_ship_date;
                    report.m_product_items = shipped_report.m_product_items;

                    MyLogger.Info($"... SC-10 post type = {report.m_sc_post_type}, tracking = {report.m_sc_tracking}");
                    MyLogger.Info($"... SC-10 shipped date = {report.m_sc_ship_date}");
                    foreach (ZProduct product in report.m_product_items)
                        MyLogger.Info($"... SC-10 qty = {product.qty}, price = {product.price}, sku = {product.sku}, item title = {product.title}");
                }

                KReportSC notshipped_report = new KReportSC();
                if (not_shipped_part != "")
                {
                    parse_sc10_notshipped_part(not_shipped_part, notshipped_report);

                    foreach (ZProduct product in notshipped_report.m_product_items)
                    {
                        report.m_product_items.Add(new ZProduct() {
                            title = product.title,
                            sku = product.sku,
                            qty = product.qty,
                            price = product.price,
                            status = ConstEnv.REPORT_ORDER_STATUS_PARTIAL_NOT_SHIPPED
                        });
                        MyLogger.Info($"... SC-10 Not-Shipped qty = {product.qty}, price = {product.price}, sku = {product.sku}, item title = {product.title}");
                    }
                }

                KReportSC cancelled_report = new KReportSC();
                if (cancelled_part != "")
                {
                    parse_sc10_cancelled_part(cancelled_part, cancelled_report);

                    foreach (ZProduct product in cancelled_report.m_product_items)
                    {
                        report.m_product_items.Add(new ZProduct()
                        {
                            title = product.title,
                            sku = product.sku,
                            qty = product.qty,
                            price = product.price,
                            status = ConstEnv.REPORT_ORDER_STATUS_PARTIAL_CANCELED
                        });
                        MyLogger.Info($"... SC-10 Cancelled qty = {product.qty}, price = {product.price}, sku = {product.sku}, item title = {product.title}");
                    }
                }
            }

            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Order Number:"))
                {
                    string temp = line.Substring("Order Number:".Length).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... SC-10 order id = {temp}");
                    continue;
                }
                if (line.IndexOf("Order Number:", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    string temp = line.Substring(line.IndexOf("Order Number:", StringComparison.CurrentCultureIgnoreCase) + "Order Number:".Length).Trim();
                    if (temp.IndexOf("Order Date:", StringComparison.CurrentCultureIgnoreCase) != -1)
                        temp = temp.Substring(0, temp.IndexOf("Order Date:", StringComparison.CurrentCultureIgnoreCase)).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... SC-10 order id = {temp}");
                    continue;
                }
            }
        }
    }
}
