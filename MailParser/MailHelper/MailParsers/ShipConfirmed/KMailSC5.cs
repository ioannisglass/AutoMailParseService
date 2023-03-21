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
        private void parse_mail_sc_5(MimeMessage mail, KReportSC report)
        {
            if (!XMailHelper.is_bodytext_existed(mail))
            {
                parse_mail_sc_5_for_htmltext(mail, report);
                return;
            }
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.SC_5;

            report.m_retailer = ConstEnv.RETAILER_HOMEDEPOT;

            MyLogger.Info($"... SC-5 m_sc_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line == "Order Number")
                {
                    string temp = lines[++i].Trim();
                    if (temp.IndexOf("Order Date") != -1)
                        continue;
                    MyLogger.Info($"... SC-5 order id = {temp}");
                    continue;
                }
                if (line.StartsWith("Carrier:"))
                {
                    string temp = line.Substring("Carrier:".Length).Trim();
                    string post_type = get_post_type(temp);
                    report.m_sc_post_type = post_type;
                    MyLogger.Info($"... SC-5 post_type = {post_type}");
                    continue;
                }
                if (line == "Tracking Number:")
                {
                    string tracking = lines[++i].Trim();
                    report.set_tracking(tracking);
                    MyLogger.Info($"... SC-5 tracking = {tracking}");
                    continue;
                }
                if (line == "Item" && lines[i + 1].Trim() == "Unit Price" && lines[i + 2].Trim() == "Qty" && lines[i + 3].Trim() == "Item Total")
                {
                    i += 4;
                    string next_line = lines[i].Trim();
                    while (i < lines.Length && next_line != "")
                    {
                        string title = "";
                        string sku = "";
                        int qty = 0;
                        float price = 0;

                        string temp = next_line;
                        if (temp.StartsWith("["))
                            temp = temp.Substring(temp.IndexOf("]") + 1).Trim();
                        if (temp.StartsWith("<"))
                            temp = temp.Substring(temp.IndexOf(">") + 1).Trim();
                        if (temp.IndexOf("<") != -1)
                            temp = temp.Substring(0, temp.IndexOf("<")).Trim();
                        title = temp;

                        temp = lines[++i].Trim();
                        if (temp.StartsWith("Store SKU #"))
                        {
                            temp = temp.Substring("Store SKU #".Length).Trim();
                            sku = temp;
                        }
                        temp = lines[++i].Trim();
                        if (temp.StartsWith("Internet #") && sku == "")
                        {
                            temp = temp.Substring("Internet #".Length).Trim();
                            sku = temp;
                        }
                        temp = lines[++i].Trim();
                        price = Str_Utils.string_to_currency(temp);
                        temp = lines[++i].Trim();
                        qty = (int)Str_Utils.string_to_float(temp);
                        temp = lines[++i].Trim();

                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        report.m_product_items.Add(product);

                        MyLogger.Info($"... SC-5 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");

                        next_line = lines[++i].Trim();
                    }
                    continue;
                }
            }
        }
        private void get_sc5_ship_date(string htmlbody, KReportSC card)
        {
            int pos;
            string temp;
            int next_pos;

            pos = htmlbody.IndexOf("Expect it", StringComparison.CurrentCultureIgnoreCase);
            if (pos != -1)
            {
                temp = htmlbody.Substring(pos + "Expect it".Length).Trim();
                temp = XMailHelper.find_html_part(temp, "b", out next_pos);
                if (temp != "" && temp.IndexOf(",") != -1)
                {
                    temp = temp.Substring(temp.IndexOf(",") + 1).Trim();
                    DateTime date = DateTime.Parse(temp);
                    card.m_sc_expected_deliver_date = date;
                }
            }
            pos = htmlbody.IndexOf("class=\"tracking_view\"", StringComparison.CurrentCultureIgnoreCase);
            if (pos != -1)
            {
                temp = htmlbody.Substring(pos).Trim();
                pos = temp.IndexOf("promise_date=");
                if (pos != -1)
                {
                    temp = temp.Substring(pos + "promise_date=".Length).Trim();
                    pos = temp.IndexOf("--");
                    if (pos == -1)
                        pos = temp.IndexOf("&");
                    temp = temp.Substring(0, pos).Trim();

                    card.m_sc_expected_deliver_date = DateTime.Parse(temp);
                }
            }
        }
        private void get_sc5_tracking_info(string htmlbody, KReportSC card)
        {
            int pos;
            string temp;
            int next_pos;

            pos = htmlbody.IndexOf("Delivery Window", StringComparison.CurrentCultureIgnoreCase);
            if (pos == -1)
                return;
            temp = htmlbody.Substring(pos);
            temp = XMailHelper.find_html_part(temp, "p", out next_pos);
            if (temp == "")
                return;
            string text = XMailHelper.html2text(temp);
            pos = text.IndexOf("Tracking Number:", StringComparison.CurrentCultureIgnoreCase);
            if (pos != -1)
            {
                string tracking = text.Substring(pos + "Tracking Number:".Length).Trim().Replace("\n", ",").Trim();
                card.set_tracking(tracking);

                text = text.Substring(0, pos);
            }
            pos = text.IndexOf("Carrier:", StringComparison.CurrentCultureIgnoreCase);
            if (pos != -1)
            {
                string post_type = text.Substring(pos + "Carrier:".Length).Trim();
                card.m_sc_post_type = get_post_type(post_type);
            }
            MyLogger.Info($"... SC-5 post_type = {card.m_sc_post_type}, tracking = {card.m_sc_tracking}");
        }
        private void get_sc5_product_info(string htmlbody, KReportSC card)
        {
            int next_pos;
            int pos;
            string temp;

            pos = htmlbody.IndexOf("class=\"desc-view\"", StringComparison.CurrentCultureIgnoreCase);
            while (pos != -1)
            {
                if (pos == -1)
                    break;
                temp = htmlbody.Substring(0, pos).Trim();
                int td_pos = temp.LastIndexOf("<td");
                if (td_pos == -1)
                    break;
                temp = htmlbody.Substring(td_pos).Trim();
                string product_part = XMailHelper.find_html_part(temp, "td", out next_pos);
                if (product_part == "")
                    break;

                string title = "";
                string sku = "";
                int qty = 0;
                float price = 0;

                temp = XMailHelper.find_html_part(product_part, "p", out next_pos);
                while (temp != "")
                {
                    string text = XMailHelper.html2text(temp);

                    if (text.IndexOf("SKU #", StringComparison.CurrentCultureIgnoreCase) != -1 || text.IndexOf("Internet #", StringComparison.CurrentCultureIgnoreCase) != -1)
                    {
                        if (text.IndexOf("SKU #", StringComparison.CurrentCultureIgnoreCase) != -1)
                            sku = text.Substring(text.IndexOf("SKU #", StringComparison.CurrentCultureIgnoreCase) + "SKU #".Length).Trim();
                        if (sku.IndexOf("Internet #", StringComparison.CurrentCultureIgnoreCase) != -1)
                            sku = sku.Substring(0, sku.IndexOf("Internet #", StringComparison.CurrentCultureIgnoreCase)).Trim();
                        if (sku == "" && text.IndexOf("Internet #", StringComparison.CurrentCultureIgnoreCase) != -1)
                            sku = text.Substring(text.IndexOf("Internet #", StringComparison.CurrentCultureIgnoreCase) + "Internet #".Length).Trim();
                    }
                    else if (text.IndexOf("Unit Price", StringComparison.CurrentCultureIgnoreCase) != -1)
                    {
                        text = text.Substring(text.IndexOf("Unit Price", StringComparison.CurrentCultureIgnoreCase) + "Unit Price".Length).Trim();
                        price = Str_Utils.string_to_currency(text);
                    }
                    else if (text.IndexOf("Qty", StringComparison.CurrentCultureIgnoreCase) != -1)
                    {
                        text = text.Substring(text.IndexOf("Qty", StringComparison.CurrentCultureIgnoreCase) + "Qty".Length).Trim();
                        qty = Str_Utils.string_to_int(text);
                    }
                    else if (text.IndexOf("Item Total", StringComparison.CurrentCultureIgnoreCase) != -1)
                    {
                        // do nothing.
                    }
                    else
                    {
                        title = text;
                    }

                    product_part = product_part.Substring(next_pos);
                    temp = XMailHelper.find_html_part(product_part, "p", out next_pos);
                }

                ZProduct product = new ZProduct();
                product.price = price;
                product.sku = sku;
                product.title = title;
                product.qty = qty;
                card.m_product_items.Add(product);

                MyLogger.Info($"... SC-5 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");

                pos = htmlbody.IndexOf("class=\"desc-view\"", pos + 1, StringComparison.CurrentCultureIgnoreCase);
            }
        }
        private void parse_mail_sc_5_for_htmltext(MimeMessage mail, KReportSC report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.SC_5;

            report.m_retailer = ConstEnv.RETAILER_HOMEDEPOT;

            MyLogger.Info($"... SC-5 m_sc_retailer = {report.m_retailer}");

            string bodytext = XMailHelper.get_bodytext(mail);
            string[] lines = bodytext.Replace("\r", "").Split('\n');
            if (lines.Length < 10)
            {
                bodytext = XMailHelper.revise_concated_bodytext(bodytext);
                lines = bodytext.Replace("\r", "").Split('\n');
            }
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line == "Order Number")
                {
                    string temp = lines[++i].Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... SC-5 order id = {temp}");
                    continue;
                }
            }

            string htmlbody = XMailHelper.get_htmltext(mail);

            // Find tracking information.
            get_sc5_tracking_info(htmlbody, report);

            // Find product information.
            get_sc5_product_info(htmlbody, report);

            // Find shipment date.
            get_sc5_ship_date(htmlbody, report);
        }
        private void analyze_mail_order_5_for_htmltext2(MimeMessage mail, KReportSC card)
        {
            string subject = XMailHelper.get_subject(mail);

            card.m_mail_type = KReportBase.MailType.SC_5;

            card.m_retailer = ConstEnv.RETAILER_HOMEDEPOT;

            MyLogger.Info($"... SC-5 m_sc_retailer = {card.m_retailer}");

            string title = "";
            string sku = "";
            int qty = 0;
            float price = 0;

            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line == "Order Number")
                {
                    string temp = lines[++i].Trim();
                    if (temp.IndexOf("Order Date") != -1)
                        continue;
                    card.set_order_id(temp);
                    MyLogger.Info($"... SC-5 order id = {temp}");
                    continue;
                }
                if (line.StartsWith("Carrier:"))
                {
                    string temp = line.Substring("Carrier:".Length).Trim();
                    string post_type = get_post_type(temp);
                    card.m_sc_post_type = post_type;
                    MyLogger.Info($"... SC-5 post_type = {post_type}");
                    continue;
                }
                if (line == "Tracking Number:")
                {
                    string tracking = lines[++i].Trim();
                    card.set_tracking(tracking);
                    MyLogger.Info($"... SC-5 tracking = {tracking}");
                    continue;
                }
                if (line.StartsWith("Tracking Number:") && line != "Tracking Number:")
                {
                    string temp = line.Substring("Tracking Number:".Length).Trim();
                    card.set_tracking(temp);
                    MyLogger.Info($"... SC-5 tracking = {temp}");
                    continue;
                }
                if (line.StartsWith("Qty") && line != "Qty")
                {
                    title = "";
                    sku = "";
                    qty = 0;
                    price = 0;

                    string temp = line.Substring("Qty".Length).Trim();
                    qty = (int)Str_Utils.string_to_float(temp);

                    temp = lines[i - 1].Trim();
                    if (temp.IndexOf("Unit Price") != -1)
                    {
                        temp = temp.Substring("Unit Price".Length).Trim();
                        price = Str_Utils.string_to_currency(temp);
                    }

                    temp = lines[i - 3].Trim();
                    if (temp.IndexOf("Store SKU #") != -1)
                    {
                        temp = temp.Substring("Store SKU #".Length).Trim();
                        sku = temp;
                    }

                    temp = lines[i - 4].Trim();
                    title = temp;

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    card.m_product_items.Add(product);

                    MyLogger.Info($"... SC-5 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                    continue;
                }
            }
        }
    }
}
