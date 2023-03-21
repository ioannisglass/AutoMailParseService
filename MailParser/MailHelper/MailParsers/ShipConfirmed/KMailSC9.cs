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
        private void parse_mail_sc_9(MimeMessage mail, KReportSC report)
        {
            if (!XMailHelper.is_bodytext_existed(mail))
            {
                parse_mail_sc_9_for_htmltext(mail, report);
                return;
            }
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.SC_9;

            report.m_retailer = ConstEnv.RETAILER_WALMART;

            MyLogger.Info($"... SC-9 m_sc_retailer = {report.m_retailer}");

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
                    report.set_order_id(temp); ;
                    MyLogger.Info($"... SC-9 order id = {temp}");
                    continue;
                }
                if (line.StartsWith("Order #:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "ORDER #:")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Order #:".Length).Trim();
                    report.set_order_id(temp); ;
                    MyLogger.Info($"... SC-9 order id = {temp}");
                    continue;
                }
                if (line.StartsWith("Shipped by"))
                {
                    string temp = line.Substring("Shipped by".Length).Trim();
                    string post_type = get_post_type(temp);
                    report.m_sc_post_type = post_type;
                    MyLogger.Info($"... SC-9 post_type = {post_type}");
                    continue;
                }
                if (line.IndexOf("tracking number") != -1)
                {
                    string temp = lines[++i].Trim();
                    string tracking = temp;
                    report.set_tracking(tracking);
                    MyLogger.Info($"... SC-9 tracking = {tracking}");
                    continue;
                }
                if (line == "Arrives by")
                {
                    CultureInfo provider = CultureInfo.InvariantCulture;
                    string temp = lines[i + 6].Trim();
                    temp = temp.Substring(temp.IndexOf(",") + 1).Trim();
                    int find = temp.IndexOf(" ", temp.IndexOf(" ") + 1);
                    temp = temp.Substring(0, find);
                    DateTime date = DateTime.ParseExact(temp, "MMM d", provider);
                    date = new DateTime(mail.Date.Year, date.Month, date.Day);
                    report.m_sc_ship_date = date;
                    MyLogger.Info($"... SC-9 Date of Shipment = {date}");
                    continue;
                }
                if (line.Replace(" ", "") == "ItemQtyTotal")
                {
                    string next_line = lines[++i].Trim();
                    while (i < lines.Length && (next_line != "" && !next_line.StartsWith("__")))
                    {
                        string title = "";
                        string sku = "";
                        int qty = 0;
                        float price = 0;

                        string temp = next_line;
                        if (temp.IndexOf("<") != -1)
                            temp = temp.Substring(0, temp.IndexOf("<"));
                        title = temp;

                        temp = lines[++i].Trim();
                        price = Str_Utils.string_to_currency(temp);

                        temp = lines[++i].Trim();
                        string qty_part = temp.Substring(0, temp.IndexOf(" "));
                        qty = Str_Utils.string_to_int(qty_part);

                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        report.m_product_items.Add(product);

                        MyLogger.Info($"... SC-9 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                        next_line = lines[++i].Trim();
                    }
                    continue;
                }
            }
        }
        private void parse_mail_sc_9_for_htmltext(MimeMessage mail, KReportSC report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.SC_9;

            report.m_retailer = ConstEnv.RETAILER_WALMART;

            MyLogger.Info($"... SC-9 m_sc_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
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
                    report.set_order_id(temp); ;
                    MyLogger.Info($"... SC-9 order id = {temp}");
                    continue;
                }
                if (line.StartsWith("Order #:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "ORDER #:")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Order #:".Length).Trim();
                    report.set_order_id(temp); ;
                    MyLogger.Info($"... SC-9 order id = {temp}");
                    continue;
                }
                if (line.StartsWith("Shipped by"))
                {
                    string temp = line.Substring("Shipped by".Length).Trim();
                    string post_type = get_post_type(temp);
                    report.m_sc_post_type = post_type;
                    MyLogger.Info($"... SC-9 post_type = {post_type}");
                    continue;
                }
                if (line == "UPS tracking number:")
                {
                    string temp = lines[++i].Trim();
                    string tracking = temp;
                    report.set_tracking(tracking);
                    report.m_sc_post_type = KReportBase.POST_TYPE_UPS;
                    MyLogger.Info($"... SC-9 post type = {report.m_sc_post_type}, tracking = {tracking}");
                    continue;
                }
                if (line.IndexOf("tracking number") != -1)
                {
                    string temp = line.Substring(0, line.IndexOf("tracking number")).Trim();
                    if (temp != "")
                    {
                        report.m_sc_post_type = get_post_type(temp);
                    }

                    temp = line.Substring(line.IndexOf("tracking number") + "tracking number".Length).Trim();
                    if (temp == "" || temp == ":")
                        temp = lines[++i].Trim();
                    string tracking = temp;
                    report.set_tracking(tracking);
                    MyLogger.Info($"... SC-9 post type = {report.m_sc_post_type}, tracking = {tracking}");
                    continue;
                }
                if (line == "Arrives by" || line == "Arrives by:")
                {
                    // from 15956

                    string[] day_of_weeks = new string[] { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" };
                    string temp = lines[i + 1].Trim();
                    foreach (string dow in day_of_weeks)
                    {
                        if (temp.StartsWith(dow + ","))
                        {
                            temp = temp.Substring(4).Trim();
                            CultureInfo provider = CultureInfo.InvariantCulture;
                            DateTime date = DateTime.ParseExact(temp, "MMM d", provider);
                            date = new DateTime(mail.Date.Year, date.Month, date.Day);
                            report.m_sc_ship_date = date;
                            MyLogger.Info($"... SC-9 Date of Shipment = {date}");
                        }
                    }
                    continue;
                }
                if (line == "Item" && lines[i + 1].Trim() == "Qty" && lines[i + 2].Trim() == "Total")
                {
                    // from 15956

                    string title = "";
                    string sku = "";
                    float price = 0;
                    int qty = 0;

                    i += 3;
                    while (lines[i].Trim()[0] != '$')
                    {
                        title += lines[i].Trim() + " ";
                        ++i;
                    }
                    title.Trim();

                    price = Str_Utils.string_to_currency(lines[i].Trim());
                    if (lines[i + 1].Trim().IndexOf("$") != -1)
                        i++;
                    qty = Str_Utils.string_to_int(lines[++i].Trim());

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    report.m_product_items.Add(product);

                    MyLogger.Info($"... SC-9 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                }
                if (line == "Item" && lines[i + 1].Trim() == "Qty" && lines[i + 2].Trim() == "Price" && lines[i + 3].Trim() == "Total")
                {
                    string title = "";
                    string sku = "";
                    float price = 0;
                    int qty = 0;

                    i += 4;
                    int k = i;
                    while (lines[k].Trim()[0] != '$')
                        k++;
                    if (k - i < 2)
                        continue;

                    for (int j = i; j < k - 1; j++)
                        title += lines[j].Trim() + " ";
                    title = title.Trim();

                    qty = Str_Utils.string_to_int(lines[k - 1].Trim());

                    price = Str_Utils.string_to_currency(lines[k].Trim());

                    i = k + 1;

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    report.m_product_items.Add(product);

                    MyLogger.Info($"... SC-9 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                }
            }
        }
    }
}
