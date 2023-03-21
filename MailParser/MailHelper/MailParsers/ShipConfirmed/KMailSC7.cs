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
        private void parse_mail_sc_7(MimeMessage mail, KReportSC report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.SC_7;

            report.m_retailer = ConstEnv.RETAILER_SAMSCLUB;

            MyLogger.Info($"... SC-7 m_sc_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Order #", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp = line.Substring("Order #".Length).Trim();
                    if (temp.IndexOf("See your order history", StringComparison.CurrentCultureIgnoreCase) != -1)
                        temp = temp.Substring(0, temp.IndexOf("See your order history", StringComparison.CurrentCultureIgnoreCase)).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... SC-7 order id = {temp}");
                    continue;
                }
                if (line == "Tracking number")
                {
                    string temp = lines[++i].Trim();
                    if (temp.IndexOf("<") != -1)
                        temp = temp.Substring(0, temp.IndexOf("<")).Trim();
                    string tracking = temp;
                    report.set_tracking(tracking);
                    MyLogger.Info($"... SC-7 tracking = {tracking}");
                    continue;
                }
                if (line.IndexOf("tracking number", StringComparison.CurrentCultureIgnoreCase) != -1 && !line.EndsWith("-->"))
                {
                    string temp = line.Substring(line.IndexOf("tracking number", StringComparison.CurrentCultureIgnoreCase) + "tracking number".Length).Trim();
                    string tracking = temp;

                    temp = line.Substring(0, line.IndexOf("tracking number", StringComparison.CurrentCultureIgnoreCase)).Trim();
                    if (temp != "")
                        report.m_sc_post_type = get_post_type(temp);

                    report.set_tracking(tracking);
                    MyLogger.Info($"... SC-7 tracking = {tracking}, post type = {report.m_sc_post_type}");
                    continue;
                }
                if (line.StartsWith("Delivered on"))
                {
                    string temp = line.Substring("Delivered on".Length).Trim();
                    DateTime date = DateTime.Parse(temp);
                    report.m_sc_ship_date = date;
                    MyLogger.Info($"... SC-7 date = {date}");
                    continue;
                }
                if (line.IndexOf("Delivered on", StringComparison.CurrentCultureIgnoreCase) != -1 && line.IndexOf("Qty:", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    string temp = line.Substring(line.IndexOf("Delivered on", StringComparison.CurrentCultureIgnoreCase) + "Delivered on".Length).Trim();
                    DateTime date = DateTime.Parse(temp);
                    report.m_sc_ship_date = date;
                    MyLogger.Info($"... SC-7 date = {date}");

                    string title = "";
                    string sku = "";
                    int qty = 0;
                    float price = 0;

                    temp = line.Substring(line.IndexOf("Qty:", StringComparison.CurrentCultureIgnoreCase) + "Qty:".Length).Trim();
                    if (temp.IndexOf("Delivered on", StringComparison.CurrentCultureIgnoreCase) != -1)
                        temp = temp.Substring(0, temp.IndexOf("Delivered on", StringComparison.CurrentCultureIgnoreCase));
                    qty = Str_Utils.string_to_int(temp);

                    temp = line.Substring(0, line.IndexOf("Qty:", StringComparison.CurrentCultureIgnoreCase)).Trim();
                    if (temp.IndexOf("Item #", StringComparison.CurrentCultureIgnoreCase) != -1)
                    {
                        sku = temp.Substring(temp.IndexOf("Item #", StringComparison.CurrentCultureIgnoreCase) + "Item #".Length).Trim();
                        temp = temp.Substring(0, temp.IndexOf("Item #", StringComparison.CurrentCultureIgnoreCase)).Trim();
                    }
                    title = temp;

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    report.m_product_items.Add(product);

                    MyLogger.Info($"... SC-7 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                    continue;
                }
                if (line.StartsWith("Qty:"))
                {
                    string title = "";
                    string sku = "";
                    int qty = 0;
                    float price = 0;

                    title = lines[i - 2].Trim();

                    string temp = lines[i - 1].Trim();
                    if (temp.IndexOf("#") != -1)
                        temp = temp.Substring(temp.IndexOf("#") + 1).Trim();
                    sku = temp;

                    temp = line.Substring("Qty:".Length).Trim();
                    qty = Str_Utils.string_to_int(temp);

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    report.m_product_items.Add(product);

                    MyLogger.Info($"... SC-7 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                    continue;
                }
            }
        }
    }
}
