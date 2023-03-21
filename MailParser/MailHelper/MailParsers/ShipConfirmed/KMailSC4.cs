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
        private void parse_mail_sc_4(MimeMessage mail, KReportSC report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.SC_4;

            report.m_retailer = ConstEnv.RETAILER_AMAZON;

            MyLogger.Info($"... SC-4 m_sc_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line == "Expected Delivery")
                {
                    CultureInfo provider = CultureInfo.InvariantCulture;
                    string temp = lines[++i].Trim();
                    if (temp.IndexOf(" - ") != -1)
                        temp = temp.Substring(0, temp.IndexOf(" - ")).Trim();
                    if (temp.IndexOf(",") != -1)
                        temp = temp.Substring(temp.IndexOf(",") + 1).Trim();
                    DateTime date = DateTime.ParseExact(temp, "MMMM d", provider);
                    date = new DateTime(mail.Date.Year, date.Month, date.Day);
                    report.m_sc_expected_deliver_date = date;
                    MyLogger.Info($"... Expected Delivery Date = {date}");
                    continue;
                }
                if (line.ToUpper() == "ARRIVING:")
                {
                    CultureInfo provider = CultureInfo.InvariantCulture;
                    string temp = lines[++i].Trim();
                    if (temp.IndexOf(" - ") != -1)
                        temp = temp.Substring(0, temp.IndexOf(" - ")).Trim();
                    if (temp.IndexOf(",") != -1)
                        temp = temp.Substring(temp.IndexOf(",") + 1).Trim();
                    DateTime date = DateTime.ParseExact(temp, "MMMM d", provider);
                    date = new DateTime(mail.Date.Year, date.Month, date.Day);
                    report.m_sc_expected_deliver_date = date;
                    MyLogger.Info($"... Expected Delivery Date = {date}");
                    continue;
                }
                if (line.EndsWith("your package will arrive:", StringComparison.CurrentCultureIgnoreCase))
                {
                    CultureInfo provider = CultureInfo.InvariantCulture;
                    string temp = lines[++i].Trim();
                    if (temp.IndexOf(" - ") != -1)
                        temp = temp.Substring(0, temp.IndexOf(" - ")).Trim();
                    if (temp.IndexOf(",") != -1)
                        temp = temp.Substring(temp.IndexOf(",") + 1).Trim();
                    DateTime date = DateTime.ParseExact(temp, "MMMM d", provider);
                    date = new DateTime(mail.Date.Year, date.Month, date.Day);
                    //card.m_sc_expected_deliver_date = date;
                    MyLogger.Info($"... Expected Delivery Date = {date}");
                    continue;
                }
                if (line.StartsWith("Order #"))
                {
                    string temp = line.Substring("Order #".Length).Trim();
                    if (temp.IndexOf("<") != -1)
                        temp = temp.Substring(0, temp.IndexOf("<")).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... order id = {temp}");
                    continue;
                }
                if (line.StartsWith("Fulfillment Order"))
                {
                    string temp = line.Substring("Fulfillment Order".Length).Trim();
                    if (temp.IndexOf(")") != -1)
                        temp = temp.Substring(0, temp.IndexOf(")")).Trim();
                    if (temp[0] == '(')
                        temp = temp.Substring(1);
                    report.set_order_id(temp);
                    MyLogger.Info($"... order id = {temp}");
                    continue;
                }
                if (line.StartsWith("Qty:"))
                {
                    string title = "";
                    string sku = "";
                    int qty = 0;
                    float price = 0;

                    string temp = line.Substring("Qty:".Length).Trim();
                    if (temp.IndexOf("|") != -1)
                        temp = temp.Substring(0, temp.IndexOf("|")).Trim();
                    qty = Str_Utils.string_to_int(temp);

                    temp = lines[i - 1].Trim();
                    if (temp[0] == '*')
                        temp = temp.Substring(1).Trim();
                    title = temp;

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    report.m_product_items.Add(product);

                    MyLogger.Info($"... qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                    continue;
                }
                if (line.Replace(" ", "") == "QtyItem")
                {
                    i += 2;
                    string next_line = lines[i].Trim();
                    while (i < lines.Length && !next_line.StartsWith("---------"))
                    {
                        string title = "";
                        string sku = "";
                        int qty = 0;
                        float price = 0;

                        string temp = next_line;

                        if (temp.StartsWith("Shipped By:"))
                        {
                            temp = temp.Substring("Shipped By:".Length).Trim();
                            string post_type = get_post_type(temp.Substring(0, temp.IndexOf(" ")).Trim());
                            report.m_sc_post_type = post_type;
                            MyLogger.Info($"... post_type = {post_type}");

                            temp = temp.Substring(temp.IndexOf(" ") + 1).Trim();
                            if (temp.StartsWith("Tracking No:"))
                            {
                                temp = temp.Substring("Tracking No:".Length).Trim();
                                string tracking = temp;
                                report.set_tracking(tracking);
                                MyLogger.Info($"... tracking = {tracking}");
                            }

                            next_line = lines[++i].Trim();
                            continue;
                        }

                        string qty_part = temp.Substring(0, temp.IndexOf(" ")).Trim();
                        qty = Str_Utils.string_to_int(qty_part);

                        temp = temp.Substring(temp.IndexOf(" ") + 1).Trim();
                        title = temp;

                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        report.m_product_items.Add(product);

                        MyLogger.Info($"... qty = {qty}, price = {price}, sku = {sku}, item title = {title}");

                        next_line = lines[++i].Trim();
                    }
                    continue;
                }
            }

            string htmlbody = XMailHelper.get_htmltext(mail);
            if (report.m_product_items.Count == 0 && htmlbody.IndexOf("Qty:", StringComparison.CurrentCultureIgnoreCase) != -1)
            {
                int next_pos;
                int qty_pos = htmlbody.IndexOf("Qty:", StringComparison.CurrentCultureIgnoreCase);
                while (qty_pos != -1)
                {
                    string title = "";
                    string sku = "";
                    int qty = 0;
                    float price = 0;

                    string temp = htmlbody.Substring(0, qty_pos);
                    int td_pos = temp.LastIndexOf("<td");
                    if (td_pos == -1)
                        break;
                    temp = htmlbody.Substring(td_pos);
                    temp = XMailHelper.find_html_part(temp, "td", out next_pos);
                    if (temp == "")
                        break;
                    temp = XMailHelper.html2text(temp);
                    temp = temp.Replace("\n", " ").Trim();
                    int qty_pos1 = temp.IndexOf("Qty:", StringComparison.CurrentCultureIgnoreCase);
                    if (qty_pos1 == -1)
                        break;

                    title = temp.Substring(0, qty_pos1).Trim();
                    temp = temp.Substring(qty_pos1 + "Qty:".Length).Trim();
                    if (temp.IndexOf("|") != -1)
                        temp = temp.Substring(0, temp.IndexOf("|")).Trim();
                    qty = Str_Utils.string_to_int(temp);

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    report.m_product_items.Add(product);

                    MyLogger.Info($"... qty = {qty}, price = {price}, sku = {sku}, item title = {title}");

                    qty_pos = htmlbody.IndexOf("Qty:", qty_pos + 1, StringComparison.CurrentCultureIgnoreCase);
                }
            }
        }
    }
}
