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
    partial class KMailBaseSC : KMailBaseParser
    {
        private void parse_mail_sc_3(MimeMessage mail, KReportSC report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.SC_3;

            report.m_retailer = ConstEnv.RETAILER_BESTBUY;

            MyLogger.Info($"... SC-3 m_sc_retailer = {report.m_retailer}");

            /**
             * There are two kind of the BestBuy mail.
             * First is the subject name is "Your order has shipped".
             * Second is the subject name is like as "Your order #BBY01-761223014906 has shipped".
             *
             **/

            int mail_kind = (subject.ToUpper() == "YOUR ORDER HAS SHIPPED") ? 1 : 2;

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.IndexOf("Your order shipped on") != -1)
                {
                    string temp = line.Substring(line.IndexOf("Your order shipped on") + "Your order shipped on".Length).Trim();
                    DateTime date = DateTime.Parse(temp);
                    report.m_sc_ship_date = date;
                    MyLogger.Info($"... SC-3 Date of Shipment = {date}");
                    continue;
                }
                if (line.StartsWith("Shipped on", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp = line.Substring("Shipped on".Length).Trim();
                    DateTime date = DateTime.Parse(temp);
                    report.m_sc_ship_date = date;
                    MyLogger.Info($"... SC-3 Date of Shipment = {date}");
                    continue;
                }
                if (line.StartsWith("Order #", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "ORDER #")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Order #".Length).Trim();
                    if (temp.IndexOf("<") != -1)
                        temp = temp.Substring(0, temp.IndexOf("<")).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... SC-3 order id = {temp}");
                    continue;
                }
                if (line.ToUpper() == "ORDER" && lines[i + 1].Trim()[0] == '#')
                {
                    string temp = lines[++i].Substring(1).Trim();
                    if (temp.IndexOf("<") != -1)
                        temp = temp.Substring(0, temp.IndexOf("<")).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... SC-3 order id = {temp}");
                    continue;
                }
                if (line.StartsWith("Order number:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "ORDER NUMBER:")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Order number:".Length).Trim();
                    if (temp.IndexOf("<") != -1)
                        temp = temp.Substring(0, temp.IndexOf("<")).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... SC-3 order id = {temp}");
                    continue;
                }
                if (line.StartsWith("Tracking #", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "TRACKING #")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Tracking #".Length).Trim();
                    if (temp.IndexOf("<") != -1)
                        temp = temp.Substring(0, temp.IndexOf("<"));
                    if (lines[i + 1].Trim().StartsWith("..."))
                        temp += lines[++i].Trim();
                    string tracking = temp;
                    report.set_tracking(tracking);
                    MyLogger.Info($"... SC-3 tracking = {tracking}");
                    continue;
                }
                if (line.ToUpper() == "YOUR TRACKING NUMBER IS:")
                {
                    string temp = lines[++i].Trim();
                    if (lines[i + 1].Trim().StartsWith("..."))
                        temp += lines[++i].Trim();
                    if (temp.StartsWith("Note:", StringComparison.CurrentCultureIgnoreCase)) // Some mail have empty tracking number.
                        continue;
                    string tracking = temp;
                    report.set_tracking(tracking);
                    MyLogger.Info($"... SC-3 tracking = {tracking}");
                    continue;
                }
                if (mail_kind == 2 && line.StartsWith("SKU:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string title = "";
                    string sku = "";
                    int qty = 0;
                    float price = 0;

                    string temp = line.Substring("SKU:".Length).Trim();
                    sku = temp;

                    temp = lines[i - 2].Trim();
                    if (temp.IndexOf("<") != -1)
                        temp = temp.Substring(0, temp.IndexOf("<"));
                    title = temp;

                    temp = lines[i + 1].Trim();
                    if (temp.ToUpper() == "QTY")
                    {
                        temp = lines[i + 2].Trim();
                        qty = Str_Utils.string_to_int(temp);
                    }

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    report.m_product_items.Add(product);

                    MyLogger.Info($"... SC-3 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                    continue;
                }
                if (mail_kind == 1 && line.StartsWith("SKU:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string title = "";
                    string sku = "";
                    int qty = 0;
                    float price = 0;

                    string temp = line.Substring("SKU:".Length).Trim();
                    sku = temp;

                    if (lines[i + 1].Trim().StartsWith("Shipped on", StringComparison.CurrentCultureIgnoreCase))
                    {
                        temp = lines[i - 1].Trim();
                        if (temp.IndexOf("<") != -1)
                            temp = temp.Substring(0, temp.IndexOf("<"));
                        title = temp;

                        temp = lines[i - 2].Trim();
                        qty = Str_Utils.string_to_int(temp);
                    }
                    else if (lines[i - 2].Trim().StartsWith("Shipped on", StringComparison.CurrentCultureIgnoreCase))
                    {
                        temp = lines[i - 3].Trim();
                        if (temp.IndexOf("<") != -1)
                            temp = temp.Substring(0, temp.IndexOf("<"));
                        title = temp;

                        temp = lines[i - 4].Trim();
                        qty = Str_Utils.string_to_int(temp);
                    }

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    report.m_product_items.Add(product);

                    MyLogger.Info($"... SC-3 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                    continue;
                }
            }

            if (mail_kind == 2 && report.m_product_items.Count == 0)
            {
                int next_pos;
                string htmlbody = XMailHelper.get_htmltext(mail);
                string temp = "";
                int find = htmlbody.IndexOf("class=\"lineItem-meta\"");
                while (find != -1)
                {
                    string title = "";
                    string sku = "";
                    int qty = 0;
                    float price = 0;

                    temp = htmlbody.Substring(0, find);
                    int td_find = temp.LastIndexOf("<td");
                    if (td_find == -1)
                        break;
                    temp = htmlbody.Substring(td_find);
                    temp = XMailHelper.find_html_part(temp, "td", out next_pos);

                    // td is divided by two part : item title&sku and qty.

                    string temp1 = XMailHelper.find_html_part(temp, "div", out next_pos);
                    temp = temp.Substring(next_pos);

                    temp1 = XMailHelper.html2text(temp1);
                    temp1 = temp1.Replace("\n", " ").Trim();
                    int model_pos = temp1.IndexOf("Model:");
                    int sku_pos = temp1.IndexOf("SKU:");
                    model_pos = (model_pos == -1) ? int.MaxValue : model_pos;
                    sku_pos = (sku_pos == -1) ? int.MaxValue : sku_pos;
                    if (model_pos == int.MaxValue && sku_pos == int.MaxValue)
                    {
                        title = temp;
                    }
                    else
                    {
                        title = temp1.Substring(0, Math.Min(model_pos, sku_pos));
                        if (sku_pos != int.MaxValue)
                        {
                            sku = temp1.Substring(sku_pos + "SKU:".Length).Trim();
                        }
                        else if (model_pos != int.MaxValue)
                        {
                            sku = temp1.Substring(temp1.IndexOf("Model:") + "Model:".Length).Trim();
                        }
                    }

                    temp1 = XMailHelper.find_html_part(temp, "div", out next_pos);
                    temp1 = XMailHelper.html2text(temp1);
                    temp1 = temp1.Replace("\n", " ").Trim();
                    if (temp1.IndexOf("QTY", StringComparison.CurrentCultureIgnoreCase) != -1)
                        temp1 = temp1.Substring(temp1.IndexOf("QTY", StringComparison.CurrentCultureIgnoreCase) + "QTY".Length).Trim();
                    qty = Str_Utils.string_to_int(temp1);

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    report.m_product_items.Add(product);

                    MyLogger.Info($"... SC-3 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");

                    find = htmlbody.IndexOf("class=\"lineItem-meta\"", find + 1);
                }
            }
        }
    }
}
