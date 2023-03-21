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
        private void parse_mail_sc_11(MimeMessage mail, KReportSC report)
        {
            if (!XMailHelper.is_bodytext_existed(mail))
            {
                analyze_mail_order_11_for_htmltext(mail, report);
                return;
            }
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.SC_11;

            report.m_retailer = ConstEnv.RETAILER_SEARS;

            MyLogger.Info($"... SC-11 m_sc_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line == "Shipping Vendor")
                {
                    string temp = lines[++i].Trim();
                    string post_type = get_post_type(temp);
                    report.m_sc_post_type = post_type;
                    MyLogger.Info($"... SC-11 post_type = {post_type}");
                    continue;
                }
                if (line == "Estimated Date:")
                {
                    string temp = lines[++i].Trim();
                    DateTime date = DateTime.Parse(temp);
                    report.m_sc_expected_deliver_date = date;
                    MyLogger.Info($"... SC-11 expected delivery date = {date}");
                    continue;
                }
                if (line.ToUpper() == "ITEM DETAILS")
                {
                    float price = 0;
                    int qty = 0;
                    string sku = "";
                    string title = "";

                    string temp = lines[++i].Trim();
                    if (temp.IndexOf("]") != -1)
                        temp = temp.Substring(temp.IndexOf("]") + 1).Trim();
                    title = temp;

                    string next_line = lines[++i].Trim();
                    while (i < lines.Length && next_line.IndexOf("UPC:", StringComparison.CurrentCultureIgnoreCase) == -1)
                    {
                        next_line = lines[++i].Trim();
                    }
                    if (next_line.IndexOf("UPC:", StringComparison.CurrentCultureIgnoreCase) != -1)
                    {
                        temp = next_line.Substring("UPC:".Length).Trim();
                        sku = temp;
                    }

                    temp = lines[++i].Trim();
                    while (i < lines.Length && !temp.StartsWith("Tracking Number:", StringComparison.CurrentCultureIgnoreCase))
                    {
                        if (temp.StartsWith("PRICE", StringComparison.CurrentCultureIgnoreCase))
                            break;

                        i++;
                        temp = lines[i].Trim();
                    }
                    if (temp.StartsWith("Tracking Number:", StringComparison.CurrentCultureIgnoreCase))
                    {
                        temp = temp.Substring("Tracking Number:".Length).Trim();
                        if (temp.IndexOf("<") != -1)
                            temp = temp.Substring(0, temp.IndexOf("<")).Trim();
                        string tracking = temp;
                        report.set_tracking(tracking);
                        MyLogger.Info($"... tracking = {tracking}");
                    }

                    temp = "";
                    next_line = lines[++i].Trim();
                    while (i < lines.Length)
                    {
                        temp = next_line.Replace(" ", "");
                        if (temp.StartsWith("PRICE", StringComparison.CurrentCultureIgnoreCase))
                            break;
                        next_line = lines[++i].Trim();
                    }
                    if (temp == "PRICEQTYITEMSUBTOTAL")
                    {
                        // first line or item price.

                        temp = lines[++i].Trim();
                        if (temp.StartsWith("Sale"))
                        {
                            temp = temp.Substring("Sale".Length).Trim();
                            price = Str_Utils.string_to_currency(temp);
                            temp = lines[++i].Trim();
                            if (temp.StartsWith("Reg."))
                            {
                                temp = temp.Substring("Reg.".Length).Trim();
                                if (temp.IndexOf(" ") != -1)
                                {
                                    temp = temp.Substring(temp.IndexOf(" ") + 1).Trim();
                                }
                            }
                        }
                        else
                        {
                            if (temp.IndexOf(" ") != -1)
                            {
                                string price_part = temp.Substring(0, temp.IndexOf(" ")).Trim();
                                price = Str_Utils.string_to_currency(price_part);
                                temp = temp.Substring(temp.IndexOf(" ") + 1).Trim();
                            }
                            else
                            {
                                price = Str_Utils.string_to_currency(temp);
                                temp = lines[++i].Trim();
                            }
                        }
                        if (temp.IndexOf(" ") != -1)
                        {
                            string qty_part = temp.Substring(0, temp.IndexOf(" ")).Trim();
                            qty = int.Parse(qty_part);
                            temp = temp.Substring(temp.IndexOf(" ") + 1).Trim();
                        }
                        else
                        {
                            qty = int.Parse(temp);
                            temp = lines[++i].Trim();
                        }
                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        report.m_product_items.Add(product);
                        MyLogger.Info($"... qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                    }
                    else if (lines[i].Trim().ToUpper() == "PRICE" && lines[i + 1].Trim().ToUpper() == "QTY" && lines[i + 2].Trim().ToUpper() == "ITEM SUBTOTAL")
                    {
                        i += 3;

                        temp = lines[i].Trim();
                        if (temp.StartsWith("Sale"))
                        {
                            temp = temp.Substring("Sale".Length).Trim();
                            price = Str_Utils.string_to_currency(temp);
                        }
                        else if (temp.IndexOf(" ") != -1)
                        {
                            string price_part = temp.Substring(0, temp.IndexOf(" ")).Trim();
                            price = Str_Utils.string_to_currency(price_part);
                        }
                        else
                        {
                            price = Str_Utils.string_to_currency(temp);
                        }
                        i++;

                        int qty_tmp;
                        while (!int.TryParse(lines[i].Trim(), out qty_tmp))
                            i++;

                        qty = Str_Utils.string_to_int(lines[i].Trim());
                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        report.m_product_items.Add(product);
                        MyLogger.Info($"... qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                        i++;
                    }
                    continue;
                }
            }
        }
        private void analyze_mail_order_11_for_htmltext(MimeMessage mail, KReportSC report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.SC_11;

            report.m_retailer = ConstEnv.RETAILER_SEARS;

            MyLogger.Info($"... SC-11 m_sc_retailer = {report.m_retailer}");

            List<string> title = new List<string>();
            List<string> sku = new List<string>();
            List<int> qty = new List<int>();
            List<float> price = new List<float>();

            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.IndexOf("your order #", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    string temp = line.Substring(line.IndexOf("your order #", StringComparison.CurrentCultureIgnoreCase) + "your order #".Length).Trim();
                    if (temp.IndexOf(" ") == -1)
                        continue;
                    temp = temp.Substring(0, temp.IndexOf(" ")).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... SC-11 order id = {temp}");
                    continue;
                }
                if (line == "Shipping Vendor")
                {
                    string temp = lines[++i].Trim();
                    string post_type = get_post_type(temp);
                    report.m_sc_post_type = post_type;
                    MyLogger.Info($"... SC-11 post_type = {post_type}");
                    continue;
                }
                if (line == "Estimated Date:")
                {
                    string temp = lines[++i].Trim();
                    DateTime date = DateTime.Parse(temp);
                    report.m_sc_expected_deliver_date = date;
                    MyLogger.Info($"... SC-11 expected delivery date = {date}");
                    continue;
                }
                if (line.ToUpper() == "ITEM DETAILS")
                {
                    string temp = lines[++i].Trim();
                    if (temp.IndexOf("]") != -1)
                        temp = temp.Substring(temp.IndexOf("]") + 1).Trim();
                    title.Add(temp);
                    continue;
                }
                if (line.StartsWith("UPC:"))
                {
                    string temp = line.Substring("UPC:".Length).Trim();
                    sku.Add(temp);
                    continue;
                }
                if (line.StartsWith("Tracking Number:"))
                {
                    string temp = line.Substring("Tracking Number:".Length).Trim();
                    report.set_tracking(temp);
                    MyLogger.Info($"... SC-11 tracking = {temp}");
                    continue;
                }
                if (line.ToUpper() == "ITEM SUBTOTAL")
                {
                    string temp = lines[++i].Trim();
                    if (!temp.StartsWith("Sale"))
                        continue;
                    temp = temp.Substring("Sale".Length).Trim();
                    price.Add(Str_Utils.string_to_currency(temp));

                    temp = lines[++i].Trim();
                    if (temp.StartsWith("Reg."))
                        temp = lines[++i].Trim();
                    qty.Add(int.Parse(temp));
                    continue;
                }
            }
            int n = Math.Min(title.Count, sku.Count);
            n = Math.Min(n, price.Count);
            n = Math.Min(n, qty.Count);

            for (int i = 0; i < n; i++)
            {
                ZProduct product = new ZProduct();
                product.price = price[i];
                product.sku = sku[i];
                product.title = title[i];
                product.qty = qty[i];
                report.m_product_items.Add(product);

                MyLogger.Info($"... SC-11 qty = {qty[i]}, price = {price[i]}, sku = {sku[i]}, item title = {title[i]}");
            }
        }
    }
}
