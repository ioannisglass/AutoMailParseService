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
        private void parse_mail_sc_1(MimeMessage mail, KReportSC report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.SC_1;

            report.m_retailer = ConstEnv.RETAILER_BASS;

            MyLogger.Info($"... SC-1 m_sc_retailer = {report.m_retailer}");

            if (XMailHelper.is_bodytext_existed(mail))
            {
                string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i].Trim();

                    if (line.StartsWith("Order number:"))
                    {
                        string temp = line.Substring("Order number:".Length).Trim();
                        report.set_order_id(temp);
                        MyLogger.Info($"... SC-1 order id = {temp}");
                        continue;
                    }
                    if (line.ToUpper() == "ORDER #:")
                    {
                        string temp = lines[++i].Trim();
                        report.set_order_id(temp);
                        MyLogger.Info($"... SC-1 order id = {temp}");
                        continue;
                    }
                    if (line.StartsWith("Item #:", StringComparison.CurrentCultureIgnoreCase)
                        && lines[i + 1].Trim().StartsWith("Qty:", StringComparison.CurrentCultureIgnoreCase)
                        && lines[i + 2].Trim().StartsWith("Description:", StringComparison.CurrentCultureIgnoreCase)
                        && lines[i + 3].Trim().StartsWith("Unit Cost:", StringComparison.CurrentCultureIgnoreCase)
                        && lines[i + 4].Trim().StartsWith("Status:", StringComparison.CurrentCultureIgnoreCase)
                        )
                    {
                        string title = "";
                        string sku = "";
                        int qty = 0;
                        float price = 0;
                        string status = "";

                        string temp = line.Substring("Item #:".Length).Trim();
                        sku = temp;

                        temp = lines[++i].Trim();
                        temp = temp.Substring("Qty:".Length).Trim();
                        qty = Str_Utils.string_to_int(temp);

                        temp = lines[++i].Trim();
                        temp = temp.Substring("Description:".Length).Trim();
                        title = temp;

                        temp = lines[++i].Trim();
                        temp = temp.Substring("Unit Cost:".Length).Trim();
                        price = Str_Utils.string_to_currency(temp);

                        temp = lines[++i].Trim(); // status
                        temp = temp.Substring("Status:".Length).Trim();
                        status = temp;

                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        product.status = status;
                        report.m_product_items.Add(product);

                        MyLogger.Info($"... SC-1 qty = {qty}, price = {price}, sku = {sku}, title = {title}, status = {status}");
                        continue;
                    }
                }
            }
            if (report.m_order_id == "" || report.m_product_items.Count == 0)
            {
                string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i].Trim();

                    if (line.StartsWith("Order number:"))
                    {
                        string temp = line.Substring("Order number:".Length).Trim();
                        report.set_order_id(temp);
                        MyLogger.Info($"... SC-1 order id = {temp}");
                        continue;
                    }
                    if (line.ToUpper() == "ORDER #:")
                    {
                        string temp = lines[++i].Trim();
                        if (temp.StartsWith("Descriptive, typographical", StringComparison.CurrentCultureIgnoreCase))
                            continue;
                        if (!char.IsDigit(temp[0])) // Bass order id is digit number.
                            continue;
                        report.set_order_id(temp);
                        MyLogger.Info($"... SC-1 order id = {temp}");
                        continue;
                    }
                    if (report.m_product_items.Count == 0)
                    {
                        if (line == "SKU" && lines[i + 1].Trim() == "Qty" && lines[i + 2].Trim() == "Description" && lines[i + 3].Trim() == "Unit Cost" && lines[i + 4].Trim() == "Status")
                        {
                            i += 5;

                            string nextline = lines[i].Trim();
                            while (i < lines.Length)
                            {
                                string title = "";
                                string sku = "";
                                int qty = 0;
                                float price = 0;

                                string temp = nextline;

                                if (temp.StartsWith("If we can be of ", StringComparison.CurrentCultureIgnoreCase))
                                    break;

                                sku = temp;

                                temp = lines[++i].Trim();
                                qty = Str_Utils.string_to_int(temp);

                                temp = lines[++i].Trim();
                                title = temp;

                                temp = lines[++i].Trim();
                                price = Str_Utils.string_to_currency(temp);

                                temp = lines[++i].Trim(); // status

                                ZProduct product = new ZProduct();
                                product.price = price;
                                product.sku = sku;
                                product.title = title;
                                product.qty = qty;
                                report.m_product_items.Add(product);

                                MyLogger.Info($"... SC-1 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                                nextline = lines[++i].Trim();
                            }
                            continue;
                        }
                    }
                }
            }
            String htmltext = XMailHelper.get_htmltext(mail);
            if (report.m_order_id == "")
            {
                if (htmltext.IndexOf("order_number=") != -1)
                {
                    string temp = htmltext.Substring(htmltext.IndexOf("order_number=") + "order_number=".Length);
                    if (temp.IndexOf("&amp;") != -1)
                    {
                        temp = temp.Substring(0, temp.IndexOf("&amp;"));
                        report.set_order_id(temp);
                        MyLogger.Info($"... SC-1 order id = {temp}");
                    }
                }
            }
            if (htmltext.IndexOf("?tracking_numbers=") != -1)
            {
                string temp1 = htmltext.Substring(0, htmltext.IndexOf("?tracking_numbers="));
                if (temp1.LastIndexOf("/") != -1)
                {
                    temp1 = temp1.Substring(temp1.LastIndexOf("/") + 1);
                    temp1 = get_post_type(temp1);
                    report.m_sc_post_type = temp1;
                    MyLogger.Info($"... SC-1 post_type = {temp1}");
                }

                string temp = htmltext.Substring(htmltext.IndexOf("?tracking_numbers=") + "?tracking_numbers=".Length);
                if (temp.IndexOf("&amp;") != -1)
                {
                    temp = temp.Substring(0, temp.IndexOf("&amp;"));
                    report.set_tracking(temp);
                    MyLogger.Info($"... SC-1 tracking = {temp}");
                }
            }
        }
    }
}
