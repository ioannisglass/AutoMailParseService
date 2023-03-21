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
        private void parse_mail_sc_12(MimeMessage mail, KReportSC report)
        {
            if (!XMailHelper.is_bodytext_existed(mail))
            {
                parse_mail_sc_12_for_htmltext(mail, report);
                return;
            }
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.SC_12;

            report.m_retailer = ConstEnv.RETAILER_LOWES;

            MyLogger.Info($"... SC-12 m_sc_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Order #"))
                {
                    string temp = line.Substring("Order #".Length).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... order id = {temp}");
                    continue;
                }
                if (line == "Pickup Item(s)" && lines[i + 1].Trim() == "QTY")
                {
                    string next_line = lines[++i].Trim();
                    while (i < lines.Length && next_line.IndexOf("Billing Information") == -1)
                    {
                        string title = "";
                        string sku = "";
                        int qty = 0;
                        float price = 0;

                        string temp = lines[++i].Trim(); // empty line
                        temp = lines[++i].Trim();
                        title = temp;

                        temp = lines[++i].Trim(); // empty line
                        temp = lines[++i].Trim();
                        if (temp.StartsWith("Item #:"))
                        {
                            temp = temp.Substring("Item #:".Length).Trim();
                            temp = temp.Substring(0, temp.IndexOf(" ")).Trim();
                            sku = temp;
                        }

                        temp = lines[++i].Trim(); // empty line
                        temp = lines[++i].Trim();
                        qty = Str_Utils.string_to_int(temp);

                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        report.m_product_items.Add(product);

                        MyLogger.Info($"... qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                        temp = lines[++i].Trim();
                        next_line = lines[++i].Trim();
                    }

                    continue;
                }
            }
        }
        private void parse_mail_sc_12_for_htmltext(MimeMessage mail, KReportSC report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.SC_12;

            report.m_retailer = ConstEnv.RETAILER_LOWES;

            MyLogger.Info($"... SC-12 m_sc_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Order #"))
                {
                    string temp = line.Substring("Order #".Length).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... order id = {temp}");
                    continue;
                }
                if (line == "Pickup Item(s)" && lines[i + 1].Trim() == "QTY")
                {
                    i += 2;
                    string next_line = lines[i].Trim();
                    while (i < lines.Length && next_line.IndexOf("Billing Information") == -1)
                    {
                        string title = "";
                        string sku = "";
                        int qty = 0;
                        float price = 0;

                        string temp = next_line;
                        title = temp;

                        temp = lines[++i].Trim();
                        if (temp.StartsWith("Item #:"))
                        {
                            temp = temp.Substring("Item #:".Length).Trim();
                            if (temp.IndexOf(" ") != -1)
                                temp = temp.Substring(0, temp.IndexOf(" ")).Trim();
                            else
                                ++i;
                            sku = temp;
                        }

                        temp = lines[++i].Trim();
                        qty = Str_Utils.string_to_int(temp);

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
        }
    }
}
