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
        private void parse_mail_sc_2(MimeMessage mail, KReportSC report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.SC_2;

            report.m_retailer = ConstEnv.RETAILER_TARGET;

            MyLogger.Info($"... SC-2 m_sc_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

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
                    MyLogger.Info($"... SC-2 order id = {temp}");
                    continue;
                }
                if (line.IndexOf("Tracking #") != -1)
                {
                    string temp = line.Substring(line.IndexOf("Tracking #") + "Tracking #".Length).Trim();
                    string tracking = temp;

                    temp = line.Substring(0, line.IndexOf("Tracking #")).Trim();
                    string post_type = get_post_type(temp);

                    report.set_tracking(tracking);
                    report.m_sc_post_type = (report.m_sc_post_type == String.Empty || report.m_sc_post_type.IndexOf(post_type) != -1) ? post_type : report.m_sc_post_type + "," + post_type;

                    MyLogger.Info($"... SC-2 post_type = {post_type}, tracking = {tracking}");
                    continue;
                }
                if (line.StartsWith("Qty:"))
                {
                    string title = "";
                    string sku = "";
                    int qty = 0;
                    float price = 0;

                    string temp = line.Substring("Qty:".Length).Trim();
                    qty = Str_Utils.string_to_int(temp);

                    temp = lines[i - 1].Trim();
                    if (temp.IndexOf("<") != -1)
                        temp = temp.Substring(0, temp.IndexOf("<")).Trim();
                    title = temp;

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    report.m_product_items.Add(product);

                    MyLogger.Info($"... SC-2 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                    continue;
                }
            }
        }
    }
}
