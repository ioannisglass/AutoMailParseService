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
    partial class KMailBaseOP : KMailBaseParser
    {
        private void parse_mail_op_21(MimeMessage mail, KReportOP report)
        {
            parse_mail_op_21_for_htmltext(mail, report);
        }
        private void parse_mail_op_21_for_htmltext(MimeMessage mail, KReportOP report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_21;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = ConstEnv.RETAILER_GOOGLEEXPRESS;

            MyLogger.Info($"... OP-21 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-21 m_op_retailer = {report.m_retailer}");

            if (!subject.StartsWith("Thanks for your order"))
                throw new Exception($"Invalid OP-21 mail. incorrect subject : {subject}");

            string order_id = subject.Substring("Thanks for your order".Length).Trim();
            if (order_id[0] == '(')
                order_id = order_id.Substring(1).Trim();
            if (order_id.EndsWith(")"))
                order_id = order_id.Substring(0, order_id.Length - 1).Trim();
            if (order_id != "")
            {
                report.set_order_id(order_id);
                MyLogger.Info($"... OP-21 order id = {order_id}");
            }

            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Thanks for your order", StringComparison.CurrentCultureIgnoreCase))
                {
                    order_id = lines[++i].Trim();
                    if (order_id[0] == '(')
                        order_id = order_id.Substring(1).Trim();
                    if (order_id.EndsWith(")"))
                        order_id = order_id.Substring(0, order_id.Length - 1).Trim();
                    if (order_id != "")
                    {
                        report.set_order_id(order_id);
                        MyLogger.Info($"... OP-21 order id = {order_id}");
                    }
                    continue;
                }
                if (line.EndsWith("Ordered", StringComparison.CurrentCultureIgnoreCase))
                {
                    string title = "";
                    string sku = "";
                    int qty = 0;
                    float price = 0;

                    string temp = line.Substring(0, line.Length - "Ordered".Length).Trim();
                    if (!int.TryParse(temp, out qty))
                        continue;
                    title = lines[i - 1].Trim();

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    report.m_product_items.Add(product);

                    MyLogger.Info($"... OP-21 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");

                    continue;
                }
                if (line.StartsWith("Arrives by", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp = line.Substring("Arrives by".Length);
                    DateTime date = DateTime.Parse(temp);
                    report.m_op_purchase_date = date;
                    MyLogger.Info($"... OP-21 purchase date = {date}");
                    continue;
                }
                if (line.StartsWith("Estimated total", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp = line.Substring("Estimated total".Length);
                    float total = Str_Utils.string_to_currency(temp);
                    report.set_total(total);
                    MyLogger.Info($"... OP-21 total = {total}");
                    continue;
                }
            }
        }
    }
}
