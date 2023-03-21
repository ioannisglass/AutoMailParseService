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
        private void parse_mail_op_2(MimeMessage mail, KReportOP report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_2;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = ConstEnv.RETAILER_TARGET;

            MyLogger.Info($"... OP-2 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-2 m_op_retailer = {report.m_retailer}");

            string title = "";
            string sku = "";
            int qty = 0;
            float price = 0;

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Order #"))
                {
                    string temp;
                    if (line == "Order #")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Order #".Length).Trim();
                    if (temp.IndexOf("<") != -1)
                        temp = temp.Substring(0, temp.IndexOf("<")).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... OP-2 order id = {temp}");
                    continue;
                }
                if (line.IndexOf("something special ready!", StringComparison.CurrentCultureIgnoreCase) != -1 || line.StartsWith("Thanks for your order,"))
                {
                    string temp = lines[++i].Trim();
                    if (temp.StartsWith("Placed "))
                    {
                        temp = temp.Substring("Placed ".Length).Trim();
                    }
                    DateTime date;
                    if (!DateTime.TryParse(temp, out date))
                        continue;
                    report.m_op_purchase_date = date;
                    MyLogger.Info($"... OP-2 order date = {date}");
                    continue;
                }
                if (line.StartsWith("Qty:"))
                {
                    string temp = lines[i - 1].Trim();
                    if (temp.IndexOf("<") != -1)
                        temp = temp.Substring(0, temp.IndexOf("<")).Trim();
                    title = temp;

                    temp = line;
                    if (temp.StartsWith("Qty:"))
                    {
                        temp = temp.Substring("Qty:".Length).Trim();
                        qty = Str_Utils.string_to_int(temp);
                    }

                    temp = lines[++i].Trim();
                    if (temp.IndexOf("/") != -1)
                        temp = temp.Substring(0, temp.IndexOf("/")).Trim();
                    price = Str_Utils.string_to_currency(temp);

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    report.m_product_items.Add(product);

                    MyLogger.Info($"... OP-2 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");

                    title = "";
                    sku = "";
                    qty = 0;
                    price = 0;

                    continue;
                }
                if (line == "Estimated Taxes")
                {
                    string temp = lines[++i].Trim();
                    float tax = Str_Utils.string_to_currency(temp);
                    float total = 0;

                    i++;

                    temp = lines[i].Trim();
                    if (temp == "Total")
                    {
                        temp = lines[++i].Trim();
                        total = Str_Utils.string_to_currency(temp);
                    }
                    report.m_tax = tax;
                    report.set_total(total);
                    MyLogger.Info($"... OP-2 tax = {tax}, total = {total}");
                    continue;
                }
                if (report.m_total == 0 && line.ToUpper() == "TOTAL")
                {
                    string temp = lines[++i].Trim();
                    float total = Str_Utils.string_to_currency(temp);
                    report.set_total(total);
                    MyLogger.Info($"... OP-2 total = {total}");
                    continue;
                }

                if (line.ToUpper() == "DELIVERS TO:")
                {
                    string temp = lines[++i].Trim();
                    string full_address = temp;
                    string state_address = XMailHelper.get_address_state_name(full_address);
                    if (state_address != "")
                    {
                        report.set_address(full_address, state_address);
                        MyLogger.Info($"... OP-2 full_address = {full_address}, state_address = {state_address}");
                    }
                    continue;
                }
            }
        }
    }
}
