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
        private void parse_mail_op_19(MimeMessage mail, KReportOP report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_19;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = ConstEnv.RETAILER_DELL;

            MyLogger.Info($"... OP-19 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-19 m_op_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line == "DPID:")
                {
                    string temp = lines[++i].Trim();
                    if (temp.StartsWith("Total:"))
                    {
                        temp = temp.Substring("Total:".Length).Trim();
                        if (temp == "")
                            temp = lines[++i].Trim();
                    }
                    report.set_order_id(temp);
                    MyLogger.Info($"... OP-19 order id = {temp}");
                }
                if (line == "Order Date:")
                {
                    string temp = lines[++i].Trim();
                    if (temp.IndexOf("Company:") != -1)
                        temp = temp.Substring(0, temp.IndexOf("Company:")).Trim();
                    DateTime date = DateTime.Parse(temp);
                    report.m_op_purchase_date = date;
                    MyLogger.Info($"... OP-19 order date = {date}");
                    continue;
                }
                if (line.Replace(" ", "") == "ItemQTYUnitPriceSubtotal")
                {
                    string next_line = lines[++i].Trim();
                    while (i < lines.Length && next_line != "" && !next_line.StartsWith("Shipping:"))
                    {
                        string title = "";
                        string sku = "";
                        int qty = 0;
                        float price = 0;

                        string temp = next_line;
                        title = temp;

                        temp = lines[++i].Trim();
                        if (temp.StartsWith("Estimated Delivery Date:"))
                            temp = lines[++i].Trim();

                        if (temp.LastIndexOf(" ") != -1)
                        {
                            string subtotal = temp.Substring(temp.LastIndexOf(" ")).Trim();

                            temp = temp.Substring(0, temp.LastIndexOf(" ")).Trim();
                            if (temp.LastIndexOf(" ") != -1)
                            {
                                string price_part = temp.Substring(temp.LastIndexOf(" ")).Trim();
                                price = Str_Utils.string_to_currency(price_part);

                                temp = temp.Substring(0, temp.LastIndexOf(" ")).Trim();
                                if (temp.LastIndexOf(" ") != -1)
                                {
                                    string qty_part = temp.Substring(temp.LastIndexOf(" ")).Trim();
                                    qty = Str_Utils.string_to_int(qty_part);
                                }
                            }
                            if (qty > 0)
                            {
                                ZProduct product = new ZProduct();
                                product.price = price;
                                product.sku = sku;
                                product.title = title;
                                product.qty = qty;
                                report.m_product_items.Add(product);

                                MyLogger.Info($"... OP-19 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                            }
                        }
                        next_line = lines[++i].Trim();
                    }

                    continue;
                }
                if (line.StartsWith("Tax:"))
                {
                    string temp = line.Substring("Tax:".Length).Trim();
                    float tax = Str_Utils.string_to_currency(temp);
                    report.m_tax = tax;
                    MyLogger.Info($"... OP-19 tax = {tax}");
                    continue;
                }
                if (line == "Total:")
                {
                    string temp = lines[++i].Trim();
                    float total = Str_Utils.string_to_currency(temp);
                    report.set_total(total);
                    MyLogger.Info($"... OP-19 total = {total}");
                    continue;
                }
                if (line.StartsWith("Method") && line.EndsWith("Total"))
                {
                    string next_line = lines[++i].Trim();
                    while (i < lines.Length && next_line != "" && next_line != "Need help?")
                    {
                        string temp = next_line;
                        if (temp.LastIndexOf(" ") != -1)
                        {
                            string payment_type = temp.Substring(0, temp.LastIndexOf(" ")).Trim();

                            temp = temp.Substring(temp.LastIndexOf(" ") + 1).Trim();
                            float card_price = Str_Utils.string_to_currency(temp);

                            ZPaymentCard c = new ZPaymentCard(payment_type, "", card_price);
                            report.add_payment_card_info(c);

                            MyLogger.Info($"... OP-19 payment_type = {payment_type}, last_digit = \"\", price = {card_price}");
                        }
                        next_line = lines[++i].Trim();
                    }
                    continue;
                }
            }
        }
    }
}
