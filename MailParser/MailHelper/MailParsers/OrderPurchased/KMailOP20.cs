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
        private void parse_mail_op_20(MimeMessage mail, KReportOP report)
        {
            if (!XMailHelper.is_bodytext_existed(mail))
            {
                parse_mail_op_20_for_htmltext(mail, report);
                return;
            }
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_20;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = ConstEnv.RETAILER_DELL;

            MyLogger.Info($"... OP-20 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-20 m_op_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Order "))
                {
                    string temp = line.Substring("Order ".Length).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... OP-20 order id = {temp}");
                }
                if (line.StartsWith("Purchased on"))
                {
                    string temp = line.Substring("Purchased on".Length).Trim();
                    DateTime date = DateTime.Parse(temp);
                    report.m_op_purchase_date = date;
                    MyLogger.Info($"... OP-20 order date = {date}");

                    temp = lines[++i].Trim();
                    if (temp == "Total")
                    {
                        temp = lines[++i].Trim();
                        float total = Str_Utils.string_to_currency(temp);
                        report.set_total(total);
                        MyLogger.Info($"... OP-20 total = {total}");
                    }
                    continue;
                }
                if (line == "Qty" && lines[i + 1].Trim() == "Total")
                {
                    i++;

                    string next_line = lines[++i].Trim();
                    while (i < lines.Length && next_line != "" && !next_line.StartsWith("Subtotal"))
                    {
                        string title = "";
                        string sku = "";
                        int qty = 0;
                        float price = 0;

                        string temp = next_line; // image link

                        temp = lines[++i].Trim();
                        title = temp;

                        temp = lines[++i].Trim();
                        price = Str_Utils.string_to_currency(temp);

                        temp = lines[++i].Trim();
                        qty = Str_Utils.string_to_int(temp);

                        temp = lines[++i].Trim(); // subtotal

                        if (qty > 0)
                        {
                            ZProduct product = new ZProduct();
                            product.price = price;
                            product.sku = sku;
                            product.title = title;
                            product.qty = qty;
                            report.m_product_items.Add(product);

                            MyLogger.Info($"... OP-20 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                        }
                        next_line = lines[++i].Trim();
                    }

                    continue;
                }
                if (line == "Estimated Tax")
                {
                    string temp = lines[++i].Trim();
                    float tax = Str_Utils.string_to_currency(temp);
                    report.m_tax = tax;
                    MyLogger.Info($"... OP-20 tax = {tax}");
                    continue;
                }
                if (line == "Payment Method")
                {
                    string next_line = lines[++i].Trim();
                    while (i < lines.Length && next_line != "")
                    {
                        string temp = next_line;
                        if (temp.LastIndexOf(" ") != -1)
                        {
                            string payment_type = temp.Substring(0, temp.LastIndexOf(" ")).Trim();

                            ZPaymentCard c = new ZPaymentCard(payment_type, "", 0);
                            report.add_payment_card_info(c);

                            MyLogger.Info($"... OP-20 payment_type = {payment_type}, last_digit = \"\", price = 0");
                        }
                        next_line = lines[++i].Trim();
                    }
                    continue;
                }
            }
        }
        private void parse_mail_op_20_for_htmltext(MimeMessage mail, KReportOP report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_20;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = ConstEnv.RETAILER_DELL;

            MyLogger.Info($"... OP-20 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-20 m_op_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            if (lines.Length < 5)
            {
                string revised_text = XMailHelper.revise_concated_bodytext(XMailHelper.get_bodytext2(mail));
                lines = revised_text.Replace("\r", "").Split('\n');
            }
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.ToUpper() == "ORDER")
                {
                    string temp = lines[++i].Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... OP-20 order id = {temp}");
                    continue;
                }
                if (line.StartsWith("Purchased on"))
                {
                    string temp = "";
                    if (line == "Purchased on")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Purchased on".Length).Trim();
                    DateTime date = DateTime.Parse(temp);
                    report.m_op_purchase_date = date;
                    MyLogger.Info($"... OP-20 order date = {date}");

                    temp = lines[++i].Trim();
                    if (temp == "Total")
                    {
                        temp = lines[++i].Trim();
                        float total = Str_Utils.string_to_currency(temp);
                        report.set_total(total);
                        MyLogger.Info($"... OP-20 total = {total}");
                    }
                    continue;
                }
                if (line == "Qty" && lines[i + 1].Trim() == "Total")
                {
                    i++;

                    string next_line = lines[++i].Trim();
                    while (i < lines.Length && next_line != "" && !next_line.StartsWith("Subtotal") && !next_line.StartsWith("Billing To"))
                    {
                        string title = "";
                        string sku = "";
                        int qty = 0;
                        float price = 0;

                        string temp = next_line; // image link
                        i++;
                        while (i < lines.Length && !lines[i].Trim().StartsWith("$"))
                        {
                            temp += " " + lines[i].Trim();
                            i++;
                        }
                        title = temp;

                        temp = lines[i].Trim();
                        price = Str_Utils.string_to_currency(temp);

                        temp = lines[++i].Trim();
                        qty = Str_Utils.string_to_int(temp);

                        if (qty > 0)
                        {
                            ZProduct product = new ZProduct();
                            product.price = price;
                            product.sku = sku;
                            product.title = title;
                            product.qty = qty;
                            report.m_product_items.Add(product);

                            MyLogger.Info($"... OP-20 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                        }

                        temp = lines[++i].Trim(); // subtotal

                        ++i;
                        while (lines[i].Trim() == "-->")
                            i++;

                        next_line = lines[i].Trim();
                    }

                    continue;
                }
                if (line == "Estimated Tax")
                {
                    string temp = lines[++i].Trim();
                    float tax = Str_Utils.string_to_currency(temp);
                    report.m_tax = tax;
                    MyLogger.Info($"... OP-20 tax = {tax}");
                    continue;
                }
                if (line == "Payment Method")
                {
                    int k = i + 1;
                    while (lines[k].Trim().IndexOf("-") != -1 && !lines[k].Trim().StartsWith("NEED HELP", StringComparison.CurrentCultureIgnoreCase))
                    {
                        string temp = lines[k].Trim();
                        string payment_type = temp.Substring(temp.LastIndexOf("-") + 1).Trim();

                        ZPaymentCard c = new ZPaymentCard(payment_type, "", 0);
                        report.add_payment_card_info(c);

                        MyLogger.Info($"... OP-20 payment_type = {payment_type}, last_digit = \"\", price = 0");
                        k++;
                    }
                    i = k - 1;
                    continue;
                }
                if (line.ToUpper() == "SHIPPING TO")
                {
                    string full_address = "";
                    string state_address = "";
                    int k = i + 1;
                    while (k < i + 8)
                    {
                        full_address += " " + lines[k].Trim();
                        state_address = XMailHelper.get_address_state_name(full_address);
                        if (state_address != "")
                            break;
                        k++;
                    }
                    full_address = full_address.Trim();
                    state_address = state_address.Trim();
                    if (state_address != "")
                    {
                        report.set_address(full_address, state_address);
                        MyLogger.Info($"... OP-20 full_address = {full_address}, state_address = {state_address}");
                    }
                    i = k;
                    continue;
                }
            }
        }
    }
}
