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
        private void parse_mail_op_15(MimeMessage mail, KReportOP report)
        {
            if (!XMailHelper.is_bodytext_existed(mail))
            {
                parse_mail_op_15_for_htmltext(mail, report);
            }
            else
            {
                parse_mail_op_15_for_bodytext(mail, report);
            }
            if (report.m_order_id == "")
                get_op15_order_id(mail, report);
        }
        private void get_op15_order_id(MimeMessage mail, KReportOP card)
        {
            string subject = XMailHelper.get_subject(mail);

            if (!subject.StartsWith("Confirmation of Staples Order: #"))
                throw new Exception($"Invalid OP-15 mail. incorrect subject : {subject}");

            string temp = subject.Substring("Confirmation of Staples Order: #".Length).Trim();
            if (temp != "")
            {
                card.set_order_id(temp);
                MyLogger.Info($"... OP-15 order id = {temp}");
            }
        }
        private void parse_mail_op_15_for_bodytext(MimeMessage mail, KReportOP report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_15;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = ConstEnv.RETAILER_STAPLES;

            MyLogger.Info($"... OP-15 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-15 m_op_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.IndexOf("Item #") != -1)
                {
                    string temp = line.Substring(line.IndexOf("Item #") + "Item #".Length);
                    string sku = temp.Trim();

                    temp = lines[i - 1].Trim();
                    if (temp.IndexOf("<") != -1)
                        temp = temp.Substring(0, temp.IndexOf("<"));
                    string title = temp.Trim();

                    int qty = 0;
                    float price = 0;
                    while (i < lines.Length)
                    {
                        string next_line = lines[++i].Trim();
                        if (next_line == "Quantity:")
                        {
                            line = lines[++i].Trim();
                            qty = Str_Utils.string_to_int(line);

                            line = lines[++i].Trim();
                            if (line == "Price:")
                            {
                                line = lines[++i].Trim();
                                price = Str_Utils.string_to_currency(line);
                            }

                            break;
                        }
                    }

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    report.m_product_items.Add(product);

                    MyLogger.Info($"... OP-15 item sku = {sku}, title = {title}, qty = {qty}, price = {price}");

                    continue;
                }
                if (line.IndexOf("ORDER NUMBER:") != -1)
                {
                    string temp = line.Substring(line.IndexOf("ORDER NUMBER:") + "ORDER NUMBER:".Length).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... OP-15 order id = {temp}");
                    continue;
                }
                if (line.IndexOf("ORDER DATE:") != -1)
                {
                    string temp = line.Substring(line.IndexOf("ORDER DATE:") + "ORDER DATE:".Length).Trim();
                    DateTime date = DateTime.Parse(temp);
                    report.m_op_purchase_date = date;
                    MyLogger.Info($"... OP-15 Date Of Purchase = {date}");
                    continue;
                }
                if (line.StartsWith("Tax:"))
                {
                    string temp = line.Substring("Tax:".Length).Trim();
                    float tax = Str_Utils.string_to_currency(temp);
                    report.m_tax = tax;
                    MyLogger.Info($"... OP-15 tax = {tax}");
                    continue;
                }
                if (line.StartsWith("Order Total:"))
                {
                    string temp = line.Substring("Order Total:".Length).Trim();
                    float total = Str_Utils.string_to_currency(temp);
                    report.set_total(total);
                    MyLogger.Info($"... OP-15 order total = {total}");
                    continue;
                }
                if (line == "Payment Method")
                {
                    string next_line;

                    while (i < lines.Length)
                    {
                        next_line = lines[++i].Trim();
                        if (next_line == "")
                            break;

                        if (next_line.IndexOf("ending in") != -1)
                        {
                            string payment_type = next_line.Substring(0, next_line.IndexOf("ending in")).Trim();
                            string temp = next_line.Substring(next_line.IndexOf("ending in") + "ending in".Length).Trim();
                            string last_4_digit = "";
                            float price = 0;
                            if (temp.IndexOf(":") != -1)
                            {
                                last_4_digit = temp.Substring(0, temp.IndexOf(":")).Trim();
                                price = Str_Utils.string_to_currency(temp.Substring(temp.IndexOf(":") + 1));
                            }
                            ZPaymentCard c = new ZPaymentCard(payment_type, last_4_digit, price);
                            report.add_payment_card_info(c);
                            MyLogger.Info($"... OP-15 card type = {payment_type}, last_4_digit = {last_4_digit}, price = {price}");
                        }
                    }
                    continue;
                }
            }
        }
        private void parse_mail_op_15_for_htmltext(MimeMessage mail, KReportOP report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_15;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = ConstEnv.RETAILER_STAPLES;

            MyLogger.Info($"... OP-15 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-15 m_op_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.IndexOf("ORDER NUMBER:") != -1)
                {
                    string temp = line.Substring(line.IndexOf("ORDER NUMBER:") + "ORDER NUMBER:".Length).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... OP-15 order id = {temp}");
                    continue;
                }
                if (line.IndexOf("ORDER DATE:") != -1)
                {
                    string temp = line.Substring(line.IndexOf("ORDER DATE:") + "ORDER DATE:".Length).Trim();
                    DateTime date = DateTime.Parse(temp);
                    report.m_op_purchase_date = date;
                    MyLogger.Info($"... OP-15 Date Of Purchase = {date}");
                    continue;
                }
                if (line.IndexOf("Order Date:") != -1)
                {
                    string temp = line.Substring(line.IndexOf("Order Date:") + "Order Date:".Length).Trim();
                    DateTime date = DateTime.Parse(temp);
                    report.m_op_purchase_date = date;
                    MyLogger.Info($"... OP-15 Date Of Purchase = {date}");
                    continue;
                }
                if (line == "Tax:")
                {
                    string temp = lines[++i].Trim();
                    float tax = Str_Utils.string_to_currency(temp);
                    report.m_tax = tax;
                    MyLogger.Info($"... OP-15 tax = {tax}");
                    continue;
                }
                if (line.StartsWith("Tax:") && line.IndexOf("$") != -1)
                {
                    string temp = line.Substring("Tax:".Length).Trim();
                    float tax = Str_Utils.string_to_currency(temp);
                    report.m_tax = tax;
                    MyLogger.Info($"... OP-15 tax = {tax}");
                    continue;
                }
                if (line == "Order Total:")
                {
                    string temp = lines[++i].Trim();
                    float total = Str_Utils.string_to_currency(temp);
                    report.set_total(total);
                    MyLogger.Info($"... OP-15 order total = {total}");
                    continue;
                }
                if (line.StartsWith("Order Total:") && line.IndexOf("$") != -1)
                {
                    string temp = line.Substring("Order Total:".Length).Trim();
                    float total = Str_Utils.string_to_currency(temp);
                    report.set_total(total);
                    MyLogger.Info($"... OP-15 order total = {total}");
                    continue;
                }
                if (string.Compare(line, "PAYMENT METHOD", true) == 0)
                {
                    string next_line;

                    while (i < lines.Length)
                    {
                        next_line = lines[++i].Trim();
                        if (next_line == "")
                            break;

                        if (next_line.IndexOf("ending in") == -1)
                            break;
                        string payment_type = next_line.Substring(0, next_line.IndexOf("ending in")).Trim();
                        string temp = next_line.Substring(next_line.IndexOf("ending in") + "ending in".Length).Trim();
                        string last_4_digit = "";
                        float price = 0;
                        if (temp.IndexOf(":") != -1)
                        {
                            last_4_digit = temp.Substring(0, temp.IndexOf(":")).Trim();
                        }
                        temp = lines[++i].Trim();
                        price = Str_Utils.string_to_currency(temp);
                        ZPaymentCard c = new ZPaymentCard(payment_type, last_4_digit, price);
                        report.add_payment_card_info(c);
                        MyLogger.Info($"... OP-15 card type = {payment_type}, last_4_digit = {last_4_digit}, price = {price}");
                    }
                    continue;
                }
                if (line.StartsWith("Ship To:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Ship To:".Length).Trim();
                    string full_address = temp;
                    string state_address = XMailHelper.get_address_state_name(full_address);
                    full_address = full_address.Trim();
                    state_address = state_address.Trim();
                    if (state_address != "")
                    {
                        report.set_address(full_address, state_address);
                        MyLogger.Info($"... OP-15 full_address = {full_address}, state_address = {state_address}");
                    }
                    continue;
                }
            }

            string html_text = mail.HtmlBody;
            if (html_text != null && html_text.IndexOf("You Paid</td>") != -1)
            {
                string temp = html_text.Substring(html_text.IndexOf("You Paid</td>"));
                temp = temp.Substring(temp.IndexOf("</tr>") + "</tr>".Length).Trim();
                int next_pos;
                if (temp.StartsWith("</table")) // from mail id 7916, 7917, 7918
                {
                    temp = XMailHelper.find_html_part(temp, "table", out next_pos);
                    temp = XMailHelper.find_html_part(temp, "table", out next_pos);
                }
                else if (temp.StartsWith("</tbody"))
                {
                    temp = XMailHelper.find_html_part(temp, "tbody", out next_pos);
                    temp = XMailHelper.find_html_part(temp, "tbody", out next_pos);
                }

                while (true)
                {
                    string title = "";
                    string sku = "";
                    float price = 0;
                    int qty = 0;

                    string tr_part = XMailHelper.find_html_part(temp, "tr", out next_pos);
                    if (tr_part == "")
                        break;
                    string tr_htmltext = XMailHelper.html2text(tr_part);
                    if (tr_htmltext == "" || tr_htmltext.StartsWith("Your warranty information", StringComparison.CurrentCultureIgnoreCase))
                        break;
                    temp = temp.Substring(next_pos);

                    string td_part = XMailHelper.find_html_part(tr_part, "td", out next_pos);
                    tr_part = tr_part.Substring(next_pos);

                    td_part = XMailHelper.find_html_part(tr_part, "td", out next_pos);
                    tr_part = tr_part.Substring(next_pos);

                    string temp1 = XMailHelper.html2text(td_part);
                    temp1 = temp1.Replace("\r\n", " ");
                    temp1 = temp1.Replace("\n", " ");

                    title = temp1.Substring(0, temp1.IndexOf("Item #")).Trim();
                    sku = temp1.Substring(temp1.IndexOf("Item #") + "Item #".Length).Trim();

                    td_part = XMailHelper.find_html_part(tr_part, "td", out next_pos);
                    tr_part = tr_part.Substring(next_pos);

                    temp1 = XMailHelper.html2text(td_part);
                    price = Str_Utils.string_to_currency(temp1);

                    td_part = XMailHelper.find_html_part(tr_part, "td", out next_pos);
                    tr_part = tr_part.Substring(next_pos);

                    temp1 = XMailHelper.html2text(td_part);
                    qty = Str_Utils.string_to_int(temp1);

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    report.m_product_items.Add(product);

                    MyLogger.Info($"... OP-15 item sku = {sku}, title = {title}, qty = {qty}, price = {price}");
                }
            }
        }
    }
}
