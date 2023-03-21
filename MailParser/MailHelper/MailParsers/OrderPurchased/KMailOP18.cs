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
        private void parse_mail_op_18(MimeMessage mail, KReportOP report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_18;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = ConstEnv.RETAILER_KOHLS;

            MyLogger.Info($"... OP-18 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-18 m_op_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Order #"))
                {
                    string temp = line.Substring("Order #".Length).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... order id = {temp}");

                    temp = lines[++i].Trim();
                    for (int y = 2000; y < 2100; y++)
                    {
                        if (temp.IndexOf(y.ToString()) != -1)
                        {
                            temp = temp.Substring(0, temp.IndexOf(y.ToString()) + y.ToString().Length).Trim();
                            DateTime date = DateTime.Parse(temp);
                            report.m_op_purchase_date = date;
                            MyLogger.Info($"... OP-18 order date = {date}");
                            break;
                        }
                    }
                    continue;
                }
                if (line.IndexOf("Qty:") != -1)
                {
                    string title = "";
                    string sku = "";
                    int qty = 0;
                    float price = 0;

                    string temp = "";
                    int k = i - 1;
                    string[] days = new string[] { "MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN" };
                    while (!days.Contains(lines[k].Trim().ToUpper()) && !lines[k].Trim().StartsWith("Your Price", StringComparison.CurrentCultureIgnoreCase))
                    {
                        if (lines[k - 1].Trim().StartsWith("Shipping Method", StringComparison.CurrentCultureIgnoreCase))
                            break;
                        if (lines[k].Trim().StartsWith("Shipping Method", StringComparison.CurrentCultureIgnoreCase))
                            break;
                        if (lines[k - 1].Trim().StartsWith("Expected Delivery", StringComparison.CurrentCultureIgnoreCase))
                            break;
                        if (lines[k].Trim().StartsWith("Expected Delivery", StringComparison.CurrentCultureIgnoreCase))
                            break;
                        if (lines[k].Trim().StartsWith("Standard will arrive", StringComparison.CurrentCultureIgnoreCase))
                            break;

                        string temp1 = lines[k].Trim();
                        if (temp1.IndexOf("<") != -1)
                            temp1 = temp1.Substring(0, temp1.IndexOf("<")).Trim();
                        temp = temp1 + " " + temp;
                        k--;
                    }
                    title = temp.Trim();

                    temp = line;
                    temp = temp.Substring(temp.IndexOf("Qty:") + "Qty:".Length).Trim();
                    qty = Str_Utils.string_to_int(temp);

                    temp = lines[++i].Trim();
                    if (temp.StartsWith("SKU #"))
                    {
                        temp = temp.Substring("SKU #".Length).Trim();
                        sku = temp;
                    }

                    k = i + 1;
                    while (!lines[k].Trim().StartsWith("Your Price:", StringComparison.CurrentCultureIgnoreCase))
                        k++;
                    temp = lines[k].Trim();
                    temp = temp.Substring("Your Price:".Length).Trim();
                    price = Str_Utils.string_to_currency(temp);
                    if (qty > 0)
                        price /= qty;

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    report.m_product_items.Add(product);

                    MyLogger.Info($"... OP-18 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");

                    qty = 0;
                    price = 0;
                    sku = "";
                    title = "";

                    continue;
                }
                if (line == "Payment Method")
                {
                    string next_line = lines[++i].Trim();
                    while (i < lines.Length && next_line.IndexOf(" x") != -1)
                    {
                        string temp = next_line.Substring(0, next_line.IndexOf(" x")).Trim();
                        string payment_type = temp;
                        temp = next_line.Substring(next_line.IndexOf(" x") + 2).Trim();
                        if (temp.IndexOf(" ") != -1)
                        {
                            string last_digit = temp.Substring(0, temp.IndexOf(" ")).Trim();
                            temp = temp.Substring(temp.IndexOf(" ") + 1).Trim();
                            float price = Str_Utils.string_to_currency(temp);

                            ZPaymentCard c = new ZPaymentCard(payment_type, last_digit, price);
                            report.add_payment_card_info(c);

                            MyLogger.Info($"... OP-18 payment_type = {payment_type}, last_digit = {last_digit}, price = {price}");
                        }
                        else
                        {
                            string last_digit = temp.Trim();
                            temp = lines[++i].Trim();
                            float price = Str_Utils.string_to_currency(temp);

                            ZPaymentCard c = new ZPaymentCard(payment_type, last_digit, price);
                            report.add_payment_card_info(c);

                            MyLogger.Info($"... OP-18 payment_type = {payment_type}, last_digit = {last_digit}, price = {price}");
                        }
                        next_line = lines[++i].Trim();
                    }
                    continue;
                }
                if (line.StartsWith("Sale Tax:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "SALE TAX:")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Sale Tax:".Length).Trim();
                    float tax = Str_Utils.string_to_currency(temp);
                    report.m_tax = tax;
                    MyLogger.Info($"... OP-18 tax = {tax}");
                    continue;
                }
                if (line.StartsWith("Tax:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "TAX:")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Tax:".Length).Trim();
                    float tax = Str_Utils.string_to_currency(temp);
                    report.m_tax = tax;
                    MyLogger.Info($"... OP-18 tax = {tax}");
                    continue;
                }
                if (line.StartsWith("Total:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "TOTAL:")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Total:".Length).Trim();
                    float total = Str_Utils.string_to_currency(temp);
                    report.set_total(total);
                    MyLogger.Info($"... OP-18 total = {total}");
                    continue;
                }
                if (line.ToUpper() == "SHIPPING" && (lines[i + 1].EndsWith("item) to:", StringComparison.CurrentCultureIgnoreCase) || lines[i + 1].EndsWith("items) to:", StringComparison.CurrentCultureIgnoreCase)))
                {
                    string full_address = "";
                    string state_address = "";
                    int k = i + 2;
                    while (k < i + 7)
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
                        MyLogger.Info($"... OP-18 full_address = {full_address}, state_address = {state_address}");
                    }
                    i = k;
                    continue;
                }
            }
        }
    }
}
