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
        private void parse_mail_op_7(MimeMessage mail, KReportOP report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_7;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = ConstEnv.RETAILER_SAMSCLUB;

            MyLogger.Info($"... OP-7 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-7 m_op_retailer = {report.m_retailer}");

            int mail_type = 0;

            List<string> title_list = new List<string>();
            List<string> sku_list = new List<string>();
            List<int> qty_list = new List<int>();
            List<float> price_list = new List<float>();

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (i < lines.Length - 4 && line.ToUpper() == "ITEM" && lines[i + 1].Trim().ToUpper() == "QTY" && lines[i + 2].Trim().ToUpper() == "ORIG. PRICE" && lines[i + 3].Trim().ToUpper() == "SUBTOTAL")
                {
                    mail_type = 1;
                    break;
                }
                if (line.ToUpper() == "ITEMQTYORIG. PRICESUBTOTAL")
                {
                    mail_type = 2;
                    break;
                }
            }
            if (mail_type == 0)
                throw new Exception("Unknown mail format.");

            lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Order #"))
                {
                    string temp;
                    if (line.ToUpper() == "ORDER #")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Order #".Length).Trim();
                    if (temp.IndexOf("See your order status", StringComparison.CurrentCultureIgnoreCase) != -1)
                        temp = temp.Substring(0, temp.IndexOf("See your order status", StringComparison.CurrentCultureIgnoreCase)).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... OP-7 order id = {temp}");
                    continue;
                }
                if (mail_type == 1 && line.StartsWith("Item #", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp = "";
                    int k = i - 1;
                    while (lines[k].Trim() != "-->" && k > 0)
                    {
                        temp = lines[k] + " " + temp;
                        k--;
                    }
                    temp = temp.Trim();
                    title_list.Add(temp);

                    temp = "";
                    if (line.ToUpper() == "ITEM #")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Item #".Length).Trim();
                    sku_list.Add(temp);
                    continue;
                }
                if (mail_type == 1 && line.StartsWith("Qty:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "QTY:")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Qty:".Length).Trim();
                    qty_list.Add(Str_Utils.string_to_int(temp));
                    continue;
                }
                if (mail_type == 1 && line.StartsWith("Orig price:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp = line.Substring("Orig price:".Length).Trim();
                    price_list.Add(Str_Utils.string_to_currency(temp));
                    continue;
                }
                if (mail_type == 2 && line.IndexOf("Qty:", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    string title = "";
                    string sku = "";
                    int qty = 0;
                    float price = 0;

                    string temp = line;
                    int k = i - 1;
                    while (!lines[k].Trim().StartsWith("-->") && k > 0)
                    {
                        temp = lines[k] + " " + temp;
                        k--;
                    }
                    temp = lines[k].Substring(3).Trim() + " " + temp;
                    temp = temp.Trim();

                    if (temp.IndexOf("Item #", StringComparison.CurrentCultureIgnoreCase) == -1)
                        continue;
                    string temp1 = temp.Substring(0, temp.IndexOf("Item #", StringComparison.CurrentCultureIgnoreCase)).Trim();
                    title = temp1;

                    temp = temp.Substring(temp.IndexOf("Item #", StringComparison.CurrentCultureIgnoreCase) + "Item #".Length).Trim();
                    if (temp.IndexOf("|") != -1)
                    {
                        sku = temp.Substring(0, temp.IndexOf("|")).Trim();
                        temp = temp.Substring(temp.IndexOf("|") + 1).Trim();
                    }
                    else if (temp.IndexOf(" ") != -1)
                    {
                        sku = temp.Substring(0, temp.IndexOf(" ")).Trim();
                        temp = temp.Substring(temp.IndexOf(" ") + 1).Trim();
                    }
                    else
                    {
                        continue;
                    }

                    if (temp.IndexOf("Qty:", StringComparison.CurrentCultureIgnoreCase) == -1)
                        continue;
                    temp = temp.Substring(temp.IndexOf("Qty:", StringComparison.CurrentCultureIgnoreCase) + "Qty:".Length).Trim();
                    if (temp.IndexOf("Orig price:", StringComparison.CurrentCultureIgnoreCase) == -1)
                        continue;
                    temp1 = temp.Substring(0, temp.IndexOf("Orig price:", StringComparison.CurrentCultureIgnoreCase));
                    qty = Str_Utils.string_to_int(temp1);

                    temp = temp.Substring(temp.IndexOf("Orig price:", StringComparison.CurrentCultureIgnoreCase) + "Orig price:".Length).Trim();
                    if (temp.IndexOf("Subtotal:", StringComparison.CurrentCultureIgnoreCase) == -1)
                        continue;
                    temp1 = temp.Substring(0, temp.IndexOf("Subtotal:", StringComparison.CurrentCultureIgnoreCase));
                    price = Str_Utils.string_to_currency(temp1);

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    report.m_product_items.Add(product);

                    MyLogger.Info($"... OP-7 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");

                }
                if (line.StartsWith("Sales tax", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "SALES TAX")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Sales tax".Length).Trim();
                    float tax = Str_Utils.string_to_currency(temp);
                    report.m_tax = tax;
                    MyLogger.Info($"... OP-7 tax = {tax}");
                    continue;
                }
                if (line.StartsWith("Pay online", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "PAY ONLINE")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Pay online".Length).Trim();
                    float total = Str_Utils.string_to_currency(temp);
                    report.set_total(total);
                    MyLogger.Info($"... OP-7 total = {total}");
                    continue;
                }
                if (line.StartsWith("Payment info", StringComparison.CurrentCultureIgnoreCase) && lines[i + 1].Trim().ToUpper() == "PAYMENT METHOD")
                {
                    string temp;

                    i += 2;
                    string next_line = lines[i].Trim();
                    while (i < lines.Length && next_line != "" && next_line.IndexOf("*") != -1)
                    {
                        temp = next_line;

                        string payment_type = temp.Substring(0, temp.IndexOf("*")).Trim();
                        temp = temp.Substring(temp.IndexOf("*") + 1).Trim();
                        if (temp.IndexOf(" ") != -1)
                        {
                            string last_digit = temp.Substring(0, temp.IndexOf(" ")).Trim();

                            temp = temp.Substring(temp.IndexOf(" ") + 1).Trim();
                            float price = Str_Utils.string_to_currency(temp);

                            ZPaymentCard c = new ZPaymentCard(payment_type, last_digit, price);
                            report.add_payment_card_info(c);

                            MyLogger.Info($"... OP-7 payment_type = {payment_type}, last_digit = {last_digit}, price = {price}");
                        }
                        else if (temp.IndexOf("$") != -1)
                        {
                            string last_digit = temp.Substring(0, temp.IndexOf("$")).Trim();

                            temp = temp.Substring(temp.IndexOf("$") + 1).Trim();
                            float price = Str_Utils.string_to_currency(temp);

                            ZPaymentCard c = new ZPaymentCard(payment_type, last_digit, price);
                            report.add_payment_card_info(c);

                            MyLogger.Info($"... OP-7 payment_type = {payment_type}, last_digit = {last_digit}, price = {price}");
                        }
                        next_line = lines[++i].Trim();
                    }
                    continue;
                }
                if (line.StartsWith("Payment method", StringComparison.CurrentCultureIgnoreCase))
                {
                    /**
                     *          Payment method
                     *          DISCOVER:xxxx-xxxx-xxxx-8273 
                     *
                     *          Payment method
                     *          Gift Card:xxxxxxxxxxxx7586
                     *          
                     *          Payment methodDISCOVER:xxxx-xxxx-xxxx-5295
                     *          
                     *          Payment methodGift Card:xxxxxxxxxxxx0794
                     **/

                    string temp;

                    if (line.ToUpper() == "PAYMENT METHOD")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Payment method".Length).Trim();

                    if (temp.IndexOf(":") == -1)
                        continue;
                    string payment_type = temp.Substring(0, temp.IndexOf(":")).Trim();
                    temp = temp.Substring(temp.IndexOf(":") + 1).Trim();

                    if (temp.IndexOf("x") != -1)
                    {
                        string last_digit = temp.Substring(temp.LastIndexOf("x") + 1).Trim();
                        if (last_digit[0] == '-')
                            last_digit = last_digit.Substring(1, last_digit.Length - 1).Trim();

                        float price = 0;

                        ZPaymentCard c = new ZPaymentCard(payment_type, last_digit, price);
                        report.add_payment_card_info(c);

                        MyLogger.Info($"... OP-7 payment_type = {payment_type}, last_digit = {last_digit}, price = {price}");
                    }

                    continue;
                }
                if (line.ToUpper() == "SHIP TO:" && lines[i - 1].Trim().ToUpper() == "SHIPPING DETAILS")
                {
                    string temp = lines[++i].Trim();
                    string full_address = "";
                    string state_address = "";
                    int k = 0;
                    while (k++ < 5)
                    {
                        full_address += " " + temp;
                        temp = lines[++i].Trim();

                        state_address = XMailHelper.get_address_state_name(full_address);
                        if (state_address != "")
                            break;
                    }
                    full_address = full_address.Trim();
                    state_address = state_address.Trim();
                    if (state_address != "")
                    {
                        report.set_address(full_address, state_address);
                        MyLogger.Info($"... OP-7 full_address = {full_address}, state_address = {state_address}");
                    }
                    continue;
                }
            }

            if (mail_type == 1)
            {
                int n = Math.Min(title_list.Count, sku_list.Count);
                n = Math.Min(n, qty_list.Count);
                n = Math.Min(n, price_list.Count);

                for (int i = 0; i < n; i++)
                {
                    ZProduct product = new ZProduct();
                    product.price = price_list[i];
                    product.sku = sku_list[i];
                    product.title = title_list[i];
                    product.qty = qty_list[i];
                    report.m_product_items.Add(product);

                    MyLogger.Info($"... OP-7 qty = {qty_list[i]}, price = {price_list[i]}, sku = {sku_list[i]}, title = {title_list[i]}");
                }
            }
        }
    }
}
