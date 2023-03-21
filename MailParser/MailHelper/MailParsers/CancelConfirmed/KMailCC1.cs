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
    partial class KMailBaseCC : KMailBaseParser
    {
        private void parse_mail_cc_1(MimeMessage mail, KReportCC card)
        {
            string subject = XMailHelper.get_subject(mail);

            card.m_mail_type = KReportBase.MailType.CC_1;

            card.m_retailer = ConstEnv.RETAILER_BASS;

            MyLogger.Info($"... CC-1 m_cc_retailer = {card.m_retailer}");

            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Order Number:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp = line.Substring("Order Number:".Length).Trim();
                    if (temp.IndexOf("Order Date:", StringComparison.CurrentCultureIgnoreCase) != -1)
                        temp = temp.Substring(0, temp.IndexOf("Order Date:", StringComparison.CurrentCultureIgnoreCase)).Trim();
                    card.set_order_id(temp);
                    MyLogger.Info($"... CC-1 order id = {temp}");
                    continue;
                }
                if (line.IndexOf("SKU:", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    string title = "";
                    string sku = "";
                    int qty = 0;
                    float price = 0;

                    string temp = line;

                    i++;
                    while (!lines[i].Trim().StartsWith("Quantity", StringComparison.CurrentCultureIgnoreCase))
                    {
                        temp += " " + lines[i].Trim();
                        i++;
                    }

                    title = temp.Substring(0, temp.IndexOf("SKU:", StringComparison.CurrentCultureIgnoreCase)).Trim();
                    sku = temp.Substring(temp.IndexOf("SKU:", StringComparison.CurrentCultureIgnoreCase) + "SKU:".Length).Trim();

                    if (lines[i].Trim().ToUpper() == "QUANTITY" && lines[i + 1].Trim().ToUpper() == "PRICE")
                        i += 2;
                    else if (lines[i].Trim().IndexOf("Price", StringComparison.CurrentCultureIgnoreCase) != -1)
                        i++;
                    else
                        continue;
                    temp = "";
                    while (lines[i].Trim().IndexOf("$") != -1)
                    {
                        temp += " " + lines[i].Trim();
                        i++;
                    }

                    string qty_part = temp.Substring(0, temp.IndexOf("$")).Trim();
                    qty = Str_Utils.string_to_int(qty_part);

                    temp = temp.Substring(temp.IndexOf("$") + 1).Trim();
                    if (temp.IndexOf("$") != -1)
                        temp = temp.Substring(0, temp.IndexOf("$")).Trim();
                    price = Str_Utils.string_to_currency(temp);

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    card.m_product_items.Add(product);

                    MyLogger.Info($"... CC-1 qty = {qty}, price = {price}, sku = {sku}, title = {title}");
                    continue;
                }
                if (line.StartsWith("Tax", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "TAX")
                        temp = lines[i + 1].Trim();
                    else
                        temp = line.Substring("Tax".Length).Trim();

                    float tax = Str_Utils.string_to_currency(temp);
                    card.m_tax = tax;

                    MyLogger.Info($"... CC-1 tax = {tax}");
                    continue;
                }
                if (line.IndexOf("Ending in ", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    string temp = line;

                    i++;
                    while (lines[i].Trim().IndexOf("Ending in ", StringComparison.CurrentCultureIgnoreCase) == -1 && !lines[i].Trim().StartsWith("We value your business", StringComparison.CurrentCultureIgnoreCase))
                    {
                        temp += " " + lines[i].Trim();
                        i++;
                    }
                    if (lines[i].Trim().IndexOf("Ending in ", StringComparison.CurrentCultureIgnoreCase) != -1)
                        i--;

                    string payment_type = temp.Substring(0, line.IndexOf("Ending in ", StringComparison.CurrentCultureIgnoreCase)).Trim();

                    temp = temp.Substring(line.IndexOf("Ending in ", StringComparison.CurrentCultureIgnoreCase) + "Ending in ".Length).Trim();
                    string last_4_digits = "";
                    float price = 0;

                    if (temp.IndexOf(" ") != -1)
                    {
                        last_4_digits = temp.Substring(0, line.IndexOf(" ")).Trim();
                        temp = temp.Substring(line.IndexOf(" ") + 1).Trim();
                        price = Str_Utils.string_to_currency(temp);
                    }

                    if (card.add_payment_card_info(payment_type, last_4_digits, price))
                        MyLogger.Info($"... CC-1 payment_type = {payment_type}, last_digit = {last_4_digits}, price = {price}");
                    continue;
                }
            }
        }
    }
}
