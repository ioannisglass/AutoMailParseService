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
        private void parse_mail_cc_11(MimeMessage mail, KReportCC card)
        {
            string subject = XMailHelper.get_subject(mail);

            card.m_mail_type = KReportBase.MailType.CC_11;

            card.m_retailer = ConstEnv.RETAILER_SEARS;

            MyLogger.Info($"... CC-11 m_cc_retailer = {card.m_retailer}");

            List<string> item_title_list = new List<string>();
            List<string> item_sku_list = new List<string>();
            List<int> item_qty_list = new List<int>();
            List<float> item_price_list = new List<float>();

            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("order#", StringComparison.CurrentCultureIgnoreCase))
                {
                    // Ex : Order# 979465994 Canceled

                    string temp = line.Substring("order#".Length).Trim();
                    if (temp.IndexOf(" ") != -1)
                        temp = temp.Substring(0, temp.IndexOf(" ")).Trim();
                    card.set_order_id(temp);
                    MyLogger.Info($"... CC-11 order id = {temp}");
                    continue;
                }
                if (line.IndexOf(" order #", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    // Ex : Your item below has been canceled from order #865362239 and you have not been charged. Please ...

                    string temp = line.Substring(line.IndexOf(" order #", StringComparison.CurrentCultureIgnoreCase) + " order #".Length).Trim();
                    if (temp.IndexOf(" ") != -1)
                        temp = temp.Substring(0, temp.IndexOf(" ")).Trim();
                    card.set_order_id(temp);
                    MyLogger.Info($"... CC-11 order id = {temp}");
                    continue;
                }
                if (line.EndsWith("order", StringComparison.CurrentCultureIgnoreCase) && lines[i + 1].Trim()[0] == '#')
                {
                    // Ex :
                    //      We really appreciate your business and thank you for placing your order. Unfortunately, order
                    //      #992750952, placed on Tue, Nov 08, 2016, has been canceled, and ...

                    string temp = line + lines[++i].Trim();
                    temp = temp.Substring(temp.IndexOf("order#", StringComparison.CurrentCultureIgnoreCase) + "order#".Length).Trim();
                    if (temp.IndexOf(" ") != -1)
                        temp = temp.Substring(0, temp.IndexOf(" ")).Trim();
                    if (temp.IndexOf(",") != -1)
                        temp = temp.Substring(0, temp.IndexOf(",")).Trim();
                    card.set_order_id(temp);
                    MyLogger.Info($"... CC-11 order id = {temp}");
                    continue;
                }

                if (line.ToUpper() == "ITEM DETAILS" && lines[i + 1].Trim().ToUpper() == "PRICE")
                {
                    int k = i + 2;
                    string temp = lines[k].Trim();
                    if (temp.StartsWith("Item#", StringComparison.CurrentCultureIgnoreCase))
                    {
                        string sku = temp.Substring("Item#".Length).Trim();
                        item_sku_list.Add(sku);
                        k++;
                    }
                    string title = lines[k].Trim();
                    item_title_list.Add(title);

                    i = k;
                    continue;
                }
                if (line.ToUpper() == "QTY" && lines[i + 1].Trim().ToUpper() == "PROMOTIONS")
                {
                    string temp = lines[i + 2].Trim();
                    int qty = Str_Utils.string_to_int(temp);
                    item_qty_list.Add(qty);

                    temp = lines[i - 1].Trim();
                    float price_tmp = Str_Utils.string_to_currency(temp);
                    item_price_list.Add(price_tmp);

                    continue;
                }

                if (line.ToUpper() == "ITEM DETAILS")
                {
                    string temp = "";
                    int k = i + 1;
                    while (!lines[k].Trim().StartsWith("Mfr#", StringComparison.CurrentCultureIgnoreCase))
                    {
                        if (lines[k].Trim().StartsWith("Part#", StringComparison.CurrentCultureIgnoreCase))
                            break;
                        if (lines[k].Trim().StartsWith("UPC:", StringComparison.CurrentCultureIgnoreCase))
                            break;
                        if (lines[k].Trim().ToUpper() == "PRICE")
                            break;
                        if (lines[k].Trim().StartsWith("Size:", StringComparison.CurrentCultureIgnoreCase))
                            break;
                        if (lines[k].Trim().StartsWith("Sold by", StringComparison.CurrentCultureIgnoreCase))
                            break;
                        if (lines[k].Trim().StartsWith("Seller Info", StringComparison.CurrentCultureIgnoreCase))
                            break;
                        if (lines[k].Trim().ToUpper() == "VIEW")
                            break;
                        if (lines[k].Trim().StartsWith("Condition:", StringComparison.CurrentCultureIgnoreCase))
                            break;

                        temp += " " + lines[k].Trim();
                        k++;
                    }
                    item_title_list.Add(temp);
                    i = k - 1;
                    continue;
                }
                if (line.StartsWith("UPC:"))
                {
                    string temp = line.Substring("UPC:".Length).Trim();
                    item_sku_list.Add(temp);
                    continue;
                }
                if (line.ToUpper() == "ITEM SUBTOTAL")
                {
                    string temp = lines[++i].Trim();
                    if (temp.StartsWith("Sale"))
                    {
                        temp = temp.Substring("Sale".Length).Trim();
                        item_price_list.Add(Str_Utils.string_to_currency(temp));
                    }
                    temp = lines[++i].Trim();
                    if (temp.StartsWith("Reg."))
                        temp = lines[++i].Trim();
                    item_qty_list.Add(int.Parse(temp));
                    continue;
                }

                if (line.StartsWith("Total Due:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "TOTAL DUE:")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Total Due:".Length).Trim();
                    card.set_total(Str_Utils.string_to_currency(temp));
                    float f = Str_Utils.string_to_currency(temp);
                    MyLogger.Info($"... CC-11 total = {f}");
                    continue;
                }
                if (line.StartsWith("Total Due", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "TOTAL DUE")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Total Due".Length).Trim();
                    card.set_total(Str_Utils.string_to_currency(temp));
                    float f = Str_Utils.string_to_currency(temp);
                    MyLogger.Info($"... CC-11 total = {f}");
                    continue;
                }

                if (line.StartsWith("Sales Tax", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "SALES TAX")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Sales Tax".Length).Trim();
                    card.m_tax = Str_Utils.string_to_currency(temp);
                    float f = Str_Utils.string_to_currency(temp);
                    MyLogger.Info($"... CC-11 tax = {f}");
                    continue;
                }

                if (line.IndexOf("ending in *", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    string last_4_digits = "";
                    float price = 0;
                    string payment_type;

                    string temp = line;

                    temp = temp.Substring(0, line.IndexOf("ending in *", StringComparison.CurrentCultureIgnoreCase)).Trim();
                    if (temp.LastIndexOf("from", StringComparison.CurrentCultureIgnoreCase) != -1)
                        temp = temp.Substring(temp.LastIndexOf("from", StringComparison.CurrentCultureIgnoreCase) + "from".Length).Trim();
                    if (temp.LastIndexOf("to", StringComparison.CurrentCultureIgnoreCase) != -1)
                        temp = temp.Substring(temp.LastIndexOf("to", StringComparison.CurrentCultureIgnoreCase) + "to".Length).Trim();
                    if (temp.LastIndexOf("for", StringComparison.CurrentCultureIgnoreCase) != -1)
                        temp = temp.Substring(temp.LastIndexOf("for", StringComparison.CurrentCultureIgnoreCase) + "for".Length).Trim();

                    payment_type = temp;

                    temp = line.Substring(line.IndexOf("ending in *", StringComparison.CurrentCultureIgnoreCase) + "ending in *".Length).Trim();
                    if (temp.LastIndexOf("*") != -1)
                        temp = temp.Substring(temp.LastIndexOf("*") + 1).Trim();
                    last_4_digits = temp;

                    temp = lines[++i].Trim();
                    price = Str_Utils.string_to_currency(temp);
                    if (price < 0)
                        price *= -1;

                    if (card.add_payment_card_info(payment_type, last_4_digits, price))
                        MyLogger.Info($"... CC-1 payment_type = {payment_type}, last_digit = {last_4_digits}, price = {price}");
                    continue;
                }
            }
            int n = Math.Min(item_title_list.Count, item_sku_list.Count);
            n = Math.Min(n, item_price_list.Count);
            n = Math.Min(n, item_qty_list.Count);

            for (int i = 0; i < n; i++)
            {
                ZProduct product = new ZProduct();
                product.price = item_price_list[i];
                product.sku = item_sku_list[i];
                product.title = item_title_list[i];
                product.qty = item_qty_list[i];
                card.m_product_items.Add(product);

                MyLogger.Info($"... CC-11 qty = {item_qty_list[i]}, price = {item_price_list[i]}, sku = {item_sku_list[i]}, title = {item_title_list[i]}");
            }
        }
    }
}
