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
        private void parse_mail_cc_3(MimeMessage mail, KReportCC card)
        {
            string subject = XMailHelper.get_subject(mail);

            card.m_mail_type = KReportBase.MailType.CC_3;

            card.m_retailer = ConstEnv.RETAILER_BESTBUY;

            MyLogger.Info($"... CC-3 m_cc_retailer = {card.m_retailer}");

            int items_type = 0;

            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.ToUpper() == "QTY" && lines[i + 1].Trim().ToUpper() == "PRODUCT DESCRIPTION")
                {
                    items_type = 1;
                    break;
                }
            }

            lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("ORDER #", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "ORDER #")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("ORDER #".Length).Trim();
                    card.set_order_id(temp);
                    MyLogger.Info($"... CC-3 order id = {temp}");
                    continue;
                }
                if (line.StartsWith("Order number:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "ORDER NUMBER:")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Order number:".Length).Trim();
                    card.set_order_id(temp);
                    MyLogger.Info($"... CC-3 order id = {temp}");
                    continue;
                }
                if (items_type == 0)
                {
                    if (line.IndexOf("SKU:", StringComparison.CurrentCultureIgnoreCase) != -1)
                    {
                        string title = "";
                        string sku = "";
                        int qty = 1;
                        float price = 0;
                        int k = i - 1;

                        string temp;

                        if (line.ToUpper() == "SKU:")
                        {
                            temp = lines[++i].Trim();
                        }
                        else if (line.StartsWith("SKU:", StringComparison.CurrentCultureIgnoreCase))
                        {
                            temp = line.Substring("SKU:".Length).Trim();
                        }
                        else
                        {
                            title = line.Substring(0, line.IndexOf("SKU:", StringComparison.CurrentCultureIgnoreCase));
                            temp = line.Substring(line.IndexOf("SKU:", StringComparison.CurrentCultureIgnoreCase) + "SKU:".Length);
                        }
                        sku = temp;

                        while (k >= 0)
                        {
                            if (lines[k].Trim().ToUpper() == "SHIPPED ITEMS")
                                break;
                            if (lines[k].Trim().ToUpper() == "SHIP TO HOME ITEMS")
                                break;
                            if (lines[k].Trim().ToUpper() == "STORE PICKUP ITEMS")
                                break;
                            if (lines[k].Trim().ToUpper() == "SERVICES & DIGITAL DOWNLOADS")
                                break;
                            if (lines[k].Trim().ToUpper() == "UNVERIFIABLE INFORMATION")
                                break;
                            if (lines[k].Trim().IndexOf("Item Canceled", StringComparison.CurrentCultureIgnoreCase) != -1)
                                break;
                            if (lines[k].Trim().IndexOf("Exceeded Quantity Limits", StringComparison.CurrentCultureIgnoreCase) != -1)
                                break;

                            title = lines[k].Trim() + " " + title;
                            k--;
                        }
                        if (title.IndexOf("Model:", StringComparison.CurrentCultureIgnoreCase) != -1)
                            title = title.Substring(0, title.IndexOf("Model:", StringComparison.CurrentCultureIgnoreCase));

                        if (lines[i + 1].Trim().IndexOf("Qty", StringComparison.CurrentCultureIgnoreCase) != -1)
                        {
                            temp = lines[++i].Trim();
                            if (temp.ToUpper() == "QTY" || temp.ToUpper() == "QTY:" || temp.ToUpper() == "QTY :")
                                temp = lines[++i].Trim();
                            else if (temp.StartsWith("qty", StringComparison.CurrentCultureIgnoreCase))
                                temp = temp.Substring("qty".Length).Trim();
                            else
                                temp = temp.Substring(temp.IndexOf("Qty", StringComparison.CurrentCultureIgnoreCase) + "qty".Length).Trim();
                            qty = Str_Utils.string_to_int(temp);
                        }

                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        card.m_product_items.Add(product);

                        MyLogger.Info($"... CC-3 qty = {qty}, price = {price}, sku = {sku}, title = {title}");
                        continue;
                    }
                }
                else if (items_type == 1)
                {
                    if (line.StartsWith("SKU:", StringComparison.CurrentCultureIgnoreCase))
                    {
                        /**
                         *      QTY
                         *      PRODUCT DESCRIPTION
                         *   +  1
                         *      HP Pavilion x360 2in1 11634 TouchScreen Laptop Intel Pentium 4GB Memory 500GB Hard Drive Smoke Silver
                         *      MODEL: F9J18UA#AB
                         *  ->  SKU:
                         *   -  5554003
                         *   +  1
                         *      Microsoft Office 365 Personal 1 Mac or PC 1 iPad or Select Windows Tablet 1 Year Subscription MacWindows
                         *      MODEL: QQ2-00309
                         *  ->  SKU:
                         *   -  5058004 
                         **/

                        string title = "";
                        string sku = "";
                        int qty = 1;
                        float price = 0;
                        int k = i - 1;

                        string temp;

                        if (line.ToUpper() == "SKU:")
                            temp = lines[++i].Trim();
                        else
                            temp = line.Substring("SKU:".Length).Trim();
                        sku = temp;

                        int qty_tmp = 0;
                        temp = "";
                        while (k >= 0 && !int.TryParse(lines[k].Trim(), out qty_tmp))
                        {
                            temp = lines[k].Trim() + " " + temp;
                            k--;
                        }
                        if (k < 0)
                            continue;

                        if (temp.IndexOf("MODEL:", StringComparison.CurrentCultureIgnoreCase) != -1)
                            temp = temp.Substring(0, temp.IndexOf("MODEL:", StringComparison.CurrentCultureIgnoreCase)).Trim();
                        title = temp;

                        qty = Str_Utils.string_to_int(lines[k].Trim());

                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        card.m_product_items.Add(product);

                        MyLogger.Info($"... CC-3 qty = {qty}, price = {price}, sku = {sku}, title = {title}");
                        continue;
                    }
                }
            }
        }
    }
}
